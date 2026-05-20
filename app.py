"""
app.py  –  Streamlit review & reconciliation UI
Run with:  streamlit run app.py
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from bank_parser import parse_bank_file, parse_debit_file, parse_all_rows, extract_comment_text
from tenant_loader import load_tenants
from matcher import match_transactions, FUZZY_THRESHOLD
from excel_updater import (apply_matches, apply_expense_matches,
                            write_raw_transactions_tab,
                            LedgerUpdater, LedgerReader, ExpenseUpdater,
                            PETTY_CASH_DEFAULT)

st.set_page_config(
    page_title="האגמית 7 – גביה",
    page_icon="🏢",
    layout="wide",
)

st.markdown("""
<style>
  body, .stApp, .stDataFrame { direction: rtl; }
  thead tr th { text-align: right !important; }
  div[data-baseweb="slider"] { direction: ltr; }
</style>
""", unsafe_allow_html=True)

st.title("🏢 האגמית 7 – התאמת תשלומים")

for key in ("matched","unmatched","tenants","ledger_path","ledger_filename","payment_month",
            "written","skipped","debits_matched","debits_unmatched","expense_written",
            "expense_skipped","raw_tab_count","raw_tab_name"):
    if key not in st.session_state:
        st.session_state[key] = None

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "⚙️ הרצה", "✅ התאמות אוטומטיות", "🔍 בדיקה ידנית", "💸 הוצאות", "📋 יומן כתיבה"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Run
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("קבצי קלט")
    col1, col2 = st.columns(2)

    with col1:
        bank_file = st.file_uploader("קובץ בנק (.xls/.xlsx/.csv)", type=["xls","xlsx","csv"])
    with col2:
        ledger_file = st.file_uploader("גיליון ניהול (.xlsx)", type=["xlsx"])

    st.divider()

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        months_2026 = [f"{m:02d}/2026" for m in range(1, 13)]
        _today_mm   = f"{__import__('datetime').date.today().month:02d}/2026"
        _def_idx    = months_2026.index(_today_mm) if _today_mm in months_2026 else 0
        month_input = st.selectbox(
            "חודש ברירת מחדל (MM/YYYY)",
            options=months_2026,
            index=_def_idx,
            help="ישמש כברירת מחדל לתנועות שחסרת בהן חותמת חודש בבנק"
        )
    with col_b:
        threshold = st.slider(
            "סף התאמה פאזי (%)", 50, 100, FUZZY_THRESHOLD,
            help="התאמות מתחת לסף יועברו לבדיקה ידנית"
        )
    with col_c:
        overwrite = st.checkbox(
            "דרוס ערכים קיימים",
            value=False,
            help="החלף סכומים קיימים בתא. נוסחאות לעולם לא יידרסו."
        )

    st.caption("💡 הגדרות גיליון קבועות: שורת כותרת 1 · עמודת דירה B · גיליון 'גביה 2026'")

    st.divider()
    run_btn = st.button("▶ הרץ התאמה", type="primary", use_container_width=True)

    if run_btn:
        roster_path = Path("tenants.csv")

        errors = []
        if not bank_file:   errors.append("נא להעלות קובץ בנק")
        if not ledger_file: errors.append("נא להעלות גיליון ניהול")
        if not roster_path.exists(): errors.append("לא נמצא קובץ tenants.csv בתיקייה הראשית")

        if errors:
            for e in errors: st.error(e)
        else:
            tmp = Path("tmp/vaad")
            tmp.mkdir(parents=True, exist_ok=True)
            bank_path = tmp / bank_file.name
            bank_path.write_bytes(bank_file.read())

            roster_path = Path(__file__).parent / "tenants.csv"
            ledger_path = tmp.resolve() / "master_ledger.xlsx"
            ledger_path.write_bytes(ledger_file.read())

            cats_path = Path(__file__).parent / "categories.csv"

            # ── 1. Parse ALL raw rows (no filtering) ─────────────────────────
            with st.spinner("קורא את כל שורות הבנק…"):
                try:
                    if not cats_path.exists():
                        st.error("לא נמצא קובץ categories.csv"); st.stop()
                    all_bank_rows = parse_all_rows(str(bank_path), str(cats_path))
                except Exception as e:
                    st.error(f"שגיאה בקריאת שורות גולמיות: {e}"); st.stop()

            # ── 2. Parse income transactions ─────────────────────────────────
            with st.spinner("מנתח הכנסות מהבנק…"):
                try:
                    transactions = parse_bank_file(str(bank_path))
                except Exception as e:
                    st.error(f"שגיאה בניתוח קובץ בנק: {e}"); st.stop()

            # ── 3. Parse expense transactions ─────────────────────────────────
            with st.spinner("מנתח הוצאות מהבנק…"):
                try:
                    all_debits       = parse_debit_file(str(bank_path), str(cats_path))
                    debits_matched   = [d for d in all_debits if d["match_method"] == "keyword"]
                    debits_unmatched = [d for d in all_debits if d["match_method"] == "unmatched"]
                except Exception as e:
                    st.warning(f"אזהרה: לא ניתן לנתח הוצאות: {e}")
                    all_debits, debits_matched, debits_unmatched = [], [], []

            # ── 4. Load tenants ───────────────────────────────────────────────
            with st.spinner("טוען רשימת דיירים…"):
                try:
                    tenants = load_tenants(str(roster_path))
                except Exception as e:
                    st.error(f"שגיאה ברשימת דיירים: {e}"); st.stop()

            # ── 5. Match income to tenants ────────────────────────────────────
            with st.spinner("מתאים תנועות…"):
                matched, unmatched = match_transactions(
                    transactions, tenants, fuzzy_threshold=threshold
                )

            # ── 6. Write reconciliation data to גביה sheet ───────────────────
            with st.spinner("כותב לגיליון גביה…"):
                written, skipped = apply_matches(
                    matched, str(ledger_path), month_input, overwrite=overwrite
                )
                try:
                    expense_written, expense_skipped = apply_expense_matches(
                        debits_matched, str(ledger_path), month_input, overwrite=overwrite
                    )
                except Exception as e:
                    st.warning(f"אזהרה: שגיאה בכתיבת הוצאות: {e}")
                    expense_written, expense_skipped = [], []

            # ── 7. Write raw transactions tab ─────────────────────────────────
            # Combines ALL bank rows (income + expenses, matched + unmatched +
            # manual-review + double-payments) into a single audit sheet.
            with st.spinner("כותב טאב תנועות גולמיות…"):
                try:
                    all_match_results = matched + unmatched   # full picture
                    raw_count, raw_sheet = write_raw_transactions_tab(
                        all_rows          = all_bank_rows,
                        all_match_results = all_match_results,
                        ledger_path       = str(ledger_path),
                        payment_month     = month_input,
                    )
                except Exception as e:
                    st.warning(f"אזהרה: שגיאה בכתיבת טאב גולמי: {e}")
                    raw_count, raw_sheet = 0, "—"

            # ── Save state ────────────────────────────────────────────────────
            st.session_state.matched          = matched
            st.session_state.unmatched        = unmatched
            st.session_state.tenants          = tenants
            st.session_state.ledger_path      = str(ledger_path)
            st.session_state.ledger_filename  = ledger_file.name
            st.session_state.payment_month    = month_input
            st.session_state.written          = written
            st.session_state.skipped          = skipped
            st.session_state.debits_matched   = debits_matched
            st.session_state.debits_unmatched = debits_unmatched
            st.session_state.expense_written  = expense_written
            st.session_state.expense_skipped  = expense_skipped
            st.session_state.raw_tab_count    = raw_count
            st.session_state.raw_tab_name     = raw_sheet

            # ── Summary metrics ───────────────────────────────────────────────
            c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
            c1.metric("שורות גולמיות",       raw_count)
            c2.metric("תנועות הכנסה",         len(transactions))
            c3.metric("התאמות אוטומטיות",     len(matched))
            c4.metric("לבדיקה ידנית",         len(unmatched))
            c5.metric("נכתבו לגביה",          len(written))
            c6.metric("הוצאות שהותאמו",       len(debits_matched))
            c7.metric("הוצאות לבדיקה",        len(debits_unmatched))

            if raw_sheet and raw_sheet != "—":
                st.success(f"✅ טאב גולמי נוצר: **{raw_sheet}** ({raw_count} שורות)")

            with open(str(ledger_path), "rb") as f:
                st.download_button(
                    "⬇️ הורד גיליון מעודכן",
                    f.read(),
                    file_name=(st.session_state.ledger_filename or "גיליון.xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – Auto-matches
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.subheader("התאמות אוטומטיות")

    if st.session_state.matched is None:
        st.info("הרץ תהליך התאמה כדי לראות תוצאות")
    else:
        matched = st.session_state.matched
        rows = []
        for r in matched:
            tx = r["transaction"]
            t  = r.get("tenant") or {}
            rows.append({
                "תאריך":        tx["date"],
                "סכום (₪)":     tx["amount"],
                "צפוי (₪)":     t.get("monthly_fee") or "—",
                "⚠️":           "✓" if r.get("amount_mismatch") else "",
                "שם שולח":      tx.get("sender_name") or "—",
                "דירה (רמז)":   tx.get("apt_hint") or "—",
                "דייר שהותאם":  t.get("tenant_name","—"),
                "דירה":         t.get("apartment","—"),
                "שיטה":         r["match_method"],
                "ביטחון":       f"{r['confidence']:.0%}",
                "נכתב":         r.get("write_message",""),
            })
        df = pd.DataFrame(rows)

        def row_colour(row):
            if row["⚠️"] == "✓":
                return ["background-color:#7a2e00"]*len(row)
            if row["שיטה"] == "apt_hint":
                return ["background-color:#1e4620"]*len(row)
            elif row["שיטה"] == "fuzzy_name":
                return ["background-color:#5c4d04"]*len(row)
            return [""]*len(row)

        st.dataframe(
            df.style.apply(row_colour, axis=1),
            use_container_width=True, height=520
        )

        c1, c2, c3 = st.columns(3)
        c1.metric("התאמת דירה (ירוק)",  sum(1 for r in matched if r["match_method"]=="apt_hint"))
        c2.metric("התאמה פאזית (צהוב)", sum(1 for r in matched if r["match_method"]=="fuzzy_name"))
        c3.metric("סה״כ גביה (₪)",       f"{sum(r['transaction']['amount'] for r in matched):,.2f}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – Manual review
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.subheader("תנועות לבדיקה ידנית")

    if st.session_state.unmatched is None:
        st.info("הרץ תהליך התאמה כדי לראות תוצאות")
    elif not st.session_state.unmatched:
        st.success("🎉 כל התנועות הותאמו אוטומטית!")
    else:
        needs_review = [r for r in (st.session_state.matched or []) if r.get("needs_manual")]
        unmatched    = needs_review + (st.session_state.unmatched or [])
        tenants      = st.session_state.tenants or []
        tenant_options = {
            f"דירה {t['apartment']:>3} – {t['tenant_name']}": t for t in tenants
        }
        st.caption(f"{len(unmatched)} תנועות ממתינות לשיוך")

        _year = "2026"
        if st.session_state.payment_month:
            _parts = str(st.session_state.payment_month).split("/")
            if len(_parts) == 2 and _parts[1].isdigit():
                _year = _parts[1]

        HEB_SHORT = {
            "01":"ינו׳","02":"פבר׳","03":"מרץ","04":"אפר׳",
            "05":"מאי", "06":"יוני","07":"יולי","08":"אוג׳",
            "09":"ספט׳","10":"אוק׳","11":"נוב׳","12":"דצמ׳",
        }
        HEB_FULL = {
            "01":"ינואר","02":"פברואר","03":"מרץ","04":"אפריל",
            "05":"מאי",  "06":"יוני",  "07":"יולי","08":"אוגוסט",
            "09":"ספטמבר","10":"אוקטובר","11":"נובמבר","12":"דצמבר",
        }

        _ledger_reader = None
        if st.session_state.ledger_path:
            try:
                _ledger_reader = LedgerReader(st.session_state.ledger_path)
            except Exception:
                pass

        def _cell_status(val, full_amount):
            try:
                v = float(val) if val not in (None, "", 0) else 0.0
            except (TypeError, ValueError):
                v = 0.0
            if v == 0:
                return "⬜", v
            elif v < full_amount:
                return f"🟡", v
            else:
                return f"🟢", v

        for idx, r in enumerate(unmatched):
            tx    = r["transaction"]
            label = tx.get("sender_name") or tx["description"]

            if f"alloc_{idx}" not in st.session_state:
                st.session_state[f"alloc_{idx}"] = 0.0
            if f"target_{idx}" not in st.session_state:
                st.session_state[f"target_{idx}"] = \
                    tx.get("payment_month") or st.session_state.payment_month

            with st.expander(
                f"📌 {tx['date']}  |  ₪{tx['amount']:.2f}  |  {label}",
                expanded=True
            ):
                col_t, col_rem = st.columns([3, 1])

                _pre = 0
                if r.get("tenant"):
                    _key = f"דירה {r['tenant']['apartment']:>3} – {r['tenant']['tenant_name']}"
                    _opts = ["— לא לשייך —"] + list(tenant_options.keys())
                    if _key in _opts:
                        _pre = _opts.index(_key)

                chosen = col_t.selectbox(
                    "דייר",
                    ["— לא לשייך —"] + list(tenant_options.keys()),
                    index=_pre,
                    key=f"sel_{idx}",
                )
                _allocated = st.session_state[f"alloc_{idx}"]
                _remaining = tx["amount"] - _allocated
                col_rem.metric("נותר לחלוקה", f"₪{_remaining:.2f}",
                               delta=f"-₪{_allocated:.2f}" if _allocated else None,
                               delta_color="inverse")

                if r.get("amount_mismatch"):
                    _exp = (r.get("tenant") or {}).get("monthly_fee", "?")
                    st.warning(f"⚠️ שולם ₪{tx['amount']:.2f} | צפוי ₪{_exp}")

                with st.expander("פרטי בנק", expanded=False):
                    st.write("**תאור מורחב:**", tx.get("raw_detail") or "—")
                    st.write("**רמז דירה:**",    tx.get("apt_hint") or "לא זוהה")
                    if r.get("tenant"):
                        st.write("**הצעת מערכת:**", r.get("match_detail", ""))

                _tenant_obj  = tenant_options.get(chosen) if chosen != "— לא לשייך —" else None
                _apt         = _tenant_obj["apartment"] if _tenant_obj else None
                _monthly_fee = (_tenant_obj.get("monthly_fee") or 350) if _tenant_obj else 350

                st.markdown("**בחר חודש יעד:**")
                _sel = st.session_state[f"target_{idx}"]

                for _row_start in (0, 6):
                    _cols6 = st.columns(6)
                    for _i, _c in enumerate(_cols6):
                        _mi   = _row_start + _i
                        _mm   = f"{_mi + 1:02d}"
                        _mstr = f"{_mm}/{_year}"
                        _raw  = (_ledger_reader.read_payment(_apt, _mstr)
                                 if _ledger_reader and _apt else None)
                        _emoji, _cv = _cell_status(_raw, _monthly_fee)
                        _is_sel = _sel == _mstr
                        _cv_str = f" {_cv:.0f}" if _cv else ""
                        _lbl    = f"{'✓ ' if _is_sel else ''}{HEB_SHORT[_mm]} {_emoji}{_cv_str}"
                        if _c.button(_lbl, key=f"p_{idx}_{_mm}",
                                     type="primary" if _is_sel else "secondary",
                                     use_container_width=True):
                            st.session_state[f"target_{idx}"] = _mstr
                            st.rerun()

                _praw       = (_ledger_reader.read_petty(_apt)
                               if _ledger_reader and _apt else None)
                _pemoji, _pv = _cell_status(_praw, PETTY_CASH_DEFAULT)
                _is_psel    = _sel == "petty"
                _plbl = (
                    f"{'✓ ' if _is_psel else ''}קופה קטנה  "
                    f"{_pemoji}{f'  ₪{_pv:.0f}' if _pv else ''}  /  ₪{PETTY_CASH_DEFAULT}"
                )
                if st.button(_plbl, key=f"p_{idx}_petty",
                             type="primary" if _is_psel else "secondary",
                             use_container_width=True):
                    st.session_state[f"petty"] = "petty"
                    st.rerun()

                st.divider()
                if not _sel:
                    st.info("בחר חודש או קופה קטנה מהרשת למעלה")
                else:
                    if _sel == "petty":
                        _target_lbl  = f"קופה קטנה  (נוכחי ₪{_pv:.0f} / ₪{PETTY_CASH_DEFAULT})"
                        _default_amt = tx["amount"]
                    else:
                        _sel_mm  = _sel.split("/")[0]
                        _cur_raw = (_ledger_reader.read_payment(_apt, _sel)
                                    if _ledger_reader and _apt else None)
                        try:
                            _cur_v = float(_cur_raw) if _cur_raw not in (None, "", 0) else 0.0
                        except (TypeError, ValueError):
                            _cur_v = 0.0
                        _cur_str    = f"  —  נוכחי ₪{_cur_v:.0f}" if _cur_v else ""
                        _target_lbl = f"{HEB_FULL.get(_sel_mm, _sel_mm)} ({_sel}){_cur_str}"
                        _default_amt = tx["amount"]

                    st.markdown(f"**יעד:** {_target_lbl}")
                    _ca, _cb = st.columns([2, 1])
                    _amt_in  = _ca.number_input(
                        "סכום להוספה (₪)",
                        value=float(_default_amt),
                        min_value=0.0,
                        step=10.0,
                        key=f"amt_{idx}",
                    )

                    if _remaining <= 0:
                        _cb.success("✓ כל הסכום חולק")
                    elif _cb.button("➕ הוסף", key=f"add_{idx}",
                                    use_container_width=True, type="primary"):
                        if not _tenant_obj:
                            st.warning("נא לבחור דייר")
                        elif not st.session_state.ledger_path:
                            st.error("אין גיליון פתוח")
                        else:
                            _u = LedgerUpdater(st.session_state.ledger_path)
                            if _sel == "petty":
                                _ok, _msg = _u.write_petty_cash(_apt, _amt_in, mode="add")
                            else:
                                _ok, _msg = _u.write_payment(_apt, _sel, _amt_in, mode="add")
                            _u.save()
                            (st.success if _ok else st.error)(_msg)
                            if _ok:
                                st.session_state[f"alloc_{idx}"] = _allocated + _amt_in
                                with open(st.session_state.ledger_path, "rb") as _f:
                                    st.download_button(
                                        "⬇️ הורד",
                                        _f.read(),
                                        file_name=(st.session_state.ledger_filename or "גיליון.xlsx"),
                                        key=f"dl_{idx}",
                                    )
                                st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Expenses
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("הוצאות – ניתוח חיובים")

    if st.session_state.debits_matched is None:
        st.info("הרץ תהליך התאמה כדי לראות הוצאות")
    else:
        debits_matched   = st.session_state.debits_matched   or []
        debits_unmatched = st.session_state.debits_unmatched or []
        overwrite_exp    = st.checkbox("דרוס ערכים קיימים בהוצאות", value=False, key="overwrite_exp")

        st.markdown("### ✅ הוצאות שהותאמו אוטומטית")
        if not debits_matched:
            st.info("לא נמצאו הוצאות מוכרות בקובץ הבנק")
        else:
            from collections import defaultdict
            grouped: dict[str, float] = defaultdict(float)
            group_rows: dict[str, list] = defaultdict(list)
            for d in debits_matched:
                grouped[d["category"]] += d["amount"]
                group_rows[d["category"]].append(d)

            rows = []
            for cat, total in grouped.items():
                txs = group_rows[cat]
                rows.append({
                    "קטגוריה":       cat,
                    "סכום כולל (₪)": f"{total:,.2f}",
                    "מס׳ תנועות":    len(txs),
                    "תאריכים":       ", ".join(t["date"] for t in txs),
                    "גורם":          txs[0].get("entity_name","—"),
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True)

            exp_written = st.session_state.expense_written or []
            exp_skipped = st.session_state.expense_skipped or []
            ca, cb = st.columns(2)
            ca.metric("נכתבו", len(exp_written))
            cb.metric("דולגו (קיים)", len(exp_skipped))

            if exp_skipped:
                with st.expander("⚠️ הוצאות שדולגו (תא כבר מכיל ערך)"):
                    st.dataframe(pd.DataFrame([{
                        "קטגוריה": e["category"],
                        "סכום":    e["amount"],
                        "סיבה":    e["write_message"],
                    } for e in exp_skipped]), use_container_width=True)

            if st.button("🔄 כתוב הוצאות מותאמות לגיליון", use_container_width=True, key="write_exp_auto"):
                if not st.session_state.ledger_path:
                    st.error("אין גיליון פתוח")
                else:
                    try:
                        ew, es = apply_expense_matches(
                            debits_matched,
                            st.session_state.ledger_path,
                            st.session_state.payment_month,
                            overwrite=overwrite_exp,
                        )
                        st.session_state.expense_written = ew
                        st.session_state.expense_skipped = es
                        st.success(f"נכתבו {len(ew)} | דולגו {len(es)}")
                        with open(st.session_state.ledger_path, "rb") as f:
                            st.download_button(
                                "⬇️ הורד גיליון מעודכן",
                                f.read(),
                                file_name=(st.session_state.ledger_filename or "גיליון.xlsx"),
                                key="dl_exp_auto"
                            )
                    except Exception as e:
                        st.error(f"שגיאה בכתיבה: {e}")

        st.divider()

        st.markdown("### 🔍 הוצאות לשיוך ידני")
        if not debits_unmatched:
            st.success("🎉 כל ההוצאות שויכו אוטומטית!")
        else:
            if st.session_state.ledger_path:
                try:
                    eu = ExpenseUpdater(st.session_state.ledger_path)
                    cat_options = eu.available_categories
                except Exception:
                    cat_options = []
            else:
                cat_options = []

            st.caption(f"{len(debits_unmatched)} הוצאות ממתינות לשיוך")
            for idx, d in enumerate(debits_unmatched):
                label = d.get("description") or d["ref"]
                with st.expander(
                    f"📌 {d['date']}  |  ₪{d['amount']:.2f}  |  {label}", expanded=True
                ):
                    col_l, col_r = st.columns([2, 1])
                    with col_l:
                        st.write("**תיאור:**",       d.get("description") or "—")
                        st.write("**תאור מורחב:**",  d.get("raw_detail")  or "—")
                    with col_r:
                        chosen_cat = st.selectbox(
                            "שייך לקטגוריה",
                            ["— לא לשייך —"] + (cat_options or ["(טען גיליון)"]),
                            key=f"exp_sel_{idx}"
                        )
                        amt = st.number_input(
                            "סכום (₪)", value=float(d["amount"]), key=f"exp_amt_{idx}"
                        )
                        btn_col1, btn_col2 = st.columns(2)
                        _write_clicked = btn_col1.button("✏️ כתוב לגיליון", key=f"exp_write_{idx}", use_container_width=True)
                        _add_clicked   = btn_col2.button("➕ הוסף לסכום",   key=f"exp_add_{idx}",   use_container_width=True)

                        def _guard(chosen, ledger):
                            if chosen == "— לא לשייך —":
                                st.warning("נא לבחור קטגוריה"); return False
                            if not ledger:
                                st.error("אין גיליון פתוח"); return False
                            return True

                        if _write_clicked or _add_clicked:
                            if _guard(chosen_cat, st.session_state.ledger_path):
                                eu2 = ExpenseUpdater(st.session_state.ledger_path)
                                if _write_clicked:
                                    ok, msg = eu2.write_expense(
                                        chosen_cat,
                                        st.session_state.payment_month,
                                        amt,
                                        overwrite=overwrite_exp,
                                    )
                                else:
                                    ok, msg = eu2.add_to_expense(
                                        chosen_cat,
                                        st.session_state.payment_month,
                                        amt,
                                    )
                                if ok and chosen_cat == "כללי":
                                    comment_text = extract_comment_text(d.get("raw_detail", ""))
                                    if comment_text:
                                        c_ok, c_msg = eu2.write_comment(
                                            st.session_state.payment_month, comment_text
                                        )
                                        msg += f"\n{'✓' if c_ok else '⚠️'} הערה: {c_msg}"
                                eu2.save()
                                (st.success if ok else st.error)(msg)
                                if ok:
                                    with open(st.session_state.ledger_path, "rb") as f:
                                        st.download_button(
                                            "⬇️ הורד גיליון מעודכן",
                                            f.read(),
                                            file_name=(st.session_state.ledger_filename or "גיליון.xlsx"),
                                            key=f"dl_exp_{idx}"
                                        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – Written log
# ══════════════════════════════════════════════════════════════════════════════
with tab5:
    st.subheader("יומן כתיבה")

    if st.session_state.written is None:
        st.info("הרץ תהליך התאמה כדי לראות יומן")
    else:
        written = st.session_state.written
        skipped = st.session_state.skipped

        # Raw tab summary
        if st.session_state.raw_tab_name:
            st.info(
                f"📋 טאב גולמי: **{st.session_state.raw_tab_name}** — "
                f"{st.session_state.raw_tab_count} שורות (כולל ידני, כפול, הוצאות)"
            )

        st.markdown(f"**נכתבו:** {len(written)}  |  **דולגו:** {len(skipped)}")

        if written:
            st.markdown("#### ✅ נכתבו בהצלחה")
            st.dataframe(pd.DataFrame([{
                "תאריך": r["transaction"]["date"],
                "סכום":  r["transaction"]["amount"],
                "דייר":  (r.get("tenant") or {}).get("tenant_name","—"),
                "דירה":  (r.get("tenant") or {}).get("apartment","—"),
                "תא":    r.get("write_message",""),
            } for r in written]), use_container_width=True)

        if skipped:
            st.markdown("#### ⚠️ דולגו")
            st.dataframe(pd.DataFrame([{
                "תאריך": r["transaction"]["date"],
                "סכום":  r["transaction"]["amount"],
                "דייר":  (r.get("tenant") or {}).get("tenant_name","—"),
                "דירה":  (r.get("tenant") or {}).get("apartment","—"),
                "סיבה":  r.get("write_message",""),
            } for r in skipped]), use_container_width=True)

        if st.session_state.ledger_path:
            original_name = st.session_state.ledger_filename or "גיליון.xlsx"
            onedrive_path = Path(r"C:\Users\danie\OneDrive") / original_name

            try:
                with open(st.session_state.ledger_path, "rb") as _fh:
                    file_bytes = _fh.read()
            except Exception as e:
                st.error(f"לא ניתן לקרוא את הקובץ: {e}")
                file_bytes = None

            if file_bytes:
                col_save, col_dl = st.columns(2)

                if col_save.button("💾 שמור לOneDrive", type="primary",
                                   use_container_width=True, key="final_save"):
                    try:
                        onedrive_path.write_bytes(file_bytes)
                        st.success(f"✓ נשמר: {onedrive_path}")
                    except Exception as e:
                        st.error(f"שגיאה בשמירה לOneDrive: {e}")

                col_dl.download_button(
                    "⬇️ הורד לדפדפן",
                    file_bytes,
                    file_name=original_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="final_dl",
                )
