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
                            PETTY_CASH_DEFAULT,
                            LedgerDashboardReader, WishlistManager,
                            PAYING_APTS, MONTHLY_FEE, TOTAL_APTS,
                            PETTY_AMOUNT, WISHLIST_STATUS_OPT)

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
            "expense_skipped","raw_tab_count","raw_tab_name","dashboard","wishlist_mgr"):
    if key not in st.session_state:
        st.session_state[key] = None

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "⚙️ הרצה", "✅ התאמות אוטומטיות", "🔍 בדיקה ידנית", "💸 הוצאות", "📋 יומן כתיבה", "📊 תכנון תקציבי"
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

            st.info("⬅️ עבור לטאב **יומן כתיבה** להורדת הגיליון המעודכן")


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
    else:
        # FIX: compute needs_review BEFORE the empty check.
        # needs_manual items live in `matched`, not `unmatched`.
        # The old code checked `not st.session_state.unmatched` first, which
        # fired the green "all matched!" banner and skipped the entire body
        # whenever unmatched=[] — even when matched contained needs_manual items.
        needs_review = [r for r in (st.session_state.matched or []) if r.get("needs_manual")]
        unmatched    = needs_review + (st.session_state.unmatched or [])
        if not unmatched:
            st.success("🎉 כל התנועות הותאמו אוטומטית!")
        else:
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
                    if tx.get("petty_cash"):
                        st.session_state[f"target_{idx}"] = "petty"   # pre-select קופה קטנה
                    else:
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
                        st.session_state[f"target_{idx}"] = "petty"
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
                st.divider()
                st.markdown("### 💾 שמירה והורדה")

                col_onedrive, col_local = st.columns(2)

                # ── Left: Save to OneDrive (atomic replace to beat sync lock) ──
                with col_onedrive:
                    st.markdown("**שמירה ל-OneDrive**")
                    st.caption(f"📁 `{onedrive_path}`")
                    if st.button(
                        "☁️ שמור ל-OneDrive",
                        type="primary",
                        use_container_width=True,
                        key="final_save",
                    ):
                        try:
                            import os, tempfile
                            # Write to a temp file in the SAME folder, then
                            # atomically replace the target. This avoids the
                            # OneDrive sync-lock race that rejects a direct write.
                            tmp_fd, tmp_path = tempfile.mkstemp(
                                dir=onedrive_path.parent,
                                suffix=".tmp",
                            )
                            try:
                                with os.fdopen(tmp_fd, "wb") as tmp_f:
                                    tmp_f.write(file_bytes)
                                os.replace(tmp_path, onedrive_path)
                                st.success(f"✓ נשמר בהצלחה:\n{onedrive_path}")
                            except Exception:
                                try:
                                    os.unlink(tmp_path)
                                except OSError:
                                    pass
                                raise
                        except Exception as e:
                            st.error(f"שגיאה בשמירה ל-OneDrive:\n{e}")

                # ── Right: Download to browser ─────────────────────────────────
                with col_local:
                    st.markdown("**הורדה מקומית**")
                    st.caption("📥 שמירה ישירה דרך הדפדפן")
                    st.download_button(
                        "⬇️ הורד לדפדפן",
                        file_bytes,
                        file_name=original_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="final_dl",
                    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 – Financial Planning Dashboard
# ══════════════════════════════════════════════════════════════════════════════
with tab6:
    st.subheader("📊 תכנון תקציבי 2026")

    # ── Ledger source: Tab 1 session OR direct upload here ───────────────────
    _tab6_upload = st.file_uploader(
        "העלה גיליון ניהול ישירות (אופציונלי — אם לא הרצת טאב 1)",
        type=["xlsx"],
        key="tab6_ledger_upload",
        help="אם כבר הרצת התאמה בטאב 1 הגיליון טעון אוטומטית — לא נדרשת העלאה חוזרת",
    )
    if _tab6_upload is not None:
        _tab6_tmp = Path("tmp/vaad")
        _tab6_tmp.mkdir(parents=True, exist_ok=True)
        _tab6_ledger_path_direct = str(_tab6_tmp.resolve() / "master_ledger_tab6.xlsx")
        Path(_tab6_ledger_path_direct).write_bytes(_tab6_upload.read())
        st.session_state["tab6_ledger_path"]     = _tab6_ledger_path_direct
        st.session_state["tab6_ledger_filename"] = _tab6_upload.name

    _resolved_ledger_path = (
        st.session_state.get("tab6_ledger_path") or st.session_state.ledger_path
    )
    _resolved_ledger_filename = (
        st.session_state.get("tab6_ledger_filename")
        or st.session_state.ledger_filename
        or "master_ledger.xlsx"
    )

    col_load, col_info = st.columns([1, 3])
    with col_load:
        load_dashboard = st.button(
            "🔄 טען נתוני דשבורד",
            type="primary",
            use_container_width=True,
            key="load_dashboard",
        )
    with col_info:
        if _resolved_ledger_path:
            _src_lbl = "טאב 1" if not st.session_state.get("tab6_ledger_path") else "העלאה ישירה"
            st.caption(
                f"✓ גיליון: `{_resolved_ledger_filename}` ({_src_lbl})  "
                f"— לחץ 'טען' לאחר כל שינוי בגיליון"
            )
        else:
            st.caption("⚠️ אין גיליון — הרץ התאמה בטאב 1 או העלה גיליון ישירות למעלה")

    if load_dashboard:
        if not _resolved_ledger_path:
            st.error("נא להעלות גיליון ניהול — ישירות כאן למעלה, או דרך טאב 1")
        else:
            with st.spinner("קורא נתונים פיננסיים מהגיליון…"):
                try:
                    st.session_state["dashboard"]    = LedgerDashboardReader(_resolved_ledger_path)
                    st.session_state["wishlist_mgr"] = WishlistManager(_resolved_ledger_path)
                    st.success("✓ נתונים נטענו בהצלחה")
                    st.rerun()
                except Exception as _e:
                    st.error(f"שגיאה בטעינת נתוני דשבורד: {_e}")

    _d = st.session_state.get("dashboard")

    if _d is None:
        st.info("לחץ על **'טען נתוני דשבורד'** להצגת הניתוח הפיננסי")
    else:
        # ═══════════════════════════════════════════════════════════════════
        # SECTION A – Operating Budget Health
        # ═══════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("## 📈 סעיף א׳ — מצב תפעולי")

        # ── A1: Top-line health metrics ────────────────────────────────────
        _ma1, _ma2, _ma3, _ma4 = st.columns(4)

        _cr = _d.coverage_ratio
        _cr_icon = "🟢" if _cr >= 1.1 else ("🟡" if _cr >= 0.9 else "🔴")
        _ma1.metric(
            f"{_cr_icon} יחס כיסוי חודשי",
            f"{_cr:.1%}",
            help=(
                "ממוצע הכנסה חודשית ÷ ממוצע הוצאות קבועות חודשיות\n\n"
                "מקור הכנסה: גביה 2026 · כל עמודות החודשים · שורות 2–61\n"
                "מקור הוצאה: הוצאות 2026 · עמודה N · שורות 2–12"
            ),
        )

        _coll_pct = (_d.ytd_collected / _d.annual_income_potential * 100) if _d.annual_income_potential else 0
        _ma2.metric(
            "📥 גביה YTD",
            f"₪{_d.ytd_collected:,.0f}",
            delta=f"{_coll_pct:.0f}% מהפוטנציאל השנתי",
            help=(
                f"סה״כ שנגבה מכל הדירות עד כה\n\n"
                f"מקור: גביה 2026 · כל עמודות החודשים · שורות 2–61\n"
                f"פוטנציאל: {PAYING_APTS} דירות × ₪{MONTHLY_FEE} × 12 = ₪{_d.annual_income_potential:,.0f}"
            ),
        )

        _ma3.metric(
            "📅 חודשים עם נתונים",
            f"{_d.months_with_data} / 12",
            help=(
                "מספר החודשים שנמצאו ערכים בגיליון ההוצאות\n"
                "ממוצעים וטווחים (min/max) מבוססים על חודשים אלה בלבד\n\n"
                "מקור: הוצאות 2026 · שורות 2–12 · ספירת תאים מלאים"
            ),
        )

        _ma4.metric(
            "🏗️ הוצאות קבועות YTD",
            f"₪{_d.fixed_ytd:,.0f}",
            help="מקור: הוצאות 2026 · שורות 2–12 · עמודה O (סה״כ שנתי)",
        )

        # ── A2: Fixed costs breakdown table ───────────────────────────────
        st.markdown("### 🏗️ הוצאות קבועות — Bucket 1")
        if _d.months_with_data <= 3:
            st.warning(
                f"⚠️ ממוצע מבוסס על **{_d.months_with_data} חודשים** בלבד — "
                f"תחזית שנתית עדיין לא בשלה. min/max ייצוב לאחר ≥6 חודשים."
            )

        _fixed_rows_display = []
        for _fr in _d.fixed_rows:
            _conf = f"⚠️ {_fr['months_count']}מ׳" if _fr["months_count"] < 6 else f"✓ {_fr['months_count']}מ׳"
            _fixed_rows_display.append({
                "קטגוריה":                _fr["category"],
                "סה״כ YTD ₪":            f"{_fr['total']:,.0f}" if _fr["total"] else "—",
                "ממוצע/חודש ₪":          f"{_fr['avg']:,.0f}"   if _fr["avg"]   else "—",
                "מינ׳ ₪":                f"{_fr['min']:,.0f}"   if _fr["min"]   else "—",
                "מקס׳ ₪":               f"{_fr['max']:,.0f}"   if _fr["max"]   else "—",
                "תחזית שנתית ₪":         f"{_fr['avg']*12:,.0f}" if _fr["avg"] else "—",
                "בסיס":                  _conf,
                "מקור":                  f"שורה {_fr['row']} · עמ׳ O, N",
            })

        # Totals row
        _fixed_rows_display.append({
            "קטגוריה":        "סה״כ",
            "סה״כ YTD ₪":     f"{_d.fixed_ytd:,.0f}",
            "ממוצע/חודש ₪":   f"{_d.fixed_avg_monthly:,.0f}",
            "מינ׳ ₪":         f"{_d.fixed_min_monthly:,.0f}",
            "מקס׳ ₪":        f"{_d.fixed_max_monthly:,.0f}",
            "תחזית שנתית ₪":  f"{_d.fixed_proj_annual:,.0f}",
            "בסיס":           f"{_d.months_with_data} חודשים",
            "מקור":           "שורות 2–12 · עמ׳ O, N",
        })

        _df_fixed = pd.DataFrame(_fixed_rows_display)

        def _style_fixed_tbl(row):
            if row["קטגוריה"] == "סה״כ":
                return [
                    "font-weight:bold;background-color:#2E4057;color:white"
                ] * len(row)
            return [""] * len(row)

        st.dataframe(
            _df_fixed.style.apply(_style_fixed_tbl, axis=1),
            use_container_width=True,
            hide_index=True,
            height=min(40 * (len(_fixed_rows_display) + 1) + 40, 520),
        )

        # Projection range note
        _pb1, _pb2, _pb3 = st.columns(3)
        _pb1.metric(
            "תחזית שנתית (ממוצע)",
            f"₪{_d.fixed_proj_annual:,.0f}",
            help="הוצאות 2026 · עמודה N (ממוצע חודשי) · שורות 2–12 · × 12",
        )
        _pb2.metric(
            "תחזית שנתית (מינימום)",
            f"₪{_d.fixed_proj_min:,.0f}",
            help="min לכל חודש שנרשם · × 12",
        )
        _pb3.metric(
            "תחזית שנתית (מקסימום)",
            f"₪{_d.fixed_proj_max:,.0f}",
            help="max לכל חודש שנרשם · × 12",
        )

        # ── A3: Payment Heatmap ────────────────────────────────────────────
        st.markdown("### 🗓️ מפת גביה — דיירים × חודשים")
        st.caption(
            "מקור: גביה 2026 · עמודות חודשים · שורות 2–61  |  "
            "🟢 שולם מלא  🟡 חלקי  ⬜ לא שולם"
        )

        _tenants_list = st.session_state.tenants or []
        _fee_map  = {t["apartment"]: t.get("monthly_fee", MONTHLY_FEE) for t in _tenants_list}
        _name_map = {t["apartment"]: t.get("tenant_name", "")          for t in _tenants_list}

        _HEB_SHORT = {
            "01":"ינו׳","02":"פבר׳","03":"מרץ","04":"אפר׳",
            "05":"מאי","06":"יוני","07":"יולי","08":"אוג׳",
            "09":"ספט׳","10":"אוק׳","11":"נוב׳","12":"דצמ׳",
        }

        _all_mm = sorted(
            {mm for months in _d.payment_grid.values() for mm in months}
        ) if _d.payment_grid else []

        if _all_mm and _d.payment_grid:
            _heat_data = {}
            for _apt in sorted(_d.payment_grid.keys()):
                _row_lbl = f"ד׳{_apt:02d} {_name_map.get(_apt,'')[:12]}"
                _fee     = _fee_map.get(_apt, MONTHLY_FEE)
                _heat_data[_row_lbl] = {
                    _HEB_SHORT.get(_mm, _mm): _d.payment_grid[_apt].get(_mm, 0.0)
                    for _mm in _all_mm
                }

            _df_heat = pd.DataFrame(_heat_data).T

            def _color_heat(val):
                try:
                    v = float(val)
                except Exception:
                    v = 0.0
                if v <= 0:
                    return "background-color:#2a2a2a;color:#666"
                elif v >= MONTHLY_FEE:
                    return "background-color:#1a5c2a;color:#cff"
                else:
                    return "background-color:#7a5800;color:#ffe"

            st.dataframe(
                _df_heat.style.map(_color_heat).format("{:.0f}"),
                use_container_width=True,
                height=min(30 * len(_d.payment_grid) + 50, 700),
            )

            # Heatmap summary
            _total_cells = sum(len(v) for v in _d.payment_grid.values())
            _paid_cells  = sum(
                1 for apt_months in _d.payment_grid.values()
                for v in apt_months.values() if v > 0
            )
            _outstanding = sum(
                max(0.0, _fee_map.get(_apt, MONTHLY_FEE) - _d.payment_grid[_apt].get(_mm, 0.0))
                for _apt in _d.payment_grid
                for _mm in _all_mm
            )
            _hc1, _hc2, _hc3 = st.columns(3)
            _hc1.metric("תאים ששולמו", f"{_paid_cells} / {_total_cells}",
                        help="גביה 2026 · ספירת תאים עם ערך > 0")
            _hc2.metric("% גביה", f"{_paid_cells / _total_cells * 100:.0f}%" if _total_cells else "—")
            _hc3.metric("חוב פוטנציאלי",  f"₪{_outstanding:,.0f}",
                        help="הפרש בין תשלום צפוי לבין שולם לכל התאים הריקים")

        # ═══════════════════════════════════════════════════════════════════
        # SECTION B – Reserve & Capital Planning
        # ═══════════════════════════════════════════════════════════════════
        st.divider()
        st.markdown("## 🏦 סעיף ב׳ — תכנון עתודות")

        _bcol1, _bcol2 = st.columns(2)

        # ── B1: Bucket 2 Waterfall ─────────────────────────────────────────
        with _bcol1:
            st.markdown("### 💧 Bucket 2 — תחזוקה שוטפת (כללי)")

            def _wf_row(label, amount, is_plus, source, width_pct=None):
                _color = "#1a7a34" if is_plus else "#b03030"
                _sign  = "+" if is_plus else "−"
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;"
                    f"align-items:center;padding:7px 12px;"
                    f"border-left:4px solid {_color};margin-bottom:3px;"
                    f"background:#1a1a2e;border-radius:2px'>"
                    f"<span style='color:#ddd'>{label}</span>"
                    f"<span style='color:{_color};font-weight:bold'>"
                    f"{_sign}₪{abs(amount):,.0f}</span></div>",
                    unsafe_allow_html=True,
                )
                st.caption(f"  ↳ {source}")

            def _wf_result(label, amount):
                _color = "#1a7a34" if amount >= 0 else "#b03030"
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;"
                    f"align-items:center;padding:10px 14px;"
                    f"border:2px solid {_color};margin-top:10px;"
                    f"background:#1a1a2e;border-radius:4px'>"
                    f"<span style='font-weight:bold;color:#eee;font-size:1.05em'>{label}</span>"
                    f"<span style='color:{_color};font-weight:bold;font-size:1.3em'>"
                    f"₪{amount:,.0f}</span></div>",
                    unsafe_allow_html=True,
                )

            _wf_row(
                "הכנסה שנתית פוטנציאלית",
                _d.annual_income_potential,
                is_plus=True,
                source=f"{PAYING_APTS} דירות × ₪{MONTHLY_FEE} × 12 | tenants.csv",
            )
            _wf_row(
                "הוצאות קבועות (תחזית ממוצע)",
                _d.fixed_proj_annual,
                is_plus=False,
                source=f"הוצאות 2026 · עמ׳ N (ממוצע) · שורות 2–12 · ×12 | טווח: ₪{_d.fixed_proj_min:,.0f}–₪{_d.fixed_proj_max:,.0f}",
            )

            _after_fixed = _d.annual_income_potential - _d.fixed_proj_annual
            st.caption(
                f"  ↔ לאחר הוצאות קבועות: **₪{_after_fixed:,.0f}** "
                f"(טווח ₪{_d.annual_income_potential - _d.fixed_proj_max:,.0f}–"
                f"₪{_d.annual_income_potential - _d.fixed_proj_min:,.0f})"
            )

            _wf_row(
                "כללי שהוצא YTD",
                _d.kelali_ytd,
                is_plus=False,
                source=f"הוצאות 2026 · שורה 13 · עמ׳ O (סה״כ שנתי) | ממוצע ₪{_d.kelali_avg:,.0f}/חודש",
            )

            _wf_result("תקציב תחזוקה פנוי 2026", _d.bucket2_remaining)

            st.markdown("")
            if _d.bucket2_monthly_rate > 0:
                _runway = _d.bucket2_months_runway
                _runway_str = f"~{_runway:.1f} חודשים" if _runway < 24 else "✓ מכוסה לשנה"
                _rate_color = "#1a7a34" if _runway > 6 else "#b03030"
                st.markdown(
                    f"<div style='padding:8px 12px;background:#1a1a2e;border-radius:4px;"
                    f"border-left:3px solid {_rate_color}'>"
                    f"<span style='color:#aaa'>קצב הוצאת כללי YTD:</span> "
                    f"<b style='color:#eee'>₪{_d.bucket2_monthly_rate:,.0f}/חודש</b><br>"
                    f"<span style='color:#aaa'>מנוחה משוערת:</span> "
                    f"<b style='color:{_rate_color}'>{_runway_str}</b></div>",
                    unsafe_allow_html=True,
                )
            else:
                st.info("אין הוצאות כללי YTD — קצב טרם ידוע")

            # כללי comments
            if _d.kelali_comments:
                with st.expander("📝 הערות כללי לפי חודש (שורה 23)", expanded=False):
                    for _mm_c, _txt in sorted(_d.kelali_comments.items()):
                        st.write(f"**{_HEB_SHORT.get(_mm_c, _mm_c)}:** {_txt}")

        # ── B2: Bucket 3 Capital Fund ──────────────────────────────────────
        with _bcol2:
            st.markdown("### 🏗️ Bucket 3 — קופה קטנה (פיתוח הון)")

            # Collection progress bar
            st.caption(
                f"מקור גביה: גביה 2026 · עמודה Q · שורות 2–61  |  "
                f"פוטנציאל: {TOTAL_APTS} × ₪{PETTY_AMOUNT:,} = ₪{_d.petty_potential:,}"
            )
            st.progress(
                min(_d.petty_collection_pct / 100, 1.0),
                text=(
                    f"נגבה: ₪{_d.petty_total_collected:,.0f} מתוך ₪{_d.petty_potential:,.0f} "
                    f"({_d.petty_collection_pct:.0f}%)  —  "
                    f"{_d.petty_apts_paid} / {TOTAL_APTS} דירות שילמו"
                ),
            )

            # Per-apt petty cash status (collapsed)
            with st.expander("🔍 גביית קופה קטנה לפי דירה", expanded=False):
                _petty_rows = []
                for _apt in sorted(_d.petty_collected.keys()):
                    _collected = _d.petty_collected[_apt]
                    _status    = "✅ שולם" if _collected >= PETTY_AMOUNT else ("🟡 חלקי" if _collected > 0 else "⬜ לא שולם")
                    _petty_rows.append({
                        "דירה":    _apt,
                        "דייר":    _name_map.get(_apt, ""),
                        "גבוי ₪":  f"{_collected:,.0f}",
                        "סטטוס":   _status,
                    })
                st.dataframe(
                    pd.DataFrame(_petty_rows),
                    use_container_width=True,
                    hide_index=True,
                    height=280,
                )

            st.markdown("")
            # Spent projects
            st.markdown("**הוצאות שבוצעו (קופה קטנה · שורות 3–5):**")
            if _d.petty_spent_projects:
                for _proj in _d.petty_spent_projects:
                    st.markdown(
                        f"<div style='display:flex;justify-content:space-between;"
                        f"align-items:center;padding:6px 12px;"
                        f"border-left:4px solid #b03030;margin-bottom:3px;"
                        f"background:#1a1a2e;border-radius:2px'>"
                        f"<span style='color:#ddd'>{_proj['name']}</span>"
                        f"<span style='color:#e06060;font-weight:bold'>"
                        f"−₪{_proj['total']:,.0f}</span></div>",
                        unsafe_allow_html=True,
                    )
                    st.caption(f"  ↳ {_proj['source']}")
            else:
                st.info("לא נמצאו הוצאות בשורות 3–5 של טאב קופה קטנה")

            # Available balance
            _bal_color = "#1a7a34" if _d.petty_available >= 0 else "#b03030"
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;"
                f"align-items:center;padding:10px 14px;"
                f"border:2px solid {_bal_color};margin:10px 0;"
                f"background:#1a1a2e;border-radius:4px'>"
                f"<span style='font-weight:bold;color:#eee;font-size:1.05em'>יתרה זמינה</span>"
                f"<span style='color:{_bal_color};font-weight:bold;font-size:1.3em'>"
                f"₪{_d.petty_available:,.0f}</span></div>",
                unsafe_allow_html=True,
            )
            st.caption(
                f"יתרה = נגבה ₪{_d.petty_total_collected:,.0f} "
                f"− הוצא ₪{_d.petty_total_spent:,.0f}"
            )

            # ── Wishlist ───────────────────────────────────────────────────
            st.markdown("---")
            st.markdown("**📋 רשימת פרויקטים מתוכננים:**")
            st.caption(f"מאוחסן בגיליון '{chr(0x05E8)}{chr(0x05E9)}{chr(0x05D9)}{chr(0x05DE)}{chr(0x05EA)} {chr(0x05E4)}{chr(0x05E8)}{chr(0x05D5)}{chr(0x05D9)}{chr(0x05E7)}{chr(0x05D8)}{chr(0x05D9)}{chr(0x05DD)}' בקובץ הלדג׳ר")

            _wm = st.session_state.get("wishlist_mgr")
            if _wm is None:
                st.info("טען נתוני דשבורד כדי לנהל את הרשימה")
            else:
                _wishlist = _wm.read()

                # Wishlist financial summary
                _planned_cost = sum(
                    i["cost"] for i in _wishlist
                    if i["status"] in ("מתוכנן", "בביצוע")
                )
                _headroom = _d.petty_available - _planned_cost
                _hc_color = "#1a7a34" if _headroom >= 0 else "#b03030"
                _hc_icon  = "✅" if _headroom >= 0 else "⚠️"

                _ws1, _ws2, _ws3 = st.columns(3)
                _ws1.metric(
                    "פרויקטים מתוכננים",
                    f"₪{_planned_cost:,.0f}",
                    help="סכום עלויות משוערות לפרויקטים במצב 'מתוכנן' ו'בביצוע' בלבד",
                )
                _ws2.metric(
                    f"{_hc_icon} מרחב פנוי",
                    f"₪{_headroom:,.0f}",
                    delta="מכוסה" if _headroom >= 0 else "חריגה",
                    delta_color="normal" if _headroom >= 0 else "inverse",
                    help="יתרה זמינה פחות סכום פרויקטים מתוכננים ובביצוע",
                )
                _ws3.metric(
                    "סה״כ פרויקטים",
                    f"{len(_wishlist)}",
                    delta=f"הושלמו: {sum(1 for i in _wishlist if i['status'] == 'הושלם')}",
                )

                # Wishlist table display
                if _wishlist:
                    _df_wish = pd.DataFrame([{
                        "שם פרויקט": i["name"],
                        "עלות (₪)":  f"{i['cost']:,.0f}",
                        "סטטוס":     i["status"],
                        "הערות":     i["notes"],
                    } for i in _wishlist])

                    def _style_wish_tbl(row):
                        s = row["סטטוס"]
                        if s == "הושלם":
                            return ["background-color:#383838;color:#888"] * len(row)
                        if s == "בביצוע":
                            return ["background-color:#1e4620;color:#cfc"] * len(row)
                        return ["background-color:#4a3e00;color:#ffe"] * len(row)

                    st.dataframe(
                        _df_wish.style.apply(_style_wish_tbl, axis=1),
                        use_container_width=True,
                        hide_index=True,
                        height=min(40 * len(_wishlist) + 50, 320),
                    )
                else:
                    st.info("הרשימה ריקה — הוסף פרויקטים בטופס למטה")

                # ── Edit / Add wishlist ────────────────────────────────────
                with st.expander(
                    "✏️ ערוך רשימת פרויקטים", expanded=(len(_wishlist) == 0)
                ):
                    _edit_items = list(_wishlist)

                    # Edit existing items
                    if _edit_items:
                        st.markdown("**עריכת פרויקטים קיימים:**")
                        _items_to_keep = []
                        for _wi, _item in enumerate(_edit_items):
                            _ec1, _ec2, _ec3, _ec4, _ec5 = st.columns([3, 2, 2, 3, 1])
                            _new_name   = _ec1.text_input("שם",   _item["name"],   key=f"wl_n_{_wi}", label_visibility="collapsed")
                            _new_cost   = _ec2.number_input("עלות", value=float(_item["cost"]), min_value=0.0, step=500.0, key=f"wl_c_{_wi}", label_visibility="collapsed")
                            _new_status = _ec3.selectbox("סטטוס", WISHLIST_STATUS_OPT,
                                                         index=WISHLIST_STATUS_OPT.index(_item["status"])
                                                         if _item["status"] in WISHLIST_STATUS_OPT else 0,
                                                         key=f"wl_s_{_wi}", label_visibility="collapsed")
                            _new_notes  = _ec4.text_input("הערות", _item["notes"], key=f"wl_no_{_wi}", label_visibility="collapsed")
                            _del_btn    = _ec5.button("🗑", key=f"wl_d_{_wi}", help="מחק שורה זו")

                            if not _del_btn:
                                _items_to_keep.append({
                                    "name":   _new_name,
                                    "cost":   _new_cost,
                                    "status": _new_status,
                                    "notes":  _new_notes,
                                })

                        if st.button("💾 שמור עריכות", use_container_width=True, key="wl_save_edits"):
                            _wm.save(_items_to_keep)
                            st.success(f"✓ נשמרו {len(_items_to_keep)} פרויקטים")
                            st.rerun()

                    # Add new project
                    st.markdown("**הוסף פרויקט חדש:**")
                    _na1, _na2, _na3, _na4 = st.columns([3, 2, 2, 3])
                    _new_proj_name   = _na1.text_input("שם פרויקט",       key="wl_new_name",   placeholder="שם הפרויקט")
                    _new_proj_cost   = _na2.number_input("עלות משוערת (₪)", min_value=0.0, step=500.0, key="wl_new_cost")
                    _new_proj_status = _na3.selectbox("סטטוס", WISHLIST_STATUS_OPT,            key="wl_new_status")
                    _new_proj_notes  = _na4.text_input("הערות",             key="wl_new_notes",  placeholder="הערות אופציונלי")

                    if st.button("➕ הוסף פרויקט לרשימה", use_container_width=True,
                                 type="primary", key="wl_add_btn"):
                        if _new_proj_name.strip():
                            _all_items = _wm.read()
                            _all_items.append({
                                "name":   _new_proj_name.strip(),
                                "cost":   _new_proj_cost,
                                "status": _new_proj_status,
                                "notes":  _new_proj_notes,
                            })
                            _wm.save(_all_items)
                            st.success(f"✓ פרויקט '{_new_proj_name}' נוסף ונשמר לגיליון")
                            st.rerun()
                        else:
                            st.warning("נא להזין שם פרויקט")

        # ── Section B bottom warning banner ───────────────────────────────
        _warnings = []
        if _d.coverage_ratio < 0.9:
            _warnings.append(
                f"⚠️ יחס כיסוי {_d.coverage_ratio:.0%} — הכנסה ממוצעת נמוכה מהוצאות קבועות"
            )
        if _d.bucket2_monthly_rate > 0 and _d.bucket2_months_runway < 3:
            _warnings.append(
                f"⚠️ תקציב תחזוקה שוטפת (כללי) צפוי להתאפס תוך {_d.bucket2_months_runway:.1f} חודשים"
            )
        if _d.bucket2_remaining < 0:
            _warnings.append(
                f"⚠️ תקציב תחזוקה שוטפת שלילי: ₪{_d.bucket2_remaining:,.0f} — הוצאות קבועות + כללי חורגות מהכנסה"
            )
        if _d.petty_available < 0:
            _warnings.append(
                f"⚠️ קופה קטנה: הוצאות (₪{_d.petty_total_spent:,.0f}) עולות על גביה (₪{_d.petty_total_collected:,.0f})"
            )
        _wm2 = st.session_state.get("wishlist_mgr")
        if _wm2:
            _wl2 = _wm2.read()
            _planned2 = sum(i["cost"] for i in _wl2 if i["status"] in ("מתוכנן", "בביצוע"))
            if _planned2 > _d.petty_available:
                _warnings.append(
                    f"⚠️ פרויקטים מתוכננים (₪{_planned2:,.0f}) עולים על יתרת קופה קטנה (₪{_d.petty_available:,.0f})"
                )

        if _warnings:
            st.divider()
            st.markdown("### 🚨 התראות")
            for _w in _warnings:
                st.error(_w)
