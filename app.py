"""
app.py  –  Streamlit review & reconciliation UI
Run with:  streamlit run app.py
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from bank_parser import parse_bank_file
from tenant_loader import load_tenants
from matcher import match_transactions, FUZZY_THRESHOLD
from excel_updater import apply_matches, LedgerUpdater

st.set_page_config(
    page_title="האגמית 7 – גביה",
    page_icon="🏢",
    layout="wide",
)

st.markdown("""
<style>
  body, .stApp, .stDataFrame { direction: rtl; }
  thead tr th { text-align: right !important; }
</style>
""", unsafe_allow_html=True)

st.title("🏢 האגמית 7 – התאמת תשלומים")

for key in ("matched","unmatched","tenants","ledger_path","payment_month","written","skipped"):
    if key not in st.session_state:
        st.session_state[key] = None

tab1, tab2, tab3, tab4 = st.tabs([
    "⚙️ הרצה", "✅ התאמות אוטומטיות", "🔍 בדיקה ידנית", "📋 יומן כתיבה"
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – Run
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("קבצי קלט")
    col1, col2, col3 = st.columns(3)

    with col1:
        bank_file = st.file_uploader("קובץ בנק (.xls/.xlsx/.csv)", type=["xls","xlsx","csv"])
    with col2:
        roster_file = st.file_uploader("רשימת דיירים (.csv)", type=["csv"])
    with col3:
        ledger_file = st.file_uploader("גיליון ניהול (.xlsx)", type=["xlsx"])

    st.divider()

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        month_input = st.text_input(
            "חודש לעדכון (MM/YYYY)",
            value=datetime.now().strftime("%m/%Y"),
            help="לדוגמה: 04/2026 לאפריל 2026"
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
        errors = []
        if not bank_file:   errors.append("נא להעלות קובץ בנק")
        if not roster_file: errors.append("נא להעלות רשימת דיירים")
        if not ledger_file: errors.append("נא להעלות גיליון ניהול")

        if errors:
            for e in errors: st.error(e)
        else:
            tmp = Path("/tmp/vaad"); tmp.mkdir(exist_ok=True)
            bank_path   = tmp / bank_file.name
            roster_path = tmp / "tenants.csv"
            ledger_path = tmp / "master_ledger.xlsx"

            bank_path.write_bytes(bank_file.read())
            roster_path.write_bytes(roster_file.read())
            ledger_path.write_bytes(ledger_file.read())

            with st.spinner("מנתח קובץ בנק…"):
                try:
                    transactions = parse_bank_file(str(bank_path))
                except Exception as e:
                    st.error(f"שגיאה בניתוח קובץ בנק: {e}"); st.stop()

            with st.spinner("טוען רשימת דיירים…"):
                try:
                    tenants = load_tenants(str(roster_path))
                except Exception as e:
                    st.error(f"שגיאה ברשימת דיירים: {e}"); st.stop()

            with st.spinner("מתאים תנועות…"):
                matched, unmatched = match_transactions(
                    transactions, tenants, fuzzy_threshold=threshold
                )

            with st.spinner("כותב לגיליון…"):
                written, skipped = apply_matches(
                    matched, str(ledger_path), month_input, overwrite=overwrite
                )

            st.session_state.matched       = matched
            st.session_state.unmatched     = unmatched
            st.session_state.tenants       = tenants
            st.session_state.ledger_path   = str(ledger_path)
            st.session_state.payment_month = month_input
            st.session_state.written       = written
            st.session_state.skipped       = skipped

            # Summary metrics
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("תנועות סה״כ",    len(transactions))
            c2.metric("התאמות אוטומטיות", len(matched))
            c3.metric("לבדיקה ידנית",   len(unmatched))
            c4.metric("נכתבו לגיליון",  len(written))

            with open(str(ledger_path), "rb") as f:
                st.download_button(
                    "⬇️ הורד גיליון מעודכן",
                    f.read(),
                    file_name=f"האגמית7_כספים_מעודכן_{month_input.replace('/','_')}.xlsx",
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
                "שם שולח":      tx.get("sender_name") or "—",
                "דירה (רמז)":   tx.get("apt_hint") or "—",
                "דייר שהותאם":  t.get("tenant_name","—"),
                "דירה":         t.get("apartment","—"),
                "שיטה":         r["match_method"],
                "ביטחון":       f"{r['confidence']:.0%}",
                "נכתב":         r.get("write_message",""),
            })
        df = pd.DataFrame(rows)

        # Colour by method
        def row_colour(row):
            if row["שיטה"] == "apt_hint":
                return ["background-color:#d4edda"]*len(row)
            elif row["שיטה"] == "fuzzy_name":
                return ["background-color:#fff3cd"]*len(row)
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
        unmatched = st.session_state.unmatched
        tenants   = st.session_state.tenants or []
        tenant_options = {
            f"דירה {t['apartment']:>3} – {t['tenant_name']}": t for t in tenants
        }
        st.caption(f"{len(unmatched)} תנועות ממתינות לשיוך")

        for idx, r in enumerate(unmatched):
            tx = r["transaction"]
            label = tx.get("sender_name") or tx["description"]
            with st.expander(f"📌 {tx['date']}  |  ₪{tx['amount']:.2f}  |  {label}", expanded=True):
                col_l, col_r = st.columns([2,1])
                with col_l:
                    st.write("**פרטי בנק:**", tx.get("raw_detail") or "—")
                    st.write("**רמז דירה:**", tx.get("apt_hint") or "לא זוהה")
                    if r.get("tenant"):
                        best = r["tenant"]
                        st.write(f"**הצעת מערכת:** דירה {best['apartment']} – {best['tenant_name']}  ({r['match_detail']})")
                with col_r:
                    chosen = st.selectbox(
                        "שייך לדייר",
                        ["— לא לשייך —"] + list(tenant_options.keys()),
                        key=f"sel_{idx}"
                    )
                    amt = st.number_input("סכום (₪)", value=float(tx["amount"]), key=f"amt_{idx}")
                    if st.button("✏️ כתוב לגיליון", key=f"write_{idx}"):
                        if chosen == "— לא לשייך —":
                            st.warning("נא לבחור דייר")
                        elif not st.session_state.ledger_path:
                            st.error("אין גיליון פתוח")
                        else:
                            tenant = tenant_options[chosen]
                            u = LedgerUpdater(st.session_state.ledger_path)
                            ok, msg = u.write_payment(
                                tenant["apartment"],
                                st.session_state.payment_month,
                                amt
                            )
                            u.save()
                            (st.success if ok else st.error)(msg)
                            if ok:
                                with open(st.session_state.ledger_path, "rb") as f:
                                    st.download_button(
                                        "⬇️ הורד גיליון מעודכן",
                                        f.read(),
                                        file_name="האגמית7_כספים_מעודכן.xlsx",
                                        key=f"dl_{idx}"
                                    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 – Written log
# ══════════════════════════════════════════════════════════════════════════════
with tab4:
    st.subheader("יומן כתיבה")

    if st.session_state.written is None:
        st.info("הרץ תהליך התאמה כדי לראות יומן")
    else:
        written = st.session_state.written
        skipped = st.session_state.skipped
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
            with open(st.session_state.ledger_path, "rb") as f:
                st.download_button(
                    "⬇️ הורד גיליון מעודכן",
                    f.read(),
                    file_name=f"האגמית7_כספים_מעודכן.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
