"""Sampling — SA 530. Uses risk level and performance materiality from the earlier pages."""
import io
import traceback

import pandas as pd
import streamlit as st

from core.sampling import (BASE_SIZES, DB_REQUIRED, FLOOR, TB_REQUIRED, TB_TOTAL, missing_columns,
                           sample_daybook, sample_trial_balance)
from core.state import FULL, LAKH, XLSX_MIME, daybook_source, file_prefix, materiality_ready, show_error, \
    to_excel_bytes
from core.ui import ICONS, page_header

page_header("Sampling", eyebrow="Step 3 · SA 530",
            intro="Upload a trial balance or a day book. Items at or above performance materiality are always "
                  "selected; the rest are sampled by value, with targeted picks for common red flags.")
if not materiality_ready():
    st.warning("Complete the risk assessment and materiality first. Sampling uses both.")
    st.page_link("views/risk.py", label="Go to risk assessment", icon=ICONS["back"])
    st.page_link("views/materiality.py", label="Go to materiality", icon=ICONS["back"])
    st.stop()

risk = st.session_state.risk_level
PM = st.session_state.performance_materiality * LAKH
OM = st.session_state.overall_materiality * LAKH
k1, k2, k3 = st.columns(3)
k1.metric("Risk level", risk, border=True)
k2.metric("Overall materiality", f"₹ {OM:,.0f}", border=True)
k3.metric("Performance materiality", f"₹ {PM:,.0f}", border=True)

with st.expander("How items are selected"):
    st.markdown(
        "- **At or above PM:** every item is checked (key items).\n"
        f"- **₹{FLOOR:,} up to PM:** sampled, weighted by value for sales, purchase and bank vouchers.\n"
        f"- **Below ₹{FLOOR:,}:** left to analytical procedures.\n"
        "- **Red flags picked on purpose:** weekend journals, round amounts and vague narrations.\n"
        f"- **Trial balance:** {TB_TOTAL[risk]} items at {risk} risk, spread across Schedule III main groups.")
    st.markdown("**Day book sample size per voucher type**")
    st.dataframe(pd.DataFrame(BASE_SIZES).T.rename_axis("Risk Level").reset_index(), hide_index=True, **FULL)

tab_tb, tab_db = st.tabs(["Trial balance", "Day book"])

# ------------------------------------------------ Trial Balance
with tab_tb:
    st.caption("Required columns: " + ", ".join(TB_REQUIRED))
    tb_file = st.file_uploader("Trial balance (.xlsx or .xls)", type=["xlsx", "xls"], key="tb_up")
    if tb_file:
        try:
            tb = pd.read_excel(tb_file)
            tb.columns = [str(c).strip() for c in tb.columns]
            missing = missing_columns(tb, TB_REQUIRED)
            if missing:
                st.error(f"Missing required columns: {', '.join(missing)}")
            else:
                sel, stats = sample_trial_balance(tb, risk, PM)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Ledgers", stats["ledgers"], border=True)
                m2.metric("Key items (≥ PM)", stats["key"], border=True)
                m3.metric("Sampled", stats["sampled"], border=True)
                m4.metric("Total selected", stats["selected"], border=True)
                fmt = {c: "{:,.2f}" for c in ["net_opening_balance", "Debit", "Credit", "net_closing_balance",
                                              "abs_closing", "abs_activity", "risk_indicator"]}
                st.dataframe(sel.style.format(fmt), **FULL)
                st.download_button("Download selected ledgers (.xlsx)", to_excel_bytes({"TB_Sample": sel}),
                                   f"{file_prefix()}TB_Sample_Selected.xlsx", XLSX_MIME, icon=ICONS["download"])
        except Exception:
            show_error("The trial balance could not be processed.", traceback.format_exc())

# ------------------------------------------------ Day Book
with tab_db:
    st.caption("Required columns: " + ", ".join(DB_REQUIRED) +
               " · optional: Day, Narrations, Vch No., Date (for red flags and de-duplication)")
    db_bytes, db_name = daybook_source("db_up", "Day book (.xlsx or .xls)")
    if db_bytes:
        try:
            db = pd.read_excel(io.BytesIO(db_bytes))
            db.columns = [str(c).strip() for c in db.columns]
            miss = missing_columns(db, DB_REQUIRED)
            if miss:
                st.error(f"Missing required columns: {', '.join(miss)}")
            else:
                sel, summ = sample_daybook(db, risk, PM)
                st.markdown("**Selection by voucher type**")
                st.dataframe(summ, **FULL, hide_index=True)
                st.markdown(f"**Selected vouchers ({len(sel):,})**")
                st.dataframe(sel.style.format({c: "{:,.2f}" for c in ["Debit Amount", "Credit Amount", "amount"]}),
                             **FULL)
                st.download_button("Download selected vouchers (.xlsx)",
                                   to_excel_bytes({"DayBook_Sample": sel, "Summary": summ}),
                                   f"{file_prefix()}DayBook_Sample_Selected.xlsx", XLSX_MIME,
                                   icon=ICONS["download"])
        except Exception:
            show_error("The day book could not be processed.", traceback.format_exc())
