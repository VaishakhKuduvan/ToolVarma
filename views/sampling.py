"""Sampling — SA 530. Uses risk level and performance materiality from the earlier pages."""
import io
import traceback

import pandas as pd
import streamlit as st

from core.sampling import (BASE_SIZES, DB_REQUIRED, FLOOR, TB_REQUIRED, TB_TOTAL, missing_columns,
                           sample_daybook, sample_trial_balance)
from core.state import FULL, LAKH, XLSX_MIME, daybook_source, file_prefix, materiality_ready, show_error, \
    to_excel_bytes

st.title("📋 Sampling — SA 530")
if not materiality_ready():
    st.warning("⚠️ Complete **1. Risk Assessment** and **2. Materiality** before sampling.")
    st.page_link("views/risk.py", label="Go to Risk Assessment", icon="⬅️")
    st.page_link("views/materiality.py", label="Go to Materiality", icon="⬅️")
    st.stop()

risk = st.session_state.risk_level
PM = st.session_state.performance_materiality * LAKH
OM = st.session_state.overall_materiality * LAKH
st.info(f"Risk: **{risk}** | OM: ₹ {OM:,.2f} | PM: ₹ {PM:,.2f}")

with st.expander("📐 Sample Size Logic & Rules"):
    st.dataframe(pd.DataFrame(BASE_SIZES).T.rename_axis("Risk Level").reset_index(), hide_index=True, **FULL)
    st.markdown(
        "- **≥ PM** → 100% checked (key items)\n"
        f"- **₹{FLOOR:,} to PM** → sampled (value-weighted / random)\n"
        f"- **< ₹{FLOOR:,}** → analytical procedures only\n"
        "- **Targeted red flags:** weekend journals, round figures, vague narrations\n"
        f"- Trial Balance sample: {TB_TOTAL[risk]} items at {risk} risk, stratified by Schedule III main group")

tab_tb, tab_db = st.tabs(["📒 Trial Balance Sampling", "📖 Day Book Sampling"])

# ------------------------------------------------ Trial Balance
with tab_tb:
    st.caption("Required columns: " + ", ".join(TB_REQUIRED))
    tb_file = st.file_uploader("Upload Trial Balance (.xlsx/.xls)", type=["xlsx", "xls"], key="tb_up")
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
                m1.metric("Ledgers in TB", stats["ledgers"])
                m2.metric("Key Items (≥ PM)", stats["key"])
                m3.metric("Sampled", stats["sampled"])
                m4.metric("Total Selected", stats["selected"])
                fmt = {c: "{:,.2f}" for c in ["net_opening_balance", "Debit", "Credit", "net_closing_balance",
                                              "abs_closing", "abs_activity", "risk_indicator"]}
                st.dataframe(sel.style.format(fmt), **FULL)
                st.download_button("⬇️ Download TB_Sample_Selected.xlsx", to_excel_bytes({"TB_Sample": sel}),
                                   f"{file_prefix()}TB_Sample_Selected.xlsx", XLSX_MIME)
        except Exception:
            show_error("Trial Balance processing failed.", traceback.format_exc())

# ------------------------------------------------ Day Book
with tab_db:
    st.caption("Required columns: " + ", ".join(DB_REQUIRED) +
               " · optional: Day, Narrations, Vch No., Date (for red flags and de-duplication)")
    db_bytes, db_name = daybook_source("db_up")
    if db_bytes:
        try:
            db = pd.read_excel(io.BytesIO(db_bytes))
            db.columns = [str(c).strip() for c in db.columns]
            miss = missing_columns(db, DB_REQUIRED)
            if miss:
                st.error(f"Missing required columns: {', '.join(miss)}")
            else:
                sel, summ = sample_daybook(db, risk, PM)
                st.markdown("**Stratum Summary**")
                st.dataframe(summ, **FULL, hide_index=True)
                st.markdown("**Selected Vouchers**")
                st.dataframe(sel.style.format({c: "{:,.2f}" for c in ["Debit Amount", "Credit Amount", "amount"]}),
                             **FULL)
                st.download_button("⬇️ Download DayBook_Sample_Selected.xlsx",
                                   to_excel_bytes({"DayBook_Sample": sel, "Summary": summ}),
                                   f"{file_prefix()}DayBook_Sample_Selected.xlsx", XLSX_MIME)
        except Exception:
            show_error("Day Book processing failed.", traceback.format_exc())
