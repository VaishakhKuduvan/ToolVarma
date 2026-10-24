"""
Audit Suite — one Streamlit app for Indian audit practitioners combining:
  * Risk Assessment (SA 315) -> Materiality (SA 320) -> Sampling (SA 530)
  * Benford's Law analysis of the Day Book (Nigrini 5-test + SA 320 tiers)
  * AI Ledger Mapping & Risk Scoring (Gemini, Schedule III)
  * External Confirmation letter generator (.docx / fillable .pdf mail-merge)

Run with:  streamlit run app.py
"""
import streamlit as st

from core.state import init_state, sidebar_status
from core.ui import ICONS, apply_global_styles

st.set_page_config(page_title="Audit Suite", page_icon=":material/account_balance:", layout="wide")
apply_global_styles()
init_state()

PAGES = {
    "": [
        st.Page("views/home.py", title="Overview", icon=ICONS["home"], default=True),
    ],
    "Engagement workflow": [
        st.Page("views/risk.py", title="Risk assessment", icon=ICONS["risk"]),
        st.Page("views/materiality.py", title="Materiality", icon=ICONS["materiality"]),
        st.Page("views/sampling.py", title="Sampling", icon=ICONS["sampling"]),
    ],
    "Analytics and tools": [
        st.Page("views/benford.py", title="Benford's Law", icon=ICONS["benford"]),
        st.Page("views/ledger_ai.py", title="AI ledger mapping", icon=ICONS["ai"]),
        st.Page("views/confirmations.py", title="External confirmations", icon=ICONS["confirmations"]),
    ],
}

page = st.navigation(PAGES)
sidebar_status()
page.run()
