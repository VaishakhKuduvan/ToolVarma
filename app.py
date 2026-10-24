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

st.set_page_config(page_title="Audit Suite", page_icon="🏛️", layout="wide")
init_state()

PAGES = {
    "Overview": [
        st.Page("views/home.py", title="Home", icon="🏠", default=True),
    ],
    "Core audit flow": [
        st.Page("views/risk.py", title="1. Risk Assessment (SA 315)", icon="🎯"),
        st.Page("views/materiality.py", title="2. Materiality (SA 320)", icon="💰"),
        st.Page("views/sampling.py", title="3. Sampling (SA 530)", icon="📋"),
    ],
    "Analytics & tools": [
        st.Page("views/benford.py", title="Benford's Law Analysis", icon="📊"),
        st.Page("views/ledger_ai.py", title="AI Ledger Mapping", icon="🤖"),
        st.Page("views/confirmations.py", title="External Confirmations", icon="✉️"),
    ],
}

page = st.navigation(PAGES)
sidebar_status()
page.run()
