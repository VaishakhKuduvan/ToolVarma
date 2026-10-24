import streamlit as st

st.title("🏛️ Audit Suite")
st.caption("Risk (SA 315) → Materiality (SA 320) → Sampling (SA 530) · Benford's Law · AI Ledger Mapping · "
           "External Confirmations | For Indian Audit Practitioners")

st.markdown(
    "Work through the **core audit flow** in order: the risk level from step 1 drives materiality in step 2, "
    "and both drive sample selection in step 3. The materiality you compute is also offered to the "
    "Benford's Law page, and a Day Book uploaded on one page can be reused on the other.")

done_risk = st.session_state.risk_level is not None
done_mat = st.session_state.overall_materiality is not None

st.subheader("Core audit flow")
c1, c2, c3 = st.columns(3)
with c1.container(border=True):
    st.markdown("#### 🎯 1. Risk Assessment")
    st.caption("SA 315 (Revised) · SA 240")
    st.write("Weighted questionnaire across 8 risk domains, producing a risk score and an audit approach.")
    st.markdown("✅ Done" if done_risk else "⬜ Not started")
    st.page_link("views/risk.py", label="Open Risk Assessment", icon="➡️")
with c2.container(border=True):
    st.markdown("#### 💰 2. Materiality")
    st.caption("SA 320")
    st.write("Recommends a benchmark from 3 years of financials, then computes OM, PM, SM and CTT.")
    st.markdown("✅ Done" if done_mat else ("⬜ Ready to start" if done_risk else "🔒 Needs step 1"))
    st.page_link("views/materiality.py", label="Open Materiality", icon="➡️")
with c3.container(border=True):
    st.markdown("#### 📋 3. Sampling")
    st.caption("SA 530")
    st.write("Selects key items, red flags and value-weighted samples from a Trial Balance and a Day Book.")
    st.markdown("⬜ Ready to start" if done_mat else "🔒 Needs steps 1 and 2")
    st.page_link("views/sampling.py", label="Open Sampling", icon="➡️")

st.subheader("Analytics & tools")
d1, d2, d3 = st.columns(3)
with d1.container(border=True):
    st.markdown("#### 📊 Benford's Law")
    st.caption("CAG of India paper · Nigrini 5-test")
    st.write("First, second, first-two, first-three and last-two digit tests per materiality tier, "
             "with an Excel workpaper of unique anomalies.")
    st.page_link("views/benford.py", label="Open Benford's Law", icon="➡️")
with d2.container(border=True):
    st.markdown("#### 🤖 AI Ledger Mapping")
    st.caption("Google Gemini · Schedule III")
    st.write("Maps each trial-balance ledger, scores its risk 1–5 and writes findings with suggested procedures.")
    st.page_link("views/ledger_ai.py", label="Open AI Ledger Mapping", icon="➡️")
with d3.container(border=True):
    st.markdown("#### ✉️ External Confirmations")
    st.caption("SA 505")
    st.write("Generates one confirmation letter per party from a .docx or fillable .pdf template and an Excel list.")
    st.page_link("views/confirmations.py", label="Open External Confirmations", icon="➡️")

st.info("Your answers and results are kept while this browser tab is open. "
        "Download the Excel outputs on each page to keep a permanent record.")
