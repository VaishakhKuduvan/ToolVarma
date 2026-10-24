import streamlit as st

from core.ui import ICONS, card_title, muted, page_header, section, step_label

page_header("Audit Suite", eyebrow="For Indian audit practitioners",
            intro="Plan the engagement, set materiality and select samples in one place, then use the analytics "
                  "tools on the same data. Each step feeds the next, and every page gives you an Excel workpaper.")

done_risk = st.session_state.risk_level is not None
done_mat = st.session_state.overall_materiality is not None


def status(done, ready, locked_text):
    if done:
        st.badge("Complete", icon=ICONS["done"], color="green")
    elif ready:
        st.badge("Ready to start", icon=ICONS["pending"], color="blue")
    else:
        st.badge(locked_text, icon=ICONS["lock"], color="gray")


section("Engagement workflow", "Work through these in order. Your answers are kept as you move between pages.")
steps = [
    ("Step 1 · SA 315, SA 240", "Risk assessment",
     "A weighted questionnaire across eight risk areas gives a risk score, level and recommended audit approach.",
     "views/risk.py", (done_risk, True, "")),
    ("Step 2 · SA 320", "Materiality",
     "Recommends a benchmark from three years of financials and calculates overall, performance and specific "
     "materiality and the clearly trivial threshold.",
     "views/materiality.py", (done_mat, done_risk, "Needs step 1")),
    ("Step 3 · SA 530", "Sampling",
     "Selects key items, red flags and value-weighted samples from a trial balance and a day book.",
     "views/sampling.py", (False, done_mat, "Needs steps 1 and 2")),
]
cols = st.columns(3)
for col, (label, title, text, path, state) in zip(cols, steps):
    with col.container(border=True, height="stretch"):
        step_label(label)
        card_title(title)
        muted(text)
        status(*state)
        st.page_link(path, label=f"Open {title.lower()}", icon=ICONS["next"])

section("Analytics and tools", "Use these on their own, or with the materiality and day book from the workflow.")
tools = [
    ("CAG of India paper · Nigrini tests", "Benford's Law",
     "First, second, first-two, first-three and last-two digit tests at each materiality level, with a "
     "workpaper of unique anomalies.", "views/benford.py", "Open Benford's Law"),
    ("Google Gemini · Schedule III", "AI ledger mapping",
     "Maps each trial balance ledger, scores its risk from 1 to 5 and writes findings with suggested procedures.",
     "views/ledger_ai.py", "Open AI ledger mapping"),
    ("SA 505", "External confirmations",
     "Generates one confirmation letter per party from a Word or fillable PDF template and an Excel list.",
     "views/confirmations.py", "Open external confirmations"),
]
cols = st.columns(3)
for col, (label, title, text, path, link) in zip(cols, tools):
    with col.container(border=True, height="stretch"):
        step_label(label)
        card_title(title)
        muted(text)
        st.page_link(path, label=link, icon=ICONS["next"])

st.caption("Work is kept only while this browser tab is open. Download each page's Excel workpaper to keep a record.")
