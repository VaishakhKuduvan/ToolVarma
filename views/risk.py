"""Risk Assessment — SA 315 (Revised). Weighted questionnaire across 8 domains."""
import traceback

import pandas as pd
import streamlit as st

from core.state import FULL, XLSX_MIME, default, file_prefix, show_error, to_excel_bytes

LOW4 = {"Low": 0.0, "Moderate": 0.34, "Elevated": 0.67, "High": 1.0}
YN = {"No": 0.0, "Yes": 1.0}

answers = []  # (section, question, answer) for the downloadable workpaper


def ask(col, section, label, options, max_pts, key, help_=None):
    """Selectbox whose options map to a fraction of max_pts. Returns (score, max)."""
    default(key, next(iter(options)))
    choice = col.selectbox(label, list(options.keys()), key=key, help=help_)
    answers.append((section, label, choice))
    return options[choice] * max_pts, max_pts


def band(pct):
    return "🔴 High" if pct >= 60 else ("🟡 Medium" if pct >= 40 else "🟢 Low")


st.title("🎯 Risk Assessment — SA 315 (Revised)")
st.caption("Answer all sections; the weighted score drives materiality and sampling.")
sections = {}

sec = "1. Team Discussion & Background"
with st.expander("1️⃣ Engagement Team Discussion & Background", expanded=True):
    c1, c2 = st.columns(2)
    items = [
        ask(c1, sec, "First-year audit?", {"No": 0.0, "Yes": 1.0}, 10, "w_s1a"),
        ask(c2, sec, "Entity type", {
            "Private Limited": 0.3, "Unlisted Public": 0.4, "Listed Company": 0.7, "NBFC": 0.9,
            "Bank": 1.0, "Insurance": 0.9, "PSU": 0.6, "Section 8": 0.4, "Partnership-LLP": 0.2}, 10, "w_s1b"),
    ]
    default("w_s1c", "Good")
    integ = c1.select_slider("Management integrity", ["Strong", "Good", "Average", "Questionable", "Weak"],
                             key="w_s1c")
    answers.append((sec, "Management integrity", integ))
    items.append(({"Strong": 0, "Good": .25, "Average": .5, "Questionable": .75, "Weak": 1}[integ] * 20, 20))
    items += [
        ask(c2, sec, "Prior material misstatements",
            {"None": 0.0, "Few-Minor": 0.33, "Multiple": 0.67, "Significant": 1.0}, 15, "w_s1d"),
        ask(c1, sec, "Predecessor auditor issues", {"N/A-None": 0.0, "Unknown": 0.5, "Yes": 1.0}, 10, "w_s1e"),
        ask(c2, sec, "CFO / key management change", YN, 10, "w_s1f"),
    ]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

sec = "2. External & Industry"
with st.expander("2️⃣ External & Industry Factors"):
    c1, c2 = st.columns(2)
    labels = ["Industry risk", "Regulatory intensity", "Recent regulatory changes",
              "Economic / macro pressure", "Financial reporting framework change", "ERP / IT migration"]
    items = [ask(c1 if i % 2 == 0 else c2, sec, l, LOW4, 10, f"w_s2_{i}") for i, l in enumerate(labels)]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

sec = "3. Nature of Entity"
with st.expander("3️⃣ Nature of the Entity"):
    c1, c2 = st.columns(2)
    spec = [("Group complexity", 12), ("Geographic spread", 10), ("M&A / restructuring", 12),
            ("Leverage intensity", 12), ("Related Party Transaction intensity", 14), ("FX / hedging exposure", 10)]
    items = [ask(c1 if i % 2 == 0 else c2, sec, l, LOW4, m, f"w_s3_{i}") for i, (l, m) in enumerate(spec)]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

sec = "4. Objectives & Business Risks"
with st.expander("4️⃣ Objectives, Strategies & Business Risks"):
    c1, c2 = st.columns(2)
    spec = [("Growth aggressiveness", 12), ("Debt covenant pressure", 14), ("KPI-linked compensation", 10),
            ("New products / markets", 10), ("Cost-cutting / restructuring", 10),
            ("Going concern indicators", 24)]
    items = [ask(c1 if i % 2 == 0 else c2, sec, l, LOW4, m, f"w_s4_{i}") for i, (l, m) in enumerate(spec)]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

sec = "5. Accounting Policies & Estimates"
with st.expander("5️⃣ Accounting Policies & Estimates"):
    c1, c2 = st.columns(2)
    spec = [("Policy changes", 10), ("Subjectivity of estimates", 15), ("Unusual / one-off transactions", 12),
            ("Complex financial instruments", 13), ("Revenue recognition complexity", 13),
            ("Aggressive accounting indication", 12)]
    items = [ask(c1 if i % 2 == 0 else c2, sec, l, LOW4, m, f"w_s5_{i}") for i, (l, m) in enumerate(spec)]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

sec = "6. Internal Control"
with st.expander("6️⃣ Internal Control (COSO 5 Components)"):
    c1, c2 = st.columns(2)
    scale = {"Strong": 1, "Adequate": 2, "Weak": 4, "Very Weak / Absent": 5}
    spec = [("Control Environment", 5), ("Risk Assessment process", 3), ("Information System / ERP", 3),
            ("Control Activities", 4), ("Monitoring", 3), ("IT General Controls", 3)]
    s = m = 0
    for i, (l, w) in enumerate(spec):
        default(f"w_s6_{i}", "Strong")
        ch = (c1 if i % 2 == 0 else c2).selectbox(l, list(scale), key=f"w_s6_{i}")
        answers.append((sec, l, ch))
        s += scale[ch] * w
        m += 5 * w
    sections[sec] = (s, m)

sec = "7. Fraud Risk (SA 240)"
with st.expander("7️⃣ Fraud Risk Indicators (SA 240 — Fraud Triangle)"):
    c1, c2 = st.columns(2)
    for k in ("w_s7a", "w_s7b", "w_s7c"):
        default(k, 2)
    default("w_s7d", "No")
    default("w_s7e", "No")
    inc = c1.slider("Incentive / Pressure (1=Low, 5=High)", 1, 5, key="w_s7a")
    opp = c2.slider("Opportunity (1=Low, 5=High)", 1, 5, key="w_s7b")
    rat = c1.slider("Rationalisation (1=Low, 5=High)", 1, 5, key="w_s7c")
    wb = c2.selectbox("Whistleblower complaints", ["No", "Yes"], key="w_s7d")
    pf = c1.selectbox("Prior fraud history", ["No", "Yes"], key="w_s7e")
    answers += [(sec, "Incentive / Pressure", inc), (sec, "Opportunity", opp), (sec, "Rationalisation", rat),
                (sec, "Whistleblower complaints", wb), (sec, "Prior fraud history", pf)]
    vals = [inc, opp, rat, 5 if wb == "Yes" else 1, 5 if pf == "Yes" else 1]
    sections[sec] = (sum(3 * v for v in vals), 75)

sec = "8. India Regulatory & Compliance"
with st.expander("8️⃣ India-Specific Regulatory & Compliance Risks"):
    c1, c2 = st.columns(2)
    items = [
        ask(c1, sec, "Prior CARO adverse observations", {"None": 0, "Minor": .4, "Multiple": .7, "Serious": 1}, 10,
            "w_s8a"),
        ask(c2, sec, "GST / TDS reconciliation status", {"Fully reconciled": 0, "Minor differences": .4,
                                                         "Significant differences": .7, "Not reconciled": 1}, 10,
            "w_s8b"),
        ask(c1, sec, "Sec 185/186 related-party loans",
            {"None": 0, "Within limits": .3, "Near limits": .6, "Breach / Unclear": 1}, 10, "w_s8c"),
        ask(c2, sec, "Material litigations / notices",
            {"None": 0, "Few-Minor": .35, "Material": .7, "Multiple Material": 1}, 10, "w_s8d"),
        ask(c1, sec, "CSR compliance", {"Not applicable": 0, "Compliant": .1, "Partial": .6, "Non-compliant": 1}, 10,
            "w_s8e"),
        ask(c2, sec, "FEMA / ODI / ECB transactions",
            {"None": 0, "Few-Routine": .35, "Significant": .7, "Complex / Delayed filings": 1}, 10, "w_s8f"),
    ]
    sections[sec] = (sum(i[0] for i in items), sum(i[1] for i in items))

if st.button("🎯 Compute Risk Assessment", type="primary"):
    try:
        tot_s = sum(v[0] for v in sections.values())
        tot_m = sum(v[1] for v in sections.values())
        pct = tot_s / tot_m * 100
        level = "High" if pct >= 60 else ("Medium" if pct >= 40 else "Low")
        rows = [{"Section": k, "Score": round(v[0], 1), "Max": v[1],
                 "%": round(v[0] / v[1] * 100, 1), "Risk Band": band(v[0] / v[1] * 100)}
                for k, v in sections.items()]
        st.session_state.risk_level = level
        st.session_state.risk_score_pct = pct
        st.session_state.risk_breakdown = pd.DataFrame(rows)
        st.session_state.risk_answers = pd.DataFrame(answers, columns=["Section", "Question", "Answer"])
    except Exception:
        show_error("Risk computation failed.", traceback.format_exc())

if st.session_state.risk_level:
    approach = {"High": "Substantive-heavy approach; extended testing, limited/no controls reliance",
                "Medium": "Balanced approach: combination of tests of controls and substantive procedures",
                "Low": "Controls reliance possible; reduced substantive testing"}[st.session_state.risk_level]
    icon = {"High": "🔴", "Medium": "🟡", "Low": "🟢"}[st.session_state.risk_level]
    m1, m2, m3 = st.columns(3)
    m1.metric("Risk Score", f"{st.session_state.risk_score_pct:.1f}%")
    m2.metric("Risk Level", f"{icon} {st.session_state.risk_level}")
    m3.metric("Audit Approach", st.session_state.risk_level, help=approach)
    st.info(f"**Recommended approach:** {approach}")
    st.dataframe(st.session_state.risk_breakdown, **FULL, hide_index=True)
    summary = pd.DataFrame([["Risk Score %", round(st.session_state.risk_score_pct, 1)],
                            ["Risk Level", st.session_state.risk_level],
                            ["Audit Approach", approach]], columns=["Item", "Value"])
    st.download_button("⬇️ Download Risk Assessment workpaper (.xlsx)",
                       to_excel_bytes({"Summary": summary, "Section_Scores": st.session_state.risk_breakdown,
                                       "Answers": st.session_state.get("risk_answers", pd.DataFrame())}),
                       f"{file_prefix()}Risk_Assessment_SA315.xlsx", XLSX_MIME)
    st.success("✅ Risk assessment saved. Proceed to **2. Materiality**.")
    st.page_link("views/materiality.py", label="Go to Materiality (SA 320)", icon="➡️")
