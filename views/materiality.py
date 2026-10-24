"""Materiality — SA 320. Benchmark recommendation from 3 years of financials, then OM / PM / SM / CTT."""
import traceback

import numpy as np
import pandas as pd
import streamlit as st

from core.state import FULL, XLSX_MIME, default, file_prefix, show_error, to_excel_bytes

st.title("💰 Materiality — SA 320")
if not st.session_state.risk_level:
    st.warning("⚠️ Please complete and compute **1. Risk Assessment** first.")
    st.page_link("views/risk.py", label="Go to Risk Assessment", icon="⬅️")
else:
    st.info(f"Carried-over Risk Level: **{st.session_state.risk_level}** "
            f"({st.session_state.risk_score_pct:.1f}%)")

ENTITY_TYPES = ["Private Limited / Others", "Listed", "Unlisted Public", "NBFC", "Bank", "Insurance",
                "PSU", "Section 8 / NPO", "Start-up / Loss-making"]

with st.expander("🏢 Entity Profile", expanded=True):
    default("w_m_entity", ENTITY_TYPES[0])
    default("w_m_users", [])
    entity_type = st.selectbox("Entity Type", ENTITY_TYPES, key="w_m_entity")
    users = st.multiselect("Primary users of FS", [
        "Investors", "Bankers", "Regulators", "Promoters", "Donors", "Government",
        "Customers/Suppliers", "Tax Authorities"], key="w_m_users")

with st.expander("📊 3-Year Financials (₹ Lakhs)", expanded=True):
    cols = st.columns(3)
    fin = {}
    for col, yr in zip(cols, ["Year T-2", "Year T-1", "Year T (Current)"]):
        col.markdown(f"**{yr}**")
        fin[yr] = {}
        for code, field, label in [("pbt", "PBT", "PBT"), ("rev", "Revenue", "Revenue / Turnover"),
                                   ("ast", "Assets", "Total Assets"), ("nw", "NW", "Net Worth"),
                                   ("exp", "Exp", "Total Expenditure")]:
            key = f"w_{code}_{yr}"
            default(key, 0.0)
            fin[yr][field] = col.number_input(label, key=key)
yrs = list(fin)
cur = fin[yrs[2]]
pbts = [fin[y]["PBT"] for y in yrs]
avg_pbt = float(np.mean(pbts))

# ---- Auto recommendation (decision tree)
volatile = (min(pbts) < 0 < max(pbts)) or (
    abs(avg_pbt) > 0 and (max(pbts) - min(pbts)) > 0.5 * abs(avg_pbt)) or (avg_pbt == 0 and max(pbts) != min(pbts))
if entity_type in ("NBFC", "Bank", "Insurance"):
    bm, pct, val = "Total Assets", 0.5, cur["Assets"]
    why = "Regulated financial entity; users (regulators, depositors, investors) focus on the balance sheet, so Total Assets is used."
elif entity_type == "Section 8 / NPO":
    bm, pct, val = "Total Expenditure", 0.5, cur["Exp"]
    why = "Not-for-profit; donors and regulators focus on how funds are spent, so Total Expenditure is used."
elif entity_type == "Start-up / Loss-making" or cur["PBT"] <= 0:
    bm, pct, val = "Revenue", 0.75, cur["Revenue"]
    why = "Start-up / loss-making or nil current-year profit; PBT is unreliable as a benchmark, so Revenue is used."
elif volatile:
    bm, pct, val = "3-Year Average PBT", 5.0, avg_pbt
    why = "PBT is volatile across the 3 years; averaging normalises one-off items (SA 320 A4)."
elif entity_type == "PSU":
    bm, pct, val = "Revenue", 0.5, cur["Revenue"]
    why = "PSU; Revenue is a stable benchmark where profit is not the primary objective."
else:
    bm, pct, val = "Current-year PBT", 5.0, cur["PBT"]
    why = "Stable profitable entity; PBT is the customary benchmark within SA 320's usual 5–10% range."

# The recommendation is used unless the manual override below replaces it
st.session_state.benchmark_choice, st.session_state.benchmark_pct = bm, pct
st.session_state.benchmark_value, st.session_state.benchmark_reason = val, why

st.markdown("### 🤖 Recommended Benchmark")
r1, r2, r3 = st.columns(3)
r1.metric("Benchmark", bm)
r2.metric("Percentage", f"{pct}%")
r3.metric("Benchmark Value (₹ Lakhs)", f"{val:,.2f}")
st.info(f"**Rationale:** {why}")

st.markdown("### 🔍 Cross-check of Alternative Benchmarks")
xr = pd.DataFrame([
    ["5% of Current PBT", 0.05 * cur["PBT"], "Stable, profit-oriented entities"],
    ["5% of 3-yr Avg PBT", 0.05 * avg_pbt, "Volatile profits"],
    ["0.5% of Revenue", 0.005 * cur["Revenue"], "Loss-making / PSU / low margin"],
    ["1% of Revenue", 0.01 * cur["Revenue"], "Loss-making / low-margin entities"],
    ["0.5% of Total Assets", 0.005 * cur["Assets"], "Banks / NBFC / Insurance"],
    ["1% of Net Worth", 0.01 * cur["NW"], "Asset-heavy / equity-focused users"],
    ["0.5% of Total Expenditure", 0.005 * cur["Exp"], "NPO / Section 8"],
], columns=["Basis", "Amount (₹ Lakhs)", "Used For"])
xr["Amount (₹ Lakhs)"] = xr["Amount (₹ Lakhs)"].map("{:,.2f}".format)
st.dataframe(xr, **FULL, hide_index=True)


def _seed_override():
    """When override is switched on, start from the current recommendation."""
    if st.session_state.w_m_override:
        st.session_state.w_ov_val = float(val)
        st.session_state.w_ov_pct = float(pct)
        st.session_state.w_ov_reason = why


default("w_m_override", False)
if st.checkbox("✏️ Manual override of benchmark", key="w_m_override", on_change=_seed_override):
    o1, o2 = st.columns(2)
    default("w_ov_val", float(val))
    default("w_ov_pct", float(pct))
    default("w_ov_reason", why)
    st.session_state.benchmark_value = o1.number_input("Benchmark value (₹ Lakhs)", key="w_ov_val")
    st.session_state.benchmark_pct = o2.number_input("Percentage (%)", step=0.25, key="w_ov_pct")
    st.session_state.benchmark_reason = st.text_area("Rationale", key="w_ov_reason")
    st.session_state.benchmark_choice = f"{bm} (manual override)"

with st.expander("🧮 Compute Materiality", expanded=True):
    current_inputs = (st.session_state.benchmark_value, st.session_state.benchmark_pct, st.session_state.risk_level)
    if not st.session_state.risk_level:
        st.warning("Complete Risk Assessment to compute materiality.")
    elif st.button("Compute Materiality", type="primary"):
        try:
            om = st.session_state.benchmark_value * st.session_state.benchmark_pct / 100
            pm_pct = {"High": 55, "Medium": 70, "Low": 80}[st.session_state.risk_level]
            pm = om * pm_pct / 100
            st.session_state.overall_materiality = om
            st.session_state.pm_pct = pm_pct
            st.session_state.performance_materiality = pm
            st.session_state.clearly_trivial = om * 0.05
            st.session_state.specific_materiality = max(om * 0.10, 1.0)
            st.session_state.materiality_inputs = current_inputs
        except Exception:
            show_error("Materiality computation failed.", traceback.format_exc())
    if st.session_state.overall_materiality is not None:
        if st.session_state.get("materiality_inputs") != current_inputs:
            st.warning("⚠️ Inputs or risk level changed since materiality was last computed — click "
                       "**Compute Materiality** again to refresh.")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Overall Materiality", f"₹ {st.session_state.overall_materiality:,.2f} L",
                  help="Benchmark value × benchmark %")
        k2.metric("Performance Materiality", f"₹ {st.session_state.performance_materiality:,.2f} L",
                  help=f"OM × {st.session_state.pm_pct}% (High 55 / Medium 70 / Low 80)")
        k3.metric("Specific Materiality", f"₹ {st.session_state.specific_materiality:,.2f} L",
                  help="max(10% of OM, ₹1 Lakh) for RPT, director remuneration, CSR etc.")
        k4.metric("Clearly Trivial", f"₹ {st.session_state.clearly_trivial:,.2f} L", help="5% of OM")
        summ = pd.DataFrame([
            ["Benchmark", st.session_state.benchmark_choice, st.session_state.benchmark_value],
            ["Benchmark %", f"{st.session_state.benchmark_pct}%", None],
            ["Overall Materiality (OM)", "Benchmark × %", st.session_state.overall_materiality],
            [f"Performance Materiality ({st.session_state.pm_pct}% of OM)", st.session_state.risk_level + " risk",
             st.session_state.performance_materiality],
            ["Specific Materiality", "max(10% OM, 1.0)", st.session_state.specific_materiality],
            ["Clearly Trivial Threshold", "5% of OM", st.session_state.clearly_trivial],
        ], columns=["Item", "Value", "Amount ₹ Lakhs"])
        shown = summ.copy()
        shown["Amount ₹ Lakhs"] = shown["Amount ₹ Lakhs"].map(lambda x: "" if x is None else f"{x:,.2f}")
        st.dataframe(shown, **FULL, hide_index=True)
        profile = pd.DataFrame([["Entity Type", entity_type], ["Primary users", ", ".join(users)],
                                ["Rationale", st.session_state.benchmark_reason]], columns=["Item", "Value"])
        financials = pd.DataFrame(fin).T.rename_axis("Year").reset_index()
        st.download_button("⬇️ Download Materiality workpaper (.xlsx)",
                           to_excel_bytes({"Materiality": summ, "Entity_Profile": profile,
                                           "Financials_Lakhs": financials, "Cross_Check": xr}),
                           f"{file_prefix()}Materiality_SA320.xlsx", XLSX_MIME)
        st.success("✅ Materiality saved. Proceed to **3. Sampling** or **Benford's Law Analysis**.")
        st.page_link("views/sampling.py", label="Go to Sampling (SA 530)", icon="➡️")
