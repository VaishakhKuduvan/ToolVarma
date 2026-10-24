"""Benford's Law Audit Analyser — Day Book edition (CAG of India paper, Nigrini 5-test + SA 320 tiers)."""
from io import BytesIO

import pandas as pd
import streamlit as st

from core.benford import (REQUIRED_COLS, TIER_LABELS, build_excel_report, make_plotly_chart,
                          plotly_to_png_bytes, run_tiers, summary_paragraph, tier_thresholds)
from core.state import FULL, XLSX_MIME, daybook_source, default, file_prefix, materiality_rupees

MAT_KEYS = {"OM": "w_bf_om", "PM": "w_bf_pm", "SM": "w_bf_sm", "CTT": "w_bf_ctt"}
MAT_FALLBACK = {"OM": 1000000.0, "PM": 750000.0, "SM": 500000.0, "CTT": 50000.0}


@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), engine="openpyxl")


def _load_from_materiality():
    for k, v in materiality_rupees().items():
        st.session_state[MAT_KEYS[k]] = round(float(v), 2)


st.title("📊 Benford's Law Audit Analyser — Daybook Edition")
st.caption("CAG of India Research Paper + SA 320 Materiality Framework | Nigrini 5-Test | Plotly Interactive Charts")

# ─── STEP 0 — MATERIALITY ───────────────────────────────────────────────────
st.subheader("🎯 Step 0 — Define Materiality Levels (SA 320)")
shared = materiality_rupees()
for k, key in MAT_KEYS.items():
    default(key, round(float(shared[k]), 2) if shared else MAT_FALLBACK[k])
if shared:
    c_info, c_btn = st.columns([3, 1])
    c_info.markdown("Materiality from the **Materiality (SA 320)** page is available (converted to ₹).")
    c_btn.button("↻ Use Materiality page values", on_click=_load_from_materiality)
else:
    st.markdown("Set thresholds here, or compute them on the **Materiality (SA 320)** page to fill them in.")

with st.expander("ℹ️ Materiality Quick Reference", expanded=False):
    st.markdown("""
    - **Overall Materiality (OM):** typically 0.5%–1% of revenue / 5% of PBT
    - **Performance Materiality (PM):** 50%–75% of OM
    - **Specific Materiality (SM):** for sensitive accounts — usually < PM
    - **Clearly Trivial Threshold (CTT):** ≈ 5% of OM
    """)

mc1, mc2, mc3, mc4 = st.columns(4)
OM = mc1.number_input("Overall Materiality (₹)", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_om")
PM = mc2.number_input("Performance Materiality (₹)", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_pm")
SM = mc3.number_input("Specific Materiality (₹)", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_sm")
CTT = mc4.number_input("Clearly Trivial Threshold (₹)", min_value=0.0, step=1000.0, format="%.2f", key="w_bf_ctt")
materiality = {"OM": OM, "PM": PM, "SM": SM, "CTT": CTT}

if not (CTT <= SM <= PM <= OM):
    st.warning("⚠️ Recommended hierarchy: CTT ≤ SM ≤ PM ≤ OM. Please review.")
if OM > 0 and CTT > 0.10 * OM:
    st.warning(f"⚠️ CTT ({CTT:,.0f}) is more than 10% of OM — typically CTT ≈ 5% of OM.")
if PM > OM:
    st.error("❌ PM cannot exceed OM.")

st.markdown("---")

# ─── FILE UPLOAD ────────────────────────────────────────────────────────────
file_bytes, file_name = daybook_source("bf_up", "📁 Upload Daybook (.xlsx)")
if file_bytes is None:
    st.info("👆 Set materiality levels above, then upload your daybook (.xlsx) to begin analysis.")
    st.stop()

try:
    with st.spinner(f"📖 Reading Excel file ({len(file_bytes) / (1024 * 1024):.1f} MB)... "
                    "This may take 30-90 seconds for large files."):
        df = load_excel_cached(file_bytes)
    st.success(f"✅ {file_name} loaded — {len(df):,} rows × {len(df.columns)} columns.")
except Exception as e:
    st.error(f"❌ Error reading Excel file: {e}")
    st.stop()

# ─── STEP 1 — VALIDATION ────────────────────────────────────────────────────
st.subheader("📋 Step 1 — Data Validation")
missing = [c for c in REQUIRED_COLS if c not in df.columns]
extra = [c for c in df.columns if c not in REQUIRED_COLS]
if missing:
    st.error(f"❌ Missing required columns: {missing}")
    st.stop()
st.success(f"✅ All {len(REQUIRED_COLS)} required columns present.")
if extra:
    st.info(f"ℹ️ Additional columns found (will be preserved): {extra}")

original_cols = list(df.columns)
df = df.reset_index(drop=True).copy()
for col in ["Debit Amount", "Credit Amount"]:
    df[col] = pd.to_numeric(df[col], errors="coerce")

bad_debit = df["Debit Amount"].isna().sum()
bad_credit = df["Credit Amount"].isna().sum()
if bad_debit or bad_credit:
    st.caption(f"ℹ️ Blank / non-numeric amounts treated as empty — Debit: {bad_debit:,}, Credit: {bad_credit:,}")

df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)  # Indian dd/mm/yyyy
bad_dates = df["Date"].isna().sum()
if bad_dates:
    st.warning(f"⚠️ {bad_dates:,} rows have invalid date format.")

n_records = len(df)
st.info(f"📦 Total daybook records: **{n_records:,}**")
if n_records < 5000:
    st.warning("⚠️ As per CAG paper, < 5,000 records may produce false-positive spikes.")

with st.expander("🔍 Preview Data (first 10 rows)"):
    st.dataframe(df.head(10), **FULL)

# ─── STEP 2 — CONFIG ────────────────────────────────────────────────────────
st.subheader("⚙️ Step 2 — Analysis Configuration")
c1, c2, c3 = st.columns(3)
vch_opts = ["(All Voucher Types)"] + sorted(df["Vch Type"].dropna().astype(str).unique().tolist())
if st.session_state.get("w_bf_vch") not in vch_opts:
    st.session_state.w_bf_vch = vch_opts[0]
default("w_bf_target", "Debit Amount")
default("w_bf_tiers", ["Full Population", "Above CTT", "Above PM"])
selected_vch = c1.selectbox("Filter by Vch Type", vch_opts, key="w_bf_vch")
target_col = c2.selectbox("Target Numeric Column", ["Debit Amount", "Credit Amount"], key="w_bf_target")
tier_choice = c3.multiselect("Materiality Tiers to Run", TIER_LABELS, key="w_bf_tiers")

st.markdown("##### 🔁 De-duplication Strategy")
default("w_bf_dedup", "strict")
dedup_mode = st.radio(
    "How should anomalies appearing in multiple sheets be handled?",
    options=["strict", "split", "none"],
    format_func=lambda x: {
        "strict": "🎯 STRICT — Each transaction appears in exactly ONE sheet "
                  "(highest tier × most specific test).",
        "split": "⚖️ SPLIT — Dedup leading-digit family separately from Last-Two "
                 "test (max 2 sheets per transaction).",
        "none": "📋 NONE — Original behaviour (transactions can repeat).",
    }[x],
    key="w_bf_dedup",
)

col_a, col_b = st.columns(2)
default("w_bf_charts", True)
default("w_bf_dev", True)
include_charts_excel = col_a.checkbox(
    "📊 Embed charts in the Excel workbook", key="w_bf_charts",
    help="Adds ~10-20 seconds to export time but produces audit-ready visual workpapers.")
show_deviation_panel = col_b.checkbox(
    "📈 Show deviation sub-panel on charts", key="w_bf_dev",
    help="Adds a secondary chart strip showing % deviation from Benford expected.")

config = (file_name, len(file_bytes), selected_vch, target_col, tuple(tier_choice), dedup_mode,
          include_charts_excel, tuple(materiality.values()))

if st.button("▶️ Run Benford Analysis", type="primary", disabled=not tier_choice):
    df_f = df[df["Vch Type"].astype(str) == selected_vch].copy() if selected_vch != "(All Voucher Types)" else df
    df_pos = df_f[df_f[target_col].notna() & (df_f[target_col] > 0)].copy()
    thresholds = tier_thresholds(materiality)

    progress = st.progress(0.0, text="Starting analysis...")
    tier_outputs, skipped = run_tiers(df_pos, target_col, thresholds, tier_choice,
                                      lambda f, t: progress.progress(min(f, 1.0), text=t))

    chart_buffers, chart_error = {}, None
    if include_charts_excel:
        for tier_label, payload in tier_outputs.items():
            for r in payload["results"]:
                if chart_error:
                    break
                progress.progress(1.0, text=f"Rendering chart: {tier_label} · {r['name']}")
                try:
                    fig = make_plotly_chart(r, tier_label, show_deviation=show_deviation_panel)
                    chart_buffers[(tier_label, r["name"])] = plotly_to_png_bytes(fig)
                except Exception as e:
                    chart_error = str(e)

    excel_bytes, excel_error = None, None
    if tier_outputs:
        progress.progress(1.0, text="📝 Building Excel workbook...")
        try:
            embed = include_charts_excel and not chart_error
            excel_bytes = build_excel_report(
                tier_outputs=tier_outputs, materiality=materiality, target_col=target_col,
                vch_filter=selected_vch, original_cols=original_cols, dedup_mode=dedup_mode,
                include_charts=embed, chart_buffers=chart_buffers if embed else None,
            ).getvalue()
        except Exception as e:
            excel_error = str(e)
    progress.empty()

    cnt_df = pd.DataFrame({
        "Tier": list(thresholds.keys()),
        "Threshold (₹)": list(thresholds.values()),
        "Records ≥ Threshold": [int((df_pos[target_col] >= t).sum()) for t in thresholds.values()],
    })
    st.session_state.bf_result = {
        "config": config, "tier_outputs": tier_outputs, "skipped": skipped, "counts": cnt_df,
        "excel": excel_bytes, "excel_error": excel_error, "chart_error": chart_error,
        "target_col": target_col, "dedup_mode": dedup_mode, "show_dev": show_deviation_panel,
    }

res = st.session_state.get("bf_result")
if res is None:
    st.info("👉 Click **Run Benford Analysis** to start.")
    st.stop()
if res["config"] != config:
    st.warning("⚠️ Settings or data changed since the last run — results below are from the previous run. "
               "Click **Run Benford Analysis** to refresh.")

st.markdown("##### 📊 Record Counts by Materiality Tier")
st.dataframe(res["counts"], **FULL, hide_index=True)

# ─── STEP 3 — RESULTS PER TIER ──────────────────────────────────────────────
for tier_label, n in res["skipped"].items():
    st.warning(f"⏭️ {tier_label} skipped — too few records ({n}) at this tier (minimum 30).")

for tier_label, payload in res["tier_outputs"].items():
    threshold = payload["threshold"]
    with st.expander(f"🧪 Tier: {tier_label}  (≥ ₹{threshold:,.2f}) — {len(payload['source_df']):,} records",
                     expanded=len(res["tier_outputs"]) == 1):
        test_tabs = st.tabs([r["name"] for r in payload["results"]])
        for tab, result in zip(test_tabs, payload["results"]):
            with tab:
                k1, k2, k3, k4, k5 = st.columns(5)
                k1.metric("n", f"{result['n']:,}")
                k2.metric("MAD", f"{result['mad']:.5f}")
                k3.metric("Chi-square", f"{result['chi2']:.2f}")
                k4.metric("p-value", f"{result['p_val']:.4f}")
                k5.metric("Conformity", result["conformity"])
                fig = make_plotly_chart(result, tier_label, show_deviation=res["show_dev"])
                st.plotly_chart(fig, **FULL, key=f"chart_{tier_label}_{result['name']}",
                                config={"displaylogo": False,
                                        "toImageButtonOptions": {"format": "png",
                                                                 "filename": f"benford_{tier_label}_{result['name']}",
                                                                 "height": 600, "width": 1200, "scale": 2}})
                with st.expander("📊 Digit-wise Table"):
                    st.dataframe(result["df"].style.format({
                        "Observed %": "{:.3f}", "Expected %": "{:.3f}", "Deviation %": "{:+.3f}",
                        "Z-Statistic": "{:.3f}", "Expected Count": "{:.2f}"}), **FULL, hide_index=True)
                st.markdown("##### 📝 Audit Narrative")
                st.markdown(summary_paragraph(result, res["target_col"], tier_label, threshold))

# ─── STEP 4 — CONSOLIDATED SUMMARY ──────────────────────────────────────────
st.markdown("---")
st.header("📑 Step 4 — Consolidated Multi-Tier Summary")
rows = [{"Tier": t, "Threshold (₹)": p["threshold"], "Test": r["name"], "n": r["n"], "MAD": round(r["mad"], 5),
         "Conformity": r["conformity"], "Spikes": len(r["spike_digits"])}
        for t, p in res["tier_outputs"].items() for r in p["results"]]
if rows:
    st.dataframe(pd.DataFrame(rows), **FULL, hide_index=True)
else:
    st.warning("No tiers produced results.")

# ─── STEP 5 — EXCEL EXPORT ──────────────────────────────────────────────────
st.markdown("---")
st.header("⬇️ Step 5 — Download Complete Audit Workpaper")
st.markdown("""
- **Materiality_Setup** — SA 320 thresholds and run settings
- **Consolidated_Summary** — All tests × tiers with post-dedup allocation
- **Master_Unique_Anomalies** — ⭐ ONE row per unique anomaly + full trigger trail
- **Charts** — All graphs (PNG) in one scrollable sheet
- **Sp_[Tier]_[Test]** sheets — De-duplicated spike transactions + inline mini-chart
- **Spike_Reference_Index** — Master index of all flagged digit buckets
""")
if res["chart_error"]:
    st.warning(f"⚠️ Charts could not be rendered for Excel, so the workbook has no images "
               f"(PNG export needs Kaleido and Google Chrome): {res['chart_error']}")
if res["excel"]:
    st.download_button("📥 Download Full Audit Workpaper (.xlsx)", res["excel"],
                       f"{file_prefix()}Benford_Audit_{res['target_col'].replace(' ', '')}_{res['dedup_mode']}.xlsx",
                       XLSX_MIME, type="primary")
elif res["excel_error"]:
    st.error(f"❌ Error building Excel: {res['excel_error']}")
else:
    st.warning("No tier outputs to export.")
