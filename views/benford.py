"""Benford's Law Audit Analyser — Day Book edition (CAG of India paper, Nigrini 5-test + SA 320 tiers)."""
from io import BytesIO

import pandas as pd
import streamlit as st

from core.benford import (REQUIRED_COLS, TIER_LABELS, build_excel_report, for_screen, make_plotly_chart,
                          plotly_to_png_bytes, run_tiers, summary_paragraph, tier_thresholds)
from core.state import FULL, XLSX_MIME, daybook_source, default, file_prefix, materiality_rupees
from core.ui import ICONS, page_header, section

MAT_KEYS = {"OM": "w_bf_om", "PM": "w_bf_pm", "SM": "w_bf_sm", "CTT": "w_bf_ctt"}
MAT_FALLBACK = {"OM": 1000000.0, "PM": 750000.0, "SM": 500000.0, "CTT": 50000.0}
SHORT_TEST = {"1. First Digit Test": "First digit", "2. Second Digit Test": "Second digit",
              "3. First-Two Digits Test": "First two digits", "4. First-Three Digits Test": "First three digits",
              "5. Last-Two Digits Test": "Last two digits"}
CONFORMITY_BADGE = {"Close Conformity": "green", "Acceptable Conformity": "blue",
                    "Marginally Acceptable": "orange", "Nonconformity": "red"}
DEDUP_HELP = {
    "strict": "Strict: each transaction appears in one sheet only (highest tier, most specific test).",
    "split": "Split: leading-digit tests and the last-two-digit test are de-duplicated separately "
             "(at most two sheets per transaction).",
    "none": "None: a transaction can appear in every sheet where it was flagged.",
}


@st.cache_data(show_spinner=False)
def load_excel_cached(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), engine="openpyxl")


def _load_from_materiality():
    for k, v in materiality_rupees().items():
        st.session_state[MAT_KEYS[k]] = round(float(v), 2)


page_header("Benford's Law analysis", eyebrow="CAG of India research paper · Nigrini tests · SA 320",
            intro="Tests whether the leading and trailing digits of day book amounts follow Benford's Law, at each "
                  "materiality level. Digit groups that deviate significantly point to vouchers worth examining.")

# ─── 1. MATERIALITY ─────────────────────────────────────────────────────────
section("1. Materiality thresholds", "Amounts in ₹. Each threshold defines a tier of transactions to test.")
shared = materiality_rupees()
for k, key in MAT_KEYS.items():
    default(key, round(float(shared[k]), 2) if shared else MAT_FALLBACK[k])
with st.container(border=True):
    if shared:
        c_info, c_btn = st.columns([3, 2], vertical_alignment="center")
        c_info.markdown("Values from the **Materiality** step are available.")
        c_btn.button("Use materiality step values", on_click=_load_from_materiality, icon=ICONS["materiality"])
    else:
        st.caption("Enter thresholds here, or calculate them in the Materiality step to fill them in.")
    mc1, mc2, mc3, mc4 = st.columns(4)
    OM = mc1.number_input("Overall materiality", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_om")
    PM = mc2.number_input("Performance materiality", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_pm")
    SM = mc3.number_input("Specific materiality", min_value=0.0, step=10000.0, format="%.2f", key="w_bf_sm")
    CTT = mc4.number_input("Clearly trivial", min_value=0.0, step=1000.0, format="%.2f", key="w_bf_ctt")
materiality = {"OM": OM, "PM": PM, "SM": SM, "CTT": CTT}

if PM > OM:
    st.error("Performance materiality cannot be higher than overall materiality.")
elif not (CTT <= SM <= PM <= OM):
    st.warning("Thresholds are usually ordered: clearly trivial ≤ specific ≤ performance ≤ overall.")
if OM > 0 and CTT > 0.10 * OM:
    st.warning(f"Clearly trivial (₹{CTT:,.0f}) is more than 10% of overall materiality; it is usually about 5%.")

# ─── 2. DAY BOOK ────────────────────────────────────────────────────────────
section("2. Day book", "Required columns: " + ", ".join(REQUIRED_COLS))
file_bytes, file_name = daybook_source("bf_up", "Day book (.xlsx)")
if file_bytes is None:
    st.stop()

try:
    with st.spinner(f"Reading {file_name} ({len(file_bytes) / (1024 * 1024):.1f} MB). Large files can take a minute."):
        df = load_excel_cached(file_bytes)
except Exception as e:
    st.error(f"The file could not be read: {e}")
    st.stop()

missing = [c for c in REQUIRED_COLS if c not in df.columns]
extra = [c for c in df.columns if c not in REQUIRED_COLS]
if missing:
    st.error("Missing required columns: " + ", ".join(missing))
    st.stop()

original_cols = list(df.columns)
df = df.reset_index(drop=True).copy()
for col in ["Debit Amount", "Credit Amount"]:
    df[col] = pd.to_numeric(df[col], errors="coerce")
df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=True)  # Indian dd/mm/yyyy
bad_dates = df["Date"].isna().sum()
n_records = len(df)

st.success(f"{file_name}: {n_records:,} records, all required columns present.")
if bad_dates:
    st.warning(f"{bad_dates:,} rows have a date that could not be read.")
if n_records < 5000:
    st.warning("Fewer than 5,000 records. The CAG paper notes small populations can show false spikes.")
with st.expander("Preview the first 10 rows"):
    if extra:
        st.caption("Extra columns kept in the output: " + ", ".join(map(str, extra)))
    st.dataframe(df.head(10), **FULL)

# ─── 3. SETTINGS ────────────────────────────────────────────────────────────
section("3. Analysis settings")
with st.container(border=True):
    c1, c2, c3 = st.columns(3)
    vch_opts = ["(All Voucher Types)"] + sorted(df["Vch Type"].dropna().astype(str).unique().tolist())
    if st.session_state.get("w_bf_vch") not in vch_opts:
        st.session_state.w_bf_vch = vch_opts[0]
    default("w_bf_target", "Debit Amount")
    default("w_bf_tiers", ["Full Population", "Above CTT", "Above PM"])
    selected_vch = c1.selectbox("Voucher type", vch_opts, key="w_bf_vch",
                                format_func=lambda v: "All voucher types" if v == "(All Voucher Types)" else v)
    target_col = c2.selectbox("Amount column", ["Debit Amount", "Credit Amount"], key="w_bf_target")
    tier_choice = c3.multiselect("Materiality tiers", TIER_LABELS, key="w_bf_tiers")

    default("w_bf_dedup", "strict")
    dedup_mode = st.radio("When a transaction is flagged by more than one test",
                          options=["strict", "split", "none"], key="w_bf_dedup", horizontal=True,
                          format_func=lambda x: {"strict": "Strict", "split": "Split", "none": "None"}[x])
    st.caption(DEDUP_HELP[dedup_mode])

    default("w_bf_charts", True)
    default("w_bf_dev", True)
    col_a, col_b = st.columns(2)
    include_charts_excel = col_a.toggle("Include charts in the Excel workpaper", key="w_bf_charts",
                                        help="Adds 30 to 90 seconds, depending on the number of tiers.")
    show_deviation_panel = col_b.toggle("Show the deviation panel under each chart", key="w_bf_dev")

config = (file_name, len(file_bytes), selected_vch, target_col, tuple(tier_choice), dedup_mode,
          include_charts_excel, tuple(materiality.values()))

if st.button("Run analysis", type="primary", icon=ICONS["run"], disabled=not tier_choice):
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
                progress.progress(1.0, text=f"Drawing chart: {tier_label}, {SHORT_TEST.get(r['name'], r['name'])}")
                try:
                    fig = make_plotly_chart(r, tier_label, show_deviation=show_deviation_panel)
                    chart_buffers[(tier_label, r["name"])] = plotly_to_png_bytes(fig)
                except Exception as e:
                    chart_error = str(e)

    excel_bytes, excel_error = None, None
    if tier_outputs:
        progress.progress(1.0, text="Building the Excel workpaper...")
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
    st.stop()

# ─── 4. RESULTS ─────────────────────────────────────────────────────────────
st.divider()
section("4. Results")
if res["config"] != config:
    st.warning("The settings or file have changed since this run. Run the analysis again to update the results.")

rows = [{"Tier": t, "Threshold (₹)": p["threshold"], "Test": SHORT_TEST.get(r["name"], r["name"]), "n": r["n"],
         "MAD": round(r["mad"], 5), "Conformity": r["conformity"], "Spikes": len(r["spike_digits"])}
        for t, p in res["tier_outputs"].items() for r in p["results"]]
total_spikes = sum(r["Spikes"] for r in rows)
nonconform = sum(r["Conformity"] == "Nonconformity" for r in rows)
k1, k2, k3 = st.columns(3)
k1.metric("Tiers tested", len(res["tier_outputs"]), border=True)
k2.metric("Tests showing nonconformity", f"{nonconform} of {len(rows)}", border=True)
k3.metric("Digit groups with spikes", total_spikes, border=True, help="Digit groups with Z above 1.96")

for tier_label, n in res["skipped"].items():
    st.info(f"{tier_label} was skipped: only {n} records, and at least 30 are needed.")

if res["excel"]:
    st.download_button("Download the Excel workpaper", res["excel"],
                       f"{file_prefix()}Benford_Audit_{res['target_col'].replace(' ', '')}_{res['dedup_mode']}.xlsx",
                       XLSX_MIME, type="primary", icon=ICONS["download"])
    st.caption("Sheets: materiality setup, consolidated summary, unique anomalies (one row per transaction), "
               "charts, flagged transactions per tier and test, and a spike index.")
elif res["excel_error"]:
    st.error(f"The Excel workpaper could not be built: {res['excel_error']}")
if res["chart_error"]:
    st.warning("Charts could not be drawn for Excel on this server, so the workpaper has no images. "
               f"Details: {res['chart_error']}")

if rows:
    with st.expander("Summary of all tests", expanded=True):
        summary = pd.DataFrame(rows)
        st.dataframe(summary.style.map(
            lambda v: "font-weight: 600" if v == "Nonconformity" else "", subset=["Conformity"]),
            **FULL, hide_index=True)

for tier_label, payload in res["tier_outputs"].items():
    threshold = payload["threshold"]
    label = tier_label if threshold == 0 else f"{tier_label} (≥ ₹{threshold:,.0f})"
    with st.container(border=True):
        st.markdown(f"**{label}** · {len(payload['source_df']):,} records")
        test_tabs = st.tabs([SHORT_TEST.get(r["name"], r["name"]) for r in payload["results"]])
        for tab, result in zip(test_tabs, payload["results"]):
            with tab:
                st.badge(result["conformity"], color=CONFORMITY_BADGE.get(result["conformity"], "gray"))
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Records", f"{result['n']:,}")
                k2.metric("MAD", f"{result['mad']:.5f}", help="Mean absolute deviation from Benford's Law")
                k3.metric("Chi-square", f"{result['chi2']:.2f}")
                k4.metric("p-value", f"{result['p_val']:.4f}")
                fig = for_screen(make_plotly_chart(result, tier_label, show_deviation=res["show_dev"]))
                st.plotly_chart(fig, **FULL, key=f"chart_{tier_label}_{result['name']}",
                                config={"displaylogo": False,
                                        "toImageButtonOptions": {"format": "png",
                                                                 "filename": f"benford_{tier_label}_{result['name']}",
                                                                 "height": 600, "width": 1200, "scale": 2}})
                st.markdown(summary_paragraph(result, res["target_col"], tier_label, threshold))
                with st.expander("Digit table"):
                    st.dataframe(result["df"].style.format({
                        "Observed %": "{:.3f}", "Expected %": "{:.3f}", "Deviation %": "{:+.3f}",
                        "Z-Statistic": "{:.3f}", "Expected Count": "{:.2f}"}), **FULL, hide_index=True)

with st.expander("Records in each tier"):
    st.dataframe(res["counts"], **FULL, hide_index=True)
