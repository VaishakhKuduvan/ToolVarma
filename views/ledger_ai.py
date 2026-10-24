"""AI Ledger Mapping, Risk Scoring & Analysis (Google Gemini). Streamlit front end for core.ledger_ai."""
import io
import logging
import os
import traceback
from pathlib import Path

import pandas as pd
import streamlit as st

from core.ledger_ai import (ANALYSIS_STYLES, AUTO_MODEL_LABEL, DEFAULT_ANALYSIS_PROMPT, DEFAULT_BATCH_SIZE,
                            MAPPING_STANDARDS, MAX_PREVIEW_ROWS, RISK_TREE_COLORS, STANDARD_IND_AS,
                            AllModelsFailed, ApiError, GeminiClient, LedgerAIEngine, SecretFilter, clean_dataframe,
                            export_workbook)
from core.state import FULL, XLSX_MIME, default, file_prefix, show_error


class ListLogHandler(logging.Handler):
    def __init__(self, sink):
        super().__init__()
        self.sink = sink

    def emit(self, record):
        self.sink.append(self.format(record))
        del self.sink[:-500]


def get_logger():
    """One logger per browser session, writing into session state for the Logs panel."""
    if "ai_logger" not in st.session_state:
        st.session_state.ai_logs = []
        log = logging.Logger("ledger_ai")  # standalone, not shared between sessions
        log.setLevel(logging.INFO)
        handler = ListLogHandler(st.session_state.ai_logs)
        handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
        st.session_state.ai_secret_filter = SecretFilter()
        handler.addFilter(st.session_state.ai_secret_filter)
        log.addHandler(handler)
        st.session_state.ai_logger = log
    return st.session_state.ai_logger


def get_client(api_key):
    log = get_logger()
    st.session_state.ai_secret_filter.secrets = {api_key}
    if st.session_state.get("ai_client_key") != api_key:
        st.session_state.ai_client = GeminiClient(api_key, log)
        st.session_state.ai_client_key = api_key
    return st.session_state.ai_client


def preferred_model():
    choice = st.session_state.w_ai_model
    return None if choice == AUTO_MODEL_LABEL else choice


def env_key():
    return os.environ.get("GEMINI_API_KEY", "")


def risk_style(df):
    if "Risk Score" not in df.columns:
        return df
    def colour(v):
        try:
            return f"background-color: {RISK_TREE_COLORS[int(v)]}; color: #1f1f1f"
        except (KeyError, TypeError, ValueError):
            return ""
    return df.style.map(colour, subset=["Risk Score"])


st.title("🤖 AI Ledger Mapping, Risk Scoring & Analysis")
st.caption("Google Gemini maps each ledger, scores its risk 1–5 and (optionally) writes findings with "
           "suggested audit procedures. The model is chosen automatically with fall-back across every "
           "Gemini model your key can use.")
log = get_logger()

# ------------------------------------------------------------------ API key
with st.expander("🔑 Gemini API key", expanded=not st.session_state.get("ai_verified")):
    default("w_ai_key", env_key())
    k1, k2 = st.columns([3, 1])
    api_key = k1.text_input("API key", type="password", key="w_ai_key",
                            help="Create a free key at https://aistudio.google.com/apikey. "
                                 "It is kept only in this browser session and is never written to files or logs.")
    k2.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
    if k2.button("Verify key", disabled=not api_key.strip()):
        client = get_client(api_key.strip())
        client.models, client.sticky = [], None
        with st.spinner("Verifying API key..."):
            try:
                models = client.list_models()
                if not models:
                    raise ApiError("The key is valid but no Gemini models with generateContent are available.")
                _, used = client.generate("Reply with the single word: OK", preferred=preferred_model())
                st.session_state.ai_verified = True
                st.session_state.ai_models = models
                log.info(f"API key verified. Test call answered by '{used}'. {len(models)} models available.")
                st.success(f"✅ Key verified — test call answered by **{used}**; {len(models)} models available.")
            except (ApiError, AllModelsFailed) as exc:
                st.session_state.ai_verified = False
                log.error(f"API key verification failed: {exc}")
                st.error(f"❌ Verification failed: {exc}")
    st.caption("Tip: set the GEMINI_API_KEY environment variable before starting the app to pre-fill this box.")

# ------------------------------------------------------------------ settings
with st.expander("⚙️ Settings", expanded=True):
    s1, s2, s3 = st.columns(3)
    models = [AUTO_MODEL_LABEL] + st.session_state.get("ai_models", [])
    if st.session_state.get("w_ai_model") not in models:
        st.session_state.w_ai_model = AUTO_MODEL_LABEL
    default("w_ai_standard", STANDARD_IND_AS)
    default("w_ai_batch", DEFAULT_BATCH_SIZE)
    s1.selectbox("Model", models, key="w_ai_model", help="Verify the key to list every model it can use.")
    s2.selectbox("Mapping standard", MAPPING_STANDARDS, key="w_ai_standard")
    s3.number_input("Ledgers per AI request", min_value=5, max_value=100, step=5, key="w_ai_batch")

# ------------------------------------------------------------------ input
st.subheader("1. Input file")
st.caption("Columns A to G with a header row: Ledger Name | Opening Dr | Opening Cr | Current Dr | Current Cr | "
           "Closing Dr | Closing Cr. Existing 'Risk Score' / 'Account Mapping' columns are overwritten.")
up = st.file_uploader("Upload trial balance (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"], key="ai_up")
raw_df = None
if up is not None:
    try:
        if Path(up.name).suffix.lower() == ".csv":
            raw = pd.read_csv(up)
        else:
            sheets = pd.ExcelFile(io.BytesIO(up.getvalue())).sheet_names
            sheet = st.selectbox("Sheet", sheets, key="ai_sheet") if len(sheets) > 1 else sheets[0]
            raw = pd.read_excel(io.BytesIO(up.getvalue()), sheet_name=sheet)
        raw = raw.drop(columns=[c for c in raw.columns if str(c).strip() in
                                ("Risk Score", "Account Mapping", "AI Rationale",
                                 "Schedule III Component", "Schedule III Main Group", "Schedule III Sub Group")])
        raw_df = clean_dataframe(raw)
        file_sig = (up.name, up.size)
        if st.session_state.get("ai_file_sig") != file_sig:  # new file: forget results of the previous one
            st.session_state.ai_file_sig = file_sig
            for k in ("ai_mapped", "ai_analysis", "ai_findings", "ai_analysis_model", "ai_prompt_used"):
                st.session_state.pop(k, None)
            st.session_state.ai_mapping_models = []
        st.success(f"✅ {len(raw_df):,} ledgers loaded from {up.name}")
        with st.expander("Preview", expanded=False):
            st.dataframe(raw_df.head(MAX_PREVIEW_ROWS), **FULL, hide_index=True)
    except ValueError as exc:
        st.error(str(exc))
    except Exception:
        show_error("Could not read the file.", traceback.format_exc())

ready = raw_df is not None and bool(api_key.strip())
if raw_df is not None and not api_key.strip():
    st.warning("Add your Gemini API key above to run the AI steps.")


def progress_cb(bar):
    return lambda done, total, msg: bar.progress(done / total if total else 0.0, text=msg)


# ------------------------------------------------------------------ mapping
st.subheader("2. Mapping & risk scoring")
if st.button("🧭 Run Mapping & Risk Scoring", type="primary", disabled=not ready):
    bar = st.progress(0.0, text="Starting...")
    try:
        engine = LedgerAIEngine(get_client(api_key.strip()), log, preferred_model(), st.session_state.w_ai_standard)
        log.info(f"Starting AI mapping for {len(raw_df)} ledgers (batch size {st.session_state.w_ai_batch})")
        st.session_state.ai_mapped = engine.map_all(raw_df, int(st.session_state.w_ai_batch), progress_cb(bar))
        st.session_state.ai_mapping_models = engine.models_used
    except (ApiError, AllModelsFailed, RuntimeError) as exc:
        log.error(str(exc))
        st.error(f"❌ Mapping failed: {exc}")
    except Exception:
        show_error("Mapping failed.", traceback.format_exc())
    bar.empty()

mapped = st.session_state.get("ai_mapped")
if mapped is not None:
    m = mapped
    risk = pd.to_numeric(m["Risk Score"], errors="coerce")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ledgers mapped", f"{(m['Account Mapping'] != 'UNMAPPED').sum():,} / {len(m):,}")
    c2.metric("High risk (4–5)", int((risk >= 4).sum()))
    c3.metric("Medium (3)", int((risk == 3).sum()))
    c4.metric("Model(s)", ", ".join(st.session_state.get("ai_mapping_models") or ["—"]))
    st.dataframe(risk_style(m), **FULL, hide_index=True)

# ------------------------------------------------------------------ analysis
st.subheader("3. AI analysis (optional)")
default("w_ai_prompt", DEFAULT_ANALYSIS_PROMPT)
default("w_ai_style", ANALYSIS_STYLES[0])
default("w_ai_automap", True)
st.text_area("Analysis prompt", key="w_ai_prompt", height=110)
a1, a2 = st.columns(2)
a1.selectbox("Answer style", ANALYSIS_STYLES, key="w_ai_style")
a2.checkbox("Run mapping first if it has not been run", key="w_ai_automap")
if st.button("🔎 Run AI Analysis", disabled=not ready or not st.session_state.w_ai_prompt.strip()):
    bar = st.progress(0.0, text="Starting...")
    try:
        client = get_client(api_key.strip())
        data = st.session_state.get("ai_mapped")
        if data is None and st.session_state.w_ai_automap:
            log.info("Mapping has not been run yet - running AI mapping before the analysis")
            mapper = LedgerAIEngine(client, log, preferred_model(), st.session_state.w_ai_standard)
            data = mapper.map_all(raw_df, int(st.session_state.w_ai_batch), progress_cb(bar))
            st.session_state.ai_mapped = data
            st.session_state.ai_mapping_models = mapper.models_used
        if data is None:
            data = raw_df
        bar.progress(1.0, text="Waiting for Gemini analysis...")
        engine = LedgerAIEngine(client, log, preferred_model(), st.session_state.w_ai_standard)
        log.info(f"Sending {len(data)} ledgers and the analysis prompt to Gemini ({st.session_state.w_ai_style})")
        text, model, fdf = engine.analyse(st.session_state.w_ai_prompt, data, st.session_state.w_ai_style)
        st.session_state.ai_analysis, st.session_state.ai_analysis_model = text, model
        st.session_state.ai_findings, st.session_state.ai_prompt_used = fdf, st.session_state.w_ai_prompt
        st.rerun()
    except (ApiError, AllModelsFailed, RuntimeError) as exc:
        log.error(str(exc))
        st.error(f"❌ Analysis failed: {exc}")
    except Exception:
        show_error("Analysis failed.", traceback.format_exc())
    bar.empty()

if st.session_state.get("ai_analysis"):
    st.caption(f"Answered by {st.session_state.ai_analysis_model}")
    findings = st.session_state.get("ai_findings")
    if findings is not None:
        # text is "OVERVIEW\n...\n\nSUMMARY\n...\n\n<n> ledger(s) with findings ..."
        blocks = st.session_state.ai_analysis.split("\n\n")
        with st.container(border=True):
            for block in blocks[:2]:
                head, _, body = block.partition("\n")
                st.markdown(f"**{head.title()}**  \n{body}")
        st.markdown(f"**Findings ({len(findings)})**")
        st.dataframe(risk_style(findings), **FULL, hide_index=True)
    else:
        with st.container(border=True):
            st.markdown(st.session_state.ai_analysis)

# ------------------------------------------------------------------ download
st.subheader("4. Download")
if mapped is None and not st.session_state.get("ai_analysis"):
    st.caption("Run the mapping and/or the analysis to enable the download.")
else:
    out = io.BytesIO()
    export_workbook(out, mapped, st.session_state.get("ai_analysis", ""), st.session_state.get("ai_prompt_used", ""),
                    st.session_state.get("ai_analysis_model", ""), st.session_state.get("ai_mapping_models"),
                    st.session_state.get("ai_findings"))
    stem = Path(up.name).stem if up is not None else "ledger"
    st.download_button("⬇️ Download processed workbook (.xlsx)", out.getvalue(),
                       f"{file_prefix()}{stem}_AI_processed.xlsx", XLSX_MIME, type="primary")

with st.expander("📜 Logs"):
    logs = st.session_state.get("ai_logs", [])
    st.code("\n".join(logs[-200:]) if logs else "No activity yet.", language=None)
