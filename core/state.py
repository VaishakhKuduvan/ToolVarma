"""
Shared session state for the Audit Suite. Values here flow between pages:
Risk (SA 315) -> Materiality (SA 320) -> Sampling (SA 530) and Benford's Law.
"""
import inspect
import io

import pandas as pd
import streamlit as st

# Compatibility: newer Streamlit uses width='stretch', older uses use_container_width
FULL = {"width": "stretch"} if "width" in inspect.signature(st.dataframe).parameters else {"use_container_width": True}
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
LAKH = 100000

DEFAULTS = {
    "risk_level": None, "risk_score_pct": None, "risk_breakdown": None,
    "overall_materiality": None, "performance_materiality": None,
    "specific_materiality": None, "clearly_trivial": None,
    "benchmark_choice": None, "benchmark_reason": "", "benchmark_pct": None,
    "benchmark_value": None, "pm_pct": None,
    # last Day Book uploaded on any page, so Sampling and Benford can share it
    "daybook_bytes": None, "daybook_name": None,
}

# Widget keys starting with this prefix keep their values when the user switches pages.
PERSIST_PREFIX = "w_"


def init_state():
    for k, v in DEFAULTS.items():
        if k not in st.session_state:
            st.session_state[k] = v
    # Streamlit drops the state of widgets that are not drawn on the current page;
    # re-assigning the keys on every run keeps questionnaire answers across pages.
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith(PERSIST_PREFIX):
            st.session_state[k] = st.session_state[k]


def default(key, value):
    """Seed a persistent widget's value once (the widget is then created without value=/index=)."""
    if key not in st.session_state:
        st.session_state[key] = value


def materiality_ready():
    return st.session_state.risk_level is not None and st.session_state.performance_materiality is not None


def materiality_rupees():
    """OM/PM/SM/CTT from the Materiality page converted from ₹ Lakhs to ₹, or None if not computed."""
    if st.session_state.overall_materiality is None:
        return None
    return {"OM": st.session_state.overall_materiality * LAKH,
            "PM": st.session_state.performance_materiality * LAKH,
            "SM": st.session_state.specific_materiality * LAKH,
            "CTT": st.session_state.clearly_trivial * LAKH}


def file_prefix():
    """'<Client>_<FY>_' from the sidebar, safe for file names (empty if not filled)."""
    parts = [st.session_state.get("w_client", "").strip(), st.session_state.get("w_fy", "").strip()]
    stem = "_".join(p for p in parts if p)
    stem = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in stem)
    return f"{stem}_" if stem else ""


def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return buf.getvalue()


def show_error(message, exc_text):
    st.error(message)
    with st.expander("Technical details"):
        st.code(exc_text)


def daybook_source(uploader_key, label="Upload Day Book (.xlsx/.xls)"):
    """File uploader that also offers the Day Book already uploaded on another page.
    Returns (bytes, name) or (None, None)."""
    up = st.file_uploader(label, type=["xlsx", "xls"], key=uploader_key)
    if up is not None:
        data = up.getvalue()
        st.session_state.daybook_bytes, st.session_state.daybook_name = data, up.name
        return data, up.name
    if st.session_state.daybook_bytes is not None:
        if st.checkbox(f"Use the Day Book already uploaded: **{st.session_state.daybook_name}**",
                       value=True, key=f"{uploader_key}_reuse"):
            return st.session_state.daybook_bytes, st.session_state.daybook_name
    return None, None


def sidebar_status():
    with st.sidebar:
        st.markdown("### 🗂️ Engagement")
        default("w_client", "")
        default("w_fy", "")
        st.text_input("Client name", key="w_client")
        st.text_input("Financial year", key="w_fy", placeholder="e.g. FY 2025-26")
        st.caption("Used to prefix downloaded file names.")
        st.markdown("### 📌 Status")
        if st.session_state.risk_level:
            icon = {"High": "🔴", "Medium": "🟡", "Low": "🟢"}[st.session_state.risk_level]
            st.markdown(f"**Risk:** {icon} {st.session_state.risk_level} ({st.session_state.risk_score_pct:.1f}%)")
        else:
            st.markdown("**Risk:** not assessed")
        if st.session_state.overall_materiality is not None:
            st.markdown(f"**OM:** ₹ {st.session_state.overall_materiality:,.2f} L  \n"
                        f"**PM:** ₹ {st.session_state.performance_materiality:,.2f} L")
        else:
            st.markdown("**Materiality:** not computed")
        if st.session_state.daybook_name:
            st.markdown(f"**Day Book:** {st.session_state.daybook_name}")
