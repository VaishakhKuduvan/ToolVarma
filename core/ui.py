"""
Shared presentation helpers so every page uses the same header, section and status patterns.
Colours come from the theme in .streamlit/config.toml; the CSS here only sets spacing and type,
so it works in both light and dark mode.
"""
import html

import streamlit as st

# Material icons (https://fonts.google.com/icons) used across the app
ICONS = {
    "home": ":material/home:",
    "risk": ":material/assignment:",
    "materiality": ":material/balance:",
    "sampling": ":material/fact_check:",
    "benford": ":material/query_stats:",
    "ai": ":material/auto_awesome:",
    "confirmations": ":material/mail:",
    "download": ":material/download:",
    "run": ":material/play_arrow:",
    "next": ":material/arrow_forward:",
    "back": ":material/arrow_back:",
    "done": ":material/check_circle:",
    "lock": ":material/lock:",
    "pending": ":material/radio_button_unchecked:",
}

# Translucent tints read correctly on both light and dark backgrounds
BAND_TINTS = {"High": "rgba(220, 38, 38, 0.16)", "Medium": "rgba(217, 119, 6, 0.16)",
              "Low": "rgba(22, 163, 74, 0.16)"}
BAND_BADGE = {"High": "red", "Medium": "orange", "Low": "green"}

_CSS = """
<style>
.block-container { padding-bottom: 4rem; max-width: 1180px; }
h1, h2, h3 { letter-spacing: -0.01em; }
.as-eyebrow { font-size: 0.75rem; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase;
              opacity: 0.65; margin-bottom: -0.6rem; }
.as-intro { font-size: 1.02rem; line-height: 1.6; max-width: 68ch; opacity: 0.82; margin: 0 0 0.5rem 0; }
.as-muted { opacity: 0.7; font-size: 0.9rem; line-height: 1.5; }
.as-card-title { font-weight: 600; font-size: 1.02rem; margin: 0.15rem 0 0.1rem 0; }
.as-step-num { font-size: 0.72rem; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase;
               opacity: 0.6; }
[data-testid="stMetricValue"] { font-variant-numeric: tabular-nums; font-weight: 600; }
[data-testid="stSidebarNav"] { padding-top: 0.5rem; }
@media (max-width: 640px) {
  .as-intro { font-size: 0.97rem; }
}
</style>
"""


def apply_global_styles():
    st.markdown(_CSS, unsafe_allow_html=True)


def page_header(title, eyebrow=None, intro=None):
    if eyebrow:
        st.markdown(f'<div class="as-eyebrow">{html.escape(eyebrow)}</div>', unsafe_allow_html=True)
    st.title(title, anchor=False)
    if intro:
        st.markdown(f'<p class="as-intro">{html.escape(intro)}</p>', unsafe_allow_html=True)


def section(title, caption=None):
    st.subheader(title, anchor=False)
    if caption:
        st.caption(caption)


def muted(text):
    st.markdown(f'<p class="as-muted">{html.escape(text)}</p>', unsafe_allow_html=True)


def card_title(text):
    st.markdown(f'<p class="as-card-title">{html.escape(text)}</p>', unsafe_allow_html=True)


def step_label(text):
    st.markdown(f'<div class="as-step-num">{html.escape(text)}</div>', unsafe_allow_html=True)


def band_label(pct):
    return "High" if pct >= 60 else ("Medium" if pct >= 40 else "Low")


def style_band_column(df, column):
    """Tint a High / Medium / Low text column so the band is readable by colour and by word."""
    return df.style.map(lambda v: f"background-color: {BAND_TINTS[v]}" if v in BAND_TINTS else "",
                        subset=[column])
