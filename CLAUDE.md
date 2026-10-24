# Audit Suite — project instructions

Streamlit app (`app.py` + `views/` pages + `core/` engines) for Indian audit practitioners,
deployed on Streamlit Community Cloud at https://varma-audit-suite.streamlit.app/ from
GitHub `VaishakhKuduvan/ToolVarma` (branch `main`; every push redeploys).

## Design and UI/UX standards (mandatory for this tool only)

These design and UI/UX instructions are mandatory whenever Claude Code is used on this tool/project.

- Follow them for every website, webpage, component, or frontend feature created or redesigned in this project.
- Treat visual quality, human-centered design, smooth interactions, responsiveness, accessibility and overall
  user experience as **mandatory acceptance criteria**, not optional suggestions.
- Do not apply these project-specific instructions automatically to unrelated tools, projects, repositories or workflows.
  When working on a different tool or project, follow that project's own instructions instead.
- These rules do not override higher-priority instructions, security requirements, or the project's existing
  technical constraints.
- If an existing project instruction conflicts with these requirements, identify the conflict and resolve it
  without silently removing existing instructions.

- Use emojis sparingly. Don't put them in front of every title, tab, button, metric and message; prefer clean
  typography, and where a visual cue is needed, a consistent icon set (e.g. Streamlit's `:material/...:` icons).

**Bottom line:** these standards are compulsory for this tool. The app must look intentionally designed by a
skilled human designer, feel polished and natural, and work smoothly — not like a generic AI-generated template.

## Technical constraints to respect when designing

- The UI is Streamlit: styling is limited to Streamlit's theme (`.streamlit/config.toml`), layout primitives and
  scoped CSS injected with `st.markdown(..., unsafe_allow_html=True)`. Keep custom CSS in one shared place.
- Pages must work on phone widths and in both light and dark themes.
- Never put client data (`.xlsx`, `.pdf`, `.docx`, etc.) in the repo; `.gitignore` already blocks these.
- On the author's Windows PC, Smart App Control blocks new native wheels: keep pandas 3.0.5 / numpy 2.5.2 /
  pyarrow 25.0.1 in the local `.venv` and do not add SciPy.
