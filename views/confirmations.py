"""External Confirmation Generator — bulk letters from a .docx ([Tag]) or fillable .pdf template + Excel list."""
import traceback

import pandas as pd
import streamlit as st

from core.confirmations import (FILE_NAME_COL, extract_placeholders, generate_zip, read_parties, template_type,
                                validate)
from core.state import FULL, XLSX_MIME, file_prefix, show_error, to_excel_bytes
from core.ui import ICONS, card_title, page_header, section

page_header("External confirmations", eyebrow="SA 505",
            intro="Create one confirmation letter per party from a Word template with [Tag] placeholders "
                  "or a fillable PDF, and an Excel list of parties. Letters download together as a ZIP.")

with st.expander("How to prepare the files", icon=":material/help:"):
    st.markdown(f"""
- **Word template (.docx):** type each placeholder in square brackets, e.g. `[Party_Name]`, `[Balance]`.
  Use letters, digits and underscores only, and type each tag in one go so Word keeps it in a single run.
- **PDF template (.pdf):** a fillable PDF; each form field name is a placeholder.
- **Excel data (.xlsx):** one column per placeholder (header = tag name, without brackets) plus a
  **`{FILE_NAME_COL}`** column giving each output file's name. One row = one letter.
""")

section("1. Upload files")
c1, c2 = st.columns(2)
with c1.container(border=True):
    card_title("Template (.docx or fillable .pdf)")
    tpl = st.file_uploader("Template", type=["docx", "pdf"], key="cf_tpl", label_visibility="collapsed")
with c2.container(border=True):
    card_title("Party list (.xlsx or .xls)")
    xls = st.file_uploader("Excel data", type=["xlsx", "xls"], key="cf_xls", label_visibility="collapsed")

placeholders, kind, df = [], None, None
if tpl is not None:
    kind = template_type(tpl.name)
    try:
        placeholders = extract_placeholders(tpl.getvalue(), kind)
        c1.caption(f"{tpl.name}: {len(placeholders)} placeholders found")
        if placeholders:
            blank = pd.DataFrame(columns=placeholders + [FILE_NAME_COL])
            c1.download_button("Blank Excel with these columns", to_excel_bytes({"Parties": blank}),
                               "Confirmation_parties_template.xlsx", XLSX_MIME,
                               icon=ICONS["download"])
    except Exception:
        show_error("Could not read the template.", traceback.format_exc())
        tpl = None
if xls is not None:
    try:
        df = read_parties(xls.getvalue())
        c2.caption(f"{xls.name}: {len(df):,} rows")
    except Exception:
        show_error("Could not read the Excel file.", traceback.format_exc())

section("2. Check the mapping")
valid = False
if tpl is None or df is None:
    st.caption("Upload both files to check that every placeholder has a matching Excel column.")
else:
    valid, facts, problems, warnings = validate(placeholders, df, kind)
    with st.container(border=True):
        for label, value in facts:
            st.markdown(f"**{label}:** {value}")
    for p in problems:
        st.error(p)
    for w in warnings:
        st.warning(w)
    if valid:
        st.success("Every placeholder has a matching column. Ready to generate.")
    with st.expander("Preview party data"):
        st.dataframe(df, **FULL, hide_index=True)

section("3. Generate letters")
if st.button("Generate letters", type="primary", icon=ICONS["run"], disabled=not valid):
    bar = st.progress(0.0, text="Starting...")
    try:
        data, ok, errors = generate_zip(tpl.getvalue(), kind, df, placeholders,
                                        lambda f, t: bar.progress(f, text=t))
        st.session_state.cf_result = {"zip": data, "ok": ok, "total": len(df), "errors": errors, "kind": kind}
    except Exception:
        show_error("Generation failed.", traceback.format_exc())
    bar.empty()

res = st.session_state.get("cf_result")
if res:
    if res["ok"] == res["total"]:
        st.success(f"Generated {res['ok']} of {res['total']} {res['kind'].upper()} letters.")
    else:
        st.warning(f"Generated {res['ok']} of {res['total']} letters. These rows failed:")
        st.code("\n".join(res["errors"]))
    if res["ok"]:
        st.download_button("Download all letters (.zip)", res["zip"], f"{file_prefix()}External_Confirmations.zip",
                           "application/zip", type="primary", icon=ICONS["download"])
