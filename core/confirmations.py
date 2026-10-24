"""
External Confirmation engine — mail-merge bulk generation from a .docx ([Tag] placeholders)
or a fillable .pdf (AcroForm fields) and an Excel sheet with one row per party.
Moved from externalconfirmation.py; output is built in memory and returned as a ZIP.
"""
import io
import re
import zipfile

import jinja2
import pandas as pd
from docxtpl import DocxTemplate
from pypdf import PdfReader, PdfWriter

FILE_NAME_COL = "File_Name"
ILLEGAL_CHARS = r'[\\/*?:"<>|]'


def bracket_env():
    """Jinja2 environment that treats single square brackets as [variable] tags."""
    return jinja2.Environment(variable_start_string="[", variable_end_string="]")


def template_type(filename):
    return filename.rsplit(".", 1)[-1].lower()


def extract_placeholders(data: bytes, kind: str):
    """Return the list of placeholder / form-field names found in the template."""
    if kind == "docx":
        doc = DocxTemplate(io.BytesIO(data))
        try:
            found = list(doc.get_undeclared_template_variables(jinja_env=bracket_env()))
        except jinja2.TemplateSyntaxError:
            found = []
        if found:
            return sorted(found)
        # Fallback: regex scan of body paragraphs and table cells for [Variable]
        doc.init_docx()
        full_text = "\n".join(p.text for p in doc.docx.paragraphs)
        for table in doc.docx.tables:
            for row in table.rows:
                for cell in row.cells:
                    full_text += " " + cell.text
        return sorted({t.strip() for t in re.findall(r"\[\s*(\w+)\s*\]", full_text)})
    fields = PdfReader(io.BytesIO(data)).get_fields()
    return list(fields.keys()) if fields else []


def read_parties(data: bytes):
    return pd.read_excel(io.BytesIO(data)).fillna("")


def validate(placeholders, df, kind):
    """Returns (is_valid, facts, problems, warnings): facts describe the files, problems block generation."""
    headers = [str(h).strip() for h in df.columns]
    missing = [p for p in placeholders if p.strip() not in headers]
    facts = [("Format", kind.upper()), ("Rows (letters)", f"{len(df):,}"),
             ("Placeholders in template", f"{len(placeholders)}")]
    if placeholders:
        facts.append(("Tags found", ", ".join(placeholders)))
    problems, warnings = [], []
    if FILE_NAME_COL not in headers:
        problems.append(f"The Excel file has no '{FILE_NAME_COL}' column.")
    if missing:
        problems.append(f"The Excel file has no column for: {', '.join(missing)}")
    if not placeholders:
        warnings.append("No placeholders were found. Write tags as [Variable_Name].")
    return not problems, facts, problems, warnings


def safe_filename(value, index):
    name = re.sub(ILLEGAL_CHARS, "", str(value)).strip()
    return name or f"Document_{index + 1}"


def render_docx(template: bytes, data: dict) -> bytes:
    doc = DocxTemplate(io.BytesIO(template))
    doc.render(data, jinja_env=bracket_env())
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


def fill_pdf(template: bytes, data: dict, placeholders) -> bytes:
    writer = PdfWriter(clone_from=PdfReader(io.BytesIO(template)))
    clean = {k: str(v) for k, v in data.items() if k in placeholders}
    for page in writer.pages:
        writer.update_page_form_field_values(page, clean, auto_regenerate=False)
    writer.set_need_appearances_writer(True)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def generate_zip(template: bytes, kind: str, df, placeholders, progress_cb=None):
    """Render one document per row. Returns (zip_bytes, success_count, errors)."""
    buf = io.BytesIO()
    errors, ok, used = [], 0, set()
    total = len(df)
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, (_, row) in enumerate(df.iterrows()):
            data = {str(k).strip(): v for k, v in row.to_dict().items()}
            name = safe_filename(data.get(FILE_NAME_COL, ""), i)
            stem, n = name, 2
            while name.lower() in used:  # keep duplicate File_Name rows from overwriting each other
                name, n = f"{stem}_{n}", n + 1
            used.add(name.lower())
            try:
                if kind == "docx":
                    zf.writestr(f"{name}.docx", render_docx(template, data))
                else:
                    zf.writestr(f"{name}.pdf", fill_pdf(template, data, placeholders))
                ok += 1
            except Exception as exc:
                errors.append(f"Row {i + 2} ({name}): {exc}")
            if progress_cb:
                progress_cb((i + 1) / total, f"Generated {i + 1} of {total}")
    return buf.getvalue(), ok, errors
