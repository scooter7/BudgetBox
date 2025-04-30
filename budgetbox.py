import io
import os
import re
import streamlit as st
import pdfplumber
import requests
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a styled Word document for manual edits.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract title and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = "Untitled Proposal"
    all_lines = []
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.splitlines()
            all_lines.extend(lines)
            for line in lines:
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break
        if "proposal" in proposal_title.lower():
            break

    page_texts = [page.extract_text() for page in pdf.pages]
    all_tables = [(i, t) for i, page in enumerate(pdf.pages) for t in (page.extract_tables() or [])]

# Total rows per page
page_totals = {}
used_lines = set()
for idx, text in enumerate(page_texts):
    if text:
        lines = text.splitlines()
        page_totals[idx] = [(i, line.strip()) for i, line in enumerate(lines)
                             if re.search(r'\btotal\b', line, re.I) and re.search(r'\$[0-9,]+\.\d{2}', line)]

def get_closest_total(page_idx):
    for i, line in page_totals.get(page_idx, []):
        if line not in used_lines:
            used_lines.add(line)
            return line
    return None

# Word document setup
doc = Document()
section = doc.sections[0]
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = Inches(17)
section.page_height = Inches(11)
doc.add_heading(proposal_title, 0)

for page_idx, raw in all_tables:
    if len(raw) < 2:
        continue
    header = [str(c).strip() if c else "" for c in raw[0]]
    rows = raw[1:]

    desc_idx = next((i for i, h in enumerate(header) if "description" in h.lower()), None)
    if desc_idx is not None:
        new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
        new_rows = []
        for row in rows:
            desc_raw = row[desc_idx] or ""
            desc_lines = str(desc_raw).split("\n")
            strategy = desc_lines[0].strip() if desc_lines else ""
            description = "\n".join(desc_lines[1:]).strip() if len(desc_lines) > 1 else ""
            rest = [row[i] for i in range(len(row)) if i != desc_idx]
            new_rows.append([strategy, description] + rest)
        header = new_header
        rows = new_rows

    table = doc.add_table(rows=1, cols=len(header))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(header):
        hdr_cells[i].text = col
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.font.bold = True
        run.font.size = Pt(10)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for row in rows:
        row_cells = table.add_row().cells
        for i, cell in enumerate(row):
            p = row_cells[i].paragraphs[0]
            p.text = str(cell)
            p.runs[0].font.size = Pt(10)
            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    total_line = get_closest_total(page_idx)
    if total_line:
        doc.add_paragraph(total_line, style='Intense Quote')

# Grand total
grand_total = None
for text in reversed(page_texts):
    if text:
        m = re.search(r'Grand Total.*?\$[0-9,]+\.\d{2}', text, re.I | re.DOTALL)
        if m:
            match = re.search(r'\$[0-9,]+\.\d{2}', m.group(0))
            if match:
                grand_total = match.group(0)
                break

if grand_total:
    doc.add_page_break()
    doc.add_heading("Grand Total", level=1)
    doc.add_paragraph(f"Total {grand_total}")

# Export Word document
buf = io.BytesIO()
doc.save(buf)
buf.seek(0)

st.success("✔️ Transformation complete!")
st.download_button(
    "📄 Download deliverable Word doc",
    data=buf,
    file_name="proposal_deliverable.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True,
)
