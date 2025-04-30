import io
import os
import re
import streamlit as st
import pdfplumber
import requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENTATION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a styled Word document.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract title and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = "Untitled Proposal"
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.splitlines()
            for line in lines:
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break
        if "proposal" in proposal_title.lower():
            break

    page_texts = [page.extract_text() for page in pdf.pages]
    all_tables = [(i, t) for i, page in enumerate(pdf.pages) for t in (page.extract_tables() or [])]

# Total line helper
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

# Word doc
doc = Document()
section = doc.sections[0]
section.orientation = WD_ORIENTATION.LANDSCAPE
section.page_width = Inches(17)
section.page_height = Inches(11)

# Centered logo
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    logo_resp = requests.get(logo_url, timeout=5)
    logo_path = "/tmp/carnegie_logo.png"
    with open(logo_path, "wb") as f:
        f.write(logo_resp.content)
    para = doc.add_paragraph()
    run = para.add_run()
    run.add_picture(logo_path, width=Inches(2.5))
    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
except:
    st.warning("Logo couldn't be loaded.")

# Title below logo
title_para = doc.add_paragraph()
run = title_para.add_run(proposal_title)
run.bold = True
run.font.size = Pt(18)
run.font.name = "DM Serif Display"
r = run._element
r.rPr.rFonts.set(qn('w:eastAsia'), 'DM Serif Display')
title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

doc.add_paragraph()

# Center these columns if present
center_cols = ["term", "start date", "end date", "monthly amount", "item total"]

def add_total_row(table, label, value):
    row_cells = table.add_row().cells
    row_cells[0].text = label
    row_cells[1].text = value
    for i, cell in enumerate(row_cells):
        p = cell.paragraphs[0]
        run = p.runs[0]
        run.font.size = Pt(10)
        run.font.name = "DM Serif Display"
        run.font.color.rgb = RGBColor(0, 0, 0)
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), 'DM Serif Display')
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT if i == 0 else WD_PARAGRAPH_ALIGNMENT.RIGHT
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Tables
for page_idx, raw in all_tables:
    if len(raw) < 2:
        continue
    header = [str(c).strip() if c else "" for c in raw[0]]
    rows = raw[1:]

    non_none = [i for i, h in enumerate(header) if h.lower() != "none" and h.strip()]
    header = [header[i] for i in non_none]
    rows = [[row[i] if i < len(row) else "" for i in non_none] for row in rows]

    desc_idx = next((i for i, h in enumerate(header) if "description" in h.lower()), None)
    if desc_idx is not None:
        new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
        new_rows = []
        for row in rows:
            desc_raw = row[desc_idx] or ""
            desc_lines = str(desc_raw).split("\n")
            strategy = desc_lines[0].strip()
            description = "\n".join(desc_lines[1:]).strip() if len(desc_lines) > 1 else ""
            rest = [row[i] for i in range(len(row)) if i != desc_idx]
            new_rows.append([strategy, description] + rest)
        header = new_header
        rows = new_rows

    table = doc.add_table(rows=1, cols=len(header), style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    hdr_cells = table.rows[0].cells
    for i, col in enumerate(header):
        hdr_cells[i].text = col
        run = hdr_cells[i].paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = "DM Serif Display"
        r = run._element
        r.rPr.rFonts.set(qn('w:eastAsia'), 'DM Serif Display')
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        # Light gray
        tc_pr = hdr_cells[i]._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'D9D9D9')
        tc_pr.append(shd)

    for row in rows:
        row_cells = table.add_row().cells
        for i, cell in enumerate(row):
            p = row_cells[i].paragraphs[0]
            p.text = str(cell)
            run = p.runs[0]
            run.font.size = Pt(10)
            run.font.name = "Barlow"
            r = run._element
            r.rPr.rFonts.set(qn('w:eastAsia'), 'Barlow')

            col_name = header[i].strip().lower()
            if col_name in center_cols:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            else:
                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

            row_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # Total row
    total_line = get_closest_total(page_idx)
    if total_line and re.search(r"\$[0-9,]+\.\d{2}", total_line):
        label = total_line.split("$")[0].strip()
        amount = "$" + total_line.split("$")[1].strip()
        add_total_row(doc.add_table(rows=0, cols=2, style="Table Grid"), label, amount)

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
    doc.add_paragraph()
    gt_table = doc.add_table(rows=0, cols=2, style="Table Grid")
    add_total_row(gt_table, "Grand Total", grand_total)

# Save to buffer
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
