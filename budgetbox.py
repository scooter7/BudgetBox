import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image as RLImage
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import re

# ─── Register fonts ───────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))

# ─── Streamlit setup ──────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Split first line = Strategy, rest = Description ──────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    return lines[0], " ".join(lines[1:])

# ─── Extract tables & totals ─────────────────────────────────────────────────
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    proposal_title = next(
        (ln for pg in page_texts for ln in pg.splitlines() if "proposal" in ln.lower()),
        "Untitled Proposal"
    ).strip()
    tables_info = []
    used_totals = set()

    def find_total(pi):
        for ln in page_texts[pi].splitlines():
            if re.search(r'\btotal\b', ln, re.I) and re.search(r'\$\d', ln) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue
            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            # Build header
            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows = []
            for row in data[1:]:
                if all(cell is None or str(cell).strip()=="" for cell in row):
                    continue
                first = next((str(cell).strip() for cell in row if cell), "")
                if first.lower() == "total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc] + rest)

            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, tbl_total))

    # Grand total
    grand_total = None
    for tx in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1)
            break

# ─── Build PDF (unchanged) ────────────────────────────────────────────────────
# [ ... PDF-building code identical to before ... ]

# ─── Build Word ───────────────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(17)
sec.page_height = Inches(11)

# Centered logo + title
try:
    p_logo = docx.add_paragraph()
    r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo), width=Inches(2))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass
p_title = docx.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]
r.font.name = "DMSerif"
r.font.size = Pt(18)
docx.add_paragraph()

# Compute Word table column widths
# Description = 45% of 17" = 7.65", Others share 55% = 9.35" across (n-1) columns
TOTAL_WIDTH = 17.0
DESC_WIDTH = 0.45 * TOTAL_WIDTH  # 7.65"
OTHER_WIDTH = (0.55 * TOTAL_WIDTH)  # 9.35"
 
for hdr, rows, tbl_total in tables_info:
    n = len(hdr)
    other_count = n - 1
    # Build table
    tbl = docx.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Apply column widths
    for i, col in enumerate(tbl.columns):
        width_inch = DESC_WIDTH if i == 1 else (OTHER_WIDTH / other_count)
        col.width = Inches(width_inch)
    # Header row
    for i, col_name in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(str(col_name))
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Body rows
    for row_data in rows:
        rc = tbl.add_row().cells
        for i, val in enumerate(row_data):
            p = rc[i].paragraphs[0]
            p.text = ""
            run = p.add_run(str(val))
            run.font.name = "Barlow"
            run.font.size = Pt(9)
    # Table total row
    if tbl_total:
        label, amount = re.split(r'\$\s*', tbl_total, 1)
        amount = "$" + amount.strip()
        rc = tbl.add_row().cells
        for i, text_val in enumerate([label] + [""]*(n-2) + [amount]):
            p = rc[i].paragraphs[0]
            p.text = ""
            run = p.add_run(text_val)
            run.font.name = "DMSerif"
            run.font.size = Pt(10)
            run.bold = True
            if i == 0:
                p.alignment = WD_TABLE_ALIGNMENT.LEFT
            elif i == n-1:
                p.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else:
                p.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx.add_paragraph()

# Grand total row
if grand_total:
    n = len(tables_info[-1][0])
    tblg = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, text_val in enumerate(["Grand Total"] + [""]*(n-2) + [grand_total]):
        cell = tblg.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(text_val)
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        if i == 0:
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
        elif i == n-1:
            p.alignment = WD_TABLE_ALIGNMENT.RIGHT
        else:
            p.alignment = WD_TABLE_ALIGNMENT.CENTER

docx.save(docx_buf)
docx_buf.seek(0)

# Download buttons
c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "📥 Download deliverable PDF",
        data=pdf_buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True
    )
with c2:
    st.download_button(
        "📥 Download deliverable DOCX",
        data=docx_buf,
        file_name="proposal_deliverable.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
