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

# Register fonts
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# Split first line = Strategy, rest = Description
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    return lines[0], " ".join(lines[1:])

# Extract tables & totals
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

            # Build header and rows
            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows = []
            for row in data[1:]:
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

# ─── Build PDF ────────────────────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11*inch,17*inch)),
    leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=36
)
title_style  = ParagraphStyle("Title",  fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body",   fontName="Barlow",  fontSize=9,  alignment=TA_LEFT)
bl_style     = ParagraphStyle("BL",     fontName="DMSerif", fontSize=10, alignment=TA_LEFT)
br_style     = ParagraphStyle("BR",     fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)

elements = []
# Logo + Title
try:
    logo = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5
    ).content
    elements.append(RLImage(io.BytesIO(logo), width=150, height=50))
except:
    pass
elements += [Spacer(1,12), Paragraph(proposal_title, title_style), Spacer(1,24)]

total_w = 17*inch - 96
for hdr, rows, tbl_total in tables_info:
    wrapped = [[Paragraph(str(h), header_style) for h in hdr]]
    for r in rows:
        wrapped.append([Paragraph(str(c), body_style) for c in r])
    col_widths = [0.45*total_w if i==1 else (0.55*total_w)/(len(hdr)-1) for i in range(len(hdr))]
    tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]))
    elements += [tbl, Spacer(1,12)]

    if tbl_total:
        lbl, val = re.split(r'\$\s*', tbl_total, 1)
        val = "$" + val.strip()
        cells = [lbl] + [""]*(len(hdr)-2) + [val]
        wrapped = [[Paragraph(str(c), bl_style) for c in cells]]
        tt = LongTable(wrapped, colWidths=col_widths)
        tt.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        elements += [tt, Spacer(1,24)]

if grand_total:
    cells = ["Grand Total"] + [""]*(len(tables_info[-1][0])-2) + [grand_total]
    wrapped = [[Paragraph(str(c), bl_style) for c in cells]]
    gt = LongTable(wrapped, colWidths=col_widths)
    gt.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

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

# Word tables with full borders
for hdr, rows, tbl_total in tables_info:
    tbl = docx.add_table(rows=1, cols=len(hdr), style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # header
    for i, col in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(str(col))
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    # body
    for rdata in rows:
        row_cells = tbl.add_row().cells
        for i, val in enumerate(rdata):
            p = row_cells[i].paragraphs[0]
            p.text = ""
            run = p.add_run(str(val))
            run.font.name = "Barlow"
            run.font.size = Pt(9)
    # table total
    if tbl_total:
        values = re.split(r'\$\s*', tbl_total, 1)
        label = values[0]
        amount = "$" + values[1].strip()
        cells = tbl.add_row().cells
        for i, text_val in enumerate([label] + [""]*(len(hdr)-2) + [amount]):
            p = cells[i].paragraphs[0]
            p.text = ""
            run = p.add_run(text_val)
            run.font.name = "DMSerif"
            run.font.size = Pt(10)
            run.bold = True
            if i == 0:
                p.alignment = WD_TABLE_ALIGNMENT.LEFT
            elif i == len(hdr)-1:
                p.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else:
                p.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx.add_paragraph()

# grand total row
if grand_total:
    hdr_len = len(tables_info[-1][0])
    tblg = docx.add_table(rows=1, cols=hdr_len, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    cells = tblg.rows[0].cells
    values = ["Grand Total"] + [""]*(hdr_len-2) + [grand_total]
    for i, text_val in enumerate(values):
        p = cells[i].paragraphs[0]
        p.text = ""
        run = p.add_run(text_val)
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        if i == 0:
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
        elif i == hdr_len-1:
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
