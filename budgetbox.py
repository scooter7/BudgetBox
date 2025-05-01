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

# ─── Simple split: first line → Strategy, rest → Description ─────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines: return "", ""
    return lines[0], " ".join(lines[1:])

# ─── Extract all tables and text once ─────────────────────────────────────────
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    # find proposal title
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
        img = page.to_image(resolution=150)
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2: continue
            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None: continue

            # build rows
            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i]
            rows = []
            for row in data[1:]:
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i in range(len(row)) if i!=desc_i]
                rows.append([strat, desc] + rest)
            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, tbl_total))

    # grand total
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
# logo + title
try:
    logo = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5
    ).content
    elements.append(RLImage(io.BytesIO(logo), width=150, height=50))
except:
    pass
elements += [Spacer(1,12), Paragraph(proposal_title, title_style), Spacer(1,24)]

# add each table
total_width = 17*inch - 96
for hdr, rows, tbl_total in tables_info:
    wrapped = [[Paragraph(str(h), header_style) for h in hdr]]
    for r in rows:
        wrapped.append([Paragraph(str(c), body_style) for c in r])
    widths = [0.45*total_width if i==1 else (0.55*total_width)/(len(hdr)-1) for i in range(len(hdr))]
    t = LongTable(wrapped, colWidths=widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]))
    elements += [t, Spacer(1,12)]
    if tbl_total:
        lbl, val = re.split(r'\$\s*', tbl_total,1)
        val = "$"+val.strip()
        row = [Paragraph(lbl, bl_style)] + [""]*(len(hdr)-2) + [Paragraph(val, br_style)]
        tt = LongTable([row], colWidths=widths)
        tt.setStyle(TableStyle([
            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
            ("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        elements += [tt, Spacer(1,24)]

# grand total
if grand_total:
    row = [Paragraph("Grand Total", bl_style)] + [""]*(len(tables_info[-1][0])-2) + [Paragraph(grand_total, br_style)]
    gt = LongTable([row], colWidths=widths)
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
# set 11x17 landscape
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(17)
sec.page_height = Inches(11)
# logo + title
try:
    logo_data = logo  # already fetched
    pic = docx.add_picture(io.BytesIO(logo_data), width=Inches(2))
    pic.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass
heading = docx.add_paragraph(proposal_title)
heading.alignment = WD_TABLE_ALIGNMENT.CENTER
heading.runs[0].font.name = "DMSerif"
heading.runs[0].font.size = Pt(18)
docx.add_paragraph()

# add tables
for hdr, rows, tbl_total in tables_info:
    tbl = docx.add_table(rows=1, cols=len(hdr))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    # header row
    for i, h in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        p = cell.paragraphs[0]
        run = p.add_run(str(h))
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    # data rows
    for r in rows:
        row_cells = tbl.add_row().cells
        for i, c in enumerate(r):
            p = row_cells[i].paragraphs[0]
            run = p.add_run(str(c))
            run.font.name = "Barlow"
            run.font.size = Pt(9)
    # table total
    if tbl_total:
        lbl, val = re.split(r'\$\s*', tbl_total,1)
        val = "$"+val.strip()
        tr = tbl.add_row().cells
        tr[0].text = lbl
        tr[-1].text = val
        for i in range(len(hdr)):
            p = tr[i].paragraphs[0]
            run = p.runs[0]
            run.font.name = "DMSerif"
            run.font.size = Pt(10)
            run.bold = True
            if i==0: p.alignment = WD_TABLE_ALIGNMENT.LEFT
            elif i==len(hdr)-1: p.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else: p.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx.add_paragraph()

# grand total
if grand_total:
    tbl = docx.add_table(rows=1, cols=len(tables_info[-1][0]))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cells = tbl.rows[0].cells
    cells[0].text = "Grand Total"
    cells[-1].text = grand_total
    for i in [0, len(cells)-1]:
        p = cells[i].paragraphs[0]
        run = p.runs[0]
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
    docx.add_paragraph()

docx.save(docx_buf)
docx_buf.seek(0)

# ─── Streamlit Download Buttons ───────────────────────────────────────────────
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
