import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

# Word‐hyperlink helper
def add_hyperlink(paragraph, url, text, font_name="Barlow", font_size=9, bold=False, align=None):
    part = paragraph.part
    rid = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), "0000FF"); rPr.append(c)
    u = OxmlElement("w:u");     u.set(qn("w:val"), "single"); rPr.append(u)
    r.append(rPr)
    t = OxmlElement("w:t"); t.text = text
    r.append(t)
    hyperlink.append(r)
    paragraph._p.append(hyperlink)
    run = paragraph.add_run()  # apply font after
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    if align is not None:
        paragraph.alignment = align
    return paragraph

# Extract tables & totals
tables_info = []
grand_total = None

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    proposal_title = next(
        (ln for pg in page_texts for ln in pg.splitlines() if "proposal" in ln.lower()),
        "Untitled Proposal"
    ).strip()

    used_totals = set()
    def find_total(pi):
        for ln in page_texts[pi].splitlines():
            if re.search(r'\btotal\b', ln, re.I) and re.search(r'\$\d', ln) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        links = page.hyperlinks
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue
            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            # build link map
            cell_links = {}
            for row_idx, row_cells in enumerate(tbl.cells):
                for col_idx, cell in enumerate(row_cells):
                    x0, top, x1, bottom = cell.bbox
                    for link in links:
                        if (link["x0"] >= x0 and link["x1"] <= x1
                            and link["top"] >= top and link["bottom"] <= bottom):
                            cell_links[(row_idx, col_idx)] = link["uri"]
                            break

            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows, hmap = [], []
            for ridx, row in enumerate(data[1:], start=1):
                if all(cell is None or not str(cell).strip() for cell in row):
                    continue
                first = next((str(cell).strip() for cell in row if cell), "")
                if first.lower() == "total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc] + rest)
                # build hyperlink map for this row
                hm = []
                for j in range(len(new_hdr)):
                    if j < 2:
                        uri = cell_links.get((ridx, desc_i))
                    else:
                        orig_j = [i for i,h in enumerate(hdr) if i!=desc_i][j-2]
                        uri = cell_links.get((ridx, orig_j))
                    hm.append(uri)
                hmap.append(hm)

            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, hmap, tbl_total))

    # grand total
    for tx in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1)
            break

# Build PDF
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11*inch,17*inch)),
    leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=36
)
title_st  = ParagraphStyle("Title",  fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
hdr_st    = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_st   = ParagraphStyle("Body",   fontName="Barlow",  fontSize=9,  alignment=TA_LEFT)
tL_st     = ParagraphStyle("TL",     fontName="DMSerif", fontSize=10, alignment=TA_LEFT)
tR_st     = ParagraphStyle("TR",     fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)

elements = []
try:
    logo = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5
    ).content
    elements.append(RLImage(io.BytesIO(logo), width=150, height=50))
except:
    pass
elements += [Spacer(1,12), Paragraph(proposal_title, title_st), Spacer(1,24)]

total_w = 17*inch - 96
for hdr, rows, hmap, tbl_total in tables_info:
    wrapped = [[Paragraph(h, hdr_st) for h in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            uri = hmap[ridx][cidx]
            txt = str(cell)
            if uri:
                p = Paragraph(f'<a href="{uri}">{txt}</a>', body_st)
            else:
                p = Paragraph(txt, body_st)
            line.append(p)
        wrapped.append(line)
    if tbl_total:
        lbl, val = re.split(r'\$\s*', tbl_total, 1)
        val = "$" + val.strip()
        tr = []
        for i in range(len(hdr)):
            if i == 0:
                tr.append(Paragraph(lbl, tL_st))
            elif i == len(hdr)-1:
                tr.append(Paragraph(val, tR_st))
            else:
                tr.append(Paragraph("", body_st))
        wrapped.append(tr)

    colws = [0.45*total_w if i==1 else (0.55*total_w)/(len(hdr)-1) for i in range(len(hdr))]
    tbl = LongTable(wrapped, colWidths=colws, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]))
    elements += [tbl, Spacer(1,24)]

if grand_total:
    hdr = tables_info[-1][0]
    colws = [0.45*total_w if i==1 else (0.55*total_w)/(len(hdr)-1) for i in range(len(hdr))]
    gr = [Paragraph("Grand Total", tL_st)] + [""]*(len(hdr)-2) + [Paragraph(grand_total, tR_st)]
    gt = LongTable([gr], colWidths=colws)
    gt.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# Build Word
docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(17)
sec.page_height = Inches(11)

try:
    p_logo = docx.add_paragraph()
    r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo), width=Inches(2))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass
p_title = docx.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]; r.font.name="DMSerif"; r.font.size=Pt(18)
docx.add_paragraph()

TOTAL_W = 17.0
for hdr, rows, hmap, tbl_total in tables_info:
    n = len(hdr)
    desc_w = 0.45 * TOTAL_W
    oth_w  = (TOTAL_W - desc_w) / (n - 1)

    tbl = docx.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, col in enumerate(tbl.columns):
        col.width = Inches(desc_w if idx==1 else oth_w)

    # header shading + text
    for i, col_name in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(str(col_name))
        run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    # body
    for ridx, row in enumerate(rows):
        rc = tbl.add_row().cells
        for cidx, val in enumerate(row):
            uri = hmap[ridx][cidx]
            p = rc[cidx].paragraphs[0]; p.text = ""
            if uri:
                add_hyperlink(p, uri, str(val), font_name="Barlow", font_size=9)
            else:
                run = p.add_run(str(val)); run.font.name="Barlow"; run.font.size=Pt(9)

    # total row
    if tbl_total:
        label, amount = re.split(r'\$\s*', tbl_total, 1)
        amount = "$" + amount.strip()
        rc = tbl.add_row().cells
        for i, text_val in enumerate([label] + [""]*(n-2) + [amount]):
            p = rc[i].paragraphs[0]; p.text = ""
            run = p.add_run(text_val); run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
            if i==0: p.alignment=WD_TABLE_ALIGNMENT.LEFT
            elif i==n-1: p.alignment=WD_TABLE_ALIGNMENT.RIGHT
            else: p.alignment=WD_TABLE_ALIGNMENT.CENTER

    docx.add_paragraph()

# grand total row
if grand_total:
    hdr = tables_info[-1][0]; n=len(hdr)
    tblg = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, text_val in enumerate(["Grand Total"] + [""]*(n-2) + [grand_total]):
        cell = tblg.rows[0].cells[idx]
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text = ""
        run = p.add_run(text_val); run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
        if idx==0: p.alignment=WD_TABLE_ALIGNMENT.LEFT
        elif idx==n-1: p.alignment=WD_TABLE_ALIGNMENT.RIGHT
        else: p.alignment=WD_TABLE_ALIGNMENT.CENTER

docx.save(docx_buf)
docx_buf.seek(0)

c1,c2 = st.columns(2)
with c1:
    st.download_button("📥 Download PDF", data=pdf_buf,    file_name="proposal.pdf", mime="application/pdf")
with c2:
    st.download_button("📥 Download DOCX",data=docx_buf, file_name="proposal.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
