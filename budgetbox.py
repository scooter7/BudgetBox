# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
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
import html
import fitz

# Register fonts
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except:
    DEFAULT_SERIF_FONT = "Times New Roman"
    DEFAULT_SANS_FONT = "Arial"

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# Split Strategy vs. Description
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return lines[0], description

# Helper to add a hyperlink run in python-docx
def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        style.font.underline = True
        style.priority = 9
        style.unhide_when_used = True
    style_el = OxmlElement('w:rStyle')
    style_el.set(qn('w:val'), 'Hyperlink')
    rPr.append(style_el)
    if font_name:
        fnt = OxmlElement('w:rFonts')
        fnt.set(qn('w:ascii'), font_name); fnt.set(qn('w:hAnsi'), font_name)
        rPr.append(fnt)
    if font_size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(font_size*2)))
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(font_size*2)))
        rPr.append(sz); rPr.append(szCs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

# Extract tables, links, and grand total
def extract_tables_and_links():
    tables = []
    grand_total = None

    table_settings = {
        "vertical_strategy": "text",
        "horizontal_strategy": "text",
        "intersection_tolerance": 3,
        "snap_tolerance": 3,
    }

    # PyMuPDF for link rectangles
    doc_fz = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_annots = [
        [(a.rect, a.uri) for a in (pg.annots() or []) if a.type[0]==1 and a.uri]
        for pg in doc_fz
    ]

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
        # Title extraction
        first = page_texts[0].splitlines() if page_texts else []
        proposal_title = next((ln for ln in first if "proposal" in ln.lower()), first[0] if first else "Untitled Proposal")
        proposal_title = proposal_title.strip()

        used_totals = set()
        def find_total(pi):
            if pi >= len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        # Page-by-page
        for pi, page in enumerate(pdf.pages):
            annots = page_annots[pi]
            # whitespace-based tables
            for tbl in page.find_tables(table_settings=table_settings):
                data = tbl.extract(x_tolerance=1, y_tolerance=1)
                if not data or len(data) < 2:
                    continue
                hdr = [str(h).strip() if h else "" for h in data[0]]
                if len(hdr) < 4 or "Start Date" not in " ".join(hdr):
                    continue
                # find Description index
                desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
                if desc_i is None:
                    desc_i = next((i for i,h in enumerate(hdr) if len(h) > 10), None)
                    if desc_i is None:
                        continue
                x0,y0,x1,y1 = tbl.bbox
                nrows = len(data)
                band = (y1 - y0) / nrows
                row_link_map = {}
                # map annotation to row
                for rect, uri in annots:
                    mid = (rect.y0 + rect.y1) / 2
                    if y0 <= mid <= y1:
                        ridx = int((mid - y0) // band)
                        if 1 <= ridx < nrows:
                            row_link_map[ridx-1] = uri

                new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
                rows_out = []
                uris_out = []
                table_total = None

                for ridx, row in enumerate(data[1:], start=1):
                    cells = [str(c).strip() if c else "" for c in row]
                    if all(not c for c in cells):
                        continue
                    first = cells[0].lower()
                    if "total" in first and any("$" in c for c in cells):
                        if table_total is None:
                            table_total = cells
                        continue
                    strat, desc = split_cell_text(cells[desc_i])
                    rest = [cells[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                    rows_out.append([strat, desc] + rest)
                    uris_out.append(row_link_map.get(ridx-1))

                if table_total is None:
                    table_total = find_total(pi)

                if rows_out:
                    tables.append((new_hdr, rows_out, uris_out, table_total))

        # grand total
        for tx in reversed(page_texts):
            m = re.search(r'Grand Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
            if m:
                grand_total = m.group(1).strip()
                break

    return proposal_title, tables, grand_total

try:
    proposal_title, tables_info, grand_total = extract_tables_and_links()
except Exception as e:
    st.error(f"Error extracting tables: {e}")
    st.stop()

# ─── Build PDF ────────────────────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((17*inch, 11*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch,
    topMargin=0.5*inch, bottomMargin=0.5*inch
)
title_style  = ParagraphStyle("Title",  fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body",   fontName=DEFAULT_SANS_FONT,  fontSize=9,  alignment=TA_LEFT, leading=11)
bl_style     = ParagraphStyle("BL",     fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, spaceBefore=6)
br_style     = ParagraphStyle("BR",     fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, spaceBefore=6)

elements = []
# Logo + Title
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp = requests.get(logo_url, timeout=10); resp.raise_for_status()
    logo = resp.content
    img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width
    w = min(5*inch, doc.width); h = w * ratio
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except:
    pass

elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), title_style), Spacer(1,24)]
total_w = doc.width

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    desc_i = hdr.index("Description")
    desc_w = total_w * 0.45
    other_w = (total_w - desc_w) / (n - 1)
    col_widths = [desc_w if i==desc_i else other_w for i in range(n)]

    wrapped = [[Paragraph(html.escape(h), header_style) for h in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            txt = html.escape(cell)
            if cidx==desc_i and uris[ridx]:
                p = Paragraph(f"{txt} <link href='{html.escape(uris[ridx])}' color='blue'>- link</link>", body_style)
            else:
                p = Paragraph(txt, body_style)
            line.append(p)
        wrapped.append(line)

    if tot:
        lbl, val = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m:
                lbl, val = m.group(1).strip(), m.group(2)
        tr = [Paragraph(lbl, bl_style)] + [Spacer(1,0)]*(n-2) + [Paragraph(val, br_style)]
        wrapped.append(tr)

    tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
        ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
        ("VALIGN",     (0,0), (-1,0),   "MIDDLE"),
        ("VALIGN",     (0,1), (-1,-1), "TOP"),
    ]
    if tot:
        style_cmds += [
            ("SPAN",  (0,-1), (-2,-1)),
            ("ALIGN", (0,-1), (-2,-1), "LEFT"),
            ("ALIGN", (-1,-1),(-1,-1),"RIGHT"),
            ("VALIGN",(0,-1), (-1,-1), "MIDDLE"),
        ]
    tbl.setStyle(TableStyle(style_cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    last_hdr = tables_info[-1][0]; n=len(last_hdr)
    di = last_hdr.index("Description")
    dw = total_w*0.45; ow = (total_w-dw)/(n-1)
    cw = [dw if i==di else ow for i in range(n)]
    row = [Paragraph("Grand Total", bl_style)] + [Spacer(1,0)]*(n-2) + [Paragraph(html.escape(grand_total), br_style)]
    gt = LongTable([row], colWidths=cw)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#E0E0E0")),
        ("GRID",      (0,0),(-1,-1), 0.25, colors.grey),
        ("VALIGN",    (0,0),(-1,-1), "MIDDLE"),
        ("SPAN",      (0,0),(-2,0)),
        ("ALIGN",     (-1,0),(-1,0), "RIGHT"),
    ]))
    elements.append(gt)

try:
    doc.build(elements)
    pdf_buf.seek(0)
except Exception as e:
    st.error(f"Error building PDF: {e}")
    pdf_buf = None

# ─── Build Word Document ─────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Inches(17); sec.page_height = Inches(11)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5)
sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)

if 'logo' in locals():
    try:
        p = docx_doc.add_paragraph(); r = p.add_run()
        img = Image.open(io.BytesIO(logo)); ratio=img.height/img.width; w_in=5
        r.add_picture(io.BytesIO(logo), width=Inches(w_in))
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    except:
        pass

p = docx_doc.add_paragraph(); p.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p.add_run(proposal_title); r.font.name = DEFAULT_SERIF_FONT; r.font.size = Pt(18); r.bold = True
docx_doc.add_paragraph()

TOTAL_W = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, uris, tot in tables_info:
    n = len(hdr); di = hdr.index("Description")
    dw = 0.45*TOTAL_W; ow = (TOTAL_W-dw)/(n-1)
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.allow_autofit=False; tbl.autofit=False
    tblPr_list = tbl._element.xpath('./w:tblPr')
    tblPr = tblPr_list[0] if tblPr_list else OxmlElement('w:tblPr')
    if not tblPr_list: tbl._element.insert(0, tblPr)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    ex = tblPr.xpath('./w:tblW')
    if ex: tblPr.remove(ex[0])
    tblPr.append(tblW)
    for i,c in enumerate(tbl.columns):
        c.width = Inches(dw if i==di else ow)

    # header
    for i,name in enumerate(hdr):
        cell = tbl.rows[0].cells[i]; tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text=""
        r=p.add_run(name); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # body
    for ridx, row in enumerate(rows):
        rc=tbl.add_row().cells
        for cidx, val in enumerate(row):
            cell=rc[cidx]; p=cell.paragraphs[0]; p.text=""
            run=p.add_run(str(val)); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            if cidx==di and ridx<len(uris) and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p, uris[ridx], "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP

    # total row
    if tot:
        trow=tbl.add_row().cells
        lbl,amt="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; amt=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl,amt=m.group(1).strip(),m.group(2)
        lc=trow[0]
        if n>1: lc.merge(trow[n-2])
        p=lc.paragraphs[0]; p.text=""; r=p.add_run(lbl); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=trow[n-1]
        p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(amt); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

    docx_doc.add_paragraph()

# grand total
if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; n=len(last_hdr); di=last_hdr.index("Description")
    dw=0.45*TOTAL_W; ow=(TOTAL_W-dw)/(n-1)
    tblg=docx_doc.add_table(rows=1,cols=n,style="Table Grid"); tblg.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblg.allow_autofit=False; tblg.autofit=False
    tblPr_list=tblg._element.xpath('./w:tblPr')
    tblPr=tblPr_list[0] if tblPr_list else OxmlElement('w:tblPr')
    if not tblPr_list: tblg._element.insert(0,tblPr)
    tblW=OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    ex=tblPr.xpath('./w:tblW'); 
    if ex: tblPr.remove(ex[0])
    tblPr.append(tblW)
    for i,c in enumerate(tblg.columns):
        c.width=Inches(dw if i==di else ow)
    cells=tblg.rows[0].cells
    lc=cells[0]
    if n>1: lc.merge(cells[n-2])
    tc=lc._tc; tcPr=tc.get_or_add_tcPr()
    sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),'E0E0E0'); tcPr.append(sh)
    p=lc.paragraphs[0]; p.text=""; r=p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]
    tc2=ac._tc; tcPr2=tc2.get_or_add_tcPr()
    sh2=OxmlElement('w:shd'); sh2.set(qn('w:fill'),'E0E0E0'); tcPr2.append(sh2)
    p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Save and download
docx_buf = io.BytesIO()
try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except:
    docx_buf = None

c1, c2 = st.columns(2)
if pdf_buf:
    c1.download_button(
        "📥 Download deliverable PDF",
        data=pdf_buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True
    )
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button(
        "📥 Download deliverable DOCX",
        data=docx_buf,
        file_name="proposal_deliverable.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
else:
    c2.error("Word document generation failed.")
