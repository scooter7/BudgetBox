# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
import docx  # Make sure docx is imported
from docx.shared import Inches, Pt, RGBColor  # Import RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE  # Import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
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
import fitz  # PyMuPDF for link annotations

# Register fonts
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except Exception as e:
    st.warning(f"Could not load custom fonts: {e}. Using system defaults.")
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

def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return lines[0], description

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
    style_element = OxmlElement('w:rStyle')
    style_element.set(qn('w:val'), 'Hyperlink')
    rPr.append(style_element)
    if font_name:
        run_font = OxmlElement('w:rFonts')
        run_font.set(qn('w:ascii'), font_name)
        run_font.set(qn('w:hAnsi'), font_name)
        rPr.append(run_font)
    if font_size:
        size = OxmlElement('w:sz'); size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size); rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

def extract_tables_and_links():
    tables = []
    grand = None

    # build whitespace-based table settings
    table_settings = {
        "vertical_strategy":   "text",
        "horizontal_strategy": "text",
        "intersection_tolerance": 3,
        "snap_tolerance": 3,
    }

    # Open source PDF to extract text + PyMuPDF annotations
    doc_fz = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_annots = [
        [(a.rect, a.uri) for a in (pg.annots() or []) if a.type[0] == 1 and a.uri]
        for pg in doc_fz
    ]

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
        title = next(
            (ln for pg in page_texts for ln in pg.splitlines() if "proposal" in ln.lower()),
            "Untitled Proposal"
        ).strip()

        used_totals = set()
        def find_total(pi):
            if pi >= len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            annots = page_annots[pi]

            for tbl in page.find_tables(table_settings=table_settings):
                data = tbl.extract(x_tolerance=1, y_tolerance=1)
                if not data or len(data) < 2:
                    continue

                hdr = [str(h).strip() if h else "" for h in data[0]]
                desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
                if desc_i is None:
                    desc_i = next((i for i,h in enumerate(hdr) if len(h) > 10), None)
                    if desc_i is None:
                        continue

                x0,y0,x1,y1 = tbl.bbox
                nrows = len(data)
                band = (y1 - y0)/nrows
                row_map = {}
                for rect, uri in annots:
                    midy = (rect.y0 + rect.y1)/2
                    if y0 <= midy <= y1:
                        ridx = int((midy - y0)//band)
                        if 1 <= ridx < nrows:
                            row_map[ridx-1] = uri

                new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
                rows_data = []
                uris = []
                table_total_info = None

                for ridx, row in enumerate(data[1:], start=1):
                    cells = [str(c).strip() if c else "" for c in row]
                    if all(not c for c in cells):
                        continue
                    first = cells[0].lower()
                    if "total" in first and any("$" in c for c in cells):
                        if table_total_info is None:
                            table_total_info = cells
                        continue

                    strat, desc = split_cell_text(cells[desc_i])
                    rest = [cells[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                    rows_data.append([strat, desc] + rest)
                    uris.append(row_map.get(ridx-1))

                if table_total_info is None:
                    table_total_info = find_total(pi)

                if rows_data:
                    tables.append((new_hdr, rows_data, uris, table_total_info))

        for tx in reversed(page_texts):
            m = re.search(r'Grand Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
            if m:
                grand = m.group(1).replace(" ", "")
                break

    return title, tables, grand

# Extract everything
try:
    proposal_title, tables_info, grand_total = extract_tables_and_links()
except Exception as e:
    st.error(f"Error extracting tables: {e}")
    st.stop()

# ─── Build PDF via ReportLab ─────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((17*inch, 11*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch,
    topMargin=0.5*inch, bottomMargin=0.5*inch
)
title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11)
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, spaceBefore=6)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, spaceBefore=6)

elements = []
# Logo + Title
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp = requests.get(logo_url, timeout=10); resp.raise_for_status()
    logo = resp.content
    img = Image.open(io.BytesIO(logo))
    ratio = img.height / img.width
    img_w = min(5*inch, doc.width)
    img_h = img_w * ratio
    elements.append(RLImage(io.BytesIO(logo), width=img_w, height=img_h))
except:
    pass

elements += [
    Spacer(1,12),
    Paragraph(html.escape(proposal_title), title_style),
    Spacer(1,24)
]

total_w = doc.width

for hdr, rows, uris, tot in tables_info:
    num_cols = len(hdr)
    desc_idx = hdr.index("Description")
    desc_w = total_w * 0.45
    other_w = (total_w - desc_w) / (num_cols - 1)
    col_widths = [desc_w if i==desc_idx else other_w for i in range(num_cols)]

    wrapped = [[Paragraph(html.escape(h), header_style) for h in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            text = html.escape(cell)
            if cidx==desc_idx and uris[ridx]:
                p = Paragraph(f"{text} <link href='{html.escape(uris[ridx])}' color='blue'>- link</link>", body_style)
            else:
                p = Paragraph(text, body_style)
            line.append(p)
        wrapped.append(line)

    if tot:
        label, val = "Total", ""
        if isinstance(tot,list):
            label = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m:
                label, val = m.group(1).strip(), m.group(2)
        total_row = [Paragraph(label, bl_style)] + [Spacer(1,0)]*(num_cols-2) + [Paragraph(val, br_style)]
        wrapped.append(total_row)

    tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("VALIGN", (0,1), (-1,-1), "TOP"),
    ]
    if tot:
        style_cmds += [
            ("SPAN", (0,-1), (-2,-1)),
            ("ALIGN", (0,-1), (-2,-1), "LEFT"),
            ("ALIGN", (-1,-1), (-1,-1), "RIGHT"),
            ("VALIGN", (0,-1), (-1,-1), "MIDDLE"),
        ]
    tbl.setStyle(TableStyle(style_cmds))
    elements += [tbl, Spacer(1,24)]

# Grand Total
if grand_total and tables_info:
    last_hdr = tables_info[-1][0]
    num_cols = len(last_hdr)
    desc_idx = last_hdr.index("Description")
    desc_w = total_w * 0.45
    other_w = (total_w - desc_w) / (num_cols - 1)
    col_widths = [desc_w if i==desc_idx else other_w for i in range(num_cols)]
    row = [Paragraph("Grand Total", bl_style)] + [Spacer(1,0)]*(num_cols-2) + [Paragraph(html.escape(grand_total), br_style)]
    gt = LongTable([row], colWidths=col_widths)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")
    ]))
    elements.append(gt)

try:
    doc.build(elements)
    pdf_buf.seek(0)
except Exception as e:
    st.error(f"Error building PDF: {e}")
    pdf_buf = None

# ─── Build Word Document ──────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Inches(17)
sec.page_height = Inches(11)
sec.left_margin = Inches(0.5)
sec.right_margin = Inches(0.5)
sec.top_margin = Inches(0.5)
sec.bottom_margin = Inches(0.5)

if 'logo' in locals():
    try:
        p_logo = docx_doc.add_paragraph()
        r_logo = p_logo.add_run()
        img = Image.open(io.BytesIO(logo))
        ratio = img.height / img.width
        w_in = 5
        r_logo.add_picture(io.BytesIO(logo), width=Inches(w_in))
        p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
    except:
        pass

p_title = docx_doc.add_paragraph()
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
run = p_title.add_run(proposal_title)
run.font.name = DEFAULT_SERIF_FONT
run.font.size = Pt(18)
run.bold = True
docx_doc.add_paragraph()

TOTAL_W_IN = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    if n == 0:
        continue
    desc_idx = hdr.index("Description")
    desc_w = 0.45 * TOTAL_W_IN
    other_w = (TOTAL_W_IN - desc_w) / (n - 1)

    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    tbl.autofit = False

    tblPr_list = tbl._element.xpath('./w:tblPr')
    tblPr = tblPr_list[0] if tblPr_list else OxmlElement('w:tblPr')
    if not tblPr_list:
        tbl._element.insert(0, tblPr)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    existing = tblPr.xpath('./w:tblW')
    if existing:
        tblPr.remove(existing[0])
    tblPr.append(tblW)

    for i,col in enumerate(tbl.columns):
        col.width = Inches(desc_w if i==desc_idx else other_w)

    # Header
    hdr_cells = tbl.rows[0].cells
    for i, name in enumerate(hdr):
        cell = hdr_cells[i]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text = ""
        r = p.add_run(name)
        r.font.name = DEFAULT_SERIF_FONT; r.font.size = Pt(10); r.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Body
    for ridx, row in enumerate(rows):
        rcells = tbl.add_row().cells
        for cidx, val in enumerate(row):
            cell = rcells[cidx]
            p = cell.paragraphs[0]; p.text = ""
            run = p.add_run(str(val))
            run.font.name = DEFAULT_SANS_FONT; run.font.size = Pt(9)
            if cidx==desc_idx and ridx < len(uris) and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p, uris[ridx], "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    # Total row
    if tot:
        trow = tbl.add_row().cells
        label, amount = "Total", ""
        if isinstance(tot,list):
            label = tot[0] or "Total"
            amount = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m:
                label, amount = m.group(1).strip(), m.group(2)
        label_cell = trow[0]
        if n>1:
            label_cell.merge(trow[n-2])
        p = label_cell.paragraphs[0]; p.text=""
        r = p.add_run(label); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.LEFT
        label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        amt_cell = trow[n-1]
        p2 = amt_cell.paragraphs[0]; p2.text=""
        r2 = p2.add_run(amount); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment = WD_TABLE_ALIGNMENT.RIGHT
        amt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    docx_doc.add_paragraph()

# Grand Total
if grand_total and tables_info:
    last_hdr = tables_info[-1][0]; n=len(last_hdr)
    desc_idx = last_hdr.index("Description")
    desc_w = 0.45 * TOTAL_W_IN; other_w = (TOTAL_W_IN - desc_w)/(n-1)
    tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblg.allow_autofit = False; tblg.autofit = False

    tblPr_list = tblg._element.xpath('./w:tblPr')
    tblPr = tblPr_list[0] if tblPr_list else OxmlElement('w:tblPr')
    if not tblPr_list: tblg._element.insert(0, tblPr)
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    existing = tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)

    for i,col in enumerate(tblg.columns):
        col.width = Inches(desc_w if i==desc_idx else other_w)

    cells = tblg.rows[0].cells
    label_cell = cells[0]
    if n>1: label_cell.merge(cells[n-2])
    tc = label_cell._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'E0E0E0'); tcPr.append(shd)
    p = label_cell.paragraphs[0]; p.text=""
    r = p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment = WD_TABLE_ALIGNMENT.LEFT
    label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    amt_cell = cells[n-1]
    tc2 = amt_cell._tc; tcPr2 = tc2.get_or_add_tcPr()
    shd2 = OxmlElement('w:shd'); shd2.set(qn('w:fill'), 'E0E0E0'); tcPr2.append(shd2)
    p2 = amt_cell.paragraphs[0]; p2.text=""
    r2 = p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment = WD_TABLE_ALIGNMENT.RIGHT
    amt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Save and offer downloads
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
