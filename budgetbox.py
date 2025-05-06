# -*- coding: utf-8 -*-
import io
import re
import html
import camelot
import pdfplumber
import requests
import streamlit as st
from PIL import Image
from docx import Document
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage

FIXED_COLUMNS = [
    "Strategy",
    "Description",
    "Term (Months)",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes"
]

try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except:
    DEFAULT_SERIF_FONT = "Times New Roman"
    DEFAULT_SANS_FONT = "Arial"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw):
    lines = [l.strip() for l in str(raw).splitlines() if l.strip()]
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

def clean_header(header_row, data_rows):
    cleaned_header = []
    idxs = []
    for i, h in enumerate(header_row):
        hstr = str(h).strip() if h is not None else ""
        if hstr.lower() == "none" or hstr == "":
            continue
        cleaned_header.append(hstr)
        idxs.append(i)
    new_data_rows = []
    for row in data_rows:
        new_data_rows.append([row[i] if i < len(row) else "" for i in idxs])
    return cleaned_header, new_data_rows

def normalize_column_name(col):
    val = (col or "").strip().lower()
    equivalents = {
        "term (months)": ["term (months)", "months", "duration"],
        "strategy": ["strategy"],
        "description": ["description", "desc"],
        "start date": ["start date", "date start", "start"],
        "end date": ["end date", "date end", "end"],
        "monthly amount": ["monthly amount", "monthly", "monthlyamt", "monthly fee", "amount/month", "monthly total"],
        "item total": ["item total", "total", "amount", "itemtotal", "subtotal"],
        "notes": ["notes", "note", "remarks", "comments"]
    }
    for canon, variants in equivalents.items():
        if val in [v.lower() for v in variants]:
            return canon
    return col.strip()

def map_table_to_columns(header, rows):
    header_map = {}
    for idx, h in enumerate(header):
        norm_h = normalize_column_name(h)
        for fix in FIXED_COLUMNS:
            if norm_h.lower() == fix.lower():
                header_map[fix] = idx
                break
    if not header_map:
        return None
    out_rows = []
    for row in rows:
        new_row = []
        for fix in FIXED_COLUMNS:
            val = row[header_map[fix]] if fix in header_map and header_map[fix] < len(row) else ""
            if val is None:
                val = ""
            new_row.append(str(val))
        out_rows.append(new_row)
    return FIXED_COLUMNS, out_rows

first_table = None
try:
    tables = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n")
    if tables and len(tables) > 0:
        df = tables[0].df
        raw = df.values.tolist()
        if len(raw) > 1:
            header1, header2 = raw[0], raw[1]
            combined = [
                header2[j] if header2[j].strip() and header2[j].lower() != "none" else header1[j]
                for j in range(len(header2))
            ]
            keep_idxs = [j for j, val in enumerate(combined) if val.strip() and val.lower() != "none"]
            new_raw = [[combined[j] for j in keep_idxs]]
            for row in raw[2:]:
                new_raw.append([row[j] for j in keep_idxs])
            first_table = new_raw
except Exception as e:
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
    first_page_lines = page_texts[0].splitlines() if page_texts else []
    pt = next((ln.strip() for ln in first_page_lines if "proposal" in ln.lower() and len(ln.strip()) > 5), None)
    if pt:
        proposal_title = pt
    elif first_page_lines:
        proposal_title = first_page_lines[0].strip()
    used_totals = set()

    def find_total(pi):
        if pi >= len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        if pi == 0 and first_table:
            data = first_table
            source = "camelot"
            links = []
            bbox = None
        else:
            source = "plumber"
            tables_found = page.find_tables()
            data = None
            for tbl in tables_found:
                d = tbl.extract(x_tolerance=1, y_tolerance=1)
                if d and len(d) >= 2:
                    data = d
                    bbox = tbl.bbox
                    break
            links = getattr(page, "hyperlinks", [])
        if not data or len(data) < 2:
            continue

        hdr, data_rows = clean_header(data[0], data[1:])

        filtered_rows = [row for row in data_rows if any(c and c.lower() != "none" for c in row)]

        table_total = None
        processed_rows = []
        for row in filtered_rows:
            first = str(row[0]).strip().lower() if len(row) > 0 else ""
            if ("total" in first or "subtotal" in first) and any("$" in str(c) for c in row):
                if table_total is None:
                    table_total = row
                continue
            processed_rows.append(row)
        if table_total is None:
            table_total = find_total(pi)

        map_result = map_table_to_columns(hdr, processed_rows)
        if not map_result:
            continue
        mapped_hdr, mapped_rows = map_result

        desc_idx = mapped_hdr.index("Description")
        desc_links = {}
        if source == "plumber" and bbox and tables_found and hasattr(page, "hyperlinks"):
            first_tbl = tables_found[0]
            for r, row_obj in enumerate(first_tbl.rows):
                if r == 0: continue
                if desc_idx < len(row_obj.cells):
                    cell = row_obj.cells[desc_idx]
                    if cell is None:
                        continue
                    x0, top, x1, bottom = cell  # <--- FIX: safe to unpack now
                    for link in links:
                        if all(k in link for k in ("x0", "x1", "top", "bottom", "uri")):
                            if not (link["x1"] < x0 or link["x0"] > x1 or link["bottom"] < top or link["top"] > bottom):
                                desc_links[r] = link["uri"]
                                break
        row_links = []
        for ridx in range(len(mapped_rows)):
            row_links.append(desc_links.get(ridx+1))

        tables_info.append((mapped_hdr, mapped_rows, row_links, table_total))

    for tx in reversed(page_texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1).replace(" ", "")
            break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((17 * inch, 11 * inch)),
    leftMargin=0.5 * inch, rightMargin=0.5 * inch,
    topMargin=0.5 * inch, bottomMargin=0.5 * inch
)
title_style = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
body_style = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11)
bl_style = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, spaceBefore=6)
br_style = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, spaceBefore=6)
elements = []
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp = requests.get(logo_url, timeout=10)
    resp.raise_for_status()
    logo = resp.content
    img = Image.open(io.BytesIO(logo))
    ratio = img.height / img.width
    w = min(5 * inch, doc.width)
    h = w * ratio
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except Exception as e:
    pass
elements += [Spacer(1, 12), Paragraph(html.escape(proposal_title), title_style), Spacer(1, 24)]

total_w = doc.width
for hdr, rows, links, tot in tables_info:
    n = len(FIXED_COLUMNS)
    desc_idx = FIXED_COLUMNS.index("Description")
    desc_w = total_w * 0.20
    other_w = (total_w - desc_w) / (n - 1) if n > 1 else total_w
    col_ws = [desc_w if i == desc_idx else other_w for i in range(n)]

    wrapped = [[Paragraph(html.escape(h), header_style) for h in FIXED_COLUMNS]]
    for i, row in enumerate(rows):
        line = []
        for j, cell in enumerate(row):
            txt = html.escape(cell)
            if j == desc_idx and i < len(links) and links[i]:
                line.append(Paragraph(f"{txt} <link href='{html.escape(links[i])}' color='blue'>- link</link>", body_style))
            else:
                line.append(Paragraph(txt, body_style))
        wrapped.append(line)
    if tot:
        lbl, val = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in str(c)), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', str(tot))
            if m:
                lbl, val = m.group(1).strip(), m.group(2)
        tr = [Paragraph(lbl, bl_style)] + [Paragraph("", body_style)] * (n - 2) + [Paragraph(val, br_style)]
        wrapped.append(tr)
    tbl = LongTable(wrapped, colWidths=col_ws, repeatRows=1)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
    ]
    if tot and n > 1:
        cmds += [("SPAN", (0, -1), (-2, -1)), ("ALIGN", (0, -1), (-2, -1), "LEFT"),
                 ("ALIGN", (-1, -1), (-1, -1), "RIGHT"), ("VALIGN", (0, -1), (-1, -1), "MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1, 24)]
if grand_total and tables_info:
    n = len(FIXED_COLUMNS)
    desc_idx = FIXED_COLUMNS.index("Description")
    desc_w = total_w * 0.20
    other_w = (total_w - desc_w) / (n - 1) if n > 1 else total_w
    col_ws = [desc_w if i == desc_idx else other_w for i in range(n)]
    row = [Paragraph("Grand Total", bl_style)] + [Paragraph("", body_style)] * (n - 2) + [
        Paragraph(html.escape(grand_total), br_style)]
    gt = LongTable([row], colWidths=col_ws)
    gt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-2, 0)),
        ("ALIGN", (-1, 0), (-1, 0), "RIGHT"),
    ]))
    elements.append(gt)
doc.build(elements)
pdf_buf.seek(0)

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
        p = docx_doc.add_paragraph()
        r = p.add_run()
        img = Image.open(io.BytesIO(logo))
        ratio = img.height / img.width
        w_in = 5
        h_in = w_in * ratio
        r.add_picture(io.BytesIO(logo), width=Inches(w_in))
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    except Exception as e:
        pass
p = docx_doc.add_paragraph()
p.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p.add_run(proposal_title)
r.font.name = DEFAULT_SERIF_FONT
r.font.size = Pt(18)
r.bold = True
docx_doc.add_paragraph()
TOTAL_W_IN = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches
for hdr, rows, links, tot in tables_info:
    n = len(FIXED_COLUMNS)
    if n == 0: continue
    desc_idx = FIXED_COLUMNS.index("Description")
    desc_w = 0.20 * TOTAL_W_IN
    other_w = (TOTAL_W_IN - desc_w) / (n - 1) if n > 1 else TOTAL_W_IN
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    tbl.autofit = False
    tblPr_list = tbl._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr = OxmlElement('w:tblPr')
        tbl._element.insert(0, tblPr)
    else:
        tblPr = tblPr_list[0]
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    existing = tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i, col in enumerate(tbl.columns):
        col.width = Inches(desc_w if i == desc_idx else other_w)
    hdr_cells = tbl.rows[0].cells
    for i, name in enumerate(FIXED_COLUMNS):
        cell = hdr_cells[i]
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'F2F2F2')
        tcPr.append(shd)
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(name)
        run.font.name = DEFAULT_SERIF_FONT
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx, row in enumerate(rows):
        rcells = tbl.add_row().cells
        for cidx, val in enumerate(row):
            cell = rcells[cidx]
            p = cell.paragraphs[0]
            p.text = ""
            run = p.add_run(str(val))
            run.font.name = DEFAULT_SANS_FONT
            run.font.size = Pt(9)
            if cidx == desc_idx and ridx < len(links) and links[ridx]:
                p.add_run(" ")
                add_hyperlink(p, links[ridx], "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        trow = tbl.add_row().cells
        lbl, amt = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"
            amt = next((c for c in reversed(tot) if "$" in str(c)), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', str(tot))
            if m:
                lbl, amt = m.group(1).strip(), m.group(2)
        lc = trow[0]
        if n > 1: lc.merge(trow[n - 2])
        p = lc.paragraphs[0]
        p.text = ""
        r = p.add_run(lbl)
        r.font.name = DEFAULT_SERIF_FONT
        r.font.size = Pt(10)
        r.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.LEFT
        lc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac = trow[n - 1]
        p2 = ac.paragraphs[0]
        p2.text = ""
        r2 = p2.add_run(amt)
        r2.font.name = DEFAULT_SERIF_FONT
        r2.font.size = Pt(10)
        r2.bold = True
        p2.alignment = WD_TABLE_ALIGNMENT.RIGHT
        ac.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()
if grand_total and tables_info:
    n = len(FIXED_COLUMNS)
    desc_idx = FIXED_COLUMNS.index("Description")
    desc_w = 0.20 * TOTAL_W_IN
    other_w = (TOTAL_W_IN - desc_w) / (n - 1) if n > 1 else TOTAL_W_IN
    tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    tblg.allow_autofit = False
    tblg.autofit = False
    tblPr_list = tblg._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr = OxmlElement('w:tblPr')
        tblg._element.insert(0, tblPr)
    else:
        tblPr = tblPr_list[0]
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    existing = tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i, col in enumerate(tblg.columns):
        col.width = Inches(desc_w if i == desc_idx else other_w)
    cells = tblg.rows[0].cells
    lc = cells[0]
    if n > 1: lc.merge(cells[n - 2])
    tc = lc._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), 'E0E0E0')
    tcPr.append(shd)
    p = lc.paragraphs[0]
    p.text = ""
    r = p.add_run("Grand Total")
    r.font.name = DEFAULT_SERIF_FONT
    r.font.size = Pt(10)
    r.bold = True
    p.alignment = WD_TABLE_ALIGNMENT.LEFT
    lc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac = cells[n - 1]
    tc2 = ac._tc
    tcPr2 = tc2.get_or_add_tcPr()
    shd2 = OxmlElement('w:shd')
    shd2.set(qn('w:fill'), 'E0E0E0')
    tcPr2.append(shd2)
    p2 = ac.paragraphs[0]
    p2.text = ""
    r2 = p2.add_run(grand_total)
    r2.font.name = DEFAULT_SERIF_FONT
    r2.font.size = Pt(10)
    r2.bold = True
    p2.alignment = WD_TABLE_ALIGNMENT.RIGHT
    ac.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1, c2 = st.columns(2)
c1.download_button(
    "📥 Download deliverable PDF",
    data=pdf_buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True
)
c2.download_button(
    "📥 Download deliverable DOCX",
    data=docx_buf,
    file_name="proposal_deliverable.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)
