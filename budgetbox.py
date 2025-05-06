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
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image as RLImage
)

# Register fonts
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except:
    DEFAULT_SERIF_FONT = "Times New Roman"
    DEFAULT_SANS_FONT = "Arial"

# Streamlit UI setup
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
    title = lines[0]
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return title, description

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK,
        is_external=True
    )
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
        size = OxmlElement('w:sz')
        size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs')
        size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size); rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

# Standard headers
STANDARD_HEADERS = [
    "Strategy", "Description", "Start Date", "End Date",
    "Term (Months)", "Monthly Amount", "Item Total", "Notes"
]

# Try Camelot on page 1
first_table = None
try:
    tables = camelot.read_pdf(
        filepath_or_buffer=io.BytesIO(pdf_bytes),
        pages="1", flavor="lattice", strip_text="\n"
    )
    if tables and tables:
        df = tables[0].df
        raw = df.values.tolist()
        if len(raw) > 1 and len(raw[0]) >= len(STANDARD_HEADERS):
            first_table = raw
except:
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [
        p.extract_text(x_tolerance=1, y_tolerance=1) or ""
        for p in pdf.pages
    ]
    first_page_lines = page_texts[0].splitlines() if page_texts else []
    potential_title = next(
        (ln.strip() for ln in first_page_lines
         if "proposal" in ln.lower() and len(ln.strip()) > 5),
        None
    )
    if potential_title:
        proposal_title = potential_title
    elif first_page_lines:
        proposal_title = first_page_lines[0].strip()

    used_totals = set()
    def find_total(pi):
        if pi >= len(page_texts):
            return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+',
                         ln, re.I) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        if pi == 0 and first_table:
            tables_found = [("camelot", first_table, None)]
            links = []
        else:
            tables_found = [
                (tbl, tbl.extract(x_tolerance=1, y_tolerance=1), tbl.bbox)
                for tbl in page.find_tables()
            ]
            links = page.hyperlinks

        for tbl_obj, data, bbox in tables_found:
            if not data or len(data) < 2:
                continue

            hdr = [str(h).strip() for h in data[0]]
            desc_i = next(
                (i for i,h in enumerate(hdr)
                 if h and "description" in h.lower()),
                None
            )
            if desc_i is None:
                desc_i = next(
                    (i for i,h in enumerate(hdr) if len(h) > 10),
                    None
                )
                if desc_i is None:
                    continue

            desc_links = {}
            if tbl_obj != "camelot":
                for r, row_obj in enumerate(tbl_obj.rows):
                    if r == 0:
                        continue
                    if desc_i < len(row_obj.cells):
                        x0, top, x1, bottom = row_obj.cells[desc_i]
                        for link in links:
                            if all(k in link for k in ("x0","x1","top","bottom","uri")):
                                if not (
                                    link["x1"] < x0
                                    or link["x0"] > x1
                                    or link["bottom"] < top
                                    or link["top"] > bottom
                                ):
                                    desc_links[r] = link["uri"]
                                    break

            header_mapping = {}
            for i, h in enumerate(hdr):
                if not h or h.lower() == "none":
                    continue
                hl = h.lower()
                if any(x in hl for x in ["strategy","service","product"]):
                    header_mapping[i] = 0
                elif any(x in hl for x in ["description","details"]):
                    header_mapping[i] = 1
                elif any(x in hl for x in ["start date","start","begin"]):
                    header_mapping[i] = 2
                elif any(x in hl for x in ["end date","end","finish"]):
                    header_mapping[i] = 3
                elif any(x in hl for x in ["term","duration","period","months"]):
                    header_mapping[i] = 4
                elif any(x in hl for x in ["monthly","per month","rate","recurring"]):
                    header_mapping[i] = 5
                elif any(x in hl for x in ["item total","subtotal","line total","amount"]):
                    header_mapping[i] = 6
                elif any(x in hl for x in ["note","comment","additional"]):
                    header_mapping[i] = 7

            rows_data = []
            row_links = []
            table_total = None

            for ridx, row in enumerate(data[1:], start=1):
                cells = [str(c).strip() for c in row]
                if not any(cells):
                    continue

                first_cell = cells[0].lower()
                if ("total" in first_cell or "subtotal" in first_cell) and any("$" in c for c in cells):
                    if table_total is None:
                        table_total = cells
                    continue

                strat, desc = split_cell_text(cells[desc_i] if desc_i < len(cells) else "")
                standardized_row = [""] * len(STANDARD_HEADERS)
                standardized_row[0] = strat
                standardized_row[1] = desc

                # Primary pass
                for i, cv in enumerate(cells):
                    if i == desc_i or not cv.strip():
                        continue
                    val = cv.strip()

                    if i in header_mapping:
                        idx = header_mapping[i]
                        if idx > 1:
                            standardized_row[idx] = val
                        continue

                    # NEW: pure integer => Term (Months)
                    if re.fullmatch(r"\d{1,3}", val):
                        standardized_row[4] = val
                        continue

                    # Dollar amounts
                    if "$" in val:
                        # If monthly empty but looks recurring => monthly
                        if not standardized_row[5] and ("month" in val.lower() or "mo" in val.lower()):
                            standardized_row[5] = val
                        # Otherwise item total
                        elif not standardized_row[6]:
                            standardized_row[6] = val
                        continue

                    # Dates
                    if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', val):
                        if not standardized_row[2]:
                            standardized_row[2] = val
                        elif not standardized_row[3]:
                            standardized_row[3] = val
                        continue

                    # Term expressions
                    if re.search(r'\b\d+\s*(?:month|mo|yr|year)', val.lower()):
                        standardized_row[4] = val
                        continue

                    # Fallback => notes
                    if not standardized_row[7]:
                        standardized_row[7] = val

                # Positional fallback
                if not any(standardized_row[2:]):
                    rem = [c for j,c in enumerate(cells) if j != desc_i and c.strip()]
                    for cv in rem:
                        v = cv.strip()
                        if re.fullmatch(r"\d{1,3}", v):
                            standardized_row[4] = v
                            continue
                        if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', v):
                            if not standardized_row[2]:
                                standardized_row[2] = v
                            elif not standardized_row[3]:
                                standardized_row[3] = v
                            continue
                        if "$" in v:
                            if not standardized_row[5] and ("month" in v.lower() or "mo" in v.lower()):
                                standardized_row[5] = v
                            elif not standardized_row[6]:
                                standardized_row[6] = v
                            continue
                        if re.search(r'\b\d+\s*(?:month|mo|yr|year)', v.lower()):
                            standardized_row[4] = v
                            continue
                        if not standardized_row[7]:
                            standardized_row[7] = v

                # **NEW SECTION**: swap & compute totals
                # If monthly is empty but item total is present, treat that as monthly
                if not standardized_row[5] and standardized_row[6]:
                    standardized_row[5] = standardized_row[6]
                    standardized_row[6] = ""

                # If we have monthly & term, but no item total, compute it
                if standardized_row[5] and standardized_row[4] and not standardized_row[6]:
                    # parse amounts/term
                    m_amt = float(standardized_row[5].replace("$","").replace(",",""))
                    mon_match = re.search(r'\d+', standardized_row[4])
                    if mon_match:
                        months = int(mon_match.group())
                        total = m_amt * months
                        standardized_row[6] = f"${total:,.2f}"

                rows_data.append(standardized_row)
                row_links.append(desc_links.get(ridx))

            if table_total is None:
                table_total = find_total(pi)
            if rows_data:
                tables_info.append((STANDARD_HEADERS, rows_data, row_links, table_total))

    # Grand Total
    for block in reversed(page_texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', block, re.I|re.S)
        if m:
            grand_total = m.group(1).replace(" ","")
            break

# Build PDF
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((17*inch,11*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch,
    topMargin=0.5*inch, bottomMargin=0.5*inch
)
ts = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT,
                    fontSize=18, alignment=TA_CENTER, spaceAfter=12)
hs = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT,
                    fontSize=10, alignment=TA_CENTER, textColor=colors.black)
bs = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT,
                    fontSize=9, alignment=TA_LEFT, leading=11)
bls = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT,
                     fontSize=10, alignment=TA_LEFT, spaceBefore=6)
brs = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT,
                     fontSize=10, alignment=TA_RIGHT, spaceBefore=6)
elements = []

logo = None
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp = requests.get(logo_url, timeout=10); resp.raise_for_status()
    logo = resp.content
    img = Image.open(io.BytesIO(logo)); ratio = img.height/img.width
    w = min(5*inch,doc.width); h = w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except:
    pass

elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
total_w = doc.width

for hdr, rows, links, tot in tables_info:
    n = len(hdr); desc_idx = 1
    # column widths
    desc_w = total_w*0.25; strat_w = total_w*0.15
    date_w = total_w*0.08; term_w = total_w*0.08
    amount_w = total_w*0.10; total_w_col = total_w*0.10
    notes_w = total_w*0.16
    col_ws = [strat_w,desc_w,date_w,date_w,term_w,amount_w,total_w_col,notes_w]

    wrapped = [[Paragraph(html.escape(h),hs) for h in hdr]]
    for i, row in enumerate(rows):
        line=[]
        for j, cell in enumerate(row):
            txt = html.escape(cell)
            if j==desc_idx and i < len(links) and links[i]:
                line.append(Paragraph(
                    f"{txt} <link href='{html.escape(links[i])}' color='blue'>- link</link>",
                    bs
                ))
            else:
                line.append(Paragraph(txt, bs))
        wrapped.append(line)
    if tot:
        lbl, val = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl,val = m.group(1).strip(), m.group(2)
        total_row = [Paragraph(lbl,bls)] + [Paragraph("",bs)]*(n-2) + [Paragraph(val,brs)]
        wrapped.append(total_row)

    tbl = LongTable(wrapped, colWidths=col_ws, repeatRows=1)
    cmds = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",      (0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",    (0,0),(-1,0),"MIDDLE"),
        ("VALIGN",    (0,1),(-1,-1),"TOP"),
    ]
    if tot and n>1:
        cmds += [
            ("SPAN",  (0,-1),(-2,-1)),
            ("ALIGN", (0,-1),(-2,-1),"LEFT"),
            ("ALIGN", (-1,-1),(-1,-1),"RIGHT"),
            ("VALIGN",(0,-1),(-1,-1),"MIDDLE"),
        ]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    n = len(STANDARD_HEADERS)
    col_ws = [total_w*0.15, total_w*0.25, total_w*0.08, total_w*0.08,
              total_w*0.08, total_w*0.10, total_w*0.10, total_w*0.16]
    row = [Paragraph("Grand Total",bls)] + [Paragraph("",bs)]*(n-2) + [Paragraph(html.escape(grand_total), brs)]
    gt = LongTable([row],colWidths=col_ws)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),
        ("ALIGN",(-1,0),(-1,0),"RIGHT"),
    ]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# Build DOCX deliverable (identical logic for column widths, etc.)
# ... [DOCX build code continues unchanged] ...

# Download buttons
c1,c2 = st.columns(2)
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
