import streamlit as st
import pdfplumber
import io
import requests
import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
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

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Capture original PDF link annotations ────────────────────────────────────
doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
page_annotations = []
for page in doc_fitz:
    annots = []
    for a in page.annots() or []:
        if a.type[0] == 1 and a.uri:
            annots.append((a.rect, a.uri))
    page_annotations.append(annots)

# ─── Helpers ───────────────────────────────────────────────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

def add_hyperlink(paragraph, url, text,
                  font_name="Barlow", font_size=9, bold=False, align=None):
    """
    Inserts a <w:hyperlink> using the built-in 'Hyperlink' style so Word
    recognizes it as a clickable link.
    """
    part = paragraph.part
    r_id = part.relate_to(
        url,
        RELATIONSHIP_TYPE.HYPERLINK,
        is_external=True
    )
    hlink = OxmlElement("w:hyperlink")
    hlink.set(qn("r:id"), r_id)

    run_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    # Use Word's default Hyperlink style (blue + underline)
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    run_elem.append(rPr)

    t = OxmlElement("w:t")
    t.text = text
    run_elem.append(t)
    hlink.append(run_elem)

    paragraph._p.append(hlink)

    # Optionally override font name/size
    run = paragraph.runs[-1]
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    if align is not None:
        paragraph.alignment = align

# ─── Extract tables + link URIs ────────────────────────────────────────────────
tables_info = []  # list of (header, rows, row_links, table_total)
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
        annots = page_annotations[pi]
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue
            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            # map link-rect → row index
            x0, y0, x1, y1 = tbl.bbox
            nrows = len(data)
            band_h = (y1 - y0) / nrows
            row_links_map = {}
            for rect, uri in annots:
                midy = (rect.y0 + rect.y1) / 2
                if y0 <= midy <= y1:
                    ridx = int((midy - y0) // band_h)
                    if 1 <= ridx < nrows:
                        row_links_map[ridx-1] = uri

            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows, row_links = [], []
            for ridx, row in enumerate(data[1:], start=1):
                if all(cell in (None, "") for cell in row):
                    continue
                first = next((str(c).strip() for c in row if c), "")
                if first.lower() == "total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc] + rest)
                row_links.append(row_links_map.get(ridx-1))

            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, row_links, tbl_total))

    # find Grand Total
    grand_total = None
    for text in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', text, re.I|re.S)
        if m:
            grand_total = m.group(1)
            break

# ─── Build PDF via ReportLab with LINK style ─────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11*inch, 17*inch)),
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
    elements.append(RLImage(io.BytesIO(logo), width=360, height=120))
except:
    pass
elements += [Spacer(1,12), Paragraph(proposal_title, title_style), Spacer(1,24)]

total_width = 17*inch - 96
for hdr, rows, row_links, tbl_total in tables_info:
    # build cell Paragraphs
    wrapped = [[Paragraph(col, header_style) for col in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            if cidx == 1 and row_links[ridx]:
                # embed <a href> so text is blue & underlined
                line.append(Paragraph(f'<a href="{row_links[ridx]}">{cell}</a>', body_style))
            else:
                line.append(Paragraph(str(cell), body_style))
        wrapped.append(line)
    # append row‐level total if present
    if tbl_total:
        lbl,val = re.split(r'\$\s*', tbl_total, 1)
        wrapped.append(
            [Paragraph(lbl, bl_style)] +
            [Paragraph("", body_style) for _ in hdr[2:-1]] +
            [Paragraph(f"${val.strip()}", br_style)]
        )

    # column widths
    col_widths = [
        0.45*total_width if i==1
        else (0.55*total_width)/(len(hdr)-1)
        for i in range(len(hdr))
    ]

    # collect table styles
    style_cmds = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]
    # add LINK annotation per cell
    for ridx, uri in enumerate(row_links):
        if uri:
            # header row is 0 → data starts at 1
            style_cmds.append(("LINK", (1, ridx+1), (1, ridx+1), uri))

    tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    elements.extend([tbl, Spacer(1,24)])

# grand total at bottom
if grand_total:
    hdr = tables_info[-1][0]
    gt_row = [Paragraph("Grand Total", bl_style)] + \
             [Paragraph("", body_style) for _ in hdr[2:-1]] + \
             [Paragraph(grand_total, br_style)]
    gt = LongTable([gt_row], colWidths=col_widths)
    gt.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"TOP")]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# ─── Build Word deliverable ───────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation  = WD_ORIENT.LANDSCAPE
sec.page_width   = Inches(17)
sec.page_height  = Inches(11)

# logo + title
try:
    p_logo = docx.add_paragraph(); r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo), width=Inches(4))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass

p_title = docx.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]
r.font.name = "DMSerif"
r.font.size = Pt(18)
docx.add_paragraph()

# tables in Word
for hdr, rows, row_links, tbl_total in tables_info:
    n = len(hdr)
    desc_w = 0.45 * 17
    oth_w  = (17 - desc_w) / (n - 1)

    tblW = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblW.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, col in enumerate(tblW.columns):
        col.width = Inches(desc_w if i == 1 else oth_w)

    # header row
    for i, col_name in enumerate(hdr):
        cell = tblW.rows[0].cells[i]
        tc    = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd   = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p     = cell.paragraphs[0]; p.text = ""
        run   = p.add_run(str(col_name))
        run.font.name = "DMSerif"; run.font.size = Pt(10); run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    # data rows
    for ridx, row in enumerate(rows):
        rc = tblW.add_row().cells
        for cidx, val in enumerate(row):
            p = rc[cidx].paragraphs[0]; p.text = ""
            if cidx == 1 and row_links[ridx]:
                # Word hyperlink
                add_hyperlink(p, row_links[ridx], str(val))
            else:
                run = p.add_run(str(val))
                run.font.name = "Barlow"; run.font.size = Pt(9)

    # table-level total row
    if tbl_total:
        label, amt = re.split(r'\$\s*', tbl_total, 1)
        amt = "$" + amt.strip()
        rc = tblW.add_row().cells
        values = [label] + [""]*(n-2) + [amt]
        for i, tv in enumerate(values):
            cell = rc[i]
            tc   = cell._tc; tcPr = tc.get_or_add_tcPr()
            shd  = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
            p    = cell.paragraphs[0]; p.text = ""
            run  = p.add_run(tv)
            run.font.name = "DMSerif"; run.font.size = Pt(10); run.bold = True
            if i == 0:
                p.alignment = WD_TABLE_ALIGNMENT.LEFT
            elif i == n-1:
                p.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else:
                p.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx.add_paragraph()

# grand total in Word
if grand_total:
    hdr = tables_info[-1][0]; n = len(hdr)
    tblG = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblG.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, tv in enumerate(["Grand Total"] + [""]*(n-2) + [grand_total]):
        cell = tblG.rows[0].cells[idx]
        tc   = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd  = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p    = cell.paragraphs[0]; p.text = ""
        run  = p.add_run(tv)
        run.font.name = "DMSerif"; run.font.size = Pt(10); run.bold = True
        if idx == 0:
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
        elif idx == n-1:
            p.alignment = WD_TABLE_ALIGNMENT.RIGHT
        else:
            p.alignment = WD_TABLE_ALIGNMENT.CENTER

docx.save(docx_buf)
docx_buf.seek(0)

# ─── Download buttons ─────────────────────────────────────────────────────────
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
