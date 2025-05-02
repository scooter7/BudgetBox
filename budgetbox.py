import streamlit as st
import pdfplumber
import io
import requests
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

# ─── Helpers ───────────────────────────────────────────────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

def add_hyperlink(paragraph, url, text):
    """
    Inserts a Word-field hyperlink (w:fldSimple) so that
    Word always shows a clickable, blue-underlined link.
    """
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), f'HYPERLINK "{url}"')
    run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    # blue + underline
    c = OxmlElement("w:color"); c.set(qn("w:val"), "0000FF"); rPr.append(c)
    u = OxmlElement("w:u");     u.set(qn("w:val"), "single"); rPr.append(u)
    run.append(rPr)
    t = OxmlElement("w:t");     t.text = text; run.append(t)
    fld.append(run)
    paragraph._p.append(fld)

# ─── Extract tables + links from source PDF ───────────────────────────────────
tables_info = []
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

    for pg_i, page in enumerate(pdf.pages):
        # pull link annotations on this page
        annots = []
        for a in page.annots() or []:
            if a.type[0] == 1 and a.uri:
                annots.append((a.rect, a.uri))

        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue
            hdr = data[0]
            # find the Description column index
            desc_i = next((i for i, h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            # map each annotation to a row
            x0, y0, x1, y1 = tbl.bbox
            band_h = (y1 - y0) / len(data)
            row_links = {}
            for rect, uri in annots:
                midy = (rect.y0 + rect.y1) / 2
                if y0 <= midy <= y1:
                    ridx = int((midy - y0) // band_h)
                    if 1 <= ridx < len(data):
                        row_links[ridx - 1] = uri

            new_hdr = ["Strategy", "Description"] + [h for i, h in enumerate(hdr) if i != desc_i and h]
            rows = []
            links = []
            for ridx, row in enumerate(data[1:], start=1):
                if all(not str(c).strip() for c in row if c):
                    continue
                first = next((str(c).strip() for c in row if c), "")
                if first.lower() == "total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i, h in enumerate(hdr) if i != desc_i and h]
                rows.append([strat, desc] + rest)
                links.append(row_links.get(ridx - 1))

            tbl_total = find_total(pg_i)
            tables_info.append((new_hdr, rows, links, tbl_total))

    # Grand total
    grand_total = None
    for txt in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', txt, re.I | re.S)
        if m:
            grand_total = m.group(1)
            break

# ─── Build PDF with ReportLab + LINKURL ────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11 * inch, 17 * inch)),
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
elements += [Spacer(1, 12), Paragraph(proposal_title, title_style), Spacer(1, 24)]

total_w = 17 * inch - 96
for hdr, rows, row_links, tbl_total in tables_info:
    wrapped = [[Paragraph(col, header_style) for col in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            if cidx == 1 and row_links[ridx]:
                # use <a href> — ReportLab recognizes this + LINKURL
                line.append(Paragraph(f'<a href="{row_links[ridx]}">{cell}</a>', body_style))
            else:
                line.append(Paragraph(str(cell), body_style))
        wrapped.append(line)

    if tbl_total:
        lbl, val = re.split(r'\$\s*', tbl_total, 1)
        wrapped.append(
            [Paragraph(lbl, bl_style)] +
            [Paragraph("", body_style) for _ in hdr[2:-1]] +
            [Paragraph(f"${val.strip()}", br_style)]
        )

    colws = [0.45 * total_w if i == 1 else (0.55 * total_w) / (len(hdr) - 1) for i in range(len(hdr))]
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("GRID",       (0, 0), (-1, -1),   0.25, colors.grey),
        ("VALIGN",     (0, 0), (-1, 0),   "MIDDLE"),
        ("VALIGN",     (0, 1), (-1, -1),"TOP"),
    ]
    # add LINKURL annotations
    for ridx, uri in enumerate(row_links):
        if uri:
            # cell at (col=1,row=ridx+1) since header is row 0
            style_cmds.append(("LINKURL", (1, ridx+1), (1, ridx+1), uri))

    tbl = LongTable(wrapped, colWidths=colws, repeatRows=1)
    tbl.setStyle(TableStyle(style_cmds))
    elements += [tbl, Spacer(1, 24)]

# grand total row
if grand_total:
    hdr = tables_info[-1][0]
    gt_row = [Paragraph("Grand Total", bl_style)] + \
             [Paragraph("", body_style) for _ in hdr[2:-1]] + \
             [Paragraph(grand_total, br_style)]
    gt = LongTable([gt_row], colWidths=colws)
    gt.setStyle(TableStyle([
        ("GRID",   (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN",(0, 0), (-1, -1),"TOP"),
    ]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# ─── Build Word deliverable ───────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = Inches(17), Inches(11)

# logo + title
try:
    p_logo = docx.add_paragraph(); r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo), width=Inches(4))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass

p_title = docx.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]; r.font.name, r.font.size = "DMSerif", Pt(18)
docx.add_paragraph()

for hdr, rows, row_links, tbl_total in tables_info:
    n = len(hdr)
    desc_w = 0.45 * 17
    oth_w  = (17 - desc_w) / (n - 1)

    tblW = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblW.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, col in enumerate(tblW.columns):
        col.width = Inches(desc_w if i == 1 else oth_w)

    # header
    for i, col_name in enumerate(hdr):
        cell = tblW.rows[0].cells[i]
        tc, tcPr = cell._tc, cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(col_name)
        run.font.name, run.font.size, run.bold = "DMSerif", Pt(10), True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    # data rows
    for ridx, row in enumerate(rows):
        rc = tblW.add_row().cells
        for cidx, val in enumerate(row):
            p = rc[cidx].paragraphs[0]; p.text = ""
            if cidx == 1 and row_links[ridx]:
                add_hyperlink(p, row_links[ridx], str(val))
            else:
                run = p.add_run(str(val))
                run.font.name, run.font.size = "Barlow", Pt(9)

    # table-level total
    if tbl_total:
        label, amt = re.split(r'\$\s*', tbl_total, 1)
        amt = "$" + amt.strip()
        rc = tblW.add_row().cells
        for i, tv in enumerate([label] + [""]*(n-2) + [amt]):
            cell = rc[i]
            tc, tcPr = cell._tc, cell._tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
            p = cell.paragraphs[0]; p.text = ""
            run = p.add_run(tv)
            run.font.name, run.font.size, run.bold = "DMSerif", Pt(10), True
            p.alignment = (
                WD_TABLE_ALIGNMENT.LEFT   if i == 0
                else WD_TABLE_ALIGNMENT.RIGHT if i == n-1
                else WD_TABLE_ALIGNMENT.CENTER
            )
    docx.add_paragraph()

# grand total in Word
if grand_total:
    hdr, n = tables_info[-1][0], len(tables_info[-1][0])
    tblG = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblG.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, tv in enumerate(["Grand Total"] + [""]*(n-2) + [grand_total]):
        cell = tblG.rows[0].cells[idx]
        tc, tcPr = cell._tc, cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(tv)
        run.font.name, run.font.size, run.bold = "DMSerif", Pt(10), True
        p.alignment = (
            WD_TABLE_ALIGNMENT.LEFT   if idx == 0
            else WD_TABLE_ALIGNMENT.RIGHT if idx == n-1
            else WD_TABLE_ALIGNMENT.CENTER
        )

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
