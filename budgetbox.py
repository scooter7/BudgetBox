# -*- coding: utf-8 -*-
import io, re, html, streamlit as st, camelot, pdfplumber, requests
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

# ── Fonts ─────────────────────────────────────────────────────────────────────
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    SERIF = "DMSerif"; SANS = "Barlow"
except:
    SERIF = "Times New Roman"; SANS = "Arial"

# ── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload your PDF and get back a landscape-formatted PDF + DOCX with 8-column tables.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ── Helpers ───────────────────────────────────────────────────────────────────
def split_cell_text(raw):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines: return "", ""
    return lines[0], " ".join(lines[1:]).replace("\n"," ").strip()

def add_hyperlink(p, url, text, font_name=None, font_size=None, bold=None):
    part = p.part
    rid  = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink"); link.set(qn("r:id"), rid)
    r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
    styles = p.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05,0x63,0xC1); style.font.underline=True
    se = OxmlElement("w:rStyle"); se.set(qn("w:val"), "Hyperlink"); rPr.append(se)
    if font_name:
        rf = OxmlElement("w:rFonts"); rf.set(qn("w:ascii"), font_name); rf.set(qn("w:hAnsi"), font_name); rPr.append(rf)
    if font_size:
        sz  = OxmlElement("w:sz");  sz.set(qn("w:val"), str(int(font_size*2)))
        szc = OxmlElement("w:szCs"); szc.set(qn("w:val"), str(int(font_size*2)))
        rPr.extend([sz,szc])
    if bold: rPr.append(OxmlElement("w:b"))
    r.append(rPr)
    t = OxmlElement("w:t"); t.set(qn("xml:space"), "preserve"); t.text = text
    r.append(t); link.append(r); p._p.append(link)

# ── Extract Table 1 via Camelot Lattice ────────────────────────────────────────
first_table = None
try:
    # If your table area is different, change these coordinates:
    # (x1,y1,x2,y2) in PDF coordinate space (origin bottom-left)
    # I measured them in Adobe Reader: this box tightly encloses the 8-col grid on page 1.
    table_areas = ["1,720,595,100"]  # LEFT=1pt, TOP=720pt, RIGHT=595pt, BOTTOM=100pt

    ct = camelot.read_pdf(
        filepath_or_buffer=io.BytesIO(pdf_bytes),
        pages="1",
        flavor="lattice",
        table_areas=table_areas,
        strip_text="\n"
    )
    if ct and ct[0].df.shape[1] >= 8:
        df = ct[0].df
        raw = df.values.tolist()
        # Merge the two header rows into one:
        hdr1, hdr2 = raw[0], raw[1]
        tail = [h.strip() for h in hdr2[2:] if h.strip()]
        header = [hdr1[0].strip(), hdr1[1].strip()] + tail
        rows = []
        for r in raw[2:]:
            cells = [c.strip() for c in r]
            vals  = cells[0:2] + cells[2:2+len(tail)]
            if any(vals):
                rows.append(vals)
        if rows:
            first_table = [header] + rows
except Exception:
    first_table = None

# ── Fallback: force pdfplumber lattice on page 1 only if Camelot really dies ──
if not first_table:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as p0:
        pg = p0.pages[0]
        tbl = pg.extract_table({
            "vertical_strategy":"lines",
            "horizontal_strategy":"lines"
        })
    if tbl and len(tbl[0]) >= 8:
        first_table = tbl

# ── Now extract ALL pages ────────────────────────────────────────────────────
tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts = [p.extract_text(x_tolerance=1,y_tolerance=1) or "" for p in pdf.pages]
    # Title:
    lines0 = texts[0].splitlines()
    pt = next((l for l in lines0 if "proposal" in l.lower()), None)
    proposal_title = pt.strip() if pt else lines0[0].strip()

    used = set()
    def find_total(pi):
        if pi >= len(texts): return None
        for ln in texts[pi].splitlines():
            if re.search(r'\b(?!grand )total\b.*?\$\s*[\d,]+', ln, re.I) and ln not in used:
                used.add(ln); return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        if pi == 0 and first_table:
            hdr, rows = first_table[0], first_table[1:]
            links = []
            tbl_links = {}
            source = "camelot"
        else:
            source = "plumber"
            hdr, rows, tbl_links = None, [], {}
            links = page.hyperlinks
            for t in page.find_tables():
                data = t.extract(x_tolerance=1,y_tolerance=1)
                if data and len(data[0])>=2:
                    hdr = [c.strip() for c in data[0]]
                    for ridx, r in enumerate(data[1:], start=1):
                        if "total" in (r[0] or "").lower() and any("$" in c for c in r):
                            continue
                        row = [c.strip() for c in r]
                        rows.append(row)
                    # collect link boxes
                    for ridx,row_obj in enumerate(t.rows, start=0):
                        if ridx==0: continue
                        cell = row_obj.cells[1]
                        x0,top,x1,bottom = cell
                        for link in links:
                            if all(k in link for k in ("x0","x1","top","bottom","uri")):
                                if not (link["x1"]<x0 or link["x0"]>x1 or link["bottom"]<top or link["top"]>bottom):
                                    tbl_links[ridx-1] = link["uri"]
                                    break
                    break

        if not hdr or len(rows)==0:
            continue

        desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), 1)
        new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i not in (0,desc_i)]
        processed = []
        link_list = []
        table_total = None

        for ridx, r in enumerate(rows):
            first = (r[0] or "").lower()
            if ("total" in first) and any("$" in c for c in r):
                if not table_total: table_total = r
                continue
            strat, desc = split_cell_text(r[desc_i] if desc_i < len(r) else "")
            rest = [r[i] for i in range(len(r)) if i not in (0,desc_i)]
            processed.append([strat,desc]+rest)
            link_list.append(tbl_links.get(ridx))

        if not table_total:
            table_total = find_total(pi)

        if processed:
            tables_info.append((new_hdr, processed, link_list, table_total))

    # Grand Total
    for t in reversed(texts):
        m = re.search(r'Grand Total.*?(\$\s*[\d,]+\.\d{2})', t, re.I)
        if m:
            grand_total = m.group(1); break

# ── Build the PDF deliverable ─────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf,
    pagesize=landscape((17*inch,11*inch)),
    leftMargin=0.5*inch,
    rightMargin=0.5*inch,
    topMargin=0.5*inch,
    bottomMargin=0.5*inch
)
ts = ParagraphStyle("Title", fontName=SERIF, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
hs = ParagraphStyle("Header", fontName=SERIF, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
bs = ParagraphStyle("Body", fontName=SANS, fontSize=9, alignment=TA_LEFT, leading=11)
bl = ParagraphStyle("BL", fontName=SERIF, fontSize=10, alignment=TA_LEFT, spaceBefore=6)
br = ParagraphStyle("BR", fontName=SERIF, fontSize=10, alignment=TA_RIGHT, spaceBefore=6)

elements = []
logo = None
try:
    url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    r = requests.get(url, timeout=10); r.raise_for_status()
    logo = r.content
    img = Image.open(io.BytesIO(logo))
    ratio = img.height/img.width
    w = min(5*inch, doc.width)
    h = w * ratio
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except:
    pass

elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
tw = doc.width

for hdr, rows, links, tot in tables_info:
    n = len(hdr)
    di = hdr.index("Description") if "Description" in hdr else 1
    w_desc = tw * 0.45
    w_other = (tw - w_desc)/(n-1) if n>1 else tw
    widths = [w_desc if i==di else w_other for i in range(n)]

    data = [[Paragraph(html.escape(c), hs) for c in hdr]]
    for i, row in enumerate(rows):
        line = []
        for j, cell in enumerate(row):
            txt = html.escape(cell)
            if j==di and links[i]:
                para = Paragraph(f"{txt} <link href='{html.escape(links[i])}' color='blue'>- link</link>", bs)
            else:
                para = Paragraph(txt, bs)
            line.append(para)
        data.append(line)

    if tot:
        label, val = "Total", ""
        if isinstance(tot, list):
            label = tot[0] or "Total"
            val   = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m:
                label, val = m.group(1).strip(), m.group(2)
        tr = [Paragraph(label, bl)] + [Paragraph("", bs)]*(n-2) + [Paragraph(val, br)]
        data.append(tr)

    tbl = LongTable(data, colWidths=widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP")
    ]
    if tot and n>1:
        style_cmds += [
            ("SPAN",(0,-1),(-2,-1)),
            ("ALIGN",(0,-1),(-2,-1),"LEFT"),
            ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),
            ("VALIGN",(0,-1),(-1,-1),"MIDDLE")
        ]
    tbl.setStyle(TableStyle(style_cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    hdr_last = tables_info[-1][0]
    n = len(hdr_last)
    di = hdr_last.index("Description") if "Description" in hdr_last else 1
    w_desc = tw * 0.45
    w_other = (tw - w_desc)/(n-1) if n>1 else tw
    widths = [w_desc if i==di else w_other for i in range(n)]

    row = [Paragraph("Grand Total", bl)] + [Paragraph("", bs)]*(n-2) + [Paragraph(html.escape(grand_total), br)]
    gt  = LongTable([row], colWidths=widths)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),
        ("ALIGN",(-1,0),(-1,0),"RIGHT")
    ]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# ── Build the DOCX deliverable ────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Inches(17)
sec.page_height= Inches(11)
sec.left_margin = Inches(0.5)
sec.right_margin= Inches(0.5)
sec.top_margin   = Inches(0.5)
sec.bottom_margin= Inches(0.5)

if logo:
    try:
        p = docx_doc.add_paragraph()
        r = p.add_run()
        img = Image.open(io.BytesIO(logo)); ratio = img.height/img.width
        w = 5; h = w * ratio
        r.add_picture(io.BytesIO(logo), width=Inches(w))
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    except:
        pass

pt = docx_doc.add_paragraph(); pt.alignment = WD_TABLE_ALIGNMENT.CENTER
rt = pt.add_run(proposal_title); rt.font.name = SERIF; rt.font.size = Pt(18); rt.bold=True
docx_doc.add_paragraph()

TW = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, links, tot in tables_info:
    n = len(hdr)
    di = hdr.index("Description") if "Description" in hdr else 1
    dw = 0.45 * TW
    ow = (TW - dw)/(n-1) if n>1 else TW

    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False; tbl.autofit = False

    # set 100% width
    tblPr_list = tbl._element.xpath("./w:tblPr")
    tblPr = tblPr_list and tblPr_list[0] or OxmlElement("w:tblPr")
    if not tblPr_list: tbl._element.insert(0, tblPr)
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"), "5000"); tblW.set(qn("w:type"), "pct")
    existing = tblPr.xpath("./w:tblW")
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)

    for i,col in enumerate(tbl.columns):
        col.width = Inches(dw if i==di else ow)

    # header
    for i, name in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "F2F2F2"); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text=""
        r = p.add_run(name); r.font.name=SERIF; r.font.size=Pt(10); r.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # rows
    for ridx, row in enumerate(rows):
        cells = tbl.add_row().cells
        for j, val in enumerate(row):
            c = cells[j]
            p = c.paragraphs[0]; p.text=""
            run = p.add_run(str(val)); run.font.name=SANS; run.font.size=Pt(9)
            if j==di and links[ridx]:
                p.add_run(" ")
                add_hyperlink(p, links[ridx], "- link", font_name=SANS, font_size=9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    if tot:
        tr = tbl.add_row().cells
        lbl, amt = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"; amt = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl, amt = m.group(1).strip(), m.group(2)
        lc = tr[0]
        if n>1: lc.merge(tr[n-2])
        p  = lc.paragraphs[0]; p.text=""
        r1 = p.add_run(lbl); r1.font.name=SERIF; r1.font.size=Pt(10); r1.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac = tr[n-1]
        p2 = ac.paragraphs[0]; p2.text=""
        r2 = p2.add_run(amt); r2.font.name=SERIF; r2.font.size=Pt(10); r2.bold=True
        p2.alignment = WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr = tables_info[-1][0]; n = len(last_hdr)
    di = last_hdr.index("Description") if "Description" in last_hdr else 1
    dw = 0.45 * TW
    ow = (TW - dw)/(n-1) if n>1 else TW

    tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment=WD_TABLE_ALIGNMENT.CENTER; tblg.allow_autofit=False; tblg.autofit=False

    tblPr_list = tblg._element.xpath("./w:tblPr")
    tblPr = tblPr_list and tblPr_list[0] or OxmlElement("w:tblPr")
    if not tblPr_list: tblg._element.insert(0,tblPr)
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"),"5000"); tblW.set(qn("w:type"),"pct")
    existing = tblPr.xpath("./w:tblW")
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)

    for i,col in enumerate(tblg.columns):
        col.width = Inches(dw if i==di else ow)

    # label cell
    lc = tblg.rows[0].cells[0]
    if n>1: lc.merge(tblg.rows[0].cells[n-2])
    tc = lc._tc; tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"),"E0E0E0"); tcPr.append(shd)
    p = lc.paragraphs[0]; p.text=""
    r = p.add_run("Grand Total"); r.font.name=SERIF; r.font.size=Pt(10); r.bold=True
    p.alignment = WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    ac = tblg.rows[0].cells[n-1]
    p2 = ac.paragraphs[0]; p2.text=""
    r2 = p2.add_run(grand_total); r2.font.name=SERIF; r2.font.size=Pt(10); r2.bold=True
    p2.alignment = WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1, c2 = st.columns(2)
if pdf_buf:
    c1.download_button("📥 Download deliverable PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf", use_container_width=True)
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button("📥 Download deliverable DOCX", data=docx_buf, file_name="proposal_deliverable.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    c2.error("Word document generation failed.")
