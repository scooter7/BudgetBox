# budgetbox.py
# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import fitz
import io
import re
import html
import requests
from PIL import Image
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from docx import Document
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))
    SERIF = "DMSerif"
    SANS  = "Barlow"
except:
    SERIF = "Times New Roman"
    SANS  = "Arial"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

def split_strategy(desc: str):
    lines = [l.strip() for l in desc.splitlines() if l.strip()]
    if not lines:
        return "", ""
    return lines[0], " ".join(lines[1:])

def add_hyperlink(p, url, text, font_name=None, font_size=None, bold=None):
    part = p.part
    rid = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    styles = p.part.document.styles
    if "Hyperlink" not in styles:
        stl = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        stl.font.color.rgb = RGBColor(0x00,0x33,0x99)
        stl.font.underline = True
    s = OxmlElement("w:rStyle")
    s.set(qn("w:val"), "Hyperlink")
    rPr.append(s)
    if font_name:
        rf = OxmlElement("w:rFonts")
        rf.set(qn("w:ascii"), font_name)
        rf.set(qn("w:hAnsi"), font_name)
        rPr.append(rf)
    if font_size:
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), str(int(font_size*2)))
        szCs = OxmlElement("w:szCs")
        szCs.set(qn("w:val"), str(int(font_size*2)))
        rPr.extend([sz, szCs])
    if bold:
        rPr.append(OxmlElement("w:b"))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    link.append(r)
    p._p.append(link)
    return docx.text.run.Run(r, p)

def extract_all():
    mz = fitz.open(stream=pdf_bytes, filetype="pdf")
    linkmaps = []
    for pg in mz:
        ann = pg.annots() or []
        linkmaps.append([(a.rect, a.uri) for a in ann if a.type[0]==1 and a.uri])
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texts = [pg.extract_text(x_tolerance=1,y_tolerance=1) or "" for pg in pdf.pages]
        first = texts[0].splitlines() if texts else []
        title = next((l for l in first if "proposal" in l.lower()), first[0] if first else "Untitled Proposal")
        tables = []
        used_totals = set()
        def find_total(pi):
            for ln in texts[pi].splitlines():
                if re.search(r'\b(?!grand)total\b.*?\$\s*\d', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None
        for pi, p in enumerate(pdf.pages):
            for tbl in p.extract_tables():
                if len(tbl)<2:
                    continue
                hdr = [str(c).strip() if c else "" for c in tbl[0]]
                if "Start Date" not in " ".join(hdr):
                    continue
                di = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
                if di is None:
                    di = 1
                new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=di]
                rows, uris, tot = [], [], None
                bbox = p.bbox
                n = len(tbl)
                band = (bbox[3]-bbox[1])/n
                lm = linkmaps[pi]
                rowlink = {}
                for rect, uri in lm:
                    mid = (rect.y0+rect.y1)/2
                    if bbox[1]<mid<bbox[3]:
                        rid = int((mid-bbox[1])//band)-1
                        if 0<=rid<len(tbl)-1:
                            rowlink[rid] = uri
                for ridx, row in enumerate(tbl[1:]):
                    if all(not (str(c).strip()) for c in row):
                        continue
                    first = str(row[0]).lower()
                    if "total" in first and any("$" in str(c) for c in row):
                        if tot is None:
                            tot = row
                        continue
                    strat, desc = split_strategy(str(row[di] or ""))
                    rest = [row[i] for i in range(len(row)) if i!=di]
                    rows.append([strat,desc]+rest)
                    uris.append(rowlink.get(ridx))
                if tot is None:
                    tot = find_total(pi)
                if rows:
                    tables.append((new_hdr, rows, uris, tot))
        gt = None
        for tx in reversed(texts):
            m = re.search(r'Grand Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I)
            if m:
                gt = m.group(1).strip()
                break
        return title, tables, gt

try:
    proposal_title, tables_info, grand_total = extract_all()
except Exception as e:
    st.error("PDF parse error: " + str(e))
    st.stop()

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11*inch,17*inch)),
    leftMargin=0.5*inch,
    rightMargin=0.5*inch,
    topMargin=0.5*inch,
    bottomMargin=0.5*inch
)
styles = {
    "title": ParagraphStyle("T", fontName=SERIF, fontSize=18, alignment=TA_CENTER, spaceAfter=12),
    "hdr":   ParagraphStyle("H", fontName=SERIF, fontSize=10, alignment=TA_CENTER),
    "body":  ParagraphStyle("B", fontName=SANS,  fontSize=9,  alignment=TA_LEFT, leading=11),
    "left":  ParagraphStyle("L", fontName=SERIF, fontSize=10, alignment=TA_LEFT),
    "right": ParagraphStyle("R", fontName=SERIF, fontSize=10, alignment=TA_RIGHT),
}
elements = []
try:
    logo = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png", timeout=5).content
    img = Image.open(io.BytesIO(logo))
    r = img.height/img.width
    w = min(5*inch, doc.width)
    h = w*r
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except:
    pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), styles["title"]), Spacer(1,24)]
W = doc.width
for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    di = hdr.index("Description")
    dw = W*0.45
    ow = (W-dw)/(n-1) if n>1 else W
    cw = [dw if i==di else ow for i in range(n)]
    data = [[Paragraph(html.escape(h), styles["hdr"]) for h in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for ci, cell in enumerate(row):
            txt = html.escape(cell)
            if ci==di and uris[ridx]:
                p = Paragraph(f"{txt} <link href='{uris[ridx]}' color='blue'>- link</link>", styles["body"])
            else:
                p = Paragraph(txt, styles["body"])
            line.append(p)
        data.append(line)
    if tot:
        lbl, val = "Total", ""
        if isinstance(tot, list):
            lbl = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m:
                lbl, val = m.group(1).strip(), m.group(2)
        row = [Paragraph(lbl, styles["left"])] + [Spacer(1,0)]*(n-2) + [Paragraph(val, styles["right"])]
        data.append(row)
    tbl = LongTable(data, colWidths=cw, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
        *(
            [
                ("SPAN",(0,-1),(-2,-1)),
                ("ALIGN",(0,-1),(-2,-1),"LEFT"),
                ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),
                ("VALIGN",(0,-1),(-1,-1),"MIDDLE"),
            ] if tot else []
        )
    ]))
    elements += [tbl, Spacer(1,24)]
if grand_total and tables_info:
    hdr0 = tables_info[-1][0]
    n = len(hdr0)
    di = hdr0.index("Description")
    dw = W*0.45
    ow = (W-dw)/(n-1) if n>1 else W
    cw = [dw if i==di else ow for i in range(n)]
    row = [Paragraph("Grand Total", styles["left"])] + [Spacer(1,0)]*(n-2) + [Paragraph(html.escape(grand_total), styles["right"])]
    gt = LongTable([row], colWidths=cw)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),
        ("ALIGN",(-1,0),(-1,0),"RIGHT"),
    ]))
    elements.append(gt)
doc.build(elements)
pdf_data = pdf_buf.getvalue()

docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(11)
sec.page_height = Inches(17)
for m in ("left_margin","right_margin","top_margin","bottom_margin"):
    setattr(sec, m, Inches(0.5))
try:
    p = docx.add_paragraph()
    r = p.add_run()
    img = Image.open(io.BytesIO(logo))
    ratio = img.height/img.width
    w = Inches(5)
    r.add_picture(io.BytesIO(logo), width=w)
    p.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass
p = docx.add_paragraph()
p.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p.add_run(proposal_title)
r.font.name = SERIF
r.font.size = Pt(18)
r.bold = True
docx.add_paragraph()
TW = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches
for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    di = hdr.index("Description")
    dw = 0.45*TW
    ow = (TW-dw)/(n-1) if n>1 else TW
    tbl = docx.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit = False
    tbl.autofit = False
    tblPr = tbl._element.xpath("./w:tblPr")[0]
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"),"5000"); tblW.set(qn("w:type"),"pct")
    ex = tblPr.xpath("./w:tblW")
    if ex:
        tblPr.remove(ex[0])
    tblPr.append(tblW)
    for i,col in enumerate(tbl.columns):
        col.width = Inches(dw if i==di else ow)
    row0 = tbl.rows[0].cells
    for i,name in enumerate(hdr):
        c = row0[i]
        tc = c._tc; tcPr = tc.get_or_add_tcPr()
        sh = OxmlElement("w:shd"); sh.set(qn("w:fill"),"F2F2F2"); tcPr.append(sh)
        p = c.paragraphs[0]; p.text = ""
        run = p.add_run(name); run.font.name = SERIF; run.font.size = Pt(10); run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER; c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        rc = tbl.add_row().cells
        for ci,cell in enumerate(row):
            c = rc[ci]
            p = c.paragraphs[0]; p.text = ""
            run = p.add_run(str(cell)); run.font.name = SANS; run.font.size = Pt(9)
            if ci==di and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p, uris[ridx], "- link", font_name=SANS, font_size=9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT; c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        tr = tbl.add_row().cells
        lbl, amt = "Total", ""
        if isinstance(tot,list):
            lbl = tot[0] or "Total"
            amt = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', str(tot))
            if m:
                lbl, amt = m.group(1).strip(), m.group(2)
        lc = tr[0]
        if n>1:
            lc.merge(tr[n-2])
        p = lc.paragraphs[0]; p.text = ""
        r = p.add_run(lbl); r.font.name=SERIF; r.font.size=Pt(10); r.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac = tr[n-1]
        p2 = ac.paragraphs[0]; p2.text = ""
        r2 = p2.add_run(amt); r2.font.name=SERIF; r2.font.size=Pt(10); r2.bold=True
        p2.alignment = WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx.add_paragraph()
if grand_total and tables_info:
    last_hdr = tables_info[-1][0]; n = len(last_hdr); di = last_hdr.index("Description")
    dw = 0.45*TW; ow = (TW-dw)/(n-1) if n>1 else TW
    tblg = docx.add_table(rows=1,cols=n,style="Table Grid"); tblg.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblg.allow_autofit=False; tblg.autofit=False
    tblPr = tblg._element.xpath("./w:tblPr")[0]
    tblW = OxmlElement("w:tblW"); tblW.set(qn("w:w"),"5000"); tblW.set(qn("w:type"),"pct")
    ex = tblPr.xpath("./w:tblW")
    if ex: tblPr.remove(ex[0])
    tblPr.append(tblW)
    for i,col in enumerate(tblg.columns):
        col.width = Inches(dw if i==di else ow)
    cells = tblg.rows[0].cells
    lc = cells[0]
    if n>1: lc.merge(cells[n-2])
    tc = lc._tc; tcPr = tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd"); sh.set(qn("w:fill"),"E0E0E0"); tcPr.append(sh)
    p = lc.paragraphs[0]; p.text = ""
    r = p.add_run("Grand Total"); r.font.name=SERIF; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac = cells[n-1]
    tc2 = ac._tc; tcPr2 = tc2.get_or_add_tcPr()
    sh2=OxmlElement("w:shd"); sh2.set(qn("w:fill"),"E0E0E0"); tcPr2.append(sh2)
    p2=ac.paragraphs[0]; p2.text=""
    r2=p2.add_run(grand_total); r2.font.name=SERIF; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
try:
    docx.save(docx_buf)
    docx_data = docx_buf.getvalue()
except:
    docx_data = None

c1, c2 = st.columns(2)
if pdf_data:
    c1.download_button("📥 Download PDF", pdf_data, "deliverable.pdf", "application/pdf")
else:
    c1.error("PDF generation failed")
if docx_data:
    c2.download_button("📥 Download DOCX", docx_data, "deliverable.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    c2.error("DOCX generation failed")
