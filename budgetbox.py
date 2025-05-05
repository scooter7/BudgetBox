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

# Register fonts
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

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    rid = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        stl = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        stl.font.color.rgb = RGBColor(0x00, 0x33, 0x99)
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
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(font_size*2)))
        szCs = OxmlElement("w:szCs"); szCs.set(qn("w:val"), str(int(font_size*2)))
        rPr.extend([sz, szCs])
    if bold:
        rPr.append(OxmlElement("w:b"))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    link.append(r)
    paragraph._p.append(link)
    return docx.text.run.Run(r, paragraph)

def extract_all():
    # build link maps via PyMuPDF
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
            for tbl in p.find_tables():
                data = tbl.extract()
                if len(data)<2: continue
                hdr = [str(c).strip() if c else "" for c in data[0]]
                if "Start Date" not in " ".join(hdr):
                    continue
                di = next((i for i,h in enumerate(hdr) if "description" in h.lower()), 1)
                new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=di]
                rows, uris, tot = [], [], None
                bbox = p.bbox
                n = len(data)
                band = (bbox[3]-bbox[1])/n
                lm = linkmaps[pi]
                rowlink = {}
                for rect, uri in lm:
                    mid = (rect.y0+rect.y1)/2
                    if bbox[1]<mid<bbox[3]:
                        rid = int((mid-bbox[1])//band)-1
                        if 0<=rid<len(data)-1:
                            rowlink[rid] = uri
                for ridx, row in enumerate(data[1:]):
                    if all(not str(c).strip() for c in row): continue
                    if ("total" in str(row[0]).lower() and any("$" in str(c) for c in row)):
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

# ---- PDF build ----
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((11*inch,17*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch,
    topMargin=0.5*inch, bottomMargin=0.5*inch
)
styles = {
    "title": ParagraphStyle("T", fontName=SERIF, fontSize=18, alignment=TA_CENTER, spaceAfter=12),
    "hdr":   ParagraphStyle("H", fontName=SERIF, fontSize=10, alignment=TA_CENTER),
    "body":  ParagraphStyle("B", fontName=SANS,  fontSize=9,  alignment=TA_LEFT, leading=11),
    "left":  ParagraphStyle("L", fontName=SERIF, fontSize=10, alignment=TA_LEFT),
    "right": ParagraphStyle("R", fontName=SERIF, fontSize=10, alignment=TA_RIGHT),
}
elements = []
# logo + title
try:
    logo = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png", timeout=5).content
    img = Image.open(io.BytesIO(logo))
    r = img.height/img.width
    w = min(5*inch, doc.width)
    h = w*r
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except:
    pass
safe_title = proposal_title or ""
elements += [
    Spacer(1,12),
    Paragraph(html.escape(safe_title), styles["title"]),
    Spacer(1,24),
]
W = doc.width
for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    di = hdr.index("Description")
    dw = W*0.45
    ow = (W-dw)/(n-1) if n>1 else W
    cw = [dw if i==di else ow for i in range(n)]
    data = [[Paragraph(html.escape(h), styles["hdr"]) for h in hdr]]
    for ridx, row in enumerate(rows):
        line=[]
        for ci,cell in enumerate(row):
            txt=html.escape(cell)
            if ci==di and uris[ridx]:
                p=Paragraph(f"{txt} <link href='{uris[ridx]}' color='blue'>- link</link>", styles["body"])
            else:
                p=Paragraph(txt, styles["body"])
            line.append(p)
        data.append(line)
    if tot:
        lbl, val = "Total",""
        if isinstance(tot,list):
            lbl = tot[0] or "Total"
            val = next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl, val = m.group(1).strip(), m.group(2)
        data.append([Paragraph(lbl, styles["left"])] + [Spacer(1,0)]*(n-2) + [Paragraph(val,styles["right"])])
    tbl=LongTable(data, colWidths=cw, repeatRows=1)
    cmds=[
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]
    if tot:
        cmds += [
            ("SPAN",(0,-1),(-2,-1)),
            ("ALIGN",(0,-1),(-2,-1),"LEFT"),
            ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),
            ("VALIGN",(0,-1),(-1,-1),"MIDDLE"),
        ]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1,24)]
if grand_total and tables_info:
    hdr0 = tables_info[-1][0]; n=len(hdr0); di=hdr0.index("Description")
    dw=W*0.45; ow=(W-dw)/(n-1) if n>1 else W
    cw=[dw if i==di else ow for i in range(n)]
    row=[Paragraph("Grand Total",styles["left"])] + [Spacer(1,0)]*(n-2) + [Paragraph(html.escape(grand_total),styles["right"])]
    gt=LongTable([row],colWidths=cw)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),
        ("ALIGN",(-1,0),(-1,0),"RIGHT"),
    ]))
    elements.append(gt)
doc.build(elements)
pdf_data = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf = pdf_buf
