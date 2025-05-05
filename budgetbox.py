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
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import re
import html

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
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded: st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines: return "", ""
    desc = " ".join(lines[1:])
    return lines[0], re.sub(r'\s+', ' ', desc).strip()

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyp = OxmlElement('w:hyperlink'); hyp.set(qn('r:id'), r_id)
    r = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05,0x63,0xC1); style.font.underline = True
    se = OxmlElement('w:rStyle'); se.set(qn('w:val'), 'Hyperlink'); rPr.append(se)
    if font_name:
        rf = OxmlElement('w:rFonts'); rf.set(qn('w:ascii'), font_name); rf.set(qn('w:hAnsi'), font_name); rPr.append(rf)
    if font_size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(font_size*2))); szc = OxmlElement('w:szCs'); szc.set(qn('w:val'), str(int(font_size*2))); rPr.extend([sz, szc])
    if bold: rPr.append(OxmlElement('w:b'))
    r.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text = text
    r.append(t); hyp.append(r); paragraph._p.append(hyp)
    return docx.text.run.Run(r, paragraph)

def reconstruct_table_from_words(page, bbox):
    x0,y0,x1,y1 = bbox
    words = [w for w in page.extract_words(x_tolerance=1, y_tolerance=1) 
             if w['x0']>=x0 and w['x1']<=x1 and w['top']>=y0 and w['bottom']<=y1]
    if not words: return None
    ys = sorted({(w['top']+w['bottom'])/2 for w in words})
    rows = []
    for y in ys:
        for r in rows:
            if abs(r[0]-y)<3:
                r.append(y); break
        else:
            rows.append([y])
    row_centers = sorted([sum(r)/len(r) for r in rows], reverse=True)
    header_words = [w for w in words if abs(((w['top']+w['bottom'])/2)-row_centers[0])<3]
    header_words.sort(key=lambda w: w['x0'])
    col_centers = [ (w['x0']+w['x1'])/2 for w in header_words ]
    table = [[[] for _ in col_centers] for _ in row_centers]
    for w in words:
        yc = (w['top']+w['bottom'])/2
        ri = min(range(len(row_centers)), key=lambda i: abs(row_centers[i]-yc))
        xc = (w['x0']+w['x1'])/2
        ci = min(range(len(col_centers)), key=lambda i: abs(col_centers[i]-xc))
        table[ri][ci].append(w['text'])
    data = [[ " ".join(cell).strip() for cell in row ] for row in table]
    if len(data)>1 and any(data[0]): return data
    return None

EXPECTED_HDR = ["Strategy","Description","Term (Months)","Start Date","End Date","Monthly Amount","Item Total","Notes"]

tables_info = []; grand_total = None; proposal_title = "Untitled Proposal"
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts = [p.extract_text() or "" for p in pdf.pages]
    for line in (texts[0] or "").splitlines():
        if "proposal" in line.lower() and len(line)>5:
            proposal_title = line.strip(); break
    def find_total(pi):
        for ln in texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\d', ln, re.I):
                return ln.strip()
    for pi, page in enumerate(pdf.pages):
        for tbl in page.find_tables():
            raw = tbl.extract(x_tolerance=1, y_tolerance=1)
            header = [str(h).strip() if h else "" for h in raw[0]]
            if len(header)<len(EXPECTED_HDR):
                rec = reconstruct_table_from_words(page, tbl.bbox)
                if rec: raw = rec
            hdr = [str(h).strip() for h in raw[0]]
            col_map = {}
            for want in EXPECTED_HDR:
                for j, got in enumerate(hdr):
                    if want.lower() in got.lower():
                        col_map[want] = j; break
                else:
                    col_map[want] = EXPECTED_HDR.index(want)
            norm_rows = []
            for row in raw[1:]:
                if all(not str(c).strip() for c in row): continue
                cells = [""]*len(EXPECTED_HDR)
                for want, j in col_map.items():
                    idx = EXPECTED_HDR.index(want)
                    cells[idx] = row[j] if j<len(row) else ""
                norm_rows.append(cells)
            if not norm_rows: continue
            tota = find_total(pi)
            tables_info.append((EXPECTED_HDR, norm_rows, tota))
    for tx in reversed(texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I)
        if m: grand_total = m.group(1); break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),
                        leftMargin=0.5*inch,rightMargin=0.5*inch,
                        topMargin=0.5*inch,bottomMargin=0.5*inch)
ts = ParagraphStyle("T",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER)
hs = ParagraphStyle("H",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
bs = ParagraphStyle("B",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT)
bl = ParagraphStyle("L",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT)
br = ParagraphStyle("R",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT)

elements = []
try:
    logo = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png").content
    img = Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    w=min(5*inch,doc.width); h=w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except: pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title),ts), Spacer(1,24)]
w_page = doc.width

for hdr, rows, tot in tables_info:
    n=len(hdr)
    desc_i=hdr.index("Description")
    desc_w=w_page*0.45; o_w=(w_page-desc_w)/(n-1)
    cw=[desc_w if i==desc_i else o_w for i in range(n)]
    data = [[Paragraph(html.escape(h),hs) for h in hdr]]
    for r in rows:
        line=[]
        for i,cell in enumerate(r):
            txt=html.escape(cell)
            if i==desc_i and txt and "-" in tot:
                p=Paragraph(f"{txt}",bs)
            else:
                p=Paragraph(txt,bs)
            line.append(p)
        data.append(line)
    if tot:
        label, val = "Total",""
        m = re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
        if m: label,val=m.group(1).strip(),m.group(2)
        row=[Paragraph(label,bl)] + [Paragraph("",bs)]*(n-2) + [Paragraph(val,br)]
        data.append(row)
    tbl = LongTable(data,colWidths=cw,repeatRows=1)
    cmds=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
          ("GRID",(0,0),(-1,-1),0.25,colors.grey),
          ("VALIGN",(0,0),(-1,0),"MIDDLE"),("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        cmds += [("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),("ALIGN",(-1,-1),(-1,-1),"RIGHT")]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total:
    hdr=tables_info[-1][0]; n=len(hdr); desc_i=hdr.index("Description")
    desc_w=w_page*0.45; o_w=(w_page-desc_w)/(n-1)
    cw=[desc_w if i==desc_i else o_w for i in range(n)]
    row=[Paragraph("Grand Total",bl)] + [Paragraph("",bs)]*(n-2) + [Paragraph(html.escape(grand_total),br)]
    gt=LongTable([row],colWidths=cw)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")
    ]))
    elements.append(gt)

doc.build(elements); pdf_buf.seek(0)

docx_buf = io.BytesIO()
docx_doc = Document()
sec=docx_doc.sections[0]; sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17); sec.page_height=Inches(11)
sec.left_margin=Inches(0.5); sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5); sec.bottom_margin=Inches(0.5)

if 'logo' in locals():
    p=docx_doc.add_paragraph(); r=p.add_run()
    img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    r.add_picture(io.BytesIO(logo),width=Inches(5)); p.alignment=WD_TABLE_ALIGNMENT.CENTER

p=docx_doc.add_paragraph(); p.alignment=WD_TABLE_ALIGNMENT.CENTER
r=p.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()

TOTAL_IN = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, tot in tables_info:
    n=len(hdr); desc_i=hdr.index("Description")
    desc_w=0.45*TOTAL_IN; o_w=(TOTAL_IN-desc_w)/(n-1)
    tbl=docx_doc.add_table(1,n,style="Table Grid"); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit=False; tbl.autofit=False
    tblPr=tbl._element.xpath('./w:tblPr')[0]
    tblW=OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    tblPr.append(tblW)
    for i,col in enumerate(tbl.columns): col.width=Inches(desc_w if i==desc_i else o_w)
    for i,name in enumerate(hdr):
        cell=tbl.rows[0].cells[i]; tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text=""; r=p.add_run(name)
        r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        rc=tbl.add_row().cells
        for i,cell in enumerate(row):
            c=rc[i]; p=c.paragraphs[0]; p.text=""; r=p.add_run(cell)
            r.font.name=DEFAULT_SANS_FONT; r.font.size=Pt(9)
            if i==desc_i and "-" in tot:
                p.add_run(" "); add_hyperlink(p, tot, "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        tr=tbl.add_row().cells
        m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
        label,val=(m.group(1).strip(),m.group(2)) if m else ("Total",tot)
        lc=tr[0]; lc.merge(tr[n-2]); p=lc.paragraphs[0]; p.text=""; r=p.add_run(label)
        r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True; p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=tr[n-1]; p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(val)
        r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True; p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total:
    hdr=tables_info[-1][0]; n=len(hdr); desc_i=hdr.index("Description")
    desc_w=0.45*TOTAL_IN; o_w=(TOTAL_IN-desc_w)/(n-1)
    tbl=docx_doc.add_table(1,n,style="Table Grid"); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit=False; tbl.autofit=False
    tblPr=tbl._element.xpath('./w:tblPr')[0]
    tblW=OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct'); tblPr.append(tblW)
    for i,col in enumerate(tbl.columns): col.width=Inches(desc_w if i==desc_i else o_w)
    cells=tbl.rows[0].cells
    lc=cells[0]; lc.merge(cells[n-2])
    tc=lc._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'E0E0E0'); tcPr.append(shd)
    p=lc.paragraphs[0]; p.text=""; r=p.add_run("Grand Total")
    r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True; p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]; tc2=ac._tc; tcPr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd'); shd2.set(qn('w:fill'),'E0E0E0'); tcPr2.append(shd2)
    p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(grand_total)
    r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True; p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO(); docx_doc.save(docx_buf); docx_buf.seek(0)
c1, c2 = st.columns(2)
if pdf_buf: c1.download_button("📥 Download PDF", data=pdf_buf, file_name="proposal.pdf", mime="application/pdf")
else: c1.error("PDF failed")
if docx_buf: c2.download_button("📥 Download DOCX", data=docx_buf, file_name="proposal.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else: c2.error("DOCX failed")
