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

def reconstruct_table_from_words(page, bbox):
    x0, top, x1, bottom = bbox
    words = page.crop((x0, top, x1, bottom)).extract_words(use_text_flow=True, keep_blank_chars=False)
    if not words:
        return None
    ys = sorted(set([(w['top'] + w['bottom']) / 2 for w in words]))
    ys_sorted = sorted(ys)
    rows = []
    for y in ys_sorted:
        row_words = [w for w in words if w['top'] <= y <= w['bottom']]
        rows.append(row_words)
    if len(rows) < 2:
        return None
    header_row = rows[0]
    xs = sorted([ (w['x0'] + w['x1'])/2 for w in header_row ])
    if len(xs) < 2:
        return None
    col_bounds = []
    xs_sorted = sorted(xs)
    for i in range(len(xs_sorted)):
        if i == 0:
            col_bounds.append((x0, (xs_sorted[0]+xs_sorted[1])/2))
        elif i == len(xs_sorted)-1:
            col_bounds.append(((xs_sorted[-2]+xs_sorted[-1])/2, x1))
        else:
            col_bounds.append(((xs_sorted[i-1]+xs_sorted[i])/2, (xs_sorted[i]+xs_sorted[i+1])/2))
    table = []
    for row in rows:
        cells = [''] * len(col_bounds)
        for w in row:
            cx = (w['x0'] + w['x1'])/2
            for ci, (cx0, cx1) in enumerate(col_bounds):
                if cx0 <= cx <= cx1:
                    cells[ci] = (cells[ci] + ' ' + w['text']).strip() if cells[ci] else w['text']
                    break
        table.append(cells)
    return table

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

def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    return lines[0], re.sub(r'\s+', ' ', " ".join(lines[1:])).strip()

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink'); hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05,0x63,0xC1); style.font.underline = True
    rStyle = OxmlElement('w:rStyle'); rStyle.set(qn('w:val'),'Hyperlink'); rPr.append(rStyle)
    if font_name:
        rf = OxmlElement('w:rFonts'); rf.set(qn('w:ascii'),font_name); rf.set(qn('w:hAnsi'),font_name); rPr.append(rf)
    if font_size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'),str(int(font_size*2)))
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'),str(int(font_size*2)))
        rPr.append(sz); rPr.append(szCs)
    if bold:
        rPr.append(OxmlElement('w:b'))
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text = text; new_run.append(t)
    hyperlink.append(new_run); paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    if page_texts:
        fl = page_texts[0].splitlines()
        for ln in fl:
            if "proposal" in ln.lower():
                proposal_title = ln.strip(); break
    used_totals = set()
    def find_total(pi):
        if pi<0 or pi>=len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\btotal\b.*\$',ln,re.I) and ln not in used_totals:
                used_totals.add(ln); return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        for tbl in page.find_tables():
            manual = reconstruct_table_from_words(page, tbl.bbox)
            data = manual if manual and len(manual)>1 and any(manual[0]) else tbl.extract()
            if not data or len(data)<2:
                continue
            hdr = [str(h).strip() if h else "" for h in data[0]]
            desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
            if desc_i is None:
                desc_i = next((i for i,h in enumerate(hdr) if len(h)>10), None)
            if desc_i is None:
                continue
            rows = []
            uris = []
            links = page.hyperlinks
            for ridx,row in enumerate(data[1:],start=1):
                cells = [str(c).strip() if c else "" for c in row]
                if all(not c for c in cells): continue
                first = cells[0].lower()
                if ("total" in first or "subtotal" in first) and any("$" in c for c in cells):
                    table_total = cells; continue
                strat, desc = split_cell_text(cells[desc_i])
                rest = [cells[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc]+rest)
                uris.append(None)
            table_total = locals().get('table_total', None) or find_total(pi)
            if rows:
                new_hdr = ["Strategy","Description"]+[h for i,h in enumerate(hdr) if i!=desc_i and h]
                tables_info.append((new_hdr, rows, uris, table_total))
    for tx in reversed(page_texts):
        m = re.search(r'Grand\s*Total.*?(\$\s*[\d,]+\.\d{2})',tx,re.I)
        if m:
            grand_total = m.group(1); break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf,pagesize=landscape((17*inch,11*inch)),
    leftMargin=0.5*inch,rightMargin=0.5*inch,topMargin=0.5*inch,bottomMargin=0.5*inch)
ts = ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hs = ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
bs = ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl = ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br = ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)
elements=[]
try:
    logo = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",timeout=5).content
    img = Image.open(io.BytesIO(logo)); r=img.height/img.width
    w=min(5*inch,doc.width); h=w*r
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except:
    logo=None
elements += [Spacer(1,12),Paragraph(html.escape(proposal_title),ts),Spacer(1,24)]
tw = doc.width
for hdr,rows,uris,tot in tables_info:
    nc=len(hdr)
    di=hdr.index("Description") if "Description" in hdr else 1
    dw=tw*0.45; ow=(tw-dw)/(nc-1) if nc>1 else tw
    cw=[dw if i==di else ow for i in range(nc)]
    data=[[Paragraph(html.escape(h),hs) for h in hdr]]
    for i,row in enumerate(rows):
        line=[]
        for j,cell in enumerate(row):
            t=html.escape(cell)
            if j==di and uris[i]:
                p=Paragraph(f"{t} <link href='{html.escape(uris[i])}' color='blue'>- link</link>",bs)
            else:
                p=Paragraph(t,bs)
            line.append(p)
        data.append(line)
    if tot:
        label,val="Total",""
        if isinstance(tot,list):
            label=tot[0] or "Total"; val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: label,val=m.group(1).strip(),m.group(2)
        tr=[Paragraph(html.escape(label),bl)]+[Spacer(1,0)]*(nc-2)+[Paragraph(val,br)]
        data.append(tr)
    tbl=LongTable(data,colWidths=cw,repeatRows=1)
    sc=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot and nc>1:
        sc+= [("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),
              ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(sc))
    elements += [tbl,Spacer(1,24)]
if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; nc=len(last_hdr)
    di=last_hdr.index("Description") if "Description" in last_hdr else 1
    dw=tw*0.45; ow=(tw-dw)/(nc-1) if nc>1 else tw
    cw=[dw if i==di else ow for i in range(nc)]
    row=[Paragraph("Grand Total",bl)]+[Spacer(1,0)]*(nc-2)+[Paragraph(html.escape(grand_total),br)]
    gt=LongTable([row],colWidths=cw)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
                            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    elements.append(gt)
doc.build(elements)
pdf_buf.seek(0)

docx_buf=io.BytesIO()
docx_doc=Document()
sec=docx_doc.sections[0];sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17);sec.page_height=Inches(11)
sec.left_margin=Inches(0.5);sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5);sec.bottom_margin=Inches(0.5)
if logo:
    p=docx_doc.add_paragraph(); r=p.add_run()
    img=Image.open(io.BytesIO(logo)); r_in=5; h_in=r_in*(img.height/img.width)
    r.add_picture(io.BytesIO(logo),width=Inches(r_in)); p.alignment=WD_TABLE_ALIGNMENT.CENTER
p=docx_doc.add_paragraph(); p.alignment=WD_TABLE_ALIGNMENT.CENTER
r=p.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()
TW=sec.page_width.inches-sec.left_margin.inches-sec.right_margin.inches
for hdr,rows,uris,tot in tables_info:
    nc=len(hdr); di=hdr.index("Description") if "Description" in hdr else 1
    dw=0.45*TW; ow=(TW-dw)/(nc-1) if nc>1 else TW
    tbl=docx_doc.add_table(rows=1,cols=nc,style="Table Grid"); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.allow_autofit=False;tbl.autofit=False
    tblPr_list=tbl._element.xpath('./w:tblPr')
    if tblPr_list: tblPr=tblPr_list[0]
    else:
        tblPr=OxmlElement('w:tblPr');tbl._element.insert(0,tblPr)
    tblW=OxmlElement('w:tblW');tblW.set(qn('w:w'),'5000');tblW.set(qn('w:type'),'pct')
    ex=tblPr.xpath('./w:tblW'); 
    if ex: tblPr.remove(ex[0])
    tblPr.append(tblW)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(dw if i==di else ow)
    hdr_cells=tbl.rows[0].cells
    for i,name in enumerate(hdr):
        c=hdr_cells[i];tc=c._tc;tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd');shd.set(qn('w:fill'),'F2F2F2');tcPr.append(shd)
        p=c.paragraphs[0];p.text=""
        run=p.add_run(name);run.font.name=DEFAULT_SERIF_FONT;run.font.size=Pt(10);run.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER;c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        cells=tbl.add_row().cells
        for j,cell in enumerate(row):
            c=cells[j];p=c.paragraphs[0];p.text=""
            r0=p.add_run(str(cell));r0.font.name=DEFAULT_SANS_FONT;r0.font.size=Pt(9)
            if j==di and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p,uris[ridx],"- link",font_name=DEFAULT_SANS_FONT,font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT;c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        tcells=tbl.add_row().cells
        label,amount="Total",""
        if isinstance(tot,list):
            label=tot[0] or "Total"; amount=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: label,amount=m.group(1).strip(),m.group(2)
        lc=tcells[0]
        if nc>1: lc.merge(tcells[nc-2])
        p0=lc.paragraphs[0];p0.text=""
        r0=p0.add_run(label);r0.font.name=DEFAULT_SERIF_FONT;r0.font.size=Pt(10);r0.bold=True
        p0.alignment=WD_TABLE_ALIGNMENT.LEFT;lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=tcells[nc-1];p1=ac.paragraphs[0];p1.text=""
        r1=p1.add_run(amount);r1.font.name=DEFAULT_SERIF_FONT;r1.font.size=Pt(10);r1.bold=True
        p1.alignment=WD_TABLE_ALIGNMENT.RIGHT;ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()
if grand_total and tables_info:
    lh=tables_info[-1][0];nc=len(lh);di=lh.index("Description") if "Description" in lh else 1
    dw=0.45*TW;ow=(TW-dw)/(nc-1) if nc>1 else TW
    tblg=docx_doc.add_table(rows=1,cols=nc,style="Table Grid");tblg.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblg.allow_autofit=False;tblg.autofit=False
    pr=tblg._element.xpath('./w:tblPr')
    if pr: tblgPr=pr[0]
    else:
        tblgPr=OxmlElement('w:tblPr');tblg._element.insert(0,tblgPr)
    w=OxmlElement('w:tblW');w.set(qn('w:w'),'5000');w.set(qn('w:type'),'pct')
    ex=tblgPr.xpath('./w:tblW'); 
    if ex: tblgPr.remove(ex[0])
    tblgPr.append(w)
    for i,col in enumerate(tblg.columns):
        col.width=Inches(dw if i==di else ow)
    cells=tblg.rows[0].cells;lc=cells[0]
    if nc>1: lc.merge(cells[nc-2])
    tc=lc._tc;tcPr=tc.get_or_add_tcPr()
    s=OxmlElement('w:shd');s.set(qn('w:fill'),'E0E0E0');tcPr.append(s)
    p=lc.paragraphs[0];p.text=""
    r=p.add_run("Grand Total");r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT;lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[nc-1];tc2=ac._tc;tcPr2=tc2.get_or_add_tcPr()
    s2=OxmlElement('w:shd');s2.set(qn('w:fill'),'E0E0E0');tcPr2.append(s2)
    p2=ac.paragraphs[0];p2.text=""
    r2=p2.add_run(grand_total);r2.font.name=DEFAULT_SERIF_FONT;r2.font.size=Pt(10);r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT;ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_doc.save(io.BytesIO())  # to reset internal pointers
docx_buf = io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1, c2 = st.columns(2)
if pdf_buf:
    c1.download_button("📥 Download deliverable PDF",pdf_buf,"proposal_deliverable.pdf","application/pdf",use_container_width=True)
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button("📥 Download deliverable DOCX",docx_buf,"proposal_deliverable.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
else:
    c2.error("Word document generation failed.")
