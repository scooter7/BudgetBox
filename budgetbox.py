# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image as PILImage
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
import re, html

# register fonts
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT  = "Barlow"
except:
    DEFAULT_SERIF_FONT = "Times New Roman"
    DEFAULT_SANS_FONT  = "Arial"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    desc = " ".join(lines[1:])
    desc = re.sub(r'\s+', ' ', desc).strip()
    return lines[0], desc

def reconstruct_table_from_words(page, bbox):
    crop = page.within_bbox(bbox)
    words = crop.extract_words(x_tolerance=1, y_tolerance=1)
    if not words:
        return None
    ys = sorted(set([w['top'] for w in words]))
    rows = []
    for y in ys:
        row = [w for w in words if abs(w['top']-y)<1]
        rows.append(sorted(row, key=lambda w:w['x0']))
    header = rows[0]
    xs = [w['x0'] for w in header]
    cols, data = [], []
    for i,start in enumerate(xs):
        end = xs[i+1] if i+1<len(xs) else bbox[2]
        cols.append((start,end))
    for row in rows:
        cells = ['']*len(cols)
        for w in row:
            x=w['x0']
            for j,(s,e) in enumerate(cols):
                if s<=x<e:
                    cells[j] += (' '+w['text']) if cells[j] else w['text']
                    break
        data.append([c.strip() for c in cells])
    return data

def extract_tables_and_links():
    tables_info=[]
    grand_total=None
    proposal_title="Untitled Proposal"
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1,y_tolerance=1) or "" for p in pdf.pages]
        if page_texts:
            for ln in page_texts[0].splitlines():
                if "proposal" in ln.lower():
                    proposal_title=ln.strip()
                    break
        used_totals=set()
        def find_total(pi):
            if pi<0 or pi>=len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?!grand\s)total\b.*?\$\s*\d', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None
        for pi,page in enumerate(pdf.pages):
            for tbl in page.find_tables():
                data = reconstruct_table_from_words(page, tbl.bbox)
                if not data or len(data)<2 or not any(data[0]):
                    data = tbl.extract(x_tolerance=1,y_tolerance=1)
                if not data or len(data)<2:
                    continue
                hdr = [str(h).strip() if h else "" for h in data[0]]
                desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
                if desc_i is None:
                    desc_i = next((i for i,h in enumerate(hdr) if len(h)>10), None)
                if desc_i is None or len(hdr)<2:
                    continue
                rows_data=[]
                table_total=None
                for row in data[1:]:
                    row = [str(c).strip() if c else "" for c in row]
                    if not any(row): continue
                    first=row[0].lower()
                    if ("total" in first or "subtotal" in first) and any("$" in c for c in row):
                        if table_total is None: table_total=row
                        continue
                    strat,desc=split_cell_text(row[desc_i])
                    rest=[row[i] for i in range(len(row)) if i!=desc_i and hdr[i]]
                    rows_data.append([strat,desc]+rest)
                if table_total is None:
                    table_total=find_total(pi)
                if rows_data:
                    tables_info.append((hdr, rows_data, table_total))
        for tx in reversed(page_texts):
            m=re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I)
            if m:
                grand_total=m.group(1).replace(" ","")
                break
    return proposal_title, tables_info, grand_total

proposal_title, tables_info, grand_total = extract_tables_and_links()

# build PDF
pdf_buf=io.BytesIO()
doc=SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),
                      leftMargin=0.5*inch,rightMargin=0.5*inch,
                      topMargin=0.5*inch,bottomMargin=0.5*inch)
ts=ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hs=ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
bs=ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl=ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br=ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)
elements=[]
# logo
try:
    logo_url="https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    r=requests.get(logo_url,timeout=5); r.raise_for_status()
    logo=r.content
    img=PILImage.open(io.BytesIO(logo)); ratio=img.height/img.width
    iw=min(5*inch,doc.width); ih=iw*ratio
    elements.append(RLImage(io.BytesIO(logo),width=iw,height=ih))
except:
    pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
w=doc.width
for hdr,rows,tot in tables_info:
    num=len(hdr)
    desc_idx=hdr.index(next(h for h in hdr if "description" in h.lower()))
    dw=w*0.45; ow=(w-dw)/(num-1) if num>1 else w
    cw=[dw if i==desc_idx else ow for i in range(num)]
    wrapped=[[Paragraph(html.escape(h),hs) for h in hdr]]
    for row in rows:
        line=[]
        for i,cell in enumerate(row):
            txt=html.escape(cell)
            p=Paragraph(txt,bs)
            line.append(p)
        wrapped.append(line)
    if tot is not None:
        label="Total"; val=""
        if isinstance(tot,list):
            label=tot[0] or "Total"
            val=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m:
                label,val=m.group(1).strip(),m.group(2)
            else:
                val=str(tot)
        total_row=[Paragraph(html.escape(label),bl)] + [Paragraph("",bs)]*(num-2) + [Paragraph(html.escape(val),br)]
        wrapped.append(total_row)
    tbl=LongTable(wrapped,colWidths=cw,repeatRows=1)
    style_cmds=[
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP"),
    ]
    if tot is not None:
        style_cmds += [
            ("SPAN",(0,-1),(-2,-1)),
            ("ALIGN",(0,-1),(-2,-1),"LEFT"),
            ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),
            ("VALIGN",(0,-1),(-1,-1),"MIDDLE"),
        ]
    tbl.setStyle(TableStyle(style_cmds))
    elements += [tbl,Spacer(1,24)]
if grand_total:
    last_hdr=tables_info[-1][0]; num=len(last_hdr)
    desc_idx=last_hdr.index(next(h for h in last_hdr if "description" in h.lower()))
    dw=w*0.45; ow=(w-dw)/(num-1) if num>1 else w
    cw=[dw if i==desc_idx else ow for i in range(num)]
    row=[Paragraph("Grand Total",bl)] + [Paragraph("",bs)]*(num-2) + [Paragraph(html.escape(grand_total),br)]
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
pdf_buf.seek(0)

# build Word
docx_buf=io.BytesIO()
docx_doc=Document()
sec=docx_doc.sections[0]
sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17); sec.page_height=Inches(11)
sec.left_margin=sec.right_margin=sec.top_margin=sec.bottom_margin=Inches(0.5)
if 'logo' in locals():
    try:
        p=docx_doc.add_paragraph(); r=p.add_run()
        img=PILImage.open(io.BytesIO(logo)); ratio=img.height/img.width
        w_in=5; h_in=w_in*ratio
        r.add_picture(io.BytesIO(logo),width=Inches(w_in))
        p.alignment=WD_TABLE_ALIGNMENT.CENTER
    except:
        pass
p=docx_doc.add_paragraph(); p.alignment=WD_TABLE_ALIGNMENT.CENTER
r=p.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()
page_w_in=sec.page_width.inches-sec.left_margin.inches-sec.right_margin.inches
for hdr,rows,tot in tables_info:
    num=len(hdr)
    desc_idx=hdr.index(next(h for h in hdr if "description" in h.lower()))
    dw_in=0.45*page_w_in; ow_in=(page_w_in-dw_in)/(num-1) if num>1 else page_w_in
    tbl=docx_doc.add_table(rows=1,cols=num,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.allow_autofit=False; tbl.autofit=False
    tblPr=tbl._element.xpath('./w:tblPr')
    if not tblPr:
        pr=OxmlElement('w:tblPr'); tbl._element.insert(0,pr)
    else:
        pr=tblPr[0]
    w_elem=OxmlElement('w:tblW'); w_elem.set(qn('w:w'),'5000'); w_elem.set(qn('w:type'),'pct')
    old=pr.xpath('./w:tblW')
    if old: pr.remove(old[0])
    pr.append(w_elem)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(dw_in if i==desc_idx else ow_in)
    # header
    for i,name in enumerate(hdr):
        cell=tbl.rows[0].cells[i]; tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text=""
        run=p.add_run(str(name)); run.font.name=DEFAULT_SERIF_FONT; run.font.size=Pt(10); run.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    # body
    for ridx,row in enumerate(rows):
        rc=tbl.add_row().cells
        for i,val in enumerate(row):
            cell=rc[i]; p=cell.paragraphs[0]; p.text=""
            run=p.add_run(str(val)); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    # total
    if tot is not None:
        tcells=tbl.add_row().cells
        label="Total"; amount=""
        if isinstance(tot,list):
            label=tot[0] or "Total"; amount=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: label,amount=m.group(1).strip(),m.group(2)
            else: amount=str(tot)
        lbl=tcells[0]
        if num>1: lbl.merge(tcells[num-2])
        p=lbl.paragraphs[0]; p.text=""
        r=p.add_run(label); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT; lbl.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        am=tcells[num-1]; p2=am.paragraphs[0]; p2.text=""
        r2=p2.add_run(amount); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; am.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total:
    hdr_last,_,_ = tables_info[-1]
    num=len(hdr_last)
    desc_idx=hdr_last.index(next(h for h in hdr_last if "description" in h.lower()))
    dw_in=0.45*page_w_in; ow_in=(page_w_in-dw_in)/(num-1) if num>1 else page_w_in
    tbl=docx_doc.add_table(rows=1,cols=num,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.allow_autofit=False; tbl.autofit=False
    pr_list=tbl._element.xpath('./w:tblPr')
    if not pr_list:
        pr=OxmlElement('w:tblPr'); tbl._element.insert(0,pr)
    else:
        pr=pr_list[0]
    w_elem=OxmlElement('w:tblW'); w_elem.set(qn('w:w'),'5000'); w_elem.set(qn('w:type'),'pct')
    old=pr.xpath('./w:tblW'); 
    if old: pr.remove(old[0])
    pr.append(w_elem)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(dw_in if i==desc_idx else ow_in)
    cells=tbl.rows[0].cells
    lbl=cells[0]
    if num>1: lbl.merge(cells[num-2])
    tc=lbl._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'E0E0E0'); tcPr.append(shd)
    p=lbl.paragraphs[0]; p.text=""
    r=p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lbl.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    am=cells[num-1]; tc2=am._tc; tcPr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd'); shd2.set(qn('w:fill'),'E0E0E0'); tcPr2.append(shd2)
    p2=am.paragraphs[0]; p2.text=""
    r2=p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; am.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf=io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1,c2 = st.columns(2)
if pdf_buf:
    c1.download_button("📥 Download deliverable PDF", data=pdf_buf.getvalue(),
                       file_name="proposal_deliverable.pdf",mime="application/pdf")
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button("📥 Download deliverable DOCX", data=docx_buf.getvalue(),
                       file_name="proposal_deliverable.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    c2.error("Word generation failed.")
