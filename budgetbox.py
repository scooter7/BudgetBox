# -*- coding: utf-8 -*-
import streamlit as st
import io
import requests
from PIL import Image
import pdfplumber
import camelot
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
    DEFAULT_SERIF_FONT="DMSerif"
    DEFAULT_SANS_FONT="Barlow"
except:
    DEFAULT_SERIF_FONT="Times New Roman"
    DEFAULT_SANS_FONT="Arial"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded=st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded: st.stop()
pdf_bytes=uploaded.read()

def split_cell_text(raw):
    lines=[l.strip() for l in raw.splitlines() if l.strip()]
    if not lines: return "",""
    return lines[0], " ".join(lines[1:])

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part=paragraph.part
    rid=part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link=OxmlElement('w:hyperlink'); link.set(qn('r:id'), rid)
    r=OxmlElement('w:r'); rPr=OxmlElement('w:rPr')
    styles=paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style=styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb=RGBColor(0x05,0x63,0xC1); style.font.underline=True
    se=OxmlElement('w:rStyle'); se.set(qn('w:val'),'Hyperlink'); rPr.append(se)
    if font_name:
        rf=OxmlElement('w:rFonts'); rf.set(qn('w:ascii'),font_name); rf.set(qn('w:hAnsi'),font_name); rPr.append(rf)
    if font_size:
        sz=OxmlElement('w:sz'); sz.set(qn('w:val'),str(int(font_size*2)))
        szcs=OxmlElement('w:szCs'); szcs.set(qn('w:val'),str(int(font_size*2)))
        rPr.append(sz); rPr.append(szcs)
    if bold: rPr.append(OxmlElement('w:b'))
    r.append(rPr)
    t=OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text=text; r.append(t)
    link.append(r); paragraph._p.append(link)
    return docx.text.run.Run(r, paragraph)

# Try Camelot for first page
first_table=None
try:
    tables=camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n")
    if tables:
        raw=tables[0].df.values.tolist()
        if len(raw)>1 and raw[0][0].strip().lower()=="strategy" and raw[1][0].strip().lower()=="description":
            hdr=["Strategy","Description"]+raw[1][1:]
            n=len(hdr)
            rows=[[r[i] if i<len(r) else "" for i in range(n)] for r in raw[2:]]
            first_table=[hdr]+rows
except:
    first_table=None

tables_info=[]
grand_total=None

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts=[p.extract_text() or "" for p in pdf.pages]
    proposal_title=next((ln for ln in page_texts[0].splitlines() if "proposal" in ln.lower()), page_texts[0].splitlines()[0] if page_texts[0].splitlines() else "Untitled Proposal").strip()
    used_totals=set()
    def find_total(pi):
        if pi>=len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    if first_table:
        hdr=first_table[0]
        rows=first_table[1:]
        uris=[None]*len(rows)
        data_rows=[[split_cell_text(r[1])[0], split_cell_text(r[1])[1]]+r[2:] for r in rows]
        tables_info.append((hdr, data_rows, uris, None))

    for pi,page in enumerate(pdf.pages):
        if pi==0 and first_table: continue
        links=page.hyperlinks
        for tbl in page.find_tables():
            data=tbl.extract()
            if not data or len(data)<2: continue
            hdr_raw=[str(h).strip() if h else "" for h in data[0]]
            desc_i=next((i for i,h in enumerate(hdr_raw) if "description" in h.lower()), None)
            if desc_i is None:
                desc_i=next((i for i,h in enumerate(hdr_raw) if len(h)>10), None)
            if desc_i is None or len(hdr_raw)<=1: continue
            desc_links={}
            for r,row_obj in enumerate(getattr(tbl,'rows',[])):
                if r==0: continue
                if hasattr(row_obj,'cells') and desc_i<len(row_obj.cells):
                    x0,top,x1,bottom=row_obj.cells[desc_i]
                    for link in links:
                        if all(k in link for k in ('x0','x1','top','bottom','uri')):
                            lx0,lt,lx1,lb=link['x0'],link['top'],link['x1'],link['bottom']
                            if not (lx1<x0 or lx0>x1 or lb<top or lt>bottom):
                                desc_links[r]=link['uri']
                                break
            new_hdr=["Strategy","Description"]+[h for i,h in enumerate(hdr_raw) if i!=desc_i and h]
            rows_data=[]; uris=[]; table_total=None
            for ridx,row in enumerate(data[1:], start=1):
                cells=[str(c).strip() if c else "" for c in row]
                if all(not c for c in cells): continue
                first=cells[0].lower()
                if ("total" in first or "subtotal" in first) and any("$" in c for c in cells):
                    if table_total is None: table_total=cells
                    continue
                strat,desc=split_cell_text(cells[desc_i])
                rest=[cells[i] for i,h in enumerate(hdr_raw) if i!=desc_i and h]
                rows_data.append([strat,desc]+rest)
                uris.append(desc_links.get(ridx))
            if table_total is None:
                table_total=find_total(pi)
            if rows_data:
                tables_info.append((new_hdr, rows_data, uris, table_total))

    for tx in reversed(page_texts):
        m=re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total=m.group(1).replace(" ","")
            break

# Build PDF
pdf_buf=io.BytesIO()
doc=SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)), leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
ts=ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hs=ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
bs=ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT)
bl=ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT)
br=ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT)
elements=[]
logo=None
try:
    logo_url="https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp=requests.get(logo_url,timeout=10); resp.raise_for_status()
    logo=resp.content
    img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    w=min(5*inch,doc.width); h=w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except: pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
total_w=doc.width

for hdr, rows, uris, tot in tables_info:
    n=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(n-1) if n>1 else total_w
    col_ws=[desc_w if i==desc_idx else other_w for i in range(n)]
    wrapped=[[Paragraph(html.escape(h),hs) for h in hdr]]
    for ridx,row in enumerate(rows):
        line=[]
        for cidx,cell in enumerate(row):
            txt=html.escape(cell)
            if cidx==desc_idx and uris[ridx]:
                p=Paragraph(f"{txt} <link href='{html.escape(uris[ridx])}'>- link</link>", bs)
            else:
                p=Paragraph(txt,bs)
            line.append(p)
        wrapped.append(line)
    if tot:
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            lbl,val=(m.group(1).strip(),m.group(2)) if m else ("Total",tot)
        row_elems=[Paragraph(lbl,bl)]+[Paragraph("",bs)]*(n-2)+[Paragraph(val,br)]
        wrapped.append(row_elems)
    tbl=LongTable(wrapped,colWidths=col_ws,repeatRows=1)
    sc=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey)]
    if tot and n>1:
        sc+=[("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),("ALIGN",(-1,-1),(-1,-1),"RIGHT")]
    tbl.setStyle(TableStyle(sc))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; n=len(last_hdr)
    desc_idx=last_hdr.index("Description") if "Description" in last_hdr else 1
    desc_w=total_w*0.45; other_w=(total_w-desc_w)/(n-1) if n>1 else total_w
    col_ws=[desc_w if i==desc_idx else other_w for i in range(n)]
    row=[Paragraph("Grand Total",bl)]+[Spacer(1,0)]*(n-2)+[Paragraph(html.escape(grand_total),br)]
    gt=LongTable([row],colWidths=col_ws)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

# Build Word
docx_buf=io.BytesIO()
docx_doc=Document()
sec=docx_doc.sections[0]
sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17); sec.page_height=Inches(11)
sec.left_margin=Inches(0.5); sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5); sec.bottom_margin=Inches(0.5)
if logo:
    p=docx_doc.add_paragraph(); r=p.add_run()
    img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    r.add_picture(io.BytesIO(logo),width=Inches(5))
    p.alignment=WD_TABLE_ALIGNMENT.CENTER
p=docx_doc.add_paragraph(); p.alignment=WD_TABLE_ALIGNMENT.CENTER
r=p.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()
TOTAL_W=sec.page_width.inches-sec.left_margin.inches-sec.right_margin.inches

for hdr, rows, uris, tot in tables_info:
    n=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=0.45*TOTAL_W; other_w=(TOTAL_W-desc_w)/(n-1) if n>1 else TOTAL_W
    tbl=docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.allow_autofit=False; tbl.autofit=False
    tpr=tbl._element.xpath('./w:tblPr')
    pr=tpr[0] if tpr else OxmlElement('w:tblPr')
    if not tpr: tbl._element.insert(0, pr)
    wtag=OxmlElement('w:tblW'); wtag.set(qn('w:w'),'5000'); wtag.set(qn('w:type'),'pct')
    ex=pr.xpath('./w:tblW')
    if ex: pr.remove(ex[0])
    pr.append(wtag)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    # header
    for i,name in enumerate(hdr):
        cell=tbl.rows[0].cells[i]; tc=cell._tc; tcpr=tc.get_or_add_tcPr()
        sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),'F2F2F2'); tcpr.append(sh)
        p=cell.paragraphs[0]; p.text=""; run=p.add_run(name)
        run.font.name=DEFAULT_SERIF_FONT; run.font.size=Pt(10); run.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    # body
    for ridx,row in enumerate(rows):
        cells=tbl.add_row().cells
        for cidx,val in enumerate(row):
            cell=cells[cidx]; p=cell.paragraphs[0]; p.text=""
            run=p.add_run(str(val)); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            if cidx==desc_idx and uris[ridx]:
                p.add_run(" "); add_hyperlink(p, uris[ridx], "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    # total
    if tot:
        cells=tbl.add_row().cells
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; amt=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            lbl,amt=(m.group(1).strip(),m.group(2)) if m else ("Total",tot)
        lc=cells[0]
        if n>1: lc.merge(cells[n-2])
        p=lc.paragraphs[0]; p.text=""; r=p.add_run(lbl)
        r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=cells[n-1]; p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(amt)
        r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; n=len(last_hdr)
    desc_idx=last_hdr.index("Description") if "Description" in last_hdr else 1
    desc_w=0.45*TOTAL_W; other_w=(TOTAL_W-desc_w)/(n-1) if n>1 else TOTAL_W
    tblg=docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tblg.alignment=WD_TABLE_ALIGNMENT.CENTER; tblg.allow_autofit=False; tblg.autofit=False
    tpr=tblg._element.xpath('./w:tblPr')
    pr=tpr[0] if tpr else OxmlElement('w:tblPr')
    if not tpr: tblg._element.insert(0,pr)
    wtag=OxmlElement('w:tblW'); wtag.set(qn('w:w'),'5000'); wtag.set(qn('w:type'),'pct')
    ex=pr.xpath('./w:tblW')
    if ex: pr.remove(ex[0])
    pr.append(wtag)
    for i,col in enumerate(tblg.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    cells=tblg.rows[0].cells
    lc=cells[0]
    if n>1: lc.merge(cells[n-2])
    tc=lc._tc; tcpr=tc.get_or_add_tcPr()
    sh=OxmlElement('w:shd'); sh.set(qn('w:fill'),'E0E0E0'); tcpr.append(sh)
    p=lc.paragraphs[0]; p.text=""; r=p.add_run("Grand Total")
    r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]; tc2=ac._tc; tcpr2=tc2.get_or_add_tcPr()
    sh2=OxmlElement('w:shd'); sh2.set(qn('w:fill'),'E0E0E0'); tcpr2.append(sh2)
    p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(grand_total)
    r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf=io.BytesIO()
docx_doc.save(docx_buf); docx_buf.seek(0)

c1,c2=st.columns(2)
if pdf_buf:
    c1.download_button("📥 Download deliverable PDF",pdf_buf,"proposal_deliverable.pdf","application/pdf",use_container_width=True)
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button("📥 Download deliverable DOCX",docx_buf,"proposal_deliverable.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
else:
    c2.error("Word document generation failed.")
