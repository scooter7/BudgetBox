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
from docx.oxml.ns import qn, nsdecls
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
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    desc = " ".join(lines[1:])
    desc = re.sub(r'\s+', ' ', desc).strip()
    return lines[0], desc

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
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
        size = OxmlElement('w:sz'); size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size); rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

def reconstruct_table_from_words(page, bbox):
    x0, top, x1, bottom = bbox
    crop = page.within_bbox((x0, top, x1, bottom))
    words = crop.extract_words()
    if not words:
        return None
    words_sorted = sorted(words, key=lambda w: w['top'])
    rows = []
    gaps = []
    for i in range(len(words_sorted) - 1):
        gaps.append(words_sorted[i+1]['top'] - words_sorted[i]['bottom'])
    row_thresh = max(gaps) * 0.5 if gaps else 5
    current_row = [words_sorted[0]]
    for w in words_sorted[1:]:
        if w['top'] - current_row[-1]['bottom'] > row_thresh:
            rows.append(current_row)
            current_row = [w]
        else:
            current_row.append(w)
    rows.append(current_row)
    if len(rows) < 2:
        return None
    header_row = rows[0]
    header_row = sorted(header_row, key=lambda w: w['x0'])
    col_edges = [w['x0'] for w in header_row] + [header_row[-1]['x1']]
    table = []
    for r in rows:
        line = [""] * (len(col_edges) - 1)
        for w in r:
            for i in range(len(col_edges)-1):
                if col_edges[i] <= w['x0'] < col_edges[i+1]:
                    sep = " " if line[i] else ""
                    line[i] += sep + w['text']
                    break
        table.append(line)
    if any(not any(c.strip() for c in row) for row in table):
        return None
    return table

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
    first = page_texts[0].splitlines() if page_texts else []
    pt = next((ln for ln in first if "proposal" in ln.lower() and len(ln.strip())>5), None)
    if pt: proposal_title = pt.strip()
    elif first: proposal_title = first[0].strip()
    used = set()
    def find_total(pi):
        if pi>=len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used:
                used.add(ln)
                return ln.strip()
        return None
    for pi, page in enumerate(pdf.pages):
        for tbl in page.find_tables():
            data = reconstruct_table_from_words(page, tbl.bbox)
            if data and len(data)>1 and any(cell.strip() for cell in data[0]):
                raw = data
            else:
                raw = tbl.extract(x_tolerance=3, y_tolerance=3)
            hdr = [str(h).strip() for h in raw[0]]
            desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
            if desc_i is None:
                desc_i = next((i for i,h in enumerate(hdr) if len(h)>10), None)
                if desc_i is None: continue
            rows_data = []
            row_links = []
            table_total = None
            for ridx,row in enumerate(raw[1:],1):
                if all(not str(c).strip() for c in row):
                    continue
                first = str(row[0]).lower()
                if ("total" in first or "subtotal" in first) and any("$" in str(c) for c in row):
                    if table_total is None:
                        table_total = row
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [str(row[i]) for i in range(len(row)) if i!=desc_i]
                rows_data.append([strat, desc] + rest)
                row_links.append(None)
            if table_total is None:
                table_total = find_total(pi)
            if rows_data:
                tables_info.append(([ "Strategy", "Description" ] + [h for i,h in enumerate(hdr) if i!=desc_i], rows_data, row_links, table_total))
    for tx in reversed(page_texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1).replace(" ","")
            break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),leftMargin=0.5*inch,rightMargin=0.5*inch,topMargin=0.5*inch,bottomMargin=0.5*inch)
title_style=ParagraphStyle("T",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hdr_style=ParagraphStyle("H",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
body_style=ParagraphStyle("B",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl_style=ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br_style=ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)
elements=[]
try:
    logo_url="https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    r=requests.get(logo_url,timeout=10);r.raise_for_status();logo=r.content
    img=Image.open(io.BytesIO(logo));rat=img.height/img.width
    w=min(5*inch,doc.width);h=w*rat;elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except: pass
elements += [Spacer(1,12),Paragraph(html.escape(proposal_title),title_style),Spacer(1,24)]
tw=doc.width
for hdr, rows, uris, tot in tables_info:
    nc=len(hdr); di=hdr.index("Description") if "Description" in hdr else 1
    dw=0.45*tw;ow=(tw-dw)/(nc-1) if nc>1 else tw
    cw=[dw if i==di else ow for i in range(nc)]
    wrapped=[[Paragraph(html.escape(h),hdr_style) for h in hdr]]
    for i,row in enumerate(rows):
        line=[]
        for j,cell in enumerate(row):
            txt=html.escape(cell)
            if j==di and uris[i]:
                p=Paragraph(f"{txt} <link href='{html.escape(uris[i])}' color='blue'>- link</link>",body_style)
            else:
                p=Paragraph(txt,body_style)
            line.append(p)
        wrapped.append(line)
    if tot:
        lbl="Total";val=""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: lbl,val=m.group(1).strip(),m.group(2)
        tr=[Paragraph(lbl,bl_style)]+[Spacer(1,0)]*(nc-2)+[Paragraph(val,br_style)]
        wrapped.append(tr)
    tbl=LongTable(wrapped,colWidths=cw,repeatRows=1)
    sc=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        sc+=("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")
    tbl.setStyle(TableStyle(sc)); elements += [tbl,Spacer(1,24)]
if grand_total and tables_info:
    last_hdr=tables_info[-1][0];nc=len(last_hdr)
    di=last_hdr.index("Description") if "Description" in last_hdr else 1
    dw=0.45*tw;ow=(tw-dw)/(nc-1) if nc>1 else tw
    cw=[dw if i==di else ow for i in range(nc)]
    gr=[Paragraph("Grand Total",bl_style)]+[Spacer(1,0)]*(nc-2)+[Paragraph(html.escape(grand_total),br_style)]
    gt=LongTable([gr],colWidths=cw)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
                            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    elements.append(gt)
doc.build(elements); pdf_buf.seek(0)

docx_buf=io.BytesIO()
docx_doc=Document()
sec=docx_doc.sections[0];sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17);sec.page_height=Inches(11)
sec.left_margin=Inches(0.5);sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5);sec.bottom_margin=Inches(0.5)
if 'logo' in locals():
    p=docx_doc.add_paragraph();r=p.add_run()
    img=Image.open(io.BytesIO(logo));rat=img.height/img.width
    wi=5;hi=wi*rat;r.add_picture(io.BytesIO(logo),width=Inches(wi));p.alignment=WD_TABLE_ALIGNMENT.CENTER
p=docx_doc.add_paragraph();p.alignment=WD_TABLE_ALIGNMENT.CENTER
r=p.add_run(proposal_title);r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(18);r.bold=True
docx_doc.add_paragraph()
TOTAL_W_IN=sec.page_width.inches-sec.left_margin.inches-sec.right_margin.inches
for hdr, rows, uris, tot in tables_info:
    n=len(hdr); di=hdr.index("Description") if "Description" in hdr else 1
    dw=0.45*TOTAL_W_IN;ow=(TOTAL_W_IN-dw)/(n-1) if n>1 else TOTAL_W_IN
    tbl=docx_doc.add_table(rows=1,cols=n,style="Table Grid");tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit=False;tbl.allow_autofit=False
    pr=tbl._element.xpath('./w:tblPr')
    if not pr:
        pr_el=OxmlElement('w:tblPr');tbl._element.insert(0,pr_el)
    else:
        pr_el=pr[0]
    tw=OxmlElement('w:tblW');tw.set(qn('w:w'),'5000');tw.set(qn('w:type'),'pct')
    ex=pr_el.xpath('./w:tblW')
    if ex:pr_el.remove(ex[0])
    pr_el.append(tw)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(dw if i==di else ow)
    for i,name in enumerate(hdr):
        cell=tbl.rows[0].cells[i];tc=cell._tc;tcPr=tc.get_or_add_tcPr()
        sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'F2F2F2');tcPr.append(sh)
        p=cell.paragraphs[0];p.text="";r=p.add_run(name)
        r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER;cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        cells=tbl.add_row().cells
        for j,val in enumerate(row):
            cell=cells[j];p=cell.paragraphs[0];p.text=""
            run=p.add_run(str(val));run.font.name=DEFAULT_SANS_FONT;run.font.size=Pt(9)
            if j==di and ridx<len(uris) and uris[ridx]:
                p.add_run(" ");add_hyperlink(p,uris[ridx],"- link",font_name=DEFAULT_SANS_FONT,font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT;cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        cells=tbl.add_row().cells
        lbl,amt="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total";amt=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m:lbl,mnt=m.group(1).strip(),m.group(2)
        lc=cells[0]
        if n>1:lc.merge(cells[n-2])
        p=lc.paragraphs[0];p.text="";r=p.add_run(lbl)
        r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT;lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=cells[n-1];p2=ac.paragraphs[0];p2.text="";r2=p2.add_run(amt)
        r2.font.name=DEFAULT_SERIF_FONT;r2.font.size=Pt(10);r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT;ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr=tables_info[-1][0];n=len(last_hdr)
    di=last_hdr.index("Description") if "Description" in last_hdr else 1
    dw=0.45*TOTAL_W_IN;ow=(TOTAL_W_IN-dw)/(n-1) if n>1 else TOTAL_W_IN
    tblg=docx_doc.add_table(rows=1,cols=n,style="Table Grid");tblg.alignment=WD_TABLE_ALIGNMENT.CENTER
    tblg.autofit=False;tblg.allow_autofit=False
    pr=tblg._element.xpath('./w:tblPr')
    if not pr:
        pr_el=OxmlElement('w:tblPr');tblg._element.insert(0,pr_el)
    else:
        pr_el=pr[0]
    tw=OxmlElement('w:tblW');tw.set(qn('w:w'),'5000');tw.set(qn('w:type'),'pct')
    ex=pr_el.xpath('./w:tblW')
    if ex:pr_el.remove(ex[0])
    pr_el.append(tw)
    for i,col in enumerate(tblg.columns):col.width=Inches(dw if i==di else ow)
    cells=tblg.rows[0].cells
    lc=cells[0]
    if n>1:lc.merge(cells[n-2])
    tc=lc._tc;tcPr=tc.get_or_add_tcPr();sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'E0E0E0');tcPr.append(sh)
    p=lc.paragraphs[0];p.text="";r=p.add_run("Grand Total")
    r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT;lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1];tc2=ac._tc;tcPr2=tc2.get_or_add_tcPr();sh2=OxmlElement('w:shd');sh2.set(qn('w:fill'),'E0E0E0');tcPr2.append(sh2)
    p2=ac.paragraphs[0];p2.text="";r2=p2.add_run(grand_total)
    r2.font.name=DEFAULT_SERIF_FONT;r2.font.size=Pt(10);r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT;ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1, c2 = st.columns(2)
if pdf_buf:
    with c1:
        st.download_button("📥 Download deliverable PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf", use_container_width=True)
else:
     with c1:
         st.error("PDF generation failed.")
if docx_buf:
    with c2:
        st.download_button("📥 Download deliverable DOCX", data=docx_buf, file_name="proposal_deliverable.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    with c2:
        st.error("Word document generation failed.")
