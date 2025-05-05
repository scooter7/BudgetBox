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
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return lines[0], description

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink'); hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05,0x63,0xC1); style.font.underline = True
        style.priority = 9; style.unhide_when_used = True
    style_element = OxmlElement('w:rStyle'); style_element.set(qn('w:val'),'Hyperlink'); rPr.append(style_element)
    if font_name:
        run_font = OxmlElement('w:rFonts'); run_font.set(qn('w:ascii'),font_name); run_font.set(qn('w:hAnsi'),font_name); rPr.append(run_font)
    if font_size:
        size = OxmlElement('w:sz'); size.set(qn('w:val'),str(int(font_size*2)))
        size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'),str(int(font_size*2)))
        rPr.append(size); rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'),'preserve'); t.text = text; new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    first_page = page_texts[0].splitlines() if page_texts else []
    title_candidate = next((ln for ln in first_page if "proposal" in ln.lower()), None)
    if title_candidate:
        proposal_title = title_candidate.strip()
    used_totals = set()
    def find_total(pi):
        if pi>=len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*\d', ln, re.I) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        for tbl in page.find_tables():
            data = tbl.extract(x_tolerance=1, y_tolerance=1)
            if not data or len(data)<2:
                continue
            hdr = [str(h).strip() if h else "" for h in data[0]]
            desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
            if desc_i is None:
                continue
            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows = []
            for row in data[1:]:
                cells = [str(c).strip() if c else "" for c in row]
                if all(not c for c in cells):
                    continue
                if cells[0].lower().startswith("total") and any("$" in c for c in cells):
                    continue
                strat, desc = split_cell_text(cells[desc_i])
                rest = [cells[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc] + rest)
            tbl_tot = find_total(pi)
            if rows:
                tables_info.append((new_hdr, rows, [None]*len(rows), tbl_tot))

    for tx in reversed(page_texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*\d[\d,]*\.\d{2})', tx, re.I)
        if m:
            grand_total = m.group(1).replace(" ", "")
            break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),
                        leftMargin=0.5*inch,rightMargin=0.5*inch,
                        topMargin=0.5*inch,bottomMargin=0.5*inch)
title_style = ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
header_style = ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
body_style = ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl_style = ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br_style = ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)

elements = []
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp = requests.get(logo_url,timeout=5); resp.raise_for_status()
    logo = resp.content
    img = Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    w=min(5*inch,doc.width); h=w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except:
    pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), title_style), Spacer(1,24)]
total_w = doc.width

for hdr, rows, uris, tot in tables_info:
    n=len(hdr)
    desc_idx = hdr.index("Description")
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(n-1) if n>1 else 0
    widths=[desc_w if i==desc_idx else other_w for i in range(n)]
    wrapped=[[Paragraph(html.escape(h),header_style) for h in hdr]]
    for ridx,row in enumerate(rows):
        line=[]
        for cidx,cell in enumerate(row):
            text=html.escape(cell)
            if cidx==desc_idx and uris[ridx]:
                p=Paragraph(f"{text} <link href='{html.escape(uris[ridx])}' color='blue'>- link</link>",body_style)
            else:
                p=Paragraph(text,body_style)
            line.append(p)
        wrapped.append(line)
    if tot:
        label="Total"; val=""
        if isinstance(tot,list):
            label=tot[0] or "Total"
            val=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m:
                label,val=m.group(1).strip(),m.group(2)
        row_el=[Paragraph(html.escape(label),bl_style)] + [Paragraph("",body_style)]*(n-2) + [Paragraph(html.escape(val),br_style)]
        wrapped.append(row_el)
    tbl=LongTable(wrapped,colWidths=widths,repeatRows=1)
    style=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
           ("GRID",(0,0),(-1,-1),0.25,colors.grey),
           ("VALIGN",(0,0),(-1,0),"MIDDLE"),
           ("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        style += [("SPAN",(0,-1),(-2,-1)),
                  ("ALIGN",(0,-1),(-2,-1),"LEFT"),
                  ("ALIGN",(-1,-1),(-1,-1),"RIGHT")]
    tbl.setStyle(TableStyle(style))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; n=len(last_hdr)
    desc_idx=last_hdr.index("Description")
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(n-1) if n>1 else total_w
    widths=[desc_w if i==desc_idx else other_w for i in range(n)]
    row=[Paragraph("Grand Total",bl_style)] + [Paragraph("",body_style)]*(n-2) + [Paragraph(html.escape(grand_total),br_style)]
    gt=LongTable([row],colWidths=widths)
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

docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Inches(17); sec.page_height = Inches(11)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5)
sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)

try:
    if logo:
        p=docx_doc.add_paragraph(); r=p.add_run()
        img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
        w_in=5; h_in=w_in*ratio
        r.add_picture(io.BytesIO(logo),width=Inches(w_in))
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass

pt=docx_doc.add_paragraph(); pt.alignment=WD_TABLE_ALIGNMENT.CENTER
r=pt.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()
TOTAL_W = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, uris, tot in tables_info:
    n=len(hdr)
    desc_idx=hdr.index("Description")
    desc_w_in=0.45*TOTAL_W
    other_w_in=(TOTAL_W-desc_w_in)/(n-1) if n>1 else TOTAL_W
    tbl = docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit=False; tbl.allow_autofit=False
    tblPr=tbl._element.xpath('./w:tblPr')
    if not tblPr:
        pr=OxmlElement('w:tblPr'); tbl._element.insert(0,pr)
    else:
        pr=tblPr[0]
    wtag=OxmlElement('w:tblW'); wtag.set(qn('w:w'),'5000'); wtag.set(qn('w:type'),'pct')
    old=pr.xpath('./w:tblW')
    if old: pr.remove(old[0])
    pr.append(wtag)
    for i,col in enumerate(tbl.columns):
        col.width = Inches(desc_w_in if i==desc_idx else other_w_in)
    hdr_cells=tbl.rows[0].cells
    for i,name in enumerate(hdr):
        cell=hdr_cells[i]; tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text=""
        run=p.add_run(name); run.font.name=DEFAULT_SERIF_FONT; run.font.size=Pt(10); run.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        cells=tbl.add_row().cells
        for cidx,val in enumerate(row):
            c=cells[cidx]; p=c.paragraphs[0]; p.text=""
            run=p.add_run(str(val)); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            if cidx==desc_idx and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p,uris[ridx],"- link",font_name=DEFAULT_SANS_FONT,font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        trow=tbl.add_row().cells
        label="Total"; amt=""
        if isinstance(tot,list):
            label=tot[0] or "Total"; amt=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: label,amt=m.group(1).strip(),m.group(2)
        lc=trow[0]
        if n>1: lc.merge(trow[n-2])
        p=lc.paragraphs[0]; p.text=""; r=p.add_run(label); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=trow[n-1]; p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(amt); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr=tables_info[-1][0]; n=len(last_hdr)
    desc_idx=last_hdr.index("Description")
    desc_w_in=0.45*TOTAL_W; other_w_in=(TOTAL_W-desc_w_in)/(n-1) if n>1 else TOTAL_W
    tblg=docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tblg.alignment=WD_TABLE_ALIGNMENT.CENTER; tblg.autofit=False; tblg.allow_autofit=False
    pr=tblg._element.xpath('./w:tblPr')
    if not pr:
        prn=OxmlElement('w:tblPr'); tblg._element.insert(0,prn)
    else:
        prn=pr[0]
    wtag=OxmlElement('w:tblW'); wtag.set(qn('w:w'),'5000'); wtag.set(qn('w:type'),'pct')
    old=prn.xpath('./w:tblW'); 
    if old: prn.remove(old[0])
    prn.append(wtag)
    for i,col in enumerate(tblg.columns):
        col.width=Inches(desc_w_in if i==desc_idx else other_w_in)
    cells=tblg.rows[0].cells
    lc=cells[0]
    if n>1: lc.merge(cells[n-2])
    tc=lc._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'E0E0E0'); tcPr.append(shd)
    p=lc.paragraphs[0]; p.text=""; r=p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]; tc2=ac._tc; tcPr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd'); shd2.set(qn('w:fill'),'E0E0E0'); tcPr2.append(shd2)
    p2=ac.paragraphs[0]; p2.text=""; r2=p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

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
