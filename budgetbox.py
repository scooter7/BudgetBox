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
    pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))
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
        sz  = OxmlElement('w:sz');   sz.set(qn('w:val'), str(int(font_size*2)))
        szCs= OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(font_size*2)))
        rPr.extend([sz, szCs])
    if bold:
        rPr.append(OxmlElement('w:b'))
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)

tables_info = []
grand_total  = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text(x_tolerance=1,y_tolerance=1) or "" for p in pdf.pages]
    first_lines = page_texts[0].splitlines() if page_texts else []
    pot = next((l for l in first_lines if "proposal" in l.lower() and len(l)>5), None)
    proposal_title = pot or (first_lines[0] if first_lines else "Untitled Proposal")
    used_totals = set()
    def find_total(pi):
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*\d', ln, re.I) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        links = page.hyperlinks
        for tbl in page.find_tables():
            data = tbl.extract(x_tolerance=1,y_tolerance=1)
            if not data or len(data)<2:
                continue
            raw0 = data[0]
            raw1 = data[1]
            if any(raw1) and all(isinstance(c,str) for c in raw1):
                hdr = []
                for a,b in zip(raw0,raw1):
                    parts = []
                    if a: parts.append(a.strip())
                    if b: parts.append(b.strip())
                    hdr.append(" ".join(parts))
                data = [hdr] + data[2:]
            else:
                hdr = [str(c).strip() if c else "" for c in raw0]
            desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
            if desc_i is None:
                continue
            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            bbox = page.bbox
            nrows = len(data)
            band = (bbox[3]-bbox[1])/nrows
            rmap = {}
            for rect,uri in links:
                mid = (rect.y0+rect.y1)/2
                if bbox[1]<mid<bbox[3]:
                    rid = int((mid-bbox[1])//band)-1
                    if 0<=rid<len(data)-1:
                        rmap[rid]=uri
            rows_data=[]
            row_links=[]
            tbl_total=None
            for ridx,row in enumerate(data[1:],start=0):
                cells=[str(c).strip() if c else "" for c in row]
                if all(not c for c in cells): continue
                first=cells[0].lower()
                if ("total" in first or "subtotal" in first) and any("$" in c for c in cells):
                    if tbl_total is None: tbl_total=cells
                    continue
                strat,desc=split_cell_text(cells[desc_i])
                rest=[cells[i] for i in range(len(cells)) if i!=desc_i]
                rows_data.append([strat,desc]+rest)
                row_links.append(rmap.get(ridx))
            if tbl_total is None:
                tbl_total=find_total(pi)
            if rows_data:
                tables_info.append((new_hdr,rows_data,row_links,tbl_total))

    for tx in reversed(page_texts):
        m=re.search(r'Grand Total.*?(\$\s*[\d,]+\.\d{2})',tx,re.I)
        if m:
            grand_total=m.group(1).strip()
            break

pdf_buf=io.BytesIO()
doc=SimpleDocTemplate(pdf_buf,pagesize=landscape((17*inch,11*inch)),leftMargin=0.5*inch,rightMargin=0.5*inch,topMargin=0.5*inch,bottomMargin=0.5*inch)
title_style  = ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
header_style = ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER)
body_style   = ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl_style     = ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br_style     = ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)
elements=[]
logo=None
try:
    logo_url="https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    resp=requests.get(logo_url,timeout=10);resp.raise_for_status()
    logo=resp.content
    img=Image.open(io.BytesIO(logo));ratio=img.height/img.width
    img_w=min(5*inch,doc.width);img_h=img_w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=img_w,height=img_h))
except:
    pass
elements+=[Spacer(1,12),Paragraph(html.escape(proposal_title),title_style),Spacer(1,24)]
total_w=doc.width

for hdr,rows,uris,tot in tables_info:
    ncols=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(ncols-1) if ncols>1 else total_w
    col_widths=[desc_w if i==desc_idx else other_w for i in range(ncols)]
    wrapped=[[Paragraph(html.escape(h),header_style) for h in hdr]]
    for ridx,row in enumerate(rows):
        line=[]
        for cidx,cell in enumerate(row):
            txt=html.escape(cell)
            if cidx==desc_idx and ridx<len(uris) and uris[ridx]:
                p=Paragraph(f"{txt} <link href='{html.escape(uris[ridx])}' color='blue'>- link</link>",body_style)
            else:
                p=Paragraph(txt,body_style)
            line.append(p)
        wrapped.append(line)
    if tot:
        lbl,val="", ""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"
            val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m:
                lbl,val=m.group(1).strip(),m.group(2)
        total_row=[Paragraph(lbl,bl_style)]+[Spacer(1,0)]*(ncols-2)+[Paragraph(val,br_style)]
        wrapped.append(total_row)
    tbl=LongTable(wrapped,colWidths=col_widths,repeatRows=1)
    cmds=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("VALIGN",(0,0),(-1,0),"MIDDLE"),("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        cmds+= [("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    elements+=[tbl,Spacer(1,24)]

if grand_total and tables_info:
    last_hdr=tables_info[-1][0]
    ncols=len(last_hdr)
    desc_idx=last_hdr.index("Description") if "Description" in last_hdr else 1
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(ncols-1) if ncols>1 else total_w
    col_widths=[desc_w if i==desc_idx else other_w for i in range(ncols)]
    row=[Paragraph("Grand Total",bl_style)]+[Spacer(1,0)]*(ncols-2)+[Paragraph(html.escape(grand_total),br_style)]
    gt=LongTable([row],colWidths=col_widths)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),("GRID",(0,0),(-1,-1),0.25,colors.grey),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

docx_buf=io.BytesIO()
docx_doc=Document()
sec=docx_doc.sections[0]
sec.orientation=WD_ORIENT.LANDSCAPE
sec.page_width=Inches(17)
sec.page_height=Inches(11)
sec.left_margin=Inches(0.5)
sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5)
sec.bottom_margin=Inches(0.5)

if logo:
    try:
        p_logo=docx_doc.add_paragraph()
        r_logo=p_logo.add_run()
        img=Image.open(io.BytesIO(logo))
        ratio=img.height/img.width
        w_in=5;h_in=w_in*ratio
        r_logo.add_picture(io.BytesIO(logo),width=Inches(w_in))
        p_logo.alignment=WD_TABLE_ALIGNMENT.CENTER
    except:
        pass

p_title=docx_doc.add_paragraph()
p_title.alignment=WD_TABLE_ALIGNMENT.CENTER
run=p_title.add_run(proposal_title)
run.font.name=DEFAULT_SERIF_FONT
run.font.size=Pt(18)
run.bold=True
docx_doc.add_paragraph()

TOTAL_W_IN=sec.page_width.inches-sec.left_margin.inches-sec.right_margin.inches

for hdr,rows,uris,tot in tables_info:
    ncols=len(hdr)
    if ncols==0: continue
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=0.45*TOTAL_W_IN
    other_w=(TOTAL_W_IN-desc_w)/(ncols-1) if ncols>1 else TOTAL_W_IN
    tbl=docx_doc.add_table(rows=1,cols=ncols,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit=False;tbl.allow_autofit=False
    tblPr_list=tbl._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr=OxmlElement('w:tblPr');tbl._element.insert(0,tblPr)
    else:
        tblPr=tblPr_list[0]
    tblW=OxmlElement('w:tblW');tblW.set(qn('w:w'),'5000');tblW.set(qn('w:type'),'pct')
    existing=tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    hdr_cells=tbl.rows[0].cells
    for i,name in enumerate(hdr):
        cell=hdr_cells[i]
        tc=cell._tc;tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd');shd.set(qn('w:fill'),'F2F2F2');tcPr.append(shd)
        p=cell.paragraphs[0];p.text=""
        r=p.add_run(name);r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx,row in enumerate(rows):
        rc=tbl.add_row().cells
        for cidx,val in enumerate(row):
            cell=rc[cidx]
            p=cell.paragraphs[0];p.text=""
            run=p.add_run(str(val));run.font.name=DEFAULT_SANS_FONT;run.font.size=Pt(9)
            if cidx==desc_idx and ridx<len(uris) and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p,uris[ridx],"- link",font_name=DEFAULT_SANS_FONT,font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        trow=tbl.add_row().cells
        label,amount="",""
        if isinstance(tot,list):
            label=tot[0] or "Total"
            amount=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: label,amount=m.group(1).strip(),m.group(2)
        label_cell=trow[0]
        if ncols>1: label_cell.merge(trow[ncols-2])
        p=label_cell.paragraphs[0];p.text=""
        r=p.add_run(label);r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT
        label_cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        amt_cell=trow[ncols-1]
        p2=amt_cell.paragraphs[0];p2.text=""
        r2=p2.add_run(amount);r2.font.name=DEFAULT_SERIF_FONT;r2.font.size=Pt(10);r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT
        amt_cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr=tables_info[-1][0];ncols=len(last_hdr)
    desc_idx=last_hdr.index("Description") if "Description" in last_hdr else 1
    desc_w=0.45*TOTAL_W_IN;other_w=(TOTAL_W_IN-desc_w)/(ncols-1) if ncols>1 else TOTAL_W_IN
    tblg=docx_doc.add_table(rows=1,cols=ncols,style="Table Grid")
    tblg.alignment=WD_TABLE_ALIGNMENT.CENTER;tblg.autofit=False;tblg.allow_autofit=False
    tblPr_list=tblg._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr=OxmlElement('w:tblPr');tblg._element.insert(0,tblPr)
    else:
        tblPr=tblPr_list[0]
    tblW=OxmlElement('w:tblW');tblW.set(qn('w:w'),'5000');tblW.set(qn('w:type'),'pct')
    existing=tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i,col in enumerate(tblg.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    cells=tblg.rows[0].cells
    label_cell=cells[0]
    if ncols>1: label_cell.merge(cells[ncols-2])
    tc=label_cell._tc;tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd');shd.set(qn('w:fill'),'E0E0E0');tcPr.append(shd)
    p=label_cell.paragraphs[0];p.text=""
    r=p.add_run("Grand Total");r.font.name=DEFAULT_SERIF_FONT;r.font.size=Pt(10);r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT
    label_cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    amt_cell=cells[ncols-1]
    tc2=amt_cell._tc;tcPr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd');shd2.set(qn('w:fill'),'E0E0E0');tcPr2.append(shd2)
    p2=amt_cell.paragraphs[0];p2.text=""
    r2=p2.add_run(grand_total);r2.font.name=DEFAULT_SERIF_FONT;r2.font.size=Pt(10);r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT
    amt_cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf=io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1,c2=st.columns(2)
with c1:
    st.download_button("📥 Download deliverable PDF",data=pdf_buf,file_name="proposal_deliverable.pdf",mime="application/pdf")
with c2:
    st.download_button("📥 Download deliverable DOCX",data=docx_buf,file_name="proposal_deliverable.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
