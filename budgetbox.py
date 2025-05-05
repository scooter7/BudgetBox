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
    desc = " ".join(lines[1:])
    return lines[0], re.sub(r'\s+', ' ', desc).strip()

def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    part = paragraph.part
    rid = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    link = OxmlElement('w:hyperlink')
    link.set(qn('r:id'), rid)
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        style.font.underline = True
        style.priority = 9
        style.unhide_when_used = True
    rs = OxmlElement('w:rStyle')
    rs.set(qn('w:val'), 'Hyperlink')
    rPr.append(rs)
    if font_name:
        rf = OxmlElement('w:rFonts')
        rf.set(qn('w:ascii'), font_name)
        rf.set(qn('w:hAnsi'), font_name)
        rPr.append(rf)
    if font_size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(font_size * 2)))
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.extend([sz, szCs])
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    r.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    r.append(t)
    link.append(r)
    paragraph._p.append(link)
    return docx.text.run.Run(r, paragraph)

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
    lines0 = texts[0].splitlines() if texts else []
    pt = next((l for l in lines0 if "proposal" in l.lower() and len(l) > 5), None)
    if pt: proposal_title = pt.strip()
    elif lines0: proposal_title = lines0[0].strip()

    used = set()
    def find_total(pi):
        if pi >= len(texts): return None
        for ln in texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used:
                used.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        links = page.hyperlinks
        for tbl in page.find_tables():
            data = tbl.extract(x_tolerance=1, y_tolerance=1)
            if not data or len(data) < 2: continue
            hdr = [str(h).strip() if h else "" for h in data[0]]
            desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
            if desc_i is None:
                desc_i = next((i for i,h in enumerate(hdr) if len(h)>10), None)
                if desc_i is None: continue

            desc_links = {}
            if hasattr(tbl, 'rows'):
                for r_idx, row_obj in enumerate(tbl.rows):
                    if r_idx==0: continue
                    if desc_i < len(row_obj.cells):
                        x0, y0, x1, y1 = row_obj.cells[desc_i]
                        for link in links:
                            if 'uri' not in link: continue
                            lx0, ly0, lx1, ly1 = link['x0'], link['top'], link['x1'], link['bottom']
                            if not (lx1<x0 or lx0>x1 or ly1<y0 or ly0>y1):
                                desc_links[r_idx] = link['uri']
                                break

            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows_data = []; uri_list = []; total_info = None

            for ridx, row in enumerate(data[1:], start=1):
                cells = [str(c).strip() if c else "" for c in row]
                if all(not c for c in cells): continue
                low = cells[0].lower()
                if ("total" in low or "subtotal" in low) and any("$" in c for c in cells):
                    if total_info is None: total_info = cells
                    continue
                strat, desc = split_cell_text(cells[desc_i] if desc_i<len(cells) else "")
                rest = [cells[i] for i,h in enumerate(hdr) if i!=desc_i and hdr[i]]
                rows_data.append([strat, desc]+rest)
                uri_list.append(desc_links.get(ridx))

            if total_info is None:
                total_info = find_total(pi)
            if rows_data:
                tables_info.append((new_hdr, rows_data, uri_list, total_info))

    for tx in reversed(texts):
        m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1).replace(" ","")
            break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),
                        leftMargin=0.5*inch, rightMargin=0.5*inch,
                        topMargin=0.5*inch, bottomMargin=0.5*inch)

title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT)
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT)

elements = []
logo = None
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    r = requests.get(logo_url, timeout=5); r.raise_for_status()
    logo = r.content
    img = Image.open(io.BytesIO(logo))
    ratio = img.height/img.width
    w = min(5*inch, doc.width)
    h = w*ratio
    elements.append(RLImage(io.BytesIO(logo), width=w, height=h))
except:
    pass

elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), title_style), Spacer(1,24)]
page_w = doc.width

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    desc_idx = hdr.index("Description") if "Description" in hdr else 1
    desc_w = page_w*0.45
    other_w = (page_w-desc_w)/(n-1) if n>1 else page_w
    widths = [desc_w if i==desc_idx else other_w for i in range(n)]

    wrapped = [[Paragraph(html.escape(h), header_style) for h in hdr]]
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            txt = html.escape(cell)
            if cidx==desc_idx and uris[ridx]:
                p = Paragraph(f"{txt} <link href='{html.escape(uris[ridx])}' color='blue'>- link</link>", body_style)
            else:
                p = Paragraph(txt, body_style)
            line.append(p)
        wrapped.append(line)

    if tot:
        lbl,val = "Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"
            val=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl,val=m.group(1).strip(),m.group(2)
        row = [Paragraph(lbl,bl_style)] + [Paragraph("",body_style)]*(n-2) + [Paragraph(val,br_style)]
        wrapped.append(row)

    tbl = LongTable(wrapped, colWidths=widths, repeatRows=1, splitByRow=1)
    cmds = [
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,0),"MIDDLE"),
        ("VALIGN",(0,1),(-1,-1),"TOP")
    ]
    if tot:
        cmds += [("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    last_hdr = tables_info[-1][0]
    n = len(last_hdr)
    desc_idx = last_hdr.index("Description") if "Description" in last_hdr else 1
    desc_w = page_w*0.45
    other_w = (page_w-desc_w)/(n-1) if n>1 else page_w
    widths = [desc_w if i==desc_idx else other_w for i in range(n)]
    row = [Paragraph("Grand Total",bl_style)] + [Paragraph("",body_style)]*(n-2) + [Paragraph(html.escape(grand_total),br_style)]
    gt = LongTable([row], colWidths=widths, splitByRow=1)
    gt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
        ("GRID",(0,0),(-1,-1),0.25,colors.grey),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")
    ]))
    elements.append(gt)

try:
    doc.build(elements)
    pdf_buf.seek(0)
except Exception as e:
    st.error(f"PDF build error: {e}")
    pdf_buf = None

docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(17); sec.page_height = Inches(11)
sec.left_margin = Inches(0.5); sec.right_margin=Inches(0.5)
sec.top_margin  = Inches(0.5); sec.bottom_margin=Inches(0.5)

if logo:
    try:
        p = docx_doc.add_paragraph(); r = p.add_run()
        img = Image.open(io.BytesIO(logo)); ratio=img.height/img.width
        wi = 5; hi = wi*ratio
        r.add_picture(io.BytesIO(logo), width=Inches(wi))
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
    except: pass

ptp = docx_doc.add_paragraph(); ptp.alignment = WD_TABLE_ALIGNMENT.CENTER
r = ptp.add_run(proposal_title); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(18); r.bold=True
docx_doc.add_paragraph()

TOTAL_W = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    if n==0: continue
    desc_idx = hdr.index("Description") if "Description" in hdr else 1
    dw = 0.45*TOTAL_W; ow = (TOTAL_W-dw)/(n-1) if n>1 else TOTAL_W
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit=False; tbl.allow_autofit=False

    tp = tbl._element.xpath('./w:tblPr')
    if not tp:
        tpr = OxmlElement('w:tblPr'); tbl._element.insert(0,tpr)
    else:
        tpr = tp[0]
    tw = OxmlElement('w:tblW'); tw.set(qn('w:w'),'5000'); tw.set(qn('w:type'),'pct')
    ex = tpr.xpath('./w:tblW')
    if ex: tpr.remove(ex[0])
    tpr.append(tw)

    for i,col in enumerate(tbl.columns):
        col.width = Inches(dw if i==desc_idx else ow)

    hdr_cells = tbl.rows[0].cells
    for i,name in enumerate(hdr):
        cell = hdr_cells[i]; tc=cell._tc; pr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); pr.append(shd)
        p = cell.paragraphs[0]; p.text=''
        run = p.add_run(name); run.font.name=DEFAULT_SERIF_FONT; run.font.size=Pt(10); run.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for ridx,row in enumerate(rows):
        rc=tbl.add_row().cells
        for cidx,val in enumerate(row):
            cell = rc[cidx]; p=cell.paragraphs[0]; p.text=''
            run = p.add_run(val); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            if cidx==desc_idx and ridx < len(uris) and uris[ridx]:
                p.add_run(" ")
                add_hyperlink(p, uris[ridx], "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    if tot:
        trow = tbl.add_row().cells
        lbl,amt="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"
            amt=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl,amt=m.group(1).strip(),m.group(2)
        lc = trow[0]
        if n>1: lc.merge(trow[n-2])
        p=lc.paragraphs[0]; p.text=''
        r=p.add_run(lbl); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac = trow[n-1]; p2=ac.paragraphs[0]; p2.text=''
        r2=p2.add_run(amt); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr = tables_info[-1][0]; n=len(last_hdr)
    desc_idx = last_hdr.index("Description") if "Description" in last_hdr else 1
    dw = 0.45*TOTAL_W; ow=(TOTAL_W-dw)/(n-1) if n>1 else TOTAL_W
    gt = docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    gt.alignment=WD_TABLE_ALIGNMENT.CENTER; gt.autofit=False; gt.allow_autofit=False

    tp = gt._element.xpath('./w:tblPr')
    if not tp:
        tpr = OxmlElement('w:tblPr'); gt._element.insert(0,tpr)
    else:
        tpr = tp[0]
    tw = OxmlElement('w:tblW'); tw.set(qn('w:w'),'5000'); tw.set(qn('w:type'),'pct')
    ex = tpr.xpath('./w:tblW')
    if ex: tpr.remove(ex[0])
    tpr.append(tw)

    for i,col in enumerate(gt.columns):
        col.width = Inches(dw if i==desc_idx else ow)

    cells = gt.rows[0].cells
    lc = cells[0]
    if n>1: lc.merge(cells[n-2])
    tc=lc._tc; pr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'E0E0E0'); pr.append(shd)
    p=lc.paragraphs[0]; p.text=''
    r=p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]; tc2=ac._tc; pr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd'); shd2.set(qn('w:fill'),'E0E0E0'); pr2.append(shd2)
    p2=ac.paragraphs[0]; p2.text=''
    r2=p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except:
    docx_buf = None

c1, c2 = st.columns(2)
if pdf_buf:
    c1.download_button("📥 Download PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf")
else:
    c1.error("PDF generation failed.")
if docx_buf:
    c2.download_button("📥 Download DOCX", data=docx_buf, file_name="proposal_deliverable.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml")
else:
    c2.error("Word document generation failed.")
