# -*- coding: utf-8 -*-
import io, re, html, camelot, pdfplumber, requests, streamlit as st
from PIL import Image
from docx import Document
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage

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
if not uploaded: st.stop()
pdf_bytes = uploaded.read()

def split_cell_text(raw):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines: return "", ""
    desc = " ".join(lines[1:]); desc = re.sub(r'\s+', ' ', desc).strip()
    return lines[0], desc

def add_hyperlink(par, url, text, font_name=None, font_size=None, bold=None):
    part = par.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyp = OxmlElement('w:hyperlink'); hyp.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r'); rPr = OxmlElement('w:rPr')
    styles = par.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1); style.font.underline = True
        style.priority = 9; style.unhide_when_used = True
    e = OxmlElement('w:rStyle'); e.set(qn('w:val'), 'Hyperlink'); rPr.append(e)
    if font_name:
        rf = OxmlElement('w:rFonts'); rf.set(qn('w:ascii'), font_name); rf.set(qn('w:hAnsi'), font_name); rPr.append(rf)
    if font_size:
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(font_size*2)))
        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(font_size*2))); rPr.append(sz); rPr.append(szCs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text; new_run.append(t)
    hyp.append(new_run); par._p.append(hyp)
    return docx.text.run.Run(new_run, par)

first_table = None
try:
    tables = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n")
    if tables and len(tables) > 0:
        df = tables[0].df; raw = df.values.tolist()
        if len(raw) > 1:
            h1, h2 = raw[0], raw[1]
            combined = []
            for j in range(len(h1)):
                if h1[j].strip() and h1[j].lower() != "none":
                    combined.append(h1[j].strip())
                else:
                    combined.append(h2[j].strip())
            keep = [j for j,v in enumerate(combined) if v and v.lower() != "none"]
            hdr = [combined[j] for j in keep]
            hdr[0], hdr[1] = "Strategy", "Description"
            new_raw = [hdr]
            for row in raw[2:]:
                new_raw.append([row[j] for j in keep])
            first_table = new_raw
except:
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
    lines = page_texts[0].splitlines() if page_texts else []
    pt = next((ln.strip() for ln in lines if "proposal" in ln.lower() and len(ln.strip())>5), None)
    if pt: proposal_title = pt
    elif lines: proposal_title = lines[0].strip()
    used = set()
    def find_total(pi):
        if pi>=len(page_texts): return None
        for ln in page_texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used:
                used.add(ln); return ln.strip()
        return None

    for pi,page in enumerate(pdf.pages):
        if pi==0 and first_table:
            data = first_table; links = []; source="camelot"
        else:
            source="plumber"; data=None
            for tbl in page.find_tables():
                d=tbl.extract(x_tolerance=1, y_tolerance=1)
                if d and len(d)>1:
                    data=d; links=page.hyperlinks; break
        if not data or len(data)<2: continue

        hdr = [str(h).strip() for h in data[0]]
        desc_i = next((i for i,h in enumerate(hdr) if "description" in h.lower()), None)
        if desc_i is None:
            desc_i = next((i for i,h in enumerate(hdr) if len(h)>10), None)
            if desc_i is None: continue

        if source=="plumber":
            vi = [0, desc_i] + [i for i,h in enumerate(hdr) if i not in (0,desc_i) and h.strip().lower()!="none"]
            vi = sorted(dict.fromkeys(vi))
            new_hdr = [hdr[i] for i in vi]
            new_hdr[0], new_hdr[1] = "Strategy","Description"
        else:
            new_hdr = data[0]

        rows, row_links = [], []
        table_total = None

        if source=="plumber":
            desc_links = {}
            tbl = page.find_tables()[0]
            for r,row_obj in enumerate(tbl.rows):
                if r==0: continue
                if desc_i < len(row_obj.cells):
                    x0,top,x1,bottom = row_obj.cells[desc_i]
                    for link in links:
                        if all(k in link for k in ("x0","x1","top","bottom","uri")):
                            if not (link["x1"]<x0 or link["x0"]>x1 or link["bottom"]<top or link["top"]>bottom):
                                desc_links[r] = link["uri"]; break

            for ridx,row in enumerate(data[1:], start=1):
                cells=[str(c).strip() for c in row]
                if not any(cells): continue
                fc=cells[0].lower()
                if ("total" in fc or "subtotal" in fc) and any("$" in c for c in cells):
                    if table_total is None: table_total = cells
                    continue
                strat,desc = split_cell_text(cells[desc_i] if desc_i<len(cells) else "")
                rest = [cells[i] for i in vi[2:]]
                rows.append([strat,desc]+rest)
                row_links.append(desc_links.get(ridx))
        else:
            for ridx,row in enumerate(data[1:], start=1):
                cells=[str(c).strip() for c in row]
                if not any(cells): continue
                fc=cells[0].lower()
                if ("total" in fc or "subtotal" in fc) and any("$" in c for c in cells):
                    if table_total is None: table_total = cells
                    continue
                strat,desc = split_cell_text(cells[1])
                rows.append([strat,desc]+cells[2:])
                row_links.append(None)

        if table_total and isinstance(table_total,list) and source=="plumber":
            table_total = [table_total[i] for i in vi]
        if table_total is None:
            table_total = find_total(pi)

        tables_info.append((new_hdr, rows, row_links, table_total))

    for tx in reversed(page_texts):
        m=re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I|re.S)
        if m:
            grand_total = m.group(1).replace(" ",""); break

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch,11*inch)),
    leftMargin=0.5*inch,rightMargin=0.5*inch,topMargin=0.5*inch,bottomMargin=0.5*inch)
ts=ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hs=ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER,textColor=colors.black)
bs=ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
bl=ParagraphStyle("BL",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_LEFT,spaceBefore=6)
br=ParagraphStyle("BR",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_RIGHT,spaceBefore=6)
elements=[]
try:
    url="https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    r=requests.get(url,timeout=10); r.raise_for_status()
    logo=r.content; img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
    w=min(5*inch,doc.width); h=w*ratio
    elements.append(RLImage(io.BytesIO(logo),width=w,height=h))
except: pass
elements += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
total_w = doc.width

for hdr, rows, links, tot in tables_info:
    n=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(n-1) if n>1 else total_w
    col_ws=[desc_w if i==desc_idx else other_w for i in range(n)]
    wrapped=[[Paragraph(html.escape(h), hs) for h in hdr]]
    for i,row in enumerate(rows):
        line=[]
        for j,cell in enumerate(row):
            txt=html.escape(cell)
            if j==desc_idx and links[i]:
                line.append(Paragraph(f"{txt} <link href='{html.escape(links[i])}' color='blue'>- link</link>", bs))
            else:
                line.append(Paragraph(txt, bs))
        wrapped.append(line)
    if tot:
        lbl="Total"; val=""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; val=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})', tot)
            if m: lbl,val=m.group(1).strip(),m.group(2)
        tr=[Paragraph(lbl,bl)]+[Paragraph("",bs)]*(n-2)+[Paragraph(val,br)]
        wrapped.append(tr)
    tbl=LongTable(wrapped,colWidths=col_ws,repeatRows=1)
    cmds=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
          ("GRID",(0,0),(-1,-1),0.25,colors.grey),
          ("VALIGN",(0,0),(-1,0),"MIDDLE"),
          ("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot and n>1:
        cmds+=[("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),
               ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    elements += [tbl, Spacer(1,24)]

if grand_total and tables_info:
    hdr=tables_info[-1][0]; n=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=total_w*0.45
    other_w=(total_w-desc_w)/(n-1) if n>1 else total_w
    col_ws=[desc_w if i==desc_idx else other_w for i in range(n)]
    row=[Paragraph("Grand Total",bl)]+[Paragraph("",bs)]*(n-2)+[Paragraph(html.escape(grand_total),br)]
    gt=LongTable([row],colWidths=col_ws)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
                            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    elements.append(gt)

doc.build(elements)
pdf_buf.seek(0)

docx_buf = io.BytesIO()
docx_doc = Document()
sec=docx_doc.sections[0]
sec.orientation=WD_ORIENT.LANDSCAPE; sec.page_width=Inches(17); sec.page_height=Inches(11)
sec.left_margin=Inches(0.5); sec.right_margin=Inches(0.5)
sec.top_margin=Inches(0.5); sec.bottom_margin=Inches(0.5)

if 'logo' in locals():
    try:
        p=docx_doc.add_paragraph(); r=p.add_run()
        img=Image.open(io.BytesIO(logo)); ratio=img.height/img.width
        w_in=5; h_in=w_in*ratio
        r.add_picture(io.BytesIO(logo), width=Inches(w_in))
        p.alignment=WD_TABLE_ALIGNMENT.CENTER
    except: pass

pt=docx_doc.add_paragraph(); pt.alignment=WD_TABLE_ALIGNMENT.CENTER
rt=pt.add_run(proposal_title); rt.font.name=DEFAULT_SERIF_FONT; rt.font.size=Pt(18); rt.bold=True
docx_doc.add_paragraph()

TOTAL_W_IN=sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, links, tot in tables_info:
    n=len(hdr)
    if n<1: continue
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=0.45*TOTAL_W_IN
    other_w=(TOTAL_W_IN-desc_w)/(n-1) if n>1 else TOTAL_W_IN
    tbl=docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; tbl.allow_autofit=False; tbl.autofit=False
    tblPr_list=tbl._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr=OxmlElement('w:tblPr'); tbl._element.insert(0,tblPr)
    else:
        tblPr=tblPr_list[0]
    tblW=OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    existing=tblPr.xpath('./w:tblW'); 
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i,col in enumerate(tbl.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    hdr_cells=tbl.rows[0].cells
    for i,name in enumerate(hdr):
        cell=hdr_cells[i]; tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p=cell.paragraphs[0]; p.text=""
        r=p.add_run(name); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for ridx, row in enumerate(rows):
        rc=tbl.add_row().cells
        for cidx,val in enumerate(row):
            cell=rc[cidx]; p=cell.paragraphs[0]; p.text=""
            run=p.add_run(str(val)); run.font.name=DEFAULT_SANS_FONT; run.font.size=Pt(9)
            if cidx==desc_idx and ridx<len(links) and links[ridx]:
                p.add_run(" "); add_hyperlink(p,links[ridx],"- link",font_name=DEFAULT_SANS_FONT,font_size=9)
            p.alignment=WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.TOP
    if tot:
        trow=tbl.add_row().cells; lbl,amt="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"; amt=next((c for c in reversed(tot) if "$" in c), "")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: lbl,amt=m.group(1).strip(),m.group(2)
        lc=trow[0]
        if n>1: lc.merge(trow[n-2])
        p=lc.paragraphs[0]; p.text=""
        r=p.add_run(lbl); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
        p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        ac=trow[n-1]; p2=ac.paragraphs[0]; p2.text=""
        r2=p2.add_run(amt); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
        p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    docx_doc.add_paragraph()

if grand_total and tables_info:
    hdr=tables_info[-1][0]; n=len(hdr)
    desc_idx=hdr.index("Description") if "Description" in hdr else 1
    desc_w=0.45*TOTAL_W_IN
    other_w=(TOTAL_W_IN-desc_w)/(n-1) if n>1 else TOTAL_W_IN
    tblg=docx_doc.add_table(rows=1,cols=n,style="Table Grid")
    tblg.alignment=WD_TABLE_ALIGNMENT.CENTER; tblg.allow_autofit=False; tblg.autofit=False
    tblPr_list=tblg._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr=OxmlElement('w:tblPr'); tblg._element.insert(0,tblPr)
    else:
        tblPr=tblPr_list[0]
    tblW=OxmlElement('w:tblW'); tblW.set(qn('w:w'),'5000'); tblW.set(qn('w:type'),'pct')
    existing=tblPr.xpath('./w:tblW')
    if existing: tblPr.remove(existing[0])
    tblPr.append(tblW)
    for i,col in enumerate(tblg.columns):
        col.width=Inches(desc_w if i==desc_idx else other_w)
    cells=tblg.rows[0].cells
    lc=cells[0]
    if n>1: lc.merge(cells[n-2])
    tc=lc._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'E0E0E0'); tcPr.append(shd)
    p=lc.paragraphs[0]; p.text=""
    r=p.add_run("Grand Total"); r.font.name=DEFAULT_SERIF_FONT; r.font.size=Pt(10); r.bold=True
    p.alignment=WD_TABLE_ALIGNMENT.LEFT; lc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    ac=cells[n-1]; tc2=ac._tc; tcPr2=tc2.get_or_add_tcPr()
    shd2=OxmlElement('w:shd'); shd2.set(qn('w:fill'),'E0E0E0'); tcPr2.append(shd2)
    p2=ac.paragraphs[0]; p2.text=""
    r2=p2.add_run(grand_total); r2.font.name=DEFAULT_SERIF_FONT; r2.font.size=Pt(10); r2.bold=True
    p2.alignment=WD_TABLE_ALIGNMENT.RIGHT; ac.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER

docx_buf = io.BytesIO()
docx_doc.save(docx_buf)
docx_buf.seek(0)

c1, c2 = st.columns(2)
c1.download_button("📥 Download deliverable PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf", use_container_width=True)
c2.download_button("📥 Download deliverable DOCX", data=docx_buf, file_name="proposal_deliverable.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
