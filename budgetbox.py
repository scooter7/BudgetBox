import streamlit as st
import pdfplumber
import io
import requests
import fitz          # PyMuPDF for extracting link annotations
from fpdf import FPDF
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import re

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Open PDF in PyMuPDF to capture link annotations ───────────────────────────
doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
page_annotations = []
for page in doc_fitz:
    annots = []
    for a in page.annots() or []:
        if a.type[0] == 1 and a.uri:
            annots.append((a.rect, a.uri))
    page_annotations.append(annots)

# ─── Helpers ───────────────────────────────────────────────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

def extract_tables_and_links():
    """Use pdfplumber to extract tables, totals, and per-row URIs."""
    tables = []
    grand = None
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text() or "" for p in pdf.pages]
        title = next(
            (ln for pg in page_texts for ln in pg.splitlines() if "proposal" in ln.lower()),
            "Untitled Proposal"
        ).strip()

        used_totals = set()
        def find_total(pi):
            for ln in page_texts[pi].splitlines():
                if re.search(r'\btotal\b', ln, re.I) and re.search(r'\$\d', ln) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            annots = page_annotations[pi]
            for tbl in page.find_tables():
                data = tbl.extract()
                if len(data) < 2:
                    continue
                hdr = data[0]
                desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
                if desc_i is None:
                    continue

                # compute per-row link by slicing table.bbox
                x0,y0,x1,y1 = tbl.bbox
                nrows = len(data)
                band = (y1-y0)/nrows
                row_map = {}
                for rect, uri in annots:
                    midy = (rect.y0+rect.y1)/2
                    if y0 <= midy <= y1:
                        ridx = int((midy - y0)//band)
                        if 1 <= ridx < nrows:
                            row_map[ridx-1] = uri

                new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
                rows, uris = [], []
                for ridx, row in enumerate(data[1:], start=1):
                    if all(cell is None or not str(cell).strip() for cell in row):
                        continue
                    first = next((str(c).strip() for c in row if c), "")
                    if first.lower()=="total":
                        continue
                    strat, desc = split_cell_text(str(row[desc_i] or ""))
                    rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                    rows.append([strat, desc] + rest)
                    uris.append(row_map.get(ridx-1))
                tot = find_total(pi)
                tables.append((new_hdr, rows, uris, tot))

        # grand total
        for tx in reversed(page_texts):
            m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', tx, re.I|re.S)
            if m:
                grand = m.group(1)
                break

    return title, tables, grand

proposal_title, tables_info, grand_total = extract_tables_and_links()

# ─── Generate PDF via fpdf2 ────────────────────────────────────────────────────
pdf = FPDF("L", "pt", (17*72, 11*72))
pdf.set_auto_page_break(False)
pdf.add_page()
# margins
pdf.set_margins(48, 48, 48)

# register fonts
pdf.add_font("DMSerif", "", "fonts/DMSerifDisplay-Regular.ttf", uni=True)
pdf.add_font("DMSerif", "B","fonts/DMSerifDisplay-Regular.ttf", uni=True)
pdf.add_font("Barlow", "", "fonts/Barlow-Regular.ttf", uni=True)

# logo + title
try:
    logo_bytes = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5,
    ).content
    pdf.image(io.BytesIO(logo_bytes), x=pdf.l_margin, y=pdf.get_y(), w=150)
    pdf.ln(60)
except:
    pdf.ln(12)

pdf.set_font("DMSerif", "", 18)
pdf.cell(0, 24, proposal_title, ln=1, align="C")
pdf.ln(12)

usable_w = 17*72 - pdf.l_margin - pdf.r_margin

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    desc_w = 0.45 * usable_w
    other_w = (usable_w - desc_w) / (n - 1)
    widths = [desc_w if i==1 else other_w for i in range(n)]
    # header
    pdf.set_fill_color(242,242,242)
    pdf.set_text_color(0,0,0)
    pdf.set_font("DMSerif","",10)
    for w, h in zip(widths, hdr):
        pdf.cell(w, 20, h, border=1, align="C", fill=True)
    pdf.ln(20)
    # rows
    for row, link in zip(rows, uris):
        pdf.set_fill_color(255,255,255)
        pdf.set_font("Barlow","",9)
        for i, (w, txt) in enumerate(zip(widths, row)):
            if i==1 and link:
                pdf.set_text_color(0,0,255)
                pdf.set_font("Barlow","U",9)
                pdf.cell(w, 18, txt, border=1, link=link)
                pdf.set_text_color(0,0,0)
                pdf.set_font("Barlow","",9)
            else:
                align = "L" if i in (0,1) else "C"
                pdf.cell(w, 18, str(txt), border=1, align=align)
        pdf.ln(18)
    # total row
    if tot:
        pdf.set_font("DMSerif","B",10)
        label, val = re.split(r'\$\s*', tot, 1)
        pdf.set_text_color(0,0,0)
        for i, w in enumerate(widths):
            if i==0:
                pdf.cell(w, 20, label, border=1, align="L")
            elif i==n-1:
                pdf.cell(w, 20, f"${val.strip()}", border=1, align="R")
            else:
                pdf.cell(w, 20, "", border=1)
        pdf.ln(24)

# grand total standalone
if grand_total:
    pdf.set_font("DMSerif","B",12)
    pdf.set_text_color(0,0,0)
    pdf.cell(0, 24, f"Grand Total {grand_total}", border=1, ln=1, align="L")

pdf_buf = io.BytesIO(pdf.output(dest="S").encode("latin1"))

# ─── Build Word (unchanged) ──────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx = Document()
sec = docx.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width  = Inches(17)
sec.page_height = Inches(11)

try:
    p_logo = docx.add_paragraph()
    r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(4))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass

p_title = docx.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]
r.font.name = "DMSerif"
r.font.size = Pt(18)
docx.add_paragraph()

for hdr, rows, uris, tot in tables_info:
    n = len(hdr)
    desc_w = 0.45 * 17
    other_w = (17 - desc_w) / (n - 1)

    tblW = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblW.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, col in enumerate(tblW.columns):
        col.width = Inches(desc_w if idx==1 else other_w)

    # header shading
    for i, col_name in enumerate(hdr):
        cell = tblW.rows[0].cells[i]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text=""
        run = p.add_run(str(col_name))
        run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    # body rows
    for ridx, row in enumerate(rows):
        rc = tblW.add_row().cells
        for cidx, val in enumerate(row):
            p = rc[cidx].paragraphs[0]; p.text=""
            if cidx==1 and uris[ridx]:
                # hyperlink
                part = p.part
                rid = part.relate_to(
                    uris[ridx],
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                    is_external=True
                )
                hl = OxmlElement("w:hyperlink"); hl.set(qn("r:id"), rid)
                r = OxmlElement("w:r"); rPr = OxmlElement("w:rPr")
                c = OxmlElement("w:color"); c.set(qn("w:val"),"0000FF"); rPr.append(c)
                u = OxmlElement("w:u");   u.set(qn("w:val"),"single"); rPr.append(u)
                r.append(rPr)
                t = OxmlElement("w:t"); t.text = str(val); r.append(t)
                hl.append(r); p._p.append(hl)
                run = p.add_run(); run.font.name="Barlow"; run.font.size=Pt(9)
            else:
                run = p.add_run(str(val)); run.font.name="Barlow"; run.font.size=Pt(9)

    # total row
    if tot:
        label, val = re.split(r'\$\s*', tot,1)
        rc = tblW.add_row().cells
        for i, txt in enumerate([label]+[""]*(n-2)+[f"${val.strip()}"]):
            p = rc[i].paragraphs[0]; p.text=""
            run = p.add_run(txt)
            run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
            if i==0:    p.alignment=WD_TABLE_ALIGNMENT.LEFT
            elif i==n-1:p.alignment=WD_TABLE_ALIGNMENT.RIGHT
            else:       p.alignment=WD_TABLE_ALIGNMENT.CENTER

    docx.add_paragraph()

# Grand total row
if grand_total:
    n = len(tables_info[-1][0])
    tblG = docx.add_table(rows=1, cols=n, style="Table Grid")
    tblG.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, txt in enumerate(["Grand Total"]+[""]*(n-2)+[grand_total]):
        cell = tblG.rows[0].cells[idx]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd'); shd.set(qn('w:fill'),'F2F2F2'); tcPr.append(shd)
        p = cell.paragraphs[0]; p.text=""
        run = p.add_run(txt)
        run.font.name="DMSerif"; run.font.size=Pt(10); run.bold=True
        if idx==0:    p.alignment=WD_TABLE_ALIGNMENT.LEFT
        elif idx==n-1:p.alignment=WD_TABLE_ALIGNMENT.RIGHT
        else:         p.alignment=WD_TABLE_ALIGNMENT.CENTER

docx.save(docx_buf := io.BytesIO())
docx_buf.seek(0)

# ─── Download buttons ─────────────────────────────────────────────────────────
c1,c2 = st.columns(2)
with c1:
    st.download_button("📥 Download deliverable PDF",
                       data=pdf_buf,
                       file_name="proposal_deliverable.pdf",
                       mime="application/pdf",
                       use_container_width=True)
with c2:
    st.download_button("📥 Download deliverable DOCX",
                       data=docx_buf,
                       file_name="proposal_deliverable.docx",
                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       use_container_width=True)
