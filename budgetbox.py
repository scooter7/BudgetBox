import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
import io
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.text.run import Run
import re

# ─── Monkey-patch to support hyperlinks ───────────────────────────────────────
def run_add_hyperlink(self, url, text=None, color="0000FF", underline=True):
    if text is None:
        text = url
    paragraph = self._parent
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    r_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    c = OxmlElement("w:color"); c.set(qn("w:val"), color); rPr.append(c)
    if underline:
        u = OxmlElement("w:u"); u.set(qn("w:val"), "single"); rPr.append(u)
    r_elem.append(rPr)
    t = OxmlElement("w:t"); t.text = text; r_elem.append(t)
    hyperlink.append(r_elem)
    paragraph._p.append(hyperlink)
    return Run(r_elem, paragraph)

Run.add_hyperlink = run_add_hyperlink

# ─── Streamlit UI setup ───────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download a Word document with preserved hyperlinks.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Extract proposal title, tables, links ────────────────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
page_links = []
for page in doc_fitz:
    links = []
    for a in page.annots() or []:
        if a.type[0] == 1 and a.uri:
            links.append((a.rect, a.uri))
    page_links.append(links)

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    proposal_title = next((ln for pg in page_texts for ln in pg.splitlines() if "proposal" in ln.lower()), "Untitled Proposal").strip()
    tables_info = []
    used_totals = set()

    def find_total(pi):
        for ln in page_texts[pi].splitlines():
            if re.search(r'\\btotal\\b', ln, re.I) and re.search(r'\\$\\d', ln) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    for pi, page in enumerate(pdf.pages):
        annots = page_links[pi]
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue
            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            x0, y0, x1, y1 = tbl.bbox
            band_h = (y1 - y0) / len(data)
            row_links = {}
            for rect, uri in annots:
                midy = (rect.y0 + rect.y1) / 2
                if y0 <= midy <= y1:
                    ridx = int((midy - y0) // band_h)
                    if 1 <= ridx < len(data):
                        row_links[ridx-1] = uri

            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i != desc_i and h]
            rows, links = [], []
            for ridx, row in enumerate(data[1:], start=1):
                if all(cell is None or not str(cell).strip() for cell in row):
                    continue
                first = next((str(c).strip() for c in row if c), "")
                if first.lower() == "total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i,h in enumerate(hdr) if i != desc_i and h]
                rows.append([strat, desc] + rest)
                links.append(row_links.get(ridx - 1))

            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, links, tbl_total))

    # Grand total
    grand_total = None
    for txt in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\\$\\d[\\d,\\,]*\\.\\d{2})', txt, re.I|re.S)
        if m:
            grand_total = m.group(1)
            break

# ─── Build Word DOCX with hyperlinks ─────────────────────────────────────────
docx_buf = io.BytesIO()
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width = Inches(17)
sec.page_height = Inches(11)

try:
    logo = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png", timeout=5).content
    p_logo = doc.add_paragraph()
    r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo), width=Inches(4))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
except:
    pass

p_title = doc.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r = p_title.runs[0]
r.font.name = "DMSerif"
r.font.size = Pt(18)
doc.add_paragraph()

for hdr, rows, row_links, tbl_total in tables_info:
    n = len(hdr)
    tbl = doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(hdr):
        cell = tbl.rows[0].cells[i]
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(h); run.bold = True; run.font.size = Pt(10)
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    for ridx, row in enumerate(rows):
        cells = tbl.add_row().cells
        for cidx, val in enumerate(row):
            p = cells[cidx].paragraphs[0]; p.text = ""
            if cidx == 1 and row_links[ridx]:
                run = p.add_run()
                run.add_hyperlink(row_links[ridx], str(val))
            else:
                run = p.add_run(str(val))
            run.font.size = Pt(9)

    if tbl_total:
        lbl, amt = re.split(r'\\$\\s*', tbl_total, 1)
        amt = "$" + amt.strip()
        rc = tbl.add_row().cells
        for i, tv in enumerate([lbl]+[""]*(n-2)+[amt]):
            p = rc[i].paragraphs[0]; p.text = ""
            run = p.add_run(tv)
            run.bold = True; run.font.size = Pt(10)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT if i == 0 else (WD_TABLE_ALIGNMENT.RIGHT if i == n-1 else WD_TABLE_ALIGNMENT.CENTER)
    doc.add_paragraph()

if grand_total:
    hdr = tables_info[-1][0]
    n = len(hdr)
    tblg = doc.add_table(rows=1, cols=n, style="Table Grid")
    tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, tv in enumerate(["Grand Total"] + [""]*(n-2) + [grand_total]):
        cell = tblg.rows[0].cells[idx]
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(tv)
        run.bold = True; run.font.size = Pt(10)
        p.alignment = WD_TABLE_ALIGNMENT.LEFT if idx == 0 else (WD_TABLE_ALIGNMENT.RIGHT if idx == n-1 else WD_TABLE_ALIGNMENT.CENTER)

doc.save(docx_buf)
docx_buf.seek(0)

st.download_button(
    "📥 Download deliverable DOCX",
    data=docx_buf,
    file_name="proposal_deliverable.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)
