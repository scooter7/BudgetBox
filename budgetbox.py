import streamlit as st
import pdfplumber
import fitz                       # PyMuPDF
import io
import requests
import re

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.text.run import Run

# ─── Monkey-patch Run to add hyperlinks ────────────────────────────────────────
def run_add_hyperlink(self, url, text=None, color="0000FF", underline=True):
    """
    Add a hyperlink to this Run’s parent paragraph.
    Usage:
        run = paragraph.add_run()
        run.add_hyperlink("https://example.com", "Click here")
    """
    if text is None:
        text = url

    paragraph = self._parent
    part = paragraph.part
    r_id = part.relate_to(
        url,
        RELATIONSHIP_TYPE.HYPERLINK,
        is_external=True
    )

    # build <w:hyperlink> element
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)

    # build run inside the hyperlink
    r_elem = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    # apply color
    c = OxmlElement("w:color")
    c.set(qn("w:val"), color)
    rPr.append(c)
    # apply underline
    if underline:
        u = OxmlElement("w:u")
        u.set(qn("w:val"), "single")
        rPr.append(u)
    r_elem.append(rPr)

    # add the text element
    t = OxmlElement("w:t")
    t.text = text
    r_elem.append(t)

    # wrap and append
    hyperlink.append(r_elem)
    paragraph._p.append(hyperlink)

    return Run(r_elem, paragraph)

Run.add_hyperlink = run_add_hyperlink

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal → Word", layout="wide")
st.title("🔄 Proposal to Word with Hyperlinks")
st.write("Upload a vertically formatted proposal PDF and download a Word file with preserved links.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Fetch logo once ───────────────────────────────────────────────────────────
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
try:
    logo_bytes = requests.get(LOGO_URL, timeout=5).content
except:
    logo_bytes = None

# ─── 1) Extract link annotations via PyMuPDF ──────────────────────────────────
doc_src = fitz.open(stream=pdf_bytes, filetype="pdf")
page_annotations = []
for page in doc_src:
    annots = []
    for a in page.annots() or []:
        if a.type[0] == 1 and a.uri:
            annots.append((a.rect, a.uri))
    page_annotations.append(annots)

# ─── 2) Extract tables & map links via pdfplumber ─────────────────────────────
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

tables_info = []
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]

    # find proposal title
    proposal_title = next(
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

            # map annotation to row
            x0,y0,x1,y1 = tbl.bbox
            band_h = (y1 - y0) / len(data)
            row_links = {}
            for rect, uri in annots:
                midy = (rect.y0 + rect.y1)/2
                if y0 <= midy <= y1:
                    ridx = int((midy - y0)//band_h)
                    if 1 <= ridx < len(data):
                        row_links[ridx-1] = uri

            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
            rows, links = [], []
            for ridx, row in enumerate(data[1:], start=1):
                if all(not str(c).strip() for c in row if c):
                    continue
                if next((str(c).strip() for c in row if c), "").lower()=="total":
                    continue
                strat, desc = split_cell_text(str(row[desc_i] or ""))
                rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                rows.append([strat, desc] + rest)
                links.append(row_links.get(ridx-1))

            tbl_total = find_total(pi)
            tables_info.append((new_hdr, rows, links, tbl_total))

    # grand total
    grand_total = None
    for txt in reversed(page_texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', txt, re.I|re.S)
        if m:
            grand_total = m.group(1)
            break

# ─── 3) Build Word document ───────────────────────────────────────────────────
doc = Document()
sec = doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = Inches(17), Inches(11)

# logo + title
if logo_bytes:
    p_logo = doc.add_paragraph()
    r_logo = p_logo.add_run()
    r_logo.add_picture(io.BytesIO(logo_bytes), width=Inches(4))
    p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER

p_title = doc.add_paragraph(proposal_title)
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
run_title = p_title.runs[0]
run_title.font.name = "DMSerif"
run_title.font.size = Pt(18)

doc.add_paragraph()

# tables
for hdr, rows, links, tbl_total in tables_info:
    n = len(hdr)
    desc_w = 0.45 * 17
    oth_w  = (17 - desc_w) / (n - 1)

    table = doc.add_table(rows=1, cols=n, style="Table Grid")
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, col in enumerate(table.columns):
        col.width = Inches(desc_w if i==1 else oth_w)

    # header
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(hdr):
        p = hdr_cells[i].paragraphs[0]
        run = p.add_run(col_name)
        run.font.name = "DMSerif"
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER

    # data rows
    for ridx, row in enumerate(rows):
        cells = table.add_row().cells
        for cidx, val in enumerate(row):
            p = cells[cidx].paragraphs[0]
            p.text = ""
            if cidx == 1 and links[ridx]:
                run = p.add_run()  # placeholder run
                run.add_hyperlink(links[ridx], str(val))
            else:
                run = p.add_run(str(val))
                run.font.name = "Barlow"
                run.font.size = Pt(9)

    # table total
    if tbl_total:
        label, amt = re.split(r'\$\s*', tbl_total, 1)
        amt = "$" + amt.strip()
        cells = table.add_row().cells
        for i, txt in enumerate([label] + [""]*(n-2) + [amt]):
            p = cells[i].paragraphs[0]
            run = p.add_run(txt)
            run.font.name = "DMSerif"
            run.font.size = Pt(10)
            run.bold = True
            if i == 0:
                p.alignment = WD_TABLE_ALIGNMENT.LEFT
            elif i == n-1:
                p.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else:
                p.alignment = WD_TABLE_ALIGNMENT.CENTER

    doc.add_paragraph()

# grand total
if grand_total:
    p_gt = doc.add_paragraph(f"Grand Total {grand_total}")
    p_gt.alignment = WD_TABLE_ALIGNMENT.RIGHT
    run_gt = p_gt.runs[0]
    run_gt.font.name = "DMSerif"
    run_gt.font.size = Pt(10)
    run_gt.bold = True

# ─── Download ────────────────────────────────────────────────────────────────
buf = io.BytesIO()
doc.save(buf)
buf.seek(0)
st.download_button(
    "📥 Download deliverable DOCX",
    data=buf,
    file_name="proposal_deliverable.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True,
)
