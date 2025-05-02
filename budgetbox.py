import streamlit as st
import pdfplumber
import fitz
import io
import requests
import re
from weasyprint import HTML
import pypandoc

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

def extract_tables_and_links(pdf_bytes):
    # Capture link annotations with PyMuPDF
    doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_annotations = []
    for page in doc_fitz:
        annots = []
        for a in page.annots() or []:
            if a.type[0] == 1 and a.uri:
                annots.append((a.rect, a.uri))
        page_annotations.append(annots)

    # Open with pdfplumber to extract text and tables
    tables_info = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # extract text lines for title + totals
        page_texts = [p.extract_text() or "" for p in pdf.pages]
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

        # parse each page
        for pi, page in enumerate(pdf.pages):
            annots = page_annotations[pi]
            for tbl in page.find_tables():
                data = tbl.extract()
                if len(data) < 2:
                    continue
                hdr = data[0]
                # find Description column index
                desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
                if desc_i is None:
                    continue

                # map link rect → row index
                x0, y0, x1, y1 = tbl.bbox
                band_h = (y1 - y0) / len(data)
                row_links = {}
                for rect, uri in annots:
                    midy = (rect.y0 + rect.y1) / 2
                    if y0 <= midy <= y1:
                        ridx = int((midy - y0) // band_h)
                        if 1 <= ridx < len(data):
                            row_links[ridx-1] = uri

                # assemble rows
                def split_cell_text(raw: str):
                    lines = [l.strip() for l in raw.splitlines() if l.strip()]
                    return (lines[0], " ".join(lines[1:])) if lines else ("", "")

                new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(hdr) if i!=desc_i and h]
                rows, links = [], []
                for ridx, row in enumerate(data[1:], start=1):
                    if all(not str(c).strip() for c in row if c):
                        continue
                    first = next((str(c).strip() for c in row if c), "")
                    if first.lower() == "total":
                        continue
                    strat, desc = split_cell_text(str(row[desc_i] or ""))
                    rest = [row[i] for i,h in enumerate(hdr) if i!=desc_i and h]
                    rows.append([strat, desc] + rest)
                    links.append(row_links.get(ridx-1))

                tbl_total = find_total(pi)
                tables_info.append({
                    "header": new_hdr,
                    "rows": rows,
                    "links": links,
                    "total": tbl_total
                })

        # find grand total
        grand_total = None
        for txt in reversed(page_texts):
            m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', txt, re.I|re.S)
            if m:
                grand_total = m.group(1)
                break

    return proposal_title, tables_info, grand_total

def build_html(proposal_title, tables_info, grand_total):
    # inline CSS with Google Fonts and table styling
    css = f"""
    @import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=Barlow&display=swap');
    body {{
      font-family: 'Barlow', sans-serif;
      margin: 1in;
    }}
    h1 {{
      font-family: 'DM Serif Display', serif;
      font-size: 24pt;
      text-align: center;
      margin-bottom: 0.5em;
    }}
    .logo {{
      display: block;
      margin: 0 auto 1em;
      max-width: 360px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 1.5em;
      font-size: 9pt;
    }}
    th, td {{
      border: 1px solid #888;
      padding: 4px;
      vertical-align: top;
    }}
    th {{
      background: #F2F2F2;
      font-family: 'DM Serif Display', serif;
      font-size: 10pt;
      text-align: center;
    }}
    td.strategy {{ width: 20%; }}
    td.description {{ width: 45%; }}
    td.total-row td {{ font-family: 'DM Serif Display', serif; font-size: 10pt; font-weight: bold; }}
    """

    html = [
        "<!DOCTYPE html>",
        "<html><head><meta charset='utf-8'>",
        f"<style>{css}</style>",
        "</head><body>",
        f"<img class='logo' src='{LOGO_URL}'/>",
        f"<h1>{proposal_title}</h1>"
    ]

    # build tables
    for tbl in tables_info:
        hdr = tbl["header"]
        rows = tbl["rows"]
        links = tbl["links"]
        total = tbl["total"]

        html.append("<table>")
        # header
        html.append("<tr>" + "".join(f"<th>{h}</th>" for h in hdr) + "</tr>")
        # data rows
        for i, row in enumerate(rows):
            html.append("<tr>")
            for j, cell in enumerate(row):
                cls = ""
                if hdr[j].lower() == "strategy":
                    cls = "strategy"
                elif hdr[j].lower() == "description":
                    cls = "description"
                if j == 1 and links[i]:
                    html.append(f"<td class='{cls}'><a href='{links[i]}'>{cell}</a></td>")
                else:
                    html.append(f"<td class='{cls}'>{cell}</td>")
            html.append("</tr>")
        # table subtotal
        if total:
            parts = re.split(r'\$\s*', total, 1)
            label = parts[0]
            amount = f"${parts[1].strip()}"
            html.append("<tr class='total-row'>")
            html.append(f"<td colspan='{len(hdr)-1}' style='text-align:left'>{label}</td>")
            html.append(f"<td style='text-align:right'>{amount}</td>")
            html.append("</tr>")

        html.append("</table>")

    # grand total
    if grand_total:
        html.append("<h2 style='text-align:right'>Grand Total " + grand_total + "</h2>")

    html.append("</body></html>")
    return "\n".join(html)

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if uploaded:
    pdf_bytes = uploaded.read()
    title, tables, grand_total = extract_tables_and_links(pdf_bytes)
    html_content = build_html(title, tables, grand_total)

    # Preview HTML if desired
    # st.write(html_content, unsafe_allow_html=True)

    # Convert to PDF
    pdf_output = HTML(string=html_content).write_pdf()

    # Convert to DOCX
    docx_bytes = pypandoc.convert_text(
        html_content,
        to="docx",
        format="html",
        extra_args=[
            "--reference-doc=template.docx"  # optional: a custom .docx template
        ]
    )
    docx_buf = io.BytesIO(docx_bytes)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "📥 Download deliverable PDF",
            data=pdf_output,
            file_name="proposal_deliverable.pdf",
            mime="application/pdf",
            use_container_width=True
        )
    with c2:
        st.download_button(
            "📥 Download deliverable DOCX",
            data=docx_buf,
            file_name="proposal_deliverable.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
