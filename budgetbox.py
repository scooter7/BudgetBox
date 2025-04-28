# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests
import base64
from xhtml2pdf import pisa

# URL for the Carnegie logo
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download a cleaned, "
    "horizontally-formatted deliverable in landscape PDF."
)

# — Upload source PDF —
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# — Extract title & all tables —
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n", 1)[0].strip()
    raw_tables = []
    for page in pdf.pages:
        raw_tables.extend(page.extract_tables() or [])

if not raw_tables:
    st.error("No tables found in the document.")
    st.stop()

# — Define expected columns —
expected_cols = [
    "Description", "Term", "Start Date",
    "End Date", "Monthly Amount", "Item Total", "Notes"
]

def process_table(raw):
    # Normalize header row
    hdr = []
    for cell in raw[0]:
        if isinstance(cell, str):
            h = cell.replace("\n", " ").strip()
            if h.lower().startswith("term"):
                h = "Term"
            hdr.append(h)
        else:
            hdr.append("")
    # Keep only non-empty headers
    keep = [i for i, h in enumerate(hdr) if h]
    headers = [hdr[i] for i in keep]
    rows = []
    for r in raw[1:]:
        rows.append([r[i] if i < len(r) else "" for i in keep])
    # Build DF and reindex to expected cols
    return pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols).fillna("")

# Process & concat all tables
dfs = [process_table(t) for t in raw_tables if len(t) > 1]
df = pd.concat(dfs, ignore_index=True)

# — Split Strategy vs. Description —
parts = df["Description"].str.split(r"\n", 1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# Preview
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# — Build HTML for PDF with fixed column widths —
# fetch logo & embed as base64
try:
    r = requests.get(LOGO_URL, timeout=5)
    r.raise_for_status()
    logo_b64 = base64.b64encode(r.content).decode()
    logo_img = (
        f'<img src="data:image/png;base64,{logo_b64}" '
        f'style="display:block;margin:0 auto 12px;width:120px;" />'
    )
except:
    logo_img = ""

html = f"""
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <style>
    @page {{ size: A4 landscape; margin: 1in; }}
    body {{ font-family: sans-serif; margin:0; padding:0; }}
    h1 {{ text-align: center; margin-bottom: 24px; }}
    table {{ width: 100%; border-collapse: collapse; table-layout: fixed; }}
    th, td {{
      border: 1px solid #ccc;
      padding: 4px;
      overflow-wrap: break-word;
    }}
    th {{ background-color: #F2F2F2; color: #000; }}

    /* Explicit column widths */
    th:nth-child(1), td:nth-child(1) {{ width: 15%; }}  /* Strategy */
    th:nth-child(2), td:nth-child(2) {{ width: 35%; }}  /* Description */
    th:nth-child(3), td:nth-child(3) {{ width: 8%; }}   /* Term */
    th:nth-child(4), td:nth-child(4) {{ width: 8%; }}   /* Start Date */
    th:nth-child(5), td:nth-child(5) {{ width: 8%; }}   /* End Date */
    th:nth-child(6), td:nth-child(6) {{ width: 8%; }}   /* Monthly Amount */
    th:nth-child(7), td:nth-child(7) {{ width: 8%; }}   /* Item Total */
    th:nth-child(8), td:nth-child(8) {{ width: 10%; }}  /* Notes */

    td {{ font-size: 10pt; }}
  </style>
</head>
<body>
  {logo_img}
  <h1>{proposal_title}</h1>
  {df.to_html(index=False, border=0)}
</body>
</html>
"""

# Render HTML → PDF
pdf_buffer = io.BytesIO()
pisa.CreatePDF(io.StringIO(html), dest=pdf_buffer)
pdf_data = pdf_buffer.getvalue()

st.success("✔️ Transformation complete!")

# — Download transformed PDF only —
st.download_button(
    "📥 Download deliverable PDF (landscape)",
    data=pdf_data,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
