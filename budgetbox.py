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
    hdr = []
    for cell in raw[0]:
        if isinstance(cell, str):
            h = cell.replace("\n", " ").strip()
            if h.lower().startswith("term"):
                h = "Term"
            hdr.append(h)
        else:
            hdr.append("")
    keep = [i for i,h in enumerate(hdr) if h]
    headers = [hdr[i] for i in keep]
    rows = []
    for r in raw[1:]:
        rows.append([r[i] if i < len(r) else "" for i in keep])
    return pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols).fillna("")

# Concatenate all tables
dfs = [process_table(t) for t in raw_tables if len(t) > 1]
df = pd.concat(dfs, ignore_index=True)

# — Split Strategy vs. Description without pandas.str.split error —
def split_desc(text):
    if not isinstance(text, str):
        return "", ""
    parts = text.split("\n", 1)
    first = parts[0].strip()
    second = parts[1].strip() if len(parts) > 1 else ""
    return first, second

df[["Strategy","Description"]] = df["Description"].apply(
    lambda t: pd.Series(split_desc(t))
)
final_cols = ["Strategy","Description"] + expected_cols[1:]
df = df[final_cols]

# Preview
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# — Build HTML for PDF with explicit col widths and logo/title —
# Fetch logo as base64
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    logo_b64 = base64.b64encode(resp.content).decode()
    logo_img = (
        f'<img src="data:image/png;base64,{logo_b64}" '
        f'style="display:block;margin:0 auto 12px;width:120px;" />'
    )
except:
    logo_img = ""

# Generate table HTML and inject <colgroup>
table_html = df.to_html(index=False, border=0)
colgroup = """
<colgroup>
  <col style="width:15%;">
  <col style="width:35%;">
  <col style="width:8%;">
  <col style="width:8%;">
  <col style="width:8%;">
  <col style="width:8%;">
  <col style="width:8%;">
</colgroup>
"""
# insert colgroup immediately after opening <table ...>
table_html = table_html.replace('>', f'>\n{colgroup}', 1)

# Compose full HTML
html = f"""
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <style>
    @page {{ size: A4 landscape; margin: 1in; }}
    body {{ font-family: sans-serif; margin:0; padding:0; }}
    h1 {{ text-align: center; margin-bottom: 24px; }}
    table {{ border-collapse: collapse; table-layout: fixed; width:100%; }}
    th, td {{
      border: 1px solid #ccc;
      padding: 4px;
      word-wrap: break-word;
    }}
    th {{
      background-color: #F2F2F2;
      color: #000;
      font-weight: bold;
    }}
    td {{ font-size: 10pt; }}
  </style>
</head>
<body>
  {logo_img}
  <h1>{proposal_title}</h1>
  {table_html}
</body>
</html>
"""

# Render to PDF
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
