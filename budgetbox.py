# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests
import base64
from weasyprint import HTML, CSS

# Constants
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download both the full PDF and "
    "a cleaned, horizontally-formatted deliverable in landscape PDF."
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
    # Build rows
    rows = []
    for r in raw[1:]:
        rows.append([r[i] if i < len(r) else "" for i in keep])
    # Create and reindex
    return pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols).fillna("")

# Concatenate all tables
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

# — Build deliverable PDF via WeasyPrint —
# 1) Fetch logo and convert to base64
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    logo_b64 = base64.b64encode(resp.content).decode()
except Exception:
    logo_b64 = ""

# 2) Compose HTML
html = f"""
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <style>
    @page {{ size: landscape; margin: 36pt; }}
    body {{ font-family: sans-serif; margin: 0; padding: 0; }}
    .logo {{ display: block; margin: 0 auto 12pt; width: 120px; }}
    h1 {{ text-align: center; margin-bottom: 24pt; }}
    table {{ width: 100%; border-collapse: collapse; }}
    th, td {{ border: 1px solid #ccc; padding: 4pt; word-wrap: break-word; }}
    th {{ background-color: #F2F2F2; color: #000; }}
    td {{ font-size: 10pt; }}
  </style>
</head>
<body>
  {f'<img class="logo" src="data:image/png;base64,{logo_b64}"/>' if logo_b64 else ""}
  <h1>{proposal_title}</h1>
  {df.to_html(index=False, border=0)}
</body>
</html>
"""

# 3) Render PDF
pdf_out = HTML(string=html).write_pdf(stylesheets=[CSS(string='@page { size: landscape; }')])

st.success("✔️ Transformation complete!")

# — Download buttons —
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "📥 Download original PDF",
        data=pdf_bytes,
        file_name=uploaded.name,
        mime="application/pdf",
        use_container_width=True,
    )
with col2:
    st.download_button(
        "📥 Download deliverable PDF (landscape)",
        data=pdf_out,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
