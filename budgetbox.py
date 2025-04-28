# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests

from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# Constants
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download both the full PDF and "
    "a cleaned, horizontally-formatted deliverable in landscape PDF."
)

# — Upload PDF ——————————————————————————————————————————————————————
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()

pdf_bytes = uploaded.read()

# — Extract title & tables ————————————————————————————————————————————
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    # Document title = first line of page 1
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n")[0].strip()

    # Gather tables from pages 1 and 2
    raw_tables = []
    for pg in pdf.pages[:2]:
        raw_tables.extend(pg.extract_tables() or [])

if not raw_tables:
    st.error("No tables found on the first two pages.")
    st.stop()

raw = raw_tables[0]

# — Define & clean expected columns ——————————————————————————————————————
expected_cols = [
    "Description",
    "Term",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes",
]

# Normalize the header row
cleaned_hdr = []
for cell in raw[0]:
    if isinstance(cell, str):
        h = cell.replace("\n", " ").strip()
        if h.lower().startswith("term"):
            h = "Term"
        cleaned_hdr.append(h)
    else:
        cleaned_hdr.append("")

# Keep only the non-empty headers
keep_idx = [i for i, h in enumerate(cleaned_hdr) if h]
header_names = [cleaned_hdr[i] for i in keep_idx]

# Build DataFrame from those columns
rows = [[r[i] for i in keep_idx] for r in raw[1:]]
df = pd.DataFrame(rows, columns=header_names)

# Subset to exactly your expected columns
df = df.loc[:, expected_cols].copy()

# — Split Strategy vs. Description ————————————————————————————————————
parts = df["Description"].str.split(r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# Reorder into final layout
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# — Preview in Streamlit —————————————————————————————————————————————
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# — Build the deliverable PDF ———————————————————————————————————————
deliverable_buf = io.BytesIO()
doc = SimpleDocTemplate(
    deliverable_buf,
    pagesize=landscape(letter),
    leftMargin=36, rightMargin=36, topMargin=72, bottomMargin=36
)
styles = getSampleStyleSheet()
elements = []

# 1) Fetch & embed Carnegie logo
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    logo_img = resp.content
    elements.append(Image(io.BytesIO(logo_img), width=120, height=40))
    elements.append(Spacer(1, 12))
except Exception as e:
    st.warning(f"Could not fetch logo: {e}")

# 2) Add document title below the logo
elements.append(Paragraph(proposal_title, styles["Title"]))
elements.append(Spacer(1, 24))

# 3) Wrap every cell in a Paragraph so long text wraps
wrapped = []
for row in [df.columns.tolist()] + df.values.tolist():
    wrapped.append([Paragraph(str(cell), styles["BodyText"]) for cell in row])

table = Table(wrapped, repeatRows=1)
table.setStyle(TableStyle([
    # Very light grey header background with black text
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
    ("TEXTCOLOR",  (0,0), (-1,0), colors.black),
    ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
    ("FONTSIZE",   (0,0), (-1,0), 12),
    ("FONTSIZE",   (0,1), (-1,-1), 10),
    ("BOTTOMPADDING",(0,0), (-1,0), 8),
    ("LEFTPADDING",(0,1), (-1,-1), 4),
    ("RIGHTPADDING",(0,1), (-1,-1), 4),
]))
elements.append(table)

doc.build(elements)
deliverable_buf.seek(0)

st.success("✔️ Transformation complete!")

# — Download buttons —————————————————————————————————————————————————
c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "📥 Download full original PDF",
        data=pdf_bytes,
        file_name=uploaded.name,
        mime="application/pdf",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "📥 Download deliverable PDF (landscape)",
        data=deliverable_buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
