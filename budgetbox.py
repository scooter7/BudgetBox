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

# --- File upload ---
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()

pdf_bytes = uploaded.read()

# --- Extract title + first two pages’ tables ---
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    # Extract first line of page-1 text for the proposal title
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n")[0].strip()

    # Gather all tables on pages 1 and 2
    raw_tables = []
    for pg in pdf.pages[:2]:
        raw_tables.extend(pg.extract_tables() or [])

if not raw_tables:
    st.error("No tables found on the first two pages.")
    st.stop()

raw = raw_tables[0]

# --- Define & clean expected columns ---
expected_cols = [
    "Description",
    "Term",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes",
]

# Normalize PDF header row (merge lines, strip, “Term (Months)”→“Term”)
cleaned_hdr = []
for cell in raw[0]:
    if isinstance(cell, str):
        h = cell.replace("\n", " ").strip()
        if h.lower().startswith("term"):
            h = "Term"
        cleaned_hdr.append(h)
    else:
        cleaned_hdr.append("")

# Keep only non-empty header columns
keep_idx = [i for i, h in enumerate(cleaned_hdr) if h]
header_names = [cleaned_hdr[i] for i in keep_idx]

# Build DataFrame from those columns
rows = [[r[i] for i in keep_idx] for r in raw[1:]]
df = pd.DataFrame(rows, columns=header_names).loc[:, expected_cols].copy()

# Drop any “Total” rows
df = df[~df["Description"].str.contains("Total", case=False, na=False)]

# Split Strategy vs Description
parts = df["Description"].str.split(r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# Reorder to final layout
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# --- Preview in Streamlit ---
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# --- Build deliverable PDF ---
deliverable_buf = io.BytesIO()
doc = SimpleDocTemplate(
    deliverable_buf,
    pagesize=landscape(letter),
    leftMargin=36, rightMargin=36, topMargin=72, bottomMargin=36
)
styles = getSampleStyleSheet()
elements = []

# Fetch and embed the Carnegie logo
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    logo_img = resp.content
    elements.append(Image(io.BytesIO(logo_img), width=120, height=40))
    elements.append(Spacer(1, 12))
except Exception as e:
    st.warning(f"Could not fetch logo from URL: {e}")

# Add proposal title
elements.append(Paragraph(proposal_title, styles["Title"]))
elements.append(Spacer(1, 24))

# Prepare wrapped table cells so long text wraps
wrapped = []
for row in [df.columns.tolist()] + df.values.tolist():
    wrapped.append([Paragraph(str(cell), styles["BodyText"]) for cell in row])

table = Table(wrapped, repeatRows=1)
table.setStyle(TableStyle([
    ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#003f5c")),
    ("TEXTCOLOR",(0,0),(-1,0),colors.whitesmoke),
    ("ALIGN",(0,0),(-1,-1),"CENTER"),
    ("GRID",(0,0),(-1,-1),0.5,colors.grey),
    ("FONTSIZE",(0,0),(-1,0),12),
    ("FONTSIZE",(0,1),(-1,-1),10),
    ("BOTTOMPADDING",(0,0),(-1,0),8),
    ("LEFTPADDING",(0,1),(-1,-1),4),
    ("RIGHTPADDING",(0,1),(-1,-1),4),
]))
elements.append(table)

doc.build(elements)
deliverable_buf.seek(0)

st.success("✔️ Transformation complete!")

# --- Download buttons ---
col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "📥 Download full original PDF",
        data=pdf_bytes,
        file_name=uploaded.name,
        mime="application/pdf",
        use_container_width=True,
    )
with col2:
    st.download_button(
        "📥 Download deliverable PDF (landscape)",
        data=deliverable_buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
