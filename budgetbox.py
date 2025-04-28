# app.py

import os
import io
import streamlit as st
import pdfplumber
import pandas as pd

from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download both the full PDF and "
    "a cleaned, horizontally-formatted deliverable in landscape PDF."
)

uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload a PDF to begin.")
    st.stop()

# Read the raw file bytes once so we can re-use for download
file_bytes = uploaded.read()

# --- Open PDF for title & tables ---
with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
    # Extract the first line of text as the proposal title
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n")[0].strip()

    # Pull out all tables on pages 1–2
    raw_tables = []
    for pg in pdf.pages[:2]:
        raw_tables.extend(pg.extract_tables() or [])

if not raw_tables:
    st.error("No tables found on the first two pages.")
    st.stop()

raw = raw_tables[0]

# --- Define expected columns & clean header row ---
expected_cols = [
    "Description",
    "Term",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes",
]

cleaned_hdr = []
for cell in raw[0]:
    if isinstance(cell, str):
        h = cell.replace("\n", " ").strip()
        if h.lower().startswith("term"):
            h = "Term"
        cleaned_hdr.append(h)
    else:
        cleaned_hdr.append("")

# Keep only non-blank headers
keep_idx = [i for i, h in enumerate(cleaned_hdr) if h]
header_names = [cleaned_hdr[i] for i in keep_idx]

# --- Build DataFrame from kept columns ---
rows = [[r[i] for i in keep_idx] for r in raw[1:]]
df = pd.DataFrame(rows, columns=header_names)

# Subset to exactly expected_cols
df = df.loc[:, expected_cols].copy()

# Drop “Total” rows
df = df[~df["Description"].str.contains("Total", case=False, na=False)]

# Split Strategy vs. Description
parts = df["Description"].str.split(r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# Final columns order
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# Preview in Streamlit
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# --- Generate deliverable PDF ---
deliverable_buf = io.BytesIO()
doc = SimpleDocTemplate(
    deliverable_buf,
    pagesize=landscape(letter),
    leftMargin=36,
    rightMargin=36,
    topMargin=72,
    bottomMargin=36,
)
styles = getSampleStyleSheet()
elements = []

# Carnegie logo (if present)
logo_path = "carnegie_logo.png"
if os.path.exists(logo_path):
    logo = Image(logo_path, width=120, height=40)
    elements.append(logo)
    elements.append(Spacer(1, 12))
else:
    st.warning(f"Logo file not found: {logo_path}")

# Proposal title
elements.append(Paragraph(proposal_title, styles["Title"]))
elements.append(Spacer(1, 24))

# Build wrapped table
wrapped_data = []
for row in [df.columns.tolist()] + df.values.tolist():
    wrapped_row = []
    for cell in row:
        wrapped_row.append(Paragraph(str(cell), styles["BodyText"]))
    wrapped_data.append(wrapped_row)

table = Table(wrapped_data, repeatRows=1)
table.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003f5c")),
    ("TEXTCOLOR",  (0,0), (-1,0), colors.whitesmoke),
    ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
    ("FONTSIZE",   (0,0), (-1,0), 12),
    ("FONTSIZE",   (0,1), (-1,-1), 10),
    ("BOTTOMPADDING", (0,0), (-1,0), 8),
    ("LEFTPADDING",   (0,1), (-1,-1), 4),
    ("RIGHTPADDING",  (0,1), (-1,-1), 4),
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
        data=file_bytes,
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
