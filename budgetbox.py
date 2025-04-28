# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download a cleaned, "
    "horizontally-formatted deliverable in landscape PDF."
)

# --- File uploader ---
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload a PDF to begin.")
    st.stop()

# --- Extract tables from pages 1–2 ---
with pdfplumber.open(uploaded) as pdf:
    raw_tables = []
    for page in pdf.pages[:2]:
        raw_tables.extend(page.extract_tables() or [])

if not raw_tables:
    st.error("No tables found on the first two pages.")
    st.stop()

raw = raw_tables[0]

# --- Define expected columns ---
expected_cols = [
    "Description",
    "Term",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes",
]

# --- Clean & normalize header row ---
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
rows = []
for r in raw[1:]:
    rows.append([r[i] for i in keep_idx])
df = pd.DataFrame(rows, columns=header_names)

# --- Subset to exactly the expected cols (avoids missing-column errors) ---
df = df.loc[:, expected_cols].copy()

# --- Drop “Total” rows ---
df = df[~df["Description"].str.contains("Total", case=False, na=False)]

# --- Split Strategy vs. Description ---
parts = df["Description"].str.split(r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# --- Reorder columns for final deliverable ---
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# --- Preview in Streamlit ---
st.subheader("Transformed Data")
st.dataframe(df, use_container_width=True)

# --- Generate landscape-oriented PDF deliverable ---
buffer = io.BytesIO()
doc = SimpleDocTemplate(
    buffer,
    pagesize=landscape(letter),
    rightMargin=20,
    leftMargin=20,
    topMargin=20,
    bottomMargin=20,
)
styles = getSampleStyleSheet()
elements = [
    Paragraph("Proposal Deliverable", styles["Title"]),
    Spacer(1, 12)
]

# Build table for ReportLab
table_data = [df.columns.tolist()] + df.values.tolist()
tbl = Table(table_data, repeatRows=1)
tbl.setStyle(TableStyle([
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003f5c")),
    ("TEXTCOLOR",  (0,0), (-1,0), colors.whitesmoke),
    ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
    ("FONTSIZE",   (0,0), (-1,0), 12),
    ("FONTSIZE",   (0,1), (-1,-1), 10),
    ("BOTTOMPADDING", (0,0), (-1,0), 8),
]))
elements.append(tbl)

doc.build(elements)
buffer.seek(0)

# --- Download button ---
st.success("✔️ Ready to download")
st.download_button(
    "📥 Download deliverable PDF (landscape)",
    data=buffer,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
