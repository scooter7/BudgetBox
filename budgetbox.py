# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests

from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    Image,
    PageBreak,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER

# URL for the Carnegie logo
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download both the full PDF and "
    "a cleaned, horizontally-formatted deliverable in landscape PDF."
)

# — Upload PDF —
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

# — Expected columns —
expected_cols = [
    "Description",
    "Term",
    "Start Date",
    "End Date",
    "Monthly Amount",
    "Item Total",
    "Notes",
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
    keep = [i for i, h in enumerate(hdr) if h]
    headers = [hdr[i] for i in keep]
    rows = []
    for row in raw[1:]:
        rows.append([row[i] if i < len(row) else None for i in keep])
    df = pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols)
    return df

# Concatenate all tables
dfs = [process_table(t) for t in raw_tables if len(t) > 1]
df = pd.concat(dfs, ignore_index=True)

# — Split Strategy vs. Description —
parts = df["Description"].fillna("").str.split(pat=r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# Preview
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# — Build deliverable PDF —
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(letter),
    leftMargin=36, rightMargin=36,
    topMargin=72, bottomMargin=36,
)
styles = getSampleStyleSheet()
title_style = styles["Title"]
title_style.alignment = TA_CENTER
body_style = styles["BodyText"]

elements = []

# Embed Carnegie logo
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    elements.append(Image(io.BytesIO(resp.content), width=120, height=40))
    elements.append(Spacer(1, 12))
except Exception as e:
    st.warning(f"Could not fetch Carnegie logo: {e}")

# Centered document title
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Wrap rows for the table
wrapped = []
for row in [df.columns.tolist()] + df.values.tolist():
    wrapped.append([Paragraph(str(cell), body_style) for cell in row])

# Break into chunks to fit the page (header + 15 data rows)
header = wrapped[0]
data_rows = wrapped[1:]
chunk_size = 15

for i in range(0, len(data_rows), chunk_size):
    chunk = [header] + data_rows[i : i + chunk_size]
    table = Table(chunk, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#F2F2F2")),
        ("TEXTCOLOR",   (0,0), (-1,0), colors.black),
        ("ALIGN",       (0,0), (-1,-1), "CENTER"),
        ("GRID",        (0,0), (-1,-1), 0.5, colors.grey),
        ("FONTSIZE",    (0,0), (-1,0), 12),
        ("FONTSIZE",    (0,1), (-1,-1), 10),
        ("BOTTOMPADDING",(0,0), (-1,0), 8),
        ("LEFTPADDING", (0,1), (-1,-1), 4),
        ("RIGHTPADDING",(0,1), (-1,-1), 4),
    ]))
    elements.append(table)
    if i + chunk_size < len(data_rows):
        elements.append(PageBreak())

doc.build(elements)
buf.seek(0)

st.success("✔️ Transformation complete!")

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
        data=buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
