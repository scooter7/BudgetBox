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

# Logo URL
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download both the full PDF and "
    "a cleaned, horizontally-formatted deliverable in landscape PDF."
)

# — Upload PDF -------------------------------------------------------------
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()

pdf_bytes = uploaded.read()

# — Extract title + all tables ------------------------------------------------
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    # Document title = first line of page 1
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n")[0].strip()

    # Gather every table on every page
    raw_tables = []
    for pg in pdf.pages:
        raw_tables.extend(pg.extract_tables() or [])

if not raw_tables:
    st.error("No tables found in the document.")
    st.stop()

# — Define the canonical columns we expect -----------------------------------
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
    # 1) Clean & normalize header row
    cleaned = []
    for cell in raw[0]:
        if isinstance(cell, str):
            h = cell.replace("\n", " ").strip()
            if h.lower().startswith("term"):
                h = "Term"
            cleaned.append(h)
        else:
            cleaned.append("")
    # 2) Which indices to keep?
    keep = [i for i,h in enumerate(cleaned) if h]
    headers = [cleaned[i] for i in keep]
    # 3) Grab data rows
    rows = []
    for row in raw[1:]:
        # guard in case row shorter/longer
        rows.append([row[i] if i < len(row) else None for i in keep])
    df = pd.DataFrame(rows, columns=headers)
    # 4) Reindex to ensure exactly expected_cols (missing → NaN)
    df = df.reindex(columns=expected_cols)
    return df

# Process & concat all tables
dfs = [process_table(rt) for rt in raw_tables if len(rt) > 1]
df = pd.concat(dfs, ignore_index=True)

# — Split Strategy vs. Description -------------------------------------------
parts = df["Description"].fillna("").str.split(r"\n", n=1, expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# Reorder for final deliverable
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# — Preview ------------------------------------------------------------------
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# — Build deliverable PDF ----------------------------------------------------
deliverable_buf = io.BytesIO()
doc = SimpleDocTemplate(
    deliverable_buf,
    pagesize=landscape(letter),
    leftMargin=36, rightMargin=36, topMargin=72, bottomMargin=36
)
styles = getSampleStyleSheet()
elements = []

# 1) Embed Carnegie logo
try:
    r = requests.get(LOGO_URL, timeout=5)
    r.raise_for_status()
    img_bytes = r.content
    elements.append(Image(io.BytesIO(img_bytes), width=120, height=40))
    elements.append(Spacer(1, 12))
except Exception as e:
    st.warning(f"Could not fetch logo: {e}")

# 2) Document title
elements.append(Paragraph(proposal_title, styles["Title"]))
elements.append(Spacer(1, 24))

# 3) Build wrapped table
wrapped = []
for row in [df.columns.tolist()] + df.values.tolist():
    wrapped.append([Paragraph(str(cell), styles["BodyText"]) for cell in row])

table = Table(wrapped, repeatRows=1)
table.setStyle(TableStyle([
    # light grey header + black text
    ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
    ("TEXTCOLOR",  (0,0), (-1,0), colors.black),
    ("ALIGN",      (0,0), (-1,-1), "CENTER"),
    ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
    ("FONTSIZE",   (0,0), (-1,0), 12),
    ("FONTSIZE",   (0,1), (-1,-1), 10),
    ("BOTTOMPADDING",(0,0),(-1,0), 8),
    ("LEFTPADDING",  (0,1),(-1,-1),4),
    ("RIGHTPADDING", (0,1),(-1,-1),4),
]))
elements.append(table)

doc.build(elements)
deliverable_buf.seek(0)

st.success("✔️ Transformation complete!")

# — Download buttons ---------------------------------------------------------
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
