# budgetbox.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests

from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch

# Define 11x17 (Tabloid) manually
tabloid = (11 * inch, 17 * inch)

from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas

# Logo URL
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and download a cleaned, horizontally-formatted deliverable in 11x17 landscape PDF."
)

# Upload PDF
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract title & tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n", 1)[0].strip()
    raw_tables = []
    for page in pdf.pages:
        raw_tables.extend(page.extract_tables() or [])

if not raw_tables:
    st.error("No tables found in the document.")
    st.stop()

# Expected columns
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
    return pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols)

dfs = [process_table(t) for t in raw_tables if len(t) > 1]
df = pd.concat(dfs, ignore_index=True)

# Split Strategy vs. Description
parts = df["Description"].fillna("").str.split(pat=r"\n", n=1, expand=True)
df["Strategy"] = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")

# Final columns
final_cols = ["Strategy", "Description"] + expected_cols[1:]
df = df[final_cols]

# Calculate totals
total_rows = len(df)
try:
    total_dollars = pd.to_numeric(df["Item Total"].str.replace(r"[^0-9.\-]", "", regex=True), errors='coerce').sum()
except Exception:
    total_dollars = 0.0

# Format dollar value
total_dollars_formatted = "${:,.2f}".format(total_dollars)

# Preview
st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# Build deliverable PDF
buf = io.BytesIO()

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        page_count = len(self.pages)
        for page in self.pages:
            self.__dict__.update(page)
            self.draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 8)
        self.drawRightString(
            1600, 20, f"Page {self._pageNumber} of {page_count}"
        )

doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(tabloid),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

styles = getSampleStyleSheet()
title_style = styles["Title"]
title_style.alignment = TA_CENTER

small_center_style = ParagraphStyle(
    'SmallCenter',
    parent=styles['BodyText'],
    fontSize=10,
    alignment=TA_CENTER,
    spaceAfter=6,
)

body_style = ParagraphStyle(
    'BodySmall',
    parent=styles['BodyText'],
    fontSize=9,
    leading=11,
)

elements = []

# Carnegie logo
try:
    resp = requests.get(LOGO_URL, timeout=5)
    resp.raise_for_status()
    elements.append(Image(io.BytesIO(resp.content), width=150, height=50))
    elements.append(Spacer(1, 12))
except Exception as e:
    st.warning(f"Could not fetch Carnegie logo: {e}")

# Centered proposal title
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 12))

# Total rows and dollars
summary_text = f"Total Rows: {total_rows} | Total Dollar Value: {total_dollars_formatted}"
elements.append(Paragraph(summary_text, small_center_style))
elements.append(Spacer(1, 24))

# Limit very large cells
MAX_CELL_LENGTH = 400

# Wrap table cells
wrapped = []
for row in [df.columns.tolist()] + df.fillna("").values.tolist():
    wrapped_row = []
    for cell in row:
        cell_text = str(cell).replace('\n', '<br/>')
        if len(cell_text) > MAX_CELL_LENGTH:
            cell_text = cell_text[:MAX_CELL_LENGTH] + "..."
        para = Paragraph(cell_text, body_style)
        wrapped_row.append(para)
    wrapped.append(wrapped_row)

# LongTable setup
table = LongTable(wrapped, repeatRows=1, splitByRow=True)
table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
    ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("BOTTOMPADDING",(0, 0), (-1, 0), 8),
    ("TOPPADDING",  (0, 0), (-1, 0), 8),
    ("BOTTOMPADDING",(0, 1), (-1, -1), 6),
    ("TOPPADDING",  (0, 1), (-1, -1), 6),
    ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ("RIGHTPADDING",(0, 0), (-1, -1), 4),
    ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
]))
elements.append(table)

doc.build(elements, canvasmaker=NumberedCanvas)
buf.seek(0)

st.success("✔️ Transformation complete!")

# Only the transformed deliverable button
st.download_button(
    "📥 Download deliverable PDF (11x17 landscape)",
    data=buf,
    file_name="proposal_deliverable_11x17.pdf",
    mime="application/pdf",
    use_container_width=True,
)
