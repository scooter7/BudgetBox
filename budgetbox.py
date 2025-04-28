# budgetbox.py

import io
import re
import streamlit as st
import pdfplumber
import pandas as pd
import requests

from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas

# Define 11x17 manually
tabloid = (11 * inch, 17 * inch)

# Logo URL
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

# Streamlit setup
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

# Initialize pdfplumber
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    # Extract title
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n", 1)[0].strip()

    # Extract all page texts
    page_texts = [page.extract_text() for page in pdf.pages]
    
    # Extract all tables
    all_raw_tables = []
    for page in pdf.pages:
        tables = page.extract_tables()
        if tables:
            all_raw_tables.extend(tables)

if not all_raw_tables:
    st.error("No tables found in the document.")
    st.stop()

# Setup PDF buffer
buf = io.BytesIO()

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
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
        self.drawRightString(1600, 20, f"Page {self._pageNumber} of {page_count}")

doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(tabloid),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

styles = getSampleStyleSheet()
title_style = styles["Title"]
title_style.alignment = TA_CENTER

header_style = ParagraphStyle(
    'Header',
    parent=styles['Heading2'],
    alignment=TA_CENTER,
    spaceAfter=12,
)

body_style = ParagraphStyle(
    'BodySmall',
    parent=styles['BodyText'],
    fontSize=9,
    leading=11,
)

bold_center_style = ParagraphStyle(
    'BoldCenter',
    parent=styles['BodyText'],
    fontSize=10,
    alignment=TA_CENTER,
    spaceAfter=12,
    spaceBefore=12,
    fontName="Helvetica-Bold"
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

# Proposal title
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Pre-scan all Total lines and index them by page
page_totals = {}

for idx, page_text in enumerate(page_texts):
    if page_text:
        lines = page_text.splitlines()
        totals = []
        for line in lines:
            if re.search(r'\btotal\b', line, re.IGNORECASE) and re.search(r'\$[0-9,]+\.\d{2}', line):
                totals.append(line.strip())
        page_totals[idx] = totals

# Helper to pop one total line per page
def pop_next_total_for_page(page_idx):
    if page_idx in page_totals and page_totals[page_idx]:
        return page_totals[page_idx].pop(0)
    return None

# Process each table
MAX_CELL_LENGTH = 400
current_page_idx = 0
tables_per_page_estimate = 2  # Assumption: about 2 tables per page, adjust if needed

for table_idx, raw_table in enumerate(all_raw_tables):
    if len(raw_table) < 2:
        continue

    wrapped = []
    for row in raw_table:
        wrapped_row = []
        for cell in row:
            cell_text = str(cell).replace('\n', '<br/>') if cell else ''
            if len(cell_text) > MAX_CELL_LENGTH:
                cell_text = cell_text[:MAX_CELL_LENGTH] + "..."
            para = Paragraph(cell_text, body_style)
            wrapped_row.append(para)
        wrapped.append(wrapped_row)

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
    elements.append(Spacer(1, 12))

    # Find matching Total under this table
    total_text = pop_next_total_for_page(current_page_idx)
    if total_text:
        elements.append(Paragraph(total_text, bold_center_style))
        elements.append(Spacer(1, 24))

    if (table_idx + 1) % tables_per_page_estimate == 0:
        current_page_idx += 1  # Move to next page after estimated number of tables

# Find Grand Total separately
grand_total_amount = None
for page_text in page_texts[::-1]:  # search last pages first
    if page_text:
        match = re.search(r'Grand Total.*?\$[0-9,]+\.\d{2}', page_text, re.IGNORECASE | re.DOTALL)
        if match:
            grand_total_amount = re.search(r'\$[0-9,]+\.\d{2}', match.group(0)).group(0)
            break

# Insert Grand Total cleanly
if grand_total_amount:
    elements.append(PageBreak())
    elements.append(Spacer(1, 36))
    elements.append(Paragraph("Grand Total", header_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Total {grand_total_amount}", bold_center_style))

doc.build(elements, canvasmaker=NumberedCanvas)
buf.seek(0)

st.success("✔️ Transformation complete!")

# Only deliverable download
st.download_button(
    "📥 Download deliverable PDF (11x17 landscape)",
    data=buf,
    file_name="proposal_deliverable_11x17.pdf",
    mime="application/pdf",
    use_container_width=True,
)
