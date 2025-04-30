import io
import os
import re
import streamlit as st
import pdfplumber
import requests

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.pdfgen import canvas

# 11x17 tabloid landscape
tabloid = (11 * inch, 17 * inch)

# Load custom fonts
FONT_DIR = "fonts"
pdfmetrics.registerFont(TTFont("DMSerif", os.path.join(FONT_DIR, "DMSerifDisplay-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Barlow", os.path.join(FONT_DIR, "Barlow-Regular.ttf")))

# Logo URL
LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

# Streamlit UI
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

# Extract text and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = (pdf.pages[0].extract_text() or "").split("\n", 1)[0].strip()
    page_texts = [page.extract_text() for page in pdf.pages]
    all_tables = []
    for i, page in enumerate(pdf.pages):
        tables = page.extract_tables()
        for t in tables:
            all_tables.append((i, t))

# Setup PDF output
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
        for p in self.pages:
            self.__dict__.update(p)
            self.draw_page_number(page_count)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, total):
        self.setFont("Helvetica", 8)
        self.drawRightString(1600, 20, f"Page {self._pageNumber} of {total}")

doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(tabloid),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

# Styles
title_style = ParagraphStyle(
    "Title",
    fontName="DMSerif",
    fontSize=18,
    alignment=TA_CENTER,
    spaceAfter=6,
)

header_style = ParagraphStyle(
    "Header",
    fontName="DMSerif",
    fontSize=11,
    alignment=TA_CENTER
)

body_style = ParagraphStyle(
    "Body",
    fontName="Barlow",
    fontSize=9,
    leading=11,
)

bold_center_style = ParagraphStyle(
    "BoldCenter",
    fontName="Barlow",
    fontSize=10,
    alignment=TA_CENTER,
    spaceAfter=12,
    spaceBefore=12
)

elements = []

# Logo and title
try:
    img_resp = requests.get(LOGO_URL, timeout=5)
    img_resp.raise_for_status()
    elements.append(Image(io.BytesIO(img_resp.content), width=150, height=50))
    elements.append(Spacer(1, 12))
except:
    st.warning("Logo couldn't be loaded.")

elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Extract all Total lines per page and track usage
page_totals = {}
used_lines = set()

for idx, text in enumerate(page_texts):
    if text:
        lines = text.splitlines()
        totals = []
        for i, line in enumerate(lines):
            if re.search(r'\btotal\b', line, re.IGNORECASE) and re.search(r'\$[0-9,]+\.\d{2}', line):
                totals.append((i, line.strip()))
        page_totals[idx] = totals

def get_closest_total(page_idx, after_line):
    if page_idx not in page_totals:
        return None
    for i, line in page_totals[page_idx]:
        if i > after_line and line not in used_lines:
            used_lines.add(line)
            return line
    return None

# Add tables and totals
for page_idx, table in all_tables:
    if len(table) < 2:
        continue
    header = table[0]
    rows = table[1:]
    wrapped = []

    wrapped.append([Paragraph(str(cell), header_style) for cell in header])
    for row in rows:
        wrapped.append([Paragraph(str(cell) if cell else "", body_style) for cell in row])

    table_obj = LongTable(wrapped, repeatRows=1, splitByRow=True)
    table_obj.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    elements.append(table_obj)
    elements.append(Spacer(1, 12))

    # Estimate last line index of this table in source text
    total_line = get_closest_total(page_idx, after_line=0)
    if total_line:
        elements.append(Paragraph(total_line, bold_center_style))
        elements.append(Spacer(1, 24))

# Grand Total detection
grand_total = None
for text in reversed(page_texts):
    if text:
        matches = re.findall(r'Grand Total.*?\$[0-9,]+\.\d{2}', text, re.IGNORECASE | re.DOTALL)
        if matches:
            match = re.search(r'\$[0-9,]+\.\d{2}', matches[-1])
            if match:
                grand_total = match.group(0)
                break

if grand_total:
    elements.append(PageBreak())
    elements.append(Paragraph("Grand Total", title_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Total {grand_total}", bold_center_style))

doc.build(elements, canvasmaker=NumberedCanvas)
buf.seek(0)

st.success("✔️ Transformation complete!")

st.download_button(
    "📥 Download deliverable PDF (11x17 landscape)",
    data=buf,
    file_name="proposal_deliverable_11x17.pdf",
    mime="application/pdf",
    use_container_width=True,
)
