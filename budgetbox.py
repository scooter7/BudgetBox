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

# 11x17 tabloid
tabloid = (11 * inch, 17 * inch)

# Register fonts
FONT_DIR = "fonts"
pdfmetrics.registerFont(TTFont("DMSerif", os.path.join(FONT_DIR, "DMSerifDisplay-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Barlow", os.path.join(FONT_DIR, "Barlow-Regular.ttf")))

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a styled, landscape 11x17 deliverable.")

# Upload
uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Awaiting PDF upload.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract title + tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    all_text_lines = []
    proposal_title = "Untitled Proposal"
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.splitlines()
            all_text_lines.extend(lines)
            for line in lines:
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break
        if "proposal" in proposal_title.lower():
            break

    page_texts = [p.extract_text() for p in pdf.pages]
    all_tables = []
    for i, page in enumerate(pdf.pages):
        tables = page.extract_tables()
        for t in tables:
            all_tables.append((i, t))

# PDF buffer
buf = io.BytesIO()

class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.pages = []

    def showPage(self):
        self.pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self.pages)
        for p in self.pages:
            self.__dict__.update(p)
            self.draw_page_number(total)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, total):
        self.setFont("Helvetica", 8)
        self.drawRightString(1600, 20, f"Page {self._pageNumber} of {total}")

# PDF doc setup
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(tabloid),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

# Styles
title_style = ParagraphStyle("Title", fontName="DMSerif", fontSize=18, alignment=TA_CENTER, spaceAfter=6)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=11, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body", fontName="Barlow", fontSize=9, leading=11)
bold_center  = ParagraphStyle("BoldCenter", fontName="Barlow", fontSize=10, alignment=TA_CENTER, spaceAfter=12)

elements = []

# Logo + Title
try:
    r = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png", timeout=5)
    r.raise_for_status()
    elements.append(Image(io.BytesIO(r.content), width=150, height=50))
    elements.append(Spacer(1, 12))
except:
    st.warning("Could not load logo.")
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Total line detection
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
    for i, line in page_totals.get(page_idx, []):
        if i > after_line and line not in used_lines:
            used_lines.add(line)
            return line
    return None

# Process tables
for page_idx, raw in all_tables:
    if len(raw) < 2:
        continue

    original_header = [str(cell).strip() if cell else "" for cell in raw[0]]
    rows = raw[1:]

    # Remove empty columns
    non_empty_cols = [i for i, h in enumerate(original_header) if h and h.lower() != "none"]
    filtered_header = [original_header[i] for i in non_empty_cols]

    # Split Description
    desc_idx = next((i for i, h in enumerate(filtered_header) if "description" in h.lower()), None)
    if desc_idx is not None:
        new_header = ["Strategy", "Description"] + [h for i, h in enumerate(filtered_header) if i != desc_idx]
        new_rows = []
        for row in rows:
            filtered_row = [row[i] if i < len(row) else "" for i in non_empty_cols]
            desc = filtered_row[desc_idx]
            parts = str(desc).split("\n", 1)
            strategy = parts[0].strip()
            description = parts[1].strip() if len(parts) > 1 else ""
            rest = [filtered_row[i] for i in range(len(filtered_row)) if i != desc_idx]
            new_rows.append([strategy, description] + rest)
        header = new_header
        rows = new_rows
    else:
        header = filtered_header
        rows = [[row[i] for i in non_empty_cols] for row in rows]

    num_cols = len(header)
    col_widths = [100, 250] + [80] * (num_cols - 2)

    wrapped = [[Paragraph(str(c), header_style) for c in header]]
    for row in rows:
        wrapped.append([Paragraph(str(c or ""), body_style) for c in row])

    t = LongTable(wrapped, repeatRows=1, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (1, 0), (-1, 0), "CENTER"),
        ("FONTNAME", (0, 1), (-1, -1), "Barlow"),  # ensure regular font for all rows except header
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 12))

    total_line = get_closest_total(page_idx, after_line=0)
    if total_line:
        elements.append(Paragraph(total_line, bold_center))
        elements.append(Spacer(1, 24))

# Grand Total
grand_total = None
for text in reversed(page_texts):
    if text:
        m = re.search(r'Grand Total.*?\$[0-9,]+\.\d{2}', text, re.IGNORECASE | re.DOTALL)
        if m:
            match = re.search(r'\$[0-9,]+\.\d{2}', m.group(0))
            if match:
                grand_total = match.group(0)
                break

if grand_total:
    elements.append(PageBreak())
    elements.append(Paragraph("Grand Total", title_style))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Total {grand_total}", bold_center))

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
