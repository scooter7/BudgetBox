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

# Page setup
TABLOID = (11 * inch, 17 * inch)
FONT_DIR = "fonts"

# Register fonts
pdfmetrics.registerFont(TTFont("DMSerif", os.path.join(FONT_DIR, "DMSerifDisplay-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Barlow", os.path.join(FONT_DIR, "Barlow-Regular.ttf")))

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a styled, landscape 11x17 deliverable.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract title and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = "Untitled Proposal"
    all_lines = []
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.splitlines()
            all_lines.extend(lines)
            for line in lines:
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break
        if "proposal" in proposal_title.lower():
            break

    page_texts = [page.extract_text() for page in pdf.pages]
    all_tables = [(i, t) for i, page in enumerate(pdf.pages) for t in (page.extract_tables() or [])]

# Total rows per page
page_totals = {}
used_lines = set()
for idx, text in enumerate(page_texts):
    if text:
        lines = text.splitlines()
        page_totals[idx] = [(i, line.strip()) for i, line in enumerate(lines)
                             if re.search(r'\btotal\b', line, re.I) and re.search(r'\$[0-9,]+\.\d{2}', line)]

def get_closest_total(page_idx):
    for i, line in page_totals.get(page_idx, []):
        if line not in used_lines:
            used_lines.add(line)
            return line
    return None

# PDF setup
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
            self.setFont("Helvetica", 8)
            self.drawRightString(1600, 20, f"Page {self._pageNumber} of {total}")
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(TABLOID),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

# Styles
title_style = ParagraphStyle("Title", fontName="DMSerif", fontSize=18, alignment=TA_CENTER, spaceAfter=6)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=11, alignment=TA_CENTER)
body_style = ParagraphStyle("Body", fontName="Barlow", fontSize=9, leading=11)
bold_center = ParagraphStyle("BoldCenter", fontName="DMSerif", fontSize=10, alignment=TA_CENTER, spaceAfter=12)

# Build PDF
elements = []
try:
    r = requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png", timeout=5)
    r.raise_for_status()
    elements.append(Image(io.BytesIO(r.content), width=150, height=50))
    elements.append(Spacer(1, 12))
except:
    st.warning("Logo couldn't be loaded.")
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

for page_idx, raw in all_tables:
    if len(raw) < 2:
        continue
    header = [str(c).strip() if c else "" for c in raw[0]]
    rows = raw[1:]

    desc_idx = next((i for i, h in enumerate(header) if "description" in h.lower()), None)
    if desc_idx is not None:
        new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
        new_rows = []
        for row in rows:
            desc_raw = row[desc_idx] or ""
            desc_lines = str(desc_raw).split("\n")
            if len(desc_lines) >= 2:
                strategy = " ".join(desc_lines[:2]).strip()
                description = "\n".join(desc_lines[2:]).strip()
            else:
                strategy = desc_lines[0].strip()
                description = ""
            rest = [row[i] for i in range(len(row)) if i != desc_idx]
            new_rows.append([strategy, description] + rest)
        header = new_header
        rows = new_rows

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
        ("FONTNAME", (0, 1), (-1, -1), "Barlow"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 12))

    total_line = get_closest_total(page_idx)
    if total_line:
        elements.append(Paragraph(total_line, bold_center))
        elements.append(Spacer(1, 24))

# Grand total
grand_total = None
for text in reversed(page_texts):
    if text:
        m = re.search(r'Grand Total.*?\$[0-9,]+\.\d{2}', text, re.I | re.DOTALL)
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

# Build PDF
doc.build(elements, canvasmaker=NumberedCanvas)
buf.seek(0)

st.success("✔️ Transformation complete!")
st.download_button(
    "📅 Download deliverable PDF (11x17 landscape)",
    data=buf,
    file_name="proposal_deliverable_11x17.pdf",
    mime="application/pdf",
    use_container_width=True,
)
