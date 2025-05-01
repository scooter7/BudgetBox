import io
import os
import re
import streamlit as st
import pdfplumber
import requests
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

# Page size
tabloid = (11 * inch, 17 * inch)

# Font registration
FONT_DIR = "fonts"
pdfmetrics.registerFont(TTFont("DMSerif", os.path.join(FONT_DIR, "DMSerifDisplay-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Barlow", os.path.join(FONT_DIR, "Barlow-Regular.ttf")))

# Streamlit UI
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a landscape PDF with cleaned formatting.")

uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# Extract text and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = "Untitled Proposal"
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            lines = text.splitlines()
            for line in lines:
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break
        if "proposal" in proposal_title.lower():
            break
    page_texts = [page.extract_text() for page in pdf.pages]
    all_tables = [(i, t) for i, page in enumerate(pdf.pages) for t in (page.extract_tables() or [])]

# Total line tracker
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

# Styles
title_style = ParagraphStyle(
    name="Title",
    fontName="DMSerif",
    fontSize=18,
    alignment=TA_CENTER,
    spaceAfter=12,
)

header_style = ParagraphStyle(
    name="Header",
    fontName="DMSerif",
    fontSize=10,
    alignment=TA_CENTER,
    leading=12
)

body_style = ParagraphStyle(
    name="Body",
    fontName="Barlow",
    fontSize=9,
    alignment=TA_LEFT,
    leading=11
)

bold_right_style = ParagraphStyle(
    name="BoldRight",
    fontName="DMSerif",
    fontSize=10,
    alignment=TA_RIGHT,
)

bold_left_style = ParagraphStyle(
    name="BoldLeft",
    fontName="DMSerif",
    fontSize=10,
    alignment=TA_LEFT,
)

# Create PDF buffer
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(tabloid),
    leftMargin=48,
    rightMargin=48,
    topMargin=48,
    bottomMargin=36
)

elements = []

# Carnegie logo
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    logo_resp = requests.get(logo_url, timeout=5)
    logo_resp.raise_for_status()
    elements.append(Image(io.BytesIO(logo_resp.content), width=150, height=50))
    elements.append(Spacer(1, 12))
except:
    st.warning("Logo couldn't be loaded.")

elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Process each table
for page_idx, raw in all_tables:
    if len(raw) < 2:
        continue
    header = [str(c).strip() if c else "" for c in raw[0]]
    rows = raw[1:]

    non_none = [i for i, h in enumerate(header) if h.lower() != "none" and h.strip()]
    header = [header[i] for i in non_none]
    rows = [[row[i] if i < len(row) else "" for i in non_none] for row in rows]

    desc_idx = next((i for i, h in enumerate(header) if "description" in h.lower()), None)
    if desc_idx is not None:
        new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
        new_rows = []
        for row in rows:
            desc_raw = row[desc_idx] or ""
            lines = str(desc_raw).split("\n")
            if len(lines) == 1:
                strategy = lines[0]
                description = ""
            elif len(lines) >= 2 and len(lines[1].strip()) <= 50:
                strategy = f"{lines[0].strip()} {lines[1].strip()}"
                description = "\n".join(lines[2:]).strip()
            else:
                strategy = lines[0].strip()
                description = "\n".join(lines[1:]).strip()
            rest = [row[i] for i in range(len(row)) if i != desc_idx]
            new_rows.append([strategy, description] + rest)
        header = new_header
        rows = new_rows

    while rows and str(rows[-1][0]).strip().lower().startswith("total") and all(
        str(x).lower() in ["", "none"] for x in rows[-1][1:]
    ):
        rows.pop()

    wrapped = [[Paragraph(col, header_style) for col in header]]
    for row in rows:
        wrapped.append([Paragraph(str(cell) or "", body_style) for cell in row])

    col_widths = []
    total_width = 17 * inch - 96
    desc_idx = next((i for i, h in enumerate(header) if "description" in h.lower()), None)
    num_cols = len(header)
    for i in range(num_cols):
        if i == desc_idx:
            col_widths.append(0.45 * total_width)
        else:
            col_widths.append((0.55 * total_width) / (num_cols - 1))

    table = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    total_line = get_closest_total(page_idx)
    if total_line and re.search(r"\$[0-9,]+\.\d{2}", total_line):
        label = total_line.split("$")[0].strip()
        amount = "$" + total_line.split("$")[1].strip()
        total_row = [[Paragraph(label, bold_left_style), Paragraph(amount, bold_right_style)]]
        total_table = LongTable(total_row, colWidths=[8.25 * inch, 8.25 * inch])
        total_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ]))
        elements.append(total_table)
        elements.append(Spacer(1, 24))

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
    label = "Grand Total"
    amount = grand_total
    total_row = [[Paragraph(label, bold_left_style), Paragraph(amount, bold_right_style)]]
    total_table = LongTable(total_row, colWidths=[8.25 * inch, 8.25 * inch])
    total_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))
    elements.append(total_table)

doc.build(elements)
buf.seek(0)

st.success("✔️ Transformation complete!")
st.download_button(
    "📥 Download deliverable PDF (landscape)",
    data=buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
