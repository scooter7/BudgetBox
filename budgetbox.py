import streamlit as st
import pdfplumber
import io
import requests
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image as RLImage
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import re

# ─── Register fonts ───────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))

# ─── Streamlit setup ──────────────────────────────────────────────────────────
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF; download a cleaned 11×17 landscape PDF.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# ─── Helper to split first line as Strategy, rest as Description ───────────────
def split_cell_text(raw_text: str):
    lines = [ln.strip() for ln in raw_text.splitlines() if ln.strip()]
    if not lines:
        return "", ""
    strategy = lines[0]
    description = " ".join(lines[1:]) if len(lines) > 1 else ""
    return strategy, description

# ─── Build the PDF ─────────────────────────────────────────────────────────────
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape((11*inch, 17*inch)),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

# Styles
title_style  = ParagraphStyle("Title",  fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body",   fontName="Barlow",  fontSize=9,  alignment=TA_LEFT)
bl_style     = ParagraphStyle("BL",     fontName="DMSerif", fontSize=10, alignment=TA_LEFT)
br_style     = ParagraphStyle("BR",     fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)

elements = []

# ─── Add logo & title ──────────────────────────────────────────────────────────
try:
    logo = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5
    ).content
    elements.append(RLImage(io.BytesIO(logo), width=150, height=50))
except:
    pass

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts = [p.extract_text() or "" for p in pdf.pages]
    # Extract proposal title
    proposal_title = next(
        (ln for pg in texts for ln in pg.splitlines() if "proposal" in ln.lower()),
        "Untitled Proposal"
    ).strip()
    elements += [Spacer(1,12), Paragraph(proposal_title, title_style), Spacer(1,24)]

    used_totals = set()
    def find_total(page_idx):
        for ln in texts[page_idx].splitlines():
            if re.search(r'\btotal\b', ln, re.I) and re.search(r'\$\d', ln) and ln not in used_totals:
                used_totals.add(ln)
                return ln.strip()
        return None

    # Process each page table
    for pi, page in enumerate(pdf.pages):
        img = page.to_image(resolution=150)  # still generate images to preserve table bbox
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data) < 2:
                continue

            header = data[0]
            desc_i = next((i for i,h in enumerate(header) if h and "description" in h.lower()), None)
            if desc_i is None:
                continue

            # Build new header row
            new_hdr = ["Strategy", "Description"] + [h for i,h in enumerate(header) if i != desc_i]
            wrapped = [[Paragraph(str(h or ""), header_style) for h in new_hdr]]

            # Process each data row
            for row in data[1:]:
                raw_desc = row[desc_i] or ""
                strat, desc = split_cell_text(str(raw_desc))
                rest = [row[i] for i in range(len(row)) if i != desc_i]
                cells = [Paragraph(strat, body_style), Paragraph(desc, body_style)] + [
                    Paragraph(str(r or ""), body_style) for r in rest
                ]
                wrapped.append(cells)

            # Compute column widths: description wide
            total_w = 17*inch - 96
            col_widths = [
                0.45*total_w if i==1 else (0.55*total_w)/(len(new_hdr)-1)
                for i in range(len(new_hdr))
            ]

            table_obj = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
            table_obj.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#F2F2F2")),
                ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
                ("VALIGN",     (0,0), (-1,0), "MIDDLE"),
                ("VALIGN",     (0,1), (-1,-1), "TOP"),
            ]))
            elements += [table_obj, Spacer(1,12)]

            # Add the table total row
            total_line = find_total(pi)
            if total_line:
                lbl, val = re.split(r'\$\s*', total_line, 1)
                val = "$" + val.strip()
                tr = (
                    [Paragraph(lbl.strip(), bl_style)]
                    + [""]*(len(new_hdr)-2)
                    + [Paragraph(val, br_style)]
                )
                ttab = LongTable([tr], colWidths=col_widths)
                ttab.setStyle(TableStyle([
                    ("GRID",   (0,0), (-1,-1), 0.25, colors.grey),
                    ("VALIGN", (0,0), (-1,-1), "TOP"),
                ]))
                elements += [ttab, Spacer(1,24)]

    # Grand total row
    gtot = None
    for tx in reversed(texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})', tx, re.I|re.S)
        if m:
            gtot = m.group(1)
            break
    if gtot:
        gr = (
            [Paragraph("Grand Total", bl_style)]
            + [""]*(len(new_hdr)-2)
            + [Paragraph(gtot, br_style)]
        )
        gtab = LongTable([gr], colWidths=col_widths)
        gtab.setStyle(TableStyle([
            ("GRID",   (0,0), (-1,-1), 0.25, colors.grey),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
        ]))
        elements.append(gtab)

# ─── Finish & Download ────────────────────────────────────────────────────────
doc.build(elements)
buf.seek(0)
st.download_button(
    "📥 Download deliverable PDF",
    data=buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
