import streamlit as st
import pdfplumber
import requests
from pdf2image import convert_from_bytes
from PIL import Image
import base64
import io
import json
from openai import OpenAI
from collections import defaultdict
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

# Fonts
FONT_DIR = "fonts"
pdfmetrics.registerFont(TTFont("DMSerif", f"{FONT_DIR}/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow", f"{FONT_DIR}/Barlow-Regular.ttf"))

# Streamlit UI
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF and download a cleaned, landscape 11x17 PDF using GPT-4 Vision.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# OpenAI client (uses Streamlit secrets)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Convert to images
images = convert_from_bytes(pdf_bytes, dpi=300)

# Helper to extract bold lines
def extract_strategy_from_image(pil_image: Image.Image) -> dict:
    buffered = io.BytesIO()
    pil_image.save(buffered, format="PNG")
    b64_img = base64.b64encode(buffered.getvalue()).decode("utf-8")

    response = client.chat.completions.create(
        model="gpt-4-vision-preview",
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": (
                            "Extract the bold portion of the text as the Strategy. "
                            "All remaining non-bold text should be listed as Description. "
                            "Only respond with JSON like: {\"Strategy\": \"...\", \"Description\": \"...\"}."
                        ),
                    },
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64_img}"}}
                ]
            }
        ],
        max_tokens=300
    )

    try:
        content = response.choices[0].message.content.strip()
        return json.loads(content)
    except:
        return {"Strategy": "", "Description": ""}

# Extract text and tables
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    proposal_title = "Untitled Proposal"
    for page in pdf.pages:
        text = page.extract_text()
        if text:
            for line in text.splitlines():
                if "proposal" in line.lower():
                    proposal_title = line.strip()
                    break

    all_tables = []
    for page_idx, page in enumerate(pdf.pages):
        tables = page.find_tables()
        for t in tables:
            all_tables.append((page_idx, t))

# Styles
title_style = ParagraphStyle("Title", fontName="DMSerif", fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER, leading=12)
body_style = ParagraphStyle("Body", fontName="Barlow", fontSize=9, alignment=TA_LEFT, leading=11)
bold_right_style = ParagraphStyle("BoldRight", fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)
bold_left_style = ParagraphStyle("BoldLeft", fontName="DMSerif", fontSize=10, alignment=TA_LEFT)

# PDF buffer
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape((11 * inch, 17 * inch)),
    leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=36
)
elements = []

# Logo + Title
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    logo_data = requests.get(logo_url, timeout=5).content
    elements.append(RLImage(io.BytesIO(logo_data), width=150, height=50))
    elements.append(Spacer(1, 12))
except:
    st.warning("Logo not loaded.")
elements.append(Paragraph(proposal_title, title_style))
elements.append(Spacer(1, 24))

# Render tables
for page_idx, table in all_tables:
    data = table.extract()
    if not data or len(data) < 2:
        continue
    header = data[0]
    rows = data[1:]
    desc_idx = next((i for i, h in enumerate(header) if h and "description" in h.lower()), None)
    if desc_idx is None:
        continue

    ncols = len(header)
    table_bbox = table.bbox
    row_height = (table_bbox[3] - table_bbox[1]) / len(data)

    new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
    wrapped = [[Paragraph(col, header_style) for col in new_header]]

    for row_idx, row in enumerate(rows):
        if desc_idx >= len(row):
            continue
        text = row[desc_idx]
        if not text or not text.strip():
            continue

        cell_x0 = table_bbox[0] + (desc_idx / ncols) * (table_bbox[2] - table_bbox[0])
        cell_x1 = table_bbox[0] + ((desc_idx + 1) / ncols) * (table_bbox[2] - table_bbox[0])
        cell_top = table_bbox[1] + row_idx * row_height
        cell_bottom = table_bbox[1] + (row_idx + 1) * row_height

        img = images[page_idx]
        scale_x = img.size[0] / pdf.pages[page_idx].width
        scale_y = img.size[1] / pdf.pages[page_idx].height

        crop = img.crop((
            int(cell_x0 * scale_x),
            int(cell_top * scale_y),
            int(cell_x1 * scale_x),
            int(cell_bottom * scale_y)
        ))

        extracted = extract_strategy_from_image(crop)
        strategy = extracted.get("Strategy", "")
        description = extracted.get("Description", "")
        rest = [row[i] for i in range(len(row)) if i != desc_idx]
        row_data = [strategy, description] + rest
        wrapped.append([Paragraph(str(c), body_style) for c in row_data])

    col_widths = []
    total_width = 17 * inch - 96
    for i in range(len(new_header)):
        col_widths.append(0.45 * total_width if i == 1 else (0.55 * total_width) / (len(new_header) - 1))

    table_obj = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
    table_obj.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(table_obj)
    elements.append(Spacer(1, 24))

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
