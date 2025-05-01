import streamlit as st
import pdfplumber
import requests
from PIL import Image
import io
import json
import base64
from openai import OpenAI
from reportlab.lib.pagesizes import landscape
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle, Paragraph,
    Spacer, Image as RLImage, PageBreak
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import re

# Fonts
FONT_DIR = "fonts"
pdfmetrics.registerFont(TTFont("DMSerif", f"{FONT_DIR}/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow", f"{FONT_DIR}/Barlow-Regular.ttf"))

# OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Streamlit UI
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# GPT-4 Vision strategy extraction
def extract_strategy_from_image(pil_image: Image.Image) -> dict:
    buffered = io.BytesIO()
    pil_image.save(buffered, format="PNG")
    b64_img = base64.b64encode(buffered.getvalue()).decode("utf-8")
    response = client.chat.completions.create(
        model="gpt-4-vision-preview",
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": (
                    "Extract the bold portion of this table cell as 'Strategy'. "
                    "The remaining regular font text is the 'Description'. "
                    "Respond with JSON: {\"Strategy\": \"...\", \"Description\": \"...\"}."
                )},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64_img}"}}
            ]
        }],
        max_tokens=300
    )
    try:
        return json.loads(response.choices[0].message.content.strip())
    except:
        return {"Strategy": "", "Description": ""}

# PDF output buffer
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape((11 * inch, 17 * inch)),
    leftMargin=48, rightMargin=48,
    topMargin=48, bottomMargin=36,
)

# Styles
title_style = ParagraphStyle("Title", fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_style = ParagraphStyle("Body", fontName="Barlow", fontSize=9, alignment=TA_LEFT)
bold_right = ParagraphStyle("BoldRight", fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)
bold_left = ParagraphStyle("BoldLeft", fontName="DMSerif", fontSize=10, alignment=TA_LEFT)

elements = []

# Logo
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    logo_data = requests.get(logo_url, timeout=5).content
    elements.append(RLImage(io.BytesIO(logo_data), width=150, height=50))
except:
    st.warning("Logo not loaded")

# Process PDF
with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    page_texts = [p.extract_text() or "" for p in pdf.pages]
    proposal_title = "Untitled Proposal"
    for text in page_texts:
        for line in text.splitlines():
            if "proposal" in line.lower():
                proposal_title = line.strip()
                break
        if "proposal" in proposal_title.lower():
            break

    elements.append(Spacer(1, 12))
    elements.append(Paragraph(proposal_title, title_style))
    elements.append(Spacer(1, 24))

    used_total_lines = set()

    def find_total_below(page_idx, start_y):
        lines = page_texts[page_idx].splitlines()
        for line in lines:
            if (re.search(r'\btotal\b', line, re.I) and
                re.search(r'\$[0-9,]+\.\d{2}', line) and
                line not in used_total_lines):
                used_total_lines.add(line)
                return line.strip()
        return None

    for page_idx, page in enumerate(pdf.pages):
        tables = page.find_tables()
        img_page = page.to_image(resolution=300)
        for table in tables:
            data = table.extract()
            if not data or len(data) < 2:
                continue
            header = data[0]
            rows = data[1:]
            desc_idx = next((i for i, h in enumerate(header) if h and "description" in h.lower()), None)
            if desc_idx is None:
                continue

            new_header = ["Strategy", "Description"] + [h for i, h in enumerate(header) if i != desc_idx]
            wrapped = [[Paragraph(h, header_style) for h in new_header]]
            row_height = (table.bbox[3] - table.bbox[1]) / len(data)

            for row_idx, row in enumerate(rows):
                if all(str(cell).lower() == "none" for cell in row):
                    continue
                x0 = table.bbox[0] + (desc_idx / len(header)) * (table.bbox[2] - table.bbox[0])
                x1 = table.bbox[0] + ((desc_idx + 1) / len(header)) * (table.bbox[2] - table.bbox[0])
                y0 = table.bbox[1] + row_idx * row_height
                y1 = table.bbox[1] + (row_idx + 1) * row_height
                cropped = img_page.crop((x0, y0, x1, y1)).original
                result = extract_strategy_from_image(cropped)
                strategy = result.get("Strategy", "")
                description = result.get("Description", "")
                rest = [row[i] for i in range(len(row)) if i != desc_idx]
                row_data = [strategy, description] + rest
                wrapped.append([Paragraph(str(c), body_style) for c in row_data])

            col_widths = []
            total_width = 17 * inch - 96
            for i in range(len(new_header)):
                col_widths.append(0.45 * total_width if i == 1 else (0.55 * total_width) / (len(new_header) - 1))

            tbl = LongTable(wrapped, colWidths=col_widths, repeatRows=1)
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("VALIGN", (0, 1), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]))
            elements.append(tbl)
            elements.append(Spacer(1, 12))

            total_line = find_total_below(page_idx, table.bbox[3])
            if total_line:
                try:
                    label, value = re.split(r'\$+', total_line, maxsplit=1)
                    label = label.strip()
                    value = "$" + value.strip()
                    row = [Paragraph(label, bold_left), "", "", "", "", "", Paragraph(value, bold_right)]
                    total_table = LongTable([row], colWidths=[2.5*inch] + [1.1*inch]*5 + [2.5*inch])
                    total_table.setStyle(TableStyle([
                        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]))
                    elements.append(total_table)
                    elements.append(Spacer(1, 24))
                except:
                    pass

    # Grand total (if any)
    grand_total = None
    for text in reversed(page_texts):
        matches = re.findall(r'Grand Total.*?\$[0-9,]+\.\d{2}', text, re.I)
        if matches:
            match = re.search(r'\$[0-9,]+\.\d{2}', matches[-1])
            if match:
                grand_total = match.group(0)
                break

    if grand_total:
        row = [Paragraph("Grand Total", bold_left), "", "", "", "", "", Paragraph(grand_total, bold_right)]
        grand_row = LongTable([row], colWidths=[2.5*inch] + [1.1*inch]*5 + [2.5*inch])
        grand_row.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements.append(grand_row)

# Build final PDF
doc.build(elements)
buf.seek(0)

st.download_button(
    "📥 Download PDF Deliverable",
    data=buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
