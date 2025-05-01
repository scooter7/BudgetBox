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
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image as RLImage
)
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import re

# ─── Configuration ─────────────────────────────────────────────────────────────

# Register fonts
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow",    "fonts/Barlow-Regular.ttf"))

# OpenAI client
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically-formatted proposal PDF; download a cleaned 11×17 landscape PDF.")

# ─── Upload PDF ────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload a PDF to begin.")
    st.stop()
pdf_bytes = uploaded.read()

# ─── GPT-4 Vision Extraction ──────────────────────────────────────────────────

def extract_strategy_from_image(pil_img: Image.Image) -> dict:
    """
    Sends a cropped table-cell image to GPT-4 Vision with a few-shot prompt
    to extract bold text as 'Strategy' and the rest as 'Description'.
    Returns {'Strategy': str, 'Description': str} or empty strings on failure.
    """
    buffered = io.BytesIO()
    pil_img.save(buffered, format="PNG")
    b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")

    # Few-shot prompt: system + example + the image
    messages = [
        {"role": "system", "content":
            "You are a JSON extractor. Given an image of a table cell, "
            "identify the text that is visually bold (that is the Strategy) "
            "and the remaining text (that is the Description). "
            "ONLY respond with valid JSON: {\"Strategy\": \"...\", \"Description\": \"...\"}."
        },
        {"role": "user", "content":
            [
                {"type": "text", "text":
                    "Example:\n\n"
                    "Cell text (bold=STRATEGY, regular=desc):\n"
                    "**Display Retargeting** Retargeting nationwide off pages\n\n"
                    "→\n"
                    ```json
                    {"Strategy":"Display Retargeting","Description":"Retargeting nationwide off pages"}
                    ```  # for clarity, not sent to model
                },
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
            ]
        }
    ]

    try:
        resp = client.chat.completions.create(
            model="gpt-4-vision-preview",
            messages=messages,
            max_tokens=150
        )
        content = resp.choices[0].message.content.strip()
        data = json.loads(content)
        # validate
        if isinstance(data.get("Strategy"), str) and isinstance(data.get("Description"), str):
            return data
    except Exception:
        pass

    # Fallback to empty
    return {"Strategy": "", "Description": ""}

# ─── Build PDF ────────────────────────────────────────────────────────────────

buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape((11*inch, 17*inch)),
    leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=36
)

# Styles
title_style   = ParagraphStyle("Title",   fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
header_style  = ParagraphStyle("Header",  fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_style    = ParagraphStyle("Body",    fontName="Barlow",  fontSize=9,  alignment=TA_LEFT)
bold_left     = ParagraphStyle("BL",      fontName="DMSerif", fontSize=10, alignment=TA_LEFT)
bold_right    = ParagraphStyle("BR",      fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)

elements = []

# Logo
try:
    logo = requests.get(
        "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",
        timeout=5
    ).content
    elements.append(RLImage(io.BytesIO(logo), width=150, height=50))
except:
    pass

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    # Extract proposal title
    texts = [p.extract_text() or "" for p in pdf.pages]
    proposal_title = next(
        (ln for page in texts for ln in page.splitlines() if "proposal" in ln.lower()),
        "Untitled Proposal"
    ).strip()
    elements.append(Spacer(1,12))
    elements.append(Paragraph(proposal_title, title_style))
    elements.append(Spacer(1,24))

    # helper to track used totals
    used = set()
    def find_total(page_idx):
        for ln in texts[page_idx].splitlines():
            if re.search(r'\btotal\b', ln, re.I) and re.search(r'\$\d', ln) and ln not in used:
                used.add(ln)
                return ln.strip()
        return None

    # Process each table
    for pi, page in enumerate(pdf.pages):
        img = page.to_image(resolution=300)
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data)<2: continue

            header = data[0]
            desc_idx = next((i for i,h in enumerate(header)
                             if h and "description" in h.lower()), None)
            if desc_idx is None: continue

            # new header row
            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(header) if i!=desc_idx]
            wrapped = [[Paragraph(str(h or ""), header_style) for h in new_hdr]]

            nrows = len(data)
            row_h = (tbl.bbox[3]-tbl.bbox[1]) / nrows

            for ridx, row in enumerate(data[1:]):
                # compute crop coords
                x0 = tbl.bbox[0] + (desc_idx/len(header))*(tbl.bbox[2]-tbl.bbox[0])
                x1 = tbl.bbox[0] + ((desc_idx+1)/len(header))*(tbl.bbox[2]-tbl.bbox[0])
                y0 = tbl.bbox[1] + ridx*row_h
                y1 = y0 + row_h
                crop = img.original.crop((
                    int(x0*img.original.width/page.width),
                    int(y0*img.original.height/page.height),
                    int(x1*img.original.width/page.width),
                    int(y1*img.original.height/page.height),
                ))

                ext = extract_strategy_from_image(crop)
                strat = ext["Strategy"]
                desc  = ext["Description"]
                rest  = [row[i] for i in range(len(row)) if i!=desc_idx]
                wrapped.append([Paragraph(strat, body_style),
                                Paragraph(desc, body_style)] +
                               [Paragraph(str(r), body_style) for r in rest])

            # column widths: description 45%, others split 55%
            total_w = 17*inch-96
            widths = [
                0.45*total_w if i==1 else (0.55*total_w)/(len(new_hdr)-1)
                for i in range(len(new_hdr))
            ]

            table_obj = LongTable(wrapped, colWidths=widths, repeatRows=1)
            table_obj.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
                ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                ("VALIGN",(0,0),(-1,0),"MIDDLE"),
                ("VALIGN",(0,1),(-1,-1),"TOP"),
            ]))
            elements.append(table_obj)
            elements.append(Spacer(1,12))

            # table total
            tot = find_total(pi)
            if tot:
                lbl,val = re.split(r'\$\s*', tot, maxsplit=1)
                val = "$"+val.strip()
                total_row = [Paragraph(lbl.strip(), bold_left)] + [""]*(len(new_hdr)-2) + [Paragraph(val, bold_right)]
                ttab = LongTable([total_row], colWidths=widths)
                ttab.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                          ("VALIGN",(0,0),(-1,-1),"TOP")]))
                elements.append(ttab)
                elements.append(Spacer(1,24))

    # grand total
    gtot = None
    for tx in reversed(texts):
        m = re.search(r'Grand Total.*?(\$\d[\d,]*\.\d{2})', tx, re.I|re.S)
        if m:
            gtot = m.group(1); break
    if gtot:
        row = [Paragraph("Grand Total", bold_left)] + [""]*(len(new_hdr)-2) + [Paragraph(gtot, bold_right)]
        gtab = LongTable([row], colWidths=widths)
        gtab.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                  ("VALIGN",(0,0),(-1,-1),"TOP")]))
        elements.append(gtab)

# Build & download
doc.build(elements)
buf.seek(0)
st.download_button(
    "📥 Download deliverable PDF",
    data=buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
