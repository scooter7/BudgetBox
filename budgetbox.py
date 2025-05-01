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

# Register custom fonts
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow",   "fonts/Barlow-Regular.ttf"))

# OpenAI client (vision-capable model)
client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# Streamlit layout
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
    """Send a table-cell image to GPT-4 Vision to split bold (Strategy) vs. regular (Description)."""
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()

    messages = [
        {"role": "system", "content":
            "You are a JSON extractor. Given an image of a table cell, "
            "return ONLY valid JSON with keys “Strategy” (the visually bold text) "
            "and “Description” (the remaining text)."
        },
        {"role": "user", "content": [
            {"type": "text", "text":
                "Example:\n"
                "**Display Retargeting** Retargeting nationwide off pages\n"
                "→ {\"Strategy\":\"Display Retargeting\",\"Description\":\"Retargeting nationwide off pages\"}"
            },
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
        ]}
    ]

    try:
        resp = client.chat.completions.create(
            model="gpt-4-vision-preview",
            messages=messages,
            max_tokens=150
        )
        data = json.loads(resp.choices[0].message.content.strip())
        return {
            "Strategy": data.get("Strategy","").strip(),
            "Description": data.get("Description","").strip()
        }
    except Exception:
        return {"Strategy": "", "Description": ""}

# ─── Build 11×17 Landscape PDF ────────────────────────────────────────────────

buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape((11*inch,17*inch)),
    leftMargin=48, rightMargin=48, topMargin=48, bottomMargin=36
)

# Styles
title_style  = ParagraphStyle("Title",  fontName="DMSerif", fontSize=18, alignment=TA_CENTER)
header_style = ParagraphStyle("Header", fontName="DMSerif", fontSize=10, alignment=TA_CENTER)
body_style   = ParagraphStyle("Body",   fontName="Barlow",  fontSize=9,  alignment=TA_LEFT)
bl_style     = ParagraphStyle("BL",     fontName="DMSerif", fontSize=10, alignment=TA_LEFT)
br_style     = ParagraphStyle("BR",     fontName="DMSerif", fontSize=10, alignment=TA_RIGHT)

elements = []

# Add logo
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
    def find_total(pi):
        for ln in texts[pi].splitlines():
            if re.search(r'\btotal\b',ln,re.I) and re.search(r'\$\d',ln) and ln not in used_totals:
                used_totals.add(ln); return ln.strip()
        return None

    # Process each table
    for pi, page in enumerate(pdf.pages):
        img = page.to_image(resolution=300)
        for tbl in page.find_tables():
            data = tbl.extract()
            if len(data)<2: continue

            hdr = data[0]
            desc_i = next((i for i,h in enumerate(hdr) if h and "description" in h.lower()), None)
            if desc_i is None: continue

            new_hdr = ["Strategy","Description"] + [h for i,h in enumerate(hdr) if i!=desc_i]
            rows_wrapped = [[Paragraph(str(h), header_style) for h in new_hdr]]

            n = len(data)
            rh = (tbl.bbox[3]-tbl.bbox[1]) / n

            for ridx,row in enumerate(data[1:]):
                # compute crop coords
                x0 = tbl.bbox[0] + desc_i/len(hdr)*(tbl.bbox[2]-tbl.bbox[0])
                x1 = tbl.bbox[0] + (desc_i+1)/len(hdr)*(tbl.bbox[2]-tbl.bbox[0])
                y0 = tbl.bbox[1] + ridx*rh
                y1 = y0 + rh
                crop = img.original.crop((
                    int(x0*img.original.width/page.width),
                    int(y0*img.original.height/page.height),
                    int(x1*img.original.width/page.width),
                    int(y1*img.original.height/page.height),
                ))
                ext = extract_strategy_from_image(crop)
                strat = ext["Strategy"]
                desc  = ext["Description"]
                rest  = [row[i] for i in range(len(row)) if i!=desc_i]
                rows_wrapped.append(
                    [Paragraph(strat,body_style), Paragraph(desc,body_style)] +
                    [Paragraph(str(r),body_style) for r in rest]
                )

            # column widths
            tw = 17*inch-96
            cw = [0.45*tw if i==1 else (0.55*tw)/(len(new_hdr)-1) for i in range(len(new_hdr))]

            table_obj = LongTable(rows_wrapped, colWidths=cw, repeatRows=1)
            table_obj.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
                ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                ("VALIGN",(0,0),(-1,0),"MIDDLE"),
                ("VALIGN",(0,1),(-1,-1),"TOP"),
            ]))
            elements += [table_obj, Spacer(1,12)]

            # table total row
            tot = find_total(pi)
            if tot:
                lbl,val = re.split(r'\$\s*', tot,1)
                val = "$"+val.strip()
                tr = [Paragraph(lbl,bl_style)] + [""]*(len(new_hdr)-2) + [Paragraph(val,br_style)]
                ttab = LongTable([tr], colWidths=cw)
                ttab.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                          ("VALIGN",(0,0),(-1,-1),"TOP")]))
                elements += [ttab, Spacer(1,24)]

    # Grand total
    gtot=None
    for tx in reversed(texts):
        m=re.search(r'Grand Total.*?(\$\d[\d,\,]*\.\d{2})',tx,re.I|re.S)
        if m: gtot=m.group(1); break
    if gtot:
        gr=[Paragraph("Grand Total",bl_style)] + [""]*(len(new_hdr)-2) + [Paragraph(gtot,br_style)]
        gtab=LongTable([gr], colWidths=cw)
        gtab.setStyle(TableStyle([("GRID",(0,0),(-1,-1),0.25,colors.grey),
                                  ("VALIGN",(0,0),(-1,-1),"TOP")]))
        elements.append(gtab)

# build & download
doc.build(elements)
buf.seek(0)
st.download_button(
    "📥 Download deliverable PDF",
    data=buf,
    file_name="proposal_deliverable.pdf",
    mime="application/pdf",
    use_container_width=True,
)
