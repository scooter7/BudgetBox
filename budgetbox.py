import io
import streamlit as st
import pdfplumber
import pandas as pd
import requests

from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle,
    Paragraph, Spacer, Image
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

LOGO_URL = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write(
    "Upload a vertically-formatted proposal PDF and get both the full PDF "
    "and a cleaned, horizontally-formatted deliverable in landscape PDF."
)

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.info("Please upload the proposal PDF to begin.")
    st.stop()

pdf_bytes = uploaded.read()

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    first_text = pdf.pages[0].extract_text() or ""
    proposal_title = first_text.split("\n")[0].strip()

    raw_tables = []
    for pg in pdf.pages:
        raw_tables.extend(pg.extract_tables() or [])

if not raw_tables:
    st.error("No tables found in the document.")
    st.stop()

expected_cols = [
    "Description", "Term", "Start Date",
    "End Date", "Monthly Amount", "Item Total", "Notes"
]

def process_table(raw):
    # normalize header
    hdr = []
    for cell in raw[0]:
        if isinstance(cell, str):
            h = cell.replace("\n", " ").strip()
            if h.lower().startswith("term"):
                h = "Term"
            hdr.append(h)
        else:
            hdr.append("")
    keep = [i for i,h in enumerate(hdr) if h]
    headers = [hdr[i] for i in keep]
    rows = [[r[i] if i<len(r) else None for i in keep] for r in raw[1:]]
    df = pd.DataFrame(rows, columns=headers).reindex(columns=expected_cols)
    return df

df = pd.concat([process_table(t) for t in raw_tables if len(t)>1], ignore_index=True)

# split Strategy / Description
parts = df["Description"].fillna("").str.split(r"\n",1,expand=True)
df["Strategy"]    = parts[0].str.strip()
df["Description"] = parts[1].str.strip().fillna("")
final_cols = ["Strategy","Description"] + expected_cols[1:]
df = df[final_cols]

st.subheader("Transformed Data Preview")
st.dataframe(df, use_container_width=True)

# build the deliverable PDF
buf = io.BytesIO()
doc = SimpleDocTemplate(
    buf,
    pagesize=landscape(letter),
    leftMargin=36, rightMargin=36, topMargin=72, bottomMargin=36
)
styles = getSampleStyleSheet()
elements = []

# logo
try:
    r = requests.get(LOGO_URL, timeout=5)
    r.raise_for_status()
    elements.append(Image(io.BytesIO(r.content), width=120, height=40))
    elements.append(Spacer(1,12))
except Exception:
    st.warning("Could not fetch Carnegie logo.")

# title
elements.append(Paragraph(proposal_title, styles["Title"]))
elements.append(Spacer(1,24))

# table w/ splitByRow
wrapped = [
    [Paragraph(str(cell), styles["BodyText"]) for cell in row]
    for row in [df.columns.tolist()] + df.values.tolist()
]
table = Table(wrapped, repeatRows=1, splitByRow=True)
table.setStyle(TableStyle([
    ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
    ("TEXTCOLOR",(0,0),(-1,0),colors.black),
    ("ALIGN",(0,0),(-1,-1),"CENTER"),
    ("GRID",(0,0),(-1,-1),0.5,colors.grey),
    ("FONTSIZE",(0,0),(-1,0),12),
    ("FONTSIZE",(0,1),(-1,-1),10),
    ("BOTTOMPADDING",(0,0),(-1,0),8),
    ("LEFTPADDING",(0,1),(-1,-1),4),
    ("RIGHTPADDING",(0,1),(-1,-1),4),
]))
elements.append(table)

doc.build(elements)
buf.seek(0)

st.success("✔️ Transformation complete!")

c1, c2 = st.columns(2)
with c1:
    st.download_button(
        "📥 Download original PDF",
        data=pdf_bytes,
        file_name=uploaded.name,
        mime="application/pdf",
        use_container_width=True,
    )
with c2:
    st.download_button(
        "📥 Download deliverable PDF (landscape)",
        data=buf,
        file_name="proposal_deliverable.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
