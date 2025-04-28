# app.py

import os
import sys
import traceback
from io import BytesIO

import streamlit as st
from pypdf import PaperSize, PdfReader, PdfWriter, Transformation
from pypdf.errors import FileNotDecryptedError
from streamlit import session_state
from streamlit_pdf_viewer import pdf_viewer

import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

from utils import helpers, init_session_states, page_config, render_sidebar

# ——— PAGE CONFIG —————————————————————————————————————————————————————————
page_config.set()

# ——— HEADER ——————————————————————————————————————————————————————————————
st.title("📄 PDF WorkDesk!")
st.write(
    "User-friendly, lightweight, and open-source tool to preview and extract content and metadata from PDFs, "
    "add or remove passwords, modify, merge, convert and compress PDFs."
)

init_session_states.init()
render_sidebar.render()

# ——— OPERATIONS ——————————————————————————————————————————————————————————

try:
    # Load PDF (with password handling)
    try:
        pdf, reader, session_state["password"], session_state["is_encrypted"] = helpers.load_pdf(key="main")
    except FileNotDecryptedError:
        pdf = "password_required"

    if pdf == "password_required":
        st.error("PDF is password protected. Please enter the password to proceed.")
    elif pdf:
        # … your existing extract-text, extract-images, tables, convert-to-Word,
        # change/add password, remove password, rotate, resize/scale, merge PDFs, reduce size …
        # (omitted here for brevity; keep exactly as you had them)

        # ——— NEW: Transform Proposal Layout ————————————————————————————————
        with st.expander("🔄 Transform Proposal Layout", expanded=False):
            uploaded_prop = st.file_uploader(
                "Upload source proposal PDF",
                type="pdf",
                key="transform"
            )
            if uploaded_prop:
                # 1) Extract tables from pages 1–2
                with pdfplumber.open(uploaded_prop) as prop_pdf:
                    raw_tables = []
                    for pg in prop_pdf.pages[:2]:
                        raw_tables.extend(pg.extract_tables() or [])

                if not raw_tables:
                    st.error("No tables found. Make sure your PDF has extractable tables.")
                else:
                    raw = raw_tables[0]

                    # 2) Clean & normalize header row, drop blank columns
                    expected_cols = [
                        "Description",
                        "Term",
                        "Start Date",
                        "End Date",
                        "Monthly Amount",
                        "Item Total",
                        "Notes",
                    ]
                    # a) normalize each header cell
                    cleaned_hdr = []
                    for cell in raw[0]:
                        if isinstance(cell, str):
                            h = cell.replace("\n", " ").strip()
                            if h.lower().startswith("term"):
                                h = "Term"
                            cleaned_hdr.append(h)
                        else:
                            cleaned_hdr.append("")
                    # b) keep only non-empty headers
                    keep_idx = [i for i, h in enumerate(cleaned_hdr) if h]
                    header_names = [cleaned_hdr[i] for i in keep_idx]

                    # 3) Build DataFrame from kept columns
                    rows = []
                    for r in raw[1:]:
                        rows.append([r[i] for i in keep_idx])
                    df = pd.DataFrame(rows, columns=header_names)

                    # 4) Subset exactly your expected_cols (avoids KeyError)
                    df = df.loc[:, expected_cols].copy()

                    # 5) Drop any “Total” rows
                    df = df[~df["Description"].str.contains("Total", case=False, na=False)]

                    # 6) Split Strategy vs. Description
                    parts = df["Description"].str.split(r"\n", n=1, expand=True)
                    df["Strategy"]    = parts[0].str.strip()
                    df["Description"] = parts[1].str.strip().fillna("")

                    # 7) Reorder into final shape
                    final_cols = ["Strategy", "Description"] + expected_cols[1:]
                    df = df[final_cols]

                    st.subheader("Transformed table")
                    st.dataframe(df)

                    # 8) Render to landscape PDF
                    buf = BytesIO()
                    doc = SimpleDocTemplate(
                        buf,
                        pagesize=landscape(letter),
                        rightMargin=20,
                        leftMargin=20,
                        topMargin=20,
                        bottomMargin=20,
                    )
                    styles = getSampleStyleSheet()
                    elems = [Paragraph("Proposal Deliverable", styles["Title"]), Spacer(1, 12)]

                    table_data = [df.columns.tolist()] + df.values.tolist()
                    tbl = Table(table_data, repeatRows=1)
                    tbl.setStyle(TableStyle([
                        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003f5c")),
                        ("TEXTCOLOR",  (0,0), (-1,0), colors.whitesmoke),
                        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
                        ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                        ("FONTSIZE",   (0,0), (-1,0), 12),
                        ("FONTSIZE",   (0,1), (-1,-1), 10),
                        ("BOTTOMPADDING", (0,0), (-1,0), 8),
                    ]))
                    elems.append(tbl)
                    doc.build(elems)
                    buf.seek(0)

                    st.success("✔️ Transformation complete!")
                    st.download_button(
                        "📥 Download transformed PDF (landscape)",
                        data=buf,
                        file_name="proposal_deliverable.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )

except Exception as e:
    st.error(
        f"""The app has encountered an error:  
`{e}`  
Please create an issue [here](https://github.com/SiddhantSadangi/pdf-workdesk/issues/new) with the traceback below."""
        , icon="🥺"
    )
    st.code(traceback.format_exc())

st.success(
    "[Star the repo](https://github.com/SiddhantSadangi/pdf-workdesk) to show your :heart:",
    icon="⭐",
)
