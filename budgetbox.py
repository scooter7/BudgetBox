# app.py

import io
import streamlit as st
import pdfplumber
import pandas as pd
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

st.set_page_config(layout="wide")
st.title("📄 PDF WorkDesk + Proposal Transformer")
st.write("Upload a proposal PDF (vertical layout) and download a cleaned, horizontal-table deliverable in landscape PDF format.")

with st.expander("🔄 Transform Proposal Layout", expanded=True):
    uploaded = st.file_uploader("Upload source proposal PDF", type="pdf")
    if uploaded:
        # 1. Extract all tables on pages 1–2
        tables = []
        with pdfplumber.open(uploaded) as pdf:
            for page in pdf.pages[:2]:
                for table in page.extract_tables():
                    tables.append(table)

        if not tables:
            st.error("No tables found. Make sure your PDF has extractable tables.")
        else:
            # 2. Build DataFrame from the first table
            raw = tables[0]
            df = pd.DataFrame(raw[1:], columns=raw[0])

            # 3. Clean columns
            df = (
                df
                .rename(columns=lambda c: c.strip())
                .loc[:, ["Description", "Term", "Start Date", "End Date", "Monthly Amount", "Item Total", "Notes"]]
            )

            # 4. Drop total rows
            df = df[~df["Description"].str.contains("Total", case=False, na=False)].copy()

            # 5. Split the top‐line (strategy) from the rest of the description
            split = df["Description"].str.split(r"\n", 1, expand=True)
            df["Strategy"]    = split[0].str.strip()
            df["Description"] = split[1].str.strip().fillna("")

            # 6. Reorder columns
            cols = ["Strategy", "Description", "Term", "Start Date", "End Date", "Monthly Amount", "Item Total", "Notes"]
            df = df[cols]

            st.subheader("Transformed table")
            st.dataframe(df)

            # 7. Render to landscape PDF
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=landscape(letter),
                                    rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()
            elems = [Paragraph("Proposal Deliverable", styles["Title"]), Spacer(1,12)]

            data = [df.columns.tolist()] + df.values.tolist()
            table = Table(data, repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003f5c")),
                ("TEXTCOLOR",  (0,0), (-1,0), colors.whitesmoke),
                ("ALIGN",      (0,0), (-1,-1), "CENTER"),
                ("GRID",       (0,0), (-1,-1), 0.5, colors.grey),
                ("FONTSIZE",   (0,0), (-1,0), 12),
                ("FONTSIZE",   (0,1), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,0), 8),
            ]))
            elems.append(table)
            doc.build(elems)
            buffer.seek(0)

            st.success("✔️ Transformation complete!")
            st.download_button(
                "📥 Download transformed PDF (landscape)",
                data=buffer,
                file_name="proposal_deliverable.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
