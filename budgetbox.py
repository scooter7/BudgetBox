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
        # 1. Pull out all tables on pages 1–2
        tables = []
        with pdfplumber.open(uploaded) as pdf:
            for page in pdf.pages[:2]:
                tables.extend(page.extract_tables() or [])

        if not tables:
            st.error("No tables found. Make sure your PDF has extractable tables.")
        else:
            raw = tables[0]

            # 2. Determine headers safely
            expected_cols = ["Description", "Term", "Start Date", "End Date",
                             "Monthly Amount", "Item Total", "Notes"]
            # If the PDF gave us a full, all‐string header row, use it (stripped); otherwise fall back:
            if len(raw[0]) == len(expected_cols) and all(isinstance(h, str) for h in raw[0]):
                headers = [h.strip() for h in raw[0]]
            else:
                headers = expected_cols

            df = pd.DataFrame(raw[1:], columns=headers)

            # 3. Now subset exactly the columns we care about (this avoids KeyErrors)
            df = df.loc[:, expected_cols].copy()

            # 4. Drop any “Total” rows
            df = df[~df["Description"].str.contains("Total", case=False, na=False)]

            # 5. Split out Strategy vs. Description
            parts = df["Description"].str.split(r"\n", n=1, expand=True)
            df["Strategy"]    = parts[0].str.strip()
            df["Description"] = parts[1].str.strip().fillna("")

            # 6. Reorder into final column order
            final_cols = ["Strategy", "Description"] + expected_cols[1:]
            df = df[final_cols]

            st.subheader("Transformed table")
            st.dataframe(df)

            # 7. Render to landscape PDF
            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=landscape(letter),
                                    rightMargin=20, leftMargin=20,
                                    topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()
            elems = [Paragraph("Proposal Deliverable", styles["Title"]), Spacer(1, 12)]

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
            buf.seek(0)

            st.success("✔️ Transformation complete!")
            st.download_button(
                "📥 Download transformed PDF (landscape)",
                data=buf,
                file_name="proposal_deliverable.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
