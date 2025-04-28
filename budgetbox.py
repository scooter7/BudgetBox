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
        with pdfplumber.open(uploaded) as pdf:
            tables = []
            for page in pdf.pages[:2]:
                tables.extend(page.extract_tables() or [])

        if not tables:
            st.error("No tables found. Make sure your PDF has extractable tables.")
        else:
            raw = tables[0]

            # 2. Parse & clean the header row, then drop any blank headers
            expected_cols = [
                "Description",
                "Term",
                "Start Date",
                "End Date",
                "Monthly Amount",
                "Item Total",
                "Notes",
            ]
            # turn raw[0] into simple strings (first line before any newline)
            cleaned_hdr = [
                (cell.split("\n")[0].strip() if isinstance(cell, str) else "")
                for cell in raw[0]
            ]
            # keep only non-blank header columns
            keep_idx = [i for i, h in enumerate(cleaned_hdr) if h]
            header_names = [cleaned_hdr[i] for i in keep_idx]

            # 3. Build DataFrame from filtered columns
            data_rows = []
            for row in raw[1:]:
                data_rows.append([row[i] for i in keep_idx])
            df = pd.DataFrame(data_rows, columns=header_names)

            # 4. Now subset to exactly your expected_cols (avoids any KeyError)
            df = df.loc[:, expected_cols].copy()

            # 5. Drop any “Total” rows
            df = df[~df["Description"].str.contains("Total", case=False, na=False)]

            # 6. Split off Strategy vs. Description
            parts = df["Description"].str.split(r"\n", n=1, expand=True)
            df["Strategy"]    = parts[0].str.strip()
            df["Description"] = parts[1].str.strip().fillna("")

            # 7. Reorder into final shape
            final_cols = ["Strategy", "Description"] + expected_cols[1:]
            df = df[final_cols]

            st.subheader("Transformed table")
            st.dataframe(df)

            # 8. Render to landscape PDF (same as before)
            buf = io.BytesIO()
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
            data = [df.columns.tolist()] + df.values.tolist()
            table = Table(data, repeatRows=1)
            table.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003f5c")),
                        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("FONTSIZE", (0, 0), (-1, 0), 12),
                        ("FONTSIZE", (0, 1), (-1, -1), 10),
                        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                    ]
                )
            )
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
