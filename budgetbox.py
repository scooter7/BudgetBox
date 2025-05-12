# -*- coding: utf-8 -*-
import io
import re
import camelot
import pdfplumber
import fitz
# Removed requests import as it's unused
import streamlit as st
# Removed PIL import as it's unused
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT # Corrected this line
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer # Removed RLImage import as it's unused

# Font registration
pdfmetrics.registerFont(TTFont("DMSerif","fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow","fonts/Barlow-Regular.ttf"))
DEFAULT_SERIF_FONT="DMSerif"
DEFAULT_SANS_FONT="Barlow"

# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download the re-formatted PDF output.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")

if not uploaded:
    st.stop()

pdf_bytes = uploaded.read()

# --- Fitz setup for rich text extraction ---
try:
    doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
except Exception as e:
    st.error(f"Error opening PDF with Fitz: {e}")
    st.stop()

def extract_rich_cell(page_number, bbox):
    """Extracts text with basic formatting (bold, line breaks) from a PDF cell bbox."""
    try:
        page = doc_fitz.load_page(page_number)
        # Use clip argument for get_text("dict") for better bounding box accuracy
        # words = page.get_text("words", clip=bbox) # Get words within the bbox

        # if not words:
        #     return ""

        # lines = {}
        # # Group words by baseline (y1)
        # for w in words:
        #     # x0, y0, x1, y1, word, block_no, line_no, word_no
        #     key = round(w[3], 1) # Use y1 as the key for line grouping
        #     lines.setdefault(key, []).append(w)

        # text_lines = []
        # # Sort lines by vertical position, then words by horizontal position
        # for key in sorted(lines.keys()):
        #     # Sort words in the line by their x0 coordinate
        #     row = sorted(lines[key], key=lambda w: w[0])
        #     pieces = []
        #     for word_info in row:
        #         t = word_info[4].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
        #         # Check flags for bold (flag 2) - need font info for more styles
        #         # Get span info to check flags requires "dict" or "rawdict"
        #         # This simplified word approach won't easily get bold.
        #         # Reverting to get_text("dict") approach but clipped.
        #         # Let's stick to the original 'extract_rich_cell' for simplicity for now,
        #         # accepting its limitations, but apply it to the notes column.
        #         # The original function is actually better suited if spans cross bbox.
        #         # Re-implementing the original logic slightly cleaned up:

                # Re-fetch with dict for span info, using clip
        d = page.get_text("dict", clip=bbox)
        spans = []
        x0_bbox,y0_bbox,x1_bbox,y1_bbox = bbox # Renamed to avoid conflict with span bbox variables
        for block in d["blocks"]:
            if block.get("type")!=0: # Text blocks only
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    # Check if span intersects bbox (more robust than simple contains)
                    sx0,sy0,sx1,sy1 = span["bbox"]
                    # Check for overlap:
                    # The span is overlapping if sx0 < x1_bbox and sx1 > x0_bbox and sy0 < y1_bbox and sy1 > y0_bbox
                    if sx0 < x1_bbox and sx1 > x0_bbox and sy0 < y1_bbox and sy1 > y0_bbox:
                        spans.append(span)

        lines_dict = {}
        for s in spans:
            # Group by baseline y1, rounded. Using s["bbox"][1] (y0) for top of span.
            # Or s["origin"][1] might be more consistent for line grouping if available.
            # Let's use the top of the span's bbox (sy0) for grouping into lines.
            key = round(s["bbox"][1], 1)
            lines_dict.setdefault(key, []).append(s)

        span_text_lines = []
        # Sort lines by their vertical position (key)
        for key in sorted(lines_dict.keys()):
            # Sort spans in line by x0
            row_spans = sorted(lines_dict[key], key=lambda s: s["bbox"][0])
            line_pieces = []
            last_x1 = x0_bbox # Start from the left edge of the bbox for space insertion logic
            for span_idx, span in enumerate(row_spans):
                # Add space if there's a gap between spans (horizontal)
                # or if it's not the first span and there's a reasonable gap
                if span_idx > 0 and span["bbox"][0] > (row_spans[span_idx-1]["bbox"][2] + 1): # Add tolerance for space
                     line_pieces.append(" ")
                t = span["text"].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                is_bold = span["flags"] & 2 # Flag for bold
                # is_italic = span["flags"] & 4 # Example if needed

                if is_bold:
                    line_pieces.append(f"<b>{t}</b>")
                # elif is_italic:
                #     line_pieces.append(f"<i>{t}</i>") # Example
                else:
                    line_pieces.append(t)
                # last_x1 = span["bbox"][2] # Update last_x1 for complex spacing (not strictly needed with current space logic)
            span_text_lines.append("".join(line_pieces))
        return "<br/>".join(span_text_lines)

    except Exception as e:
        # st.warning(f"Error in extract_rich_cell for bbox {bbox} on page {page_number}: {e}") # Detailed for debugging
        return "" # Return empty string on error

# --- Headers ---
HEADERS = [
    "Description",
    "Start Date",
    "End Date",
    "Term (Months)",
    "Monthly Amount",
    "Item Total",
    "Notes"
]

# --- Table Extraction Logic ---
first_table = None
# Attempt Camelot extraction for the first page
try:
    tables_camelot = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n", line_scale=40) # Added line_scale
    if tables_camelot:
        raw = tables_camelot[0].df.values.tolist()
        if len(raw) > 1 and len(raw[0]) >= 6: # Check against a reasonable number of expected cols
             header_row_text = "".join([str(h).lower() for h in raw[0]])
             if any(kw in header_row_text for kw in ['description', 'date', 'term', 'amount', 'total', 'notes']):
                first_table = raw
except Exception as e:
    # st.warning(f"Camelot failed on first page: {e}") # Optional: for debugging
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # Store page text along with 0-indexed page number for Fitz compatibility
        texts = [(p.page_number -1, p.extract_text(x_tolerance=1, y_tolerance=1, layout=True) or "") for p in pdf.pages]
        first_page_idx, first_text = texts[0] if texts else (0, "")
        first_lines = first_text.splitlines()

        # Extract Proposal Title
        pot = next((l.strip() for l in first_lines if "proposal" in l.lower() and len(l.strip())>10), None)
        if pot:
            proposal_title = pot
        elif first_lines:
            proposal_title = next((l.strip() for l in first_lines if l.strip()), "Untitled Proposal")

        used_total_lines = set()
        def find_total(page_idx_for_texts_list):
            """Finds a 'Total' line on a specific page using the 0-indexed page number."""
            if page_idx_for_texts_list >= len(texts):
                return None
            actual_page_num_for_key, text_content = texts[page_idx_for_texts_list] # actual_page_num is 0-indexed
            for l in text_content.splitlines():
                line_key = (actual_page_num_for_key, l)
                if re.search(r'\b(?<!grand\s)total\b.*?\$\s*[\d,.]+',l, re.I) and line_key not in used_total_lines:
                    used_total_lines.add(line_key)
                    return l.strip()
            return None

        # Process tables page by page
        for pi_pdfplumber, page in enumerate(pdf.pages): # pi_pdfplumber is 0-indexed
            # pdfplumber page.page_number is 1-based, Fitz page index is 0-based
            current_page_idx_fitz = page.page_number - 1
            page_content = ""
            for idx, content in texts:
                if idx == current_page_idx_fitz:
                    page_content = content
                    break

            if current_page_idx_fitz == 0 and first_table:
                found = [("camelot", first_table, None, None)]
                page_links = [] # No link info from Camelot
            else:
                # Enhanced table finding settings for pdfplumber
                table_settings = {
                    "vertical_strategy": "lines_strict", # or "text"
                    "horizontal_strategy": "lines_strict", # or "text"
                    "explicit_vertical_lines": page.curves + page.edges, # Consider curves as potential table lines
                    "explicit_horizontal_lines": page.curves + page.edges,
                    "snap_tolerance": 5, # Increased tolerance
                    "join_tolerance": 5,
                    "min_words_vertical": 2, # Adjust based on expected cell content
                    "min_words_horizontal": 1
                }
                found_tables = page.find_tables(table_settings=table_settings)
                if not found_tables: # Fallback to default if strict fails
                    found_tables = page.find_tables()

                found = [(tbl, tbl.extract(x_tolerance=1, y_tolerance=1), tbl.bbox, tbl.rows) for tbl in found_tables]
                page_links = page.hyperlinks

            for tbl_obj, data, table_bbox, rows_obj in found:
                if not data or len(data) < 2:
                    continue

                hdr = [str(h).strip().replace('\n', ' ') for h in data[0]]
                desc_i = next((i for i, h_text in enumerate(hdr) if "description" in h_text.lower()), None)
                notes_i = next((i for i, h_text in enumerate(hdr) if any(x in h_text.lower() for x in ["note", "comment"])), None)
                start_date_i = next((i for i, h_text in enumerate(hdr) if "start date" in h_text.lower()), None)
                end_date_i = next((i for i, h_text in enumerate(hdr) if "end date" in h_text.lower()), None)
                term_i = next((i for i, h_text in enumerate(hdr) if "term" in h_text.lower() or ("month" in h_text.lower() and "amount" not in h_text.lower())), None)
                monthly_i = next((i for i, h_text in enumerate(hdr) if "monthly" in h_text.lower()), None)
                total_i = next((i for i, h_text in enumerate(hdr) if "total" in h_text.lower() and "grand" not in h_text.lower() and "month" not in h_text.lower()), None)

                if desc_i is None: desc_i = 0 # Default to first column

                desc_links = {}
                if tbl_obj != "camelot" and rows_obj and desc_i is not None:
                    for r_idx, row_obj in enumerate(rows_obj): # r_idx is 0-based index for rows_obj
                        if r_idx == 0: continue # Skip header row_obj
                        if desc_i < len(row_obj.cells) and row_obj.cells[desc_i]:
                            cell_bbox = row_obj.cells[desc_i]
                            x0_c, top_c, x1_c, bottom_c = cell_bbox
                            for link_item in page_links:
                                if all(k in link_item for k in ("x0", "x1", "top", "bottom", "uri")):
                                    if not (link_item["x1"] < x0_c or link_item["x0"] > x1_c or link_item["bottom"] < top_c or link_item["top"] > bottom_c):
                                        desc_links[r_idx] = link_item["uri"] # Use r_idx (0-based for rows_obj)
                                        break

                table_rows_data = []
                row_links_ordered = []
                table_total_row_content = None

                for ridx_data, row_content_list in enumerate(data[1:], start=1): # ridx_data is 1-based for data list
                    cells = [str(c).strip() if c is not None else "" for c in row_content_list]
                    if not any(cells): continue

                    first_cell_lower = cells[0].lower() if cells else ""
                    is_total_row_check = ("total" in first_cell_lower or "subtotal" in first_cell_lower) and any("$" in c for c in cells)
                    if is_total_row_check:
                        if table_total_row_content is None:
                            table_total_row_content = cells
                        continue

                    new_row_values = [""] * len(HEADERS)

                    # Description
                    if desc_i is not None and desc_i < len(cells):
                        raw_desc_text = cells[desc_i]
                        # rows_obj is 0-indexed, ridx_data is 1-indexed for data rows
                        # So, for rows_obj, the index is ridx_data
                        if tbl_obj != "camelot" and rows_obj and ridx_data < len(rows_obj) and \
                           desc_i < len(rows_obj[ridx_data].cells) and rows_obj[ridx_data].cells[desc_i]:
                             current_cell_bbox = rows_obj[ridx_data].cells[desc_i]
                             rich_desc_text = extract_rich_cell(current_page_idx_fitz, current_cell_bbox)
                             new_row_values[0] = rich_desc_text or raw_desc_text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")
                        else:
                            new_row_values[0] = raw_desc_text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")

                    # Notes
                    if notes_i is not None and notes_i < len(cells):
                        raw_notes_text = cells[notes_i]
                        if tbl_obj != "camelot" and rows_obj and ridx_data < len(rows_obj) and \
                           notes_i < len(rows_obj[ridx_data].cells) and rows_obj[ridx_data].cells[notes_i]:
                            current_cell_bbox = rows_obj[ridx_data].cells[notes_i]
                            rich_notes_text = extract_rich_cell(current_page_idx_fitz, current_cell_bbox)
                            new_row_values[6] = rich_notes_text or raw_notes_text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")
                        else:
                             new_row_values[6] = raw_notes_text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")

                    # Assign by header index first
                    if start_date_i is not None and start_date_i < len(cells): new_row_values[1] = cells[start_date_i]
                    if end_date_i is not None and end_date_i < len(cells):   new_row_values[2] = cells[end_date_i]
                    if term_i is not None and term_i < len(cells):       new_row_values[3] = cells[term_i]
                    if monthly_i is not None and monthly_i < len(cells):    new_row_values[4] = cells[monthly_i]
                    if total_i is not None and total_i < len(cells):      new_row_values[5] = cells[total_i]

                    # Fallback guessing for unassigned columns
                    for cell_idx, cell_val in enumerate(cells):
                        if not cell_val: continue
                        is_assigned_by_header = (desc_i == cell_idx and new_row_values[0]) or \
                                                (notes_i == cell_idx and new_row_values[6]) or \
                                                (start_date_i == cell_idx and new_row_values[1]) or \
                                                (end_date_i == cell_idx and new_row_values[2]) or \
                                                (term_i == cell_idx and new_row_values[3]) or \
                                                (monthly_i == cell_idx and new_row_values[4]) or \
                                                (total_i == cell_idx and new_row_values[5])
                        if is_assigned_by_header: continue

                        if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', cell_val): # Date
                            if not new_row_values[1]: new_row_values[1] = cell_val
                            elif not new_row_values[2]: new_row_values[2] = cell_val
                        elif re.fullmatch(r"\d{1,3}", cell_val.strip()) or ("month" in cell_val.lower() and "amount" not in cell_val.lower()) or "mo" in cell_val.lower() : # Term
                            if not new_row_values[3]: new_row_values[3] = cell_val.replace("months","").replace("month","").strip()
                        elif "$" in cell_val: # Currency
                            if not new_row_values[4] and ("monthly" in hdr[cell_idx].lower() if cell_idx < len(hdr) else False): new_row_values[4] = cell_val
                            elif not new_row_values[5] and ("total" in hdr[cell_idx].lower() if cell_idx < len(hdr) else False): new_row_values[5] = cell_val
                            elif not new_row_values[4]: new_row_values[4] = cell_val # Fallback for monthly
                            elif not new_row_values[5]: new_row_values[5] = cell_val # Fallback for item total
                        # Avoid assigning to notes here if notes_i was None initially and desc_i got it by default
                        elif notes_i is None and desc_i == cell_idx : continue # if notes_i was never found, don't overwrite desc
                        elif not new_row_values[6] and len(cell_val) > 3: # Fallback notes (if not already description/other main field)
                             if not any (x in cell_val.lower() for x in ["date", "term", "$", "month"]):
                                new_row_values[6] = cell_val.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")


                    if any(new_row_values[i].strip().replace('\n',' ') == HEADERS[i] for i in range(len(HEADERS)) if new_row_values[i]):
                        continue
                    table_rows_data.append(new_row_values)
                    # Link corresponds to ridx_data (1-based index for data rows, also for rows_obj after header)
                    row_links_ordered.append(desc_links.get(ridx_data))


                if table_total_row_content is None:
                    table_total_row_content = find_total(current_page_idx_fitz)

                if table_rows_data:
                    tables_info.append((HEADERS, table_rows_data, row_links_ordered, table_total_row_content))

        # Find Grand Total from the end of the document
        for page_idx_fitz, blk_text in reversed(texts): # page_idx_fitz is 0-indexed
            m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', blk_text, re.I | re.S)
            if m:
                grand_total = m.group(1).replace(" ", "")
                break
except pdfplumber.PDFSyntaxError as e:
    st.error(f"PDFPlumber Error: Failed to process PDF. It might be corrupted or password-protected. Error: {e}")
    st.stop()
except Exception as e:
    st.error(f"An unexpected error occurred during PDF processing: {e}")
    st.exception(e) # Shows full traceback in Streamlit for debugging
    st.stop()


# --- PDF Generation ---
if not tables_info:
    st.warning("No tables suitable for reformatting were found in the uploaded PDF.")
    st.stop()

pdf_buf=io.BytesIO()
doc=SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch, 11*inch)),
                     leftMargin=0.5*inch, rightMargin=0.5*inch,
                     topMargin=0.5*inch, bottomMargin=0.5*inch)

# Styles
ts = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
hs = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black, spaceAfter=6)
bs = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=12)
bs_right = ParagraphStyle("BodyRight", parent=bs, alignment=TA_RIGHT)
bs_center = ParagraphStyle("BodyCenter", parent=bs, alignment=TA_CENTER)

story = [Spacer(1, 12), Paragraph(proposal_title, ts), Spacer(1, 24)]
table_width = doc.width

col_widths = [
    table_width * 0.30, # Description
    table_width * 0.08, # Start Date
    table_width * 0.08, # End Date
    table_width * 0.06, # Term (Months)
    table_width * 0.10, # Monthly Amount
    table_width * 0.10, # Item Total
    table_width * 0.28  # Notes
]

for current_headers, current_rows, current_links, current_total_info in tables_info:
    n_cols = len(current_headers)
    # Ensure col_widths matches header length if dynamic headers were ever introduced
    current_col_widths = col_widths[:n_cols] if len(col_widths) >= n_cols else [table_width / n_cols] * n_cols


    header_row_styled = [Paragraph(h_text, hs) for h_text in current_headers]
    table_data_styled = [header_row_styled]

    for i, row_data_list in enumerate(current_rows):
        styled_row_elements = []
        for j, cell_text_val in enumerate(row_data_list):
            cell_style_to_use = bs
            if j == 1 or j == 2 or j == 3: # Start Date, End Date, Term
                cell_style_to_use = bs_center
            elif j == 4 or j == 5: # Monthly Amount, Item Total
                cell_style_to_use = bs_right

            # Add hyperlink if available for the description column (index 0)
            if j == 0 and i < len(current_links) and current_links[i]:
                 linked_text_val = cell_text_val + f" <link href='{current_links[i]}' color='blue'>[link]</link>"
                 styled_row_elements.append(Paragraph(linked_text_val, cell_style_to_use))
            else:
                 styled_row_elements.append(Paragraph(cell_text_val, cell_style_to_use))
        table_data_styled.append(styled_row_elements)

    if current_total_info:
        total_label_text = "Total"
        total_value_text = ""
        if isinstance(current_total_info, list):
             total_label_text = next((c for c in current_total_info if c and '$' not in c and c.strip().lower() not in ["total", "subtotal"]), "Total")
             # If label is still "Total", look for a more descriptive one like "Subtotal"
             if total_label_text == "Total":
                 total_label_text = next((c for c in current_total_info if c and ('total' in c.lower() or 'subtotal' in c.lower())), "Total")

             total_value_text = next((c for c in reversed(current_total_info) if "$" in c), "")
        elif isinstance(current_total_info, str):
            m = re.match(r'(.*?)\s*(\$\s*[\d,]+\.\d{2})', current_total_info)
            if m:
                total_label_text, total_value_text = m.group(1).strip(), m.group(2).strip()
            else:
                total_label_text = re.sub(r'\$\s*[\d,.]+', '', current_total_info).strip() or "Total"
                val_match = re.search(r'(\$\s*[\d,.]+\.\d{2})', current_total_info)
                if val_match: total_value_text = val_match.group(1)


        total_row_styled = [Paragraph(total_label_text, bs)] + \
                           [Paragraph("", bs)] * (n_cols - 2) + \
                           [Paragraph(total_value_text, bs_right)]
        table_data_styled.append(total_row_styled)

    tbl_reportlab = LongTable(table_data_styled, colWidths=current_col_widths, repeatRows=1)

    style_cmds_list = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ("ALIGN", (1, 1), (1, -1), "CENTER"), # Start Date
        ("ALIGN", (2, 1), (2, -1), "CENTER"), # End Date
        ("ALIGN", (3, 1), (3, -1), "CENTER"), # Term
        ("ALIGN", (4, 1), (4, -1), "RIGHT"),  # Monthly Amount
        ("ALIGN", (5, 1), (5, -1), "RIGHT"),  # Item Total
    ]

    if current_total_info:
        style_cmds_list.extend([
            ("SPAN", (0, -1), (-2, -1)),
            ("ALIGN", (0, -1), (-2, -1), "RIGHT"), # Total label align
            ("ALIGN", (-1, -1), (-1, -1), "RIGHT"), # Total value align
            ("VALIGN", (0, -1), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAEAEA")),
        ])

    tbl_reportlab.setStyle(TableStyle(style_cmds_list))
    story.extend([tbl_reportlab, Spacer(1, 24)])

if grand_total:
    grand_total_row_styled = [Paragraph("Grand Total", bs)] + \
                             [Paragraph("", bs)] * (len(HEADERS) - 2) + \
                             [Paragraph(grand_total, bs_right)]

    gt_table = LongTable([grand_total_row_styled], colWidths=col_widths) # Use main col_widths
    gt_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D0D0D0")), # Darker background for GT
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black), # Stronger grid for GT
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-2, 0)),
        ("ALIGN", (0, 0), (-2, 0), "RIGHT"), # Grand Total label align
        ("ALIGN", (-1, 0), (-1, 0), "RIGHT"), # Grand Total value align
        ("TEXTCOLOR", (0,0), (-1,-1), colors.black),
        ("FONTNAME", (0,0), (-1,-1), DEFAULT_SANS_FONT), # Ensure font
        ("FONTSIZE", (0,0), (-1,-1), 10), # Slightly larger font for GT
    ]))
    story.append(gt_table)

try:
    doc.build(story)
    pdf_buf.seek(0)
    st.download_button(
        "📥 Download Transformed PDF",
        data=pdf_buf,
        file_name="transformed_proposal.pdf",
        mime="application/pdf",
        use_container_width=True
        )
except Exception as e:
    st.error(f"Error building final PDF with ReportLab: {e}")
    st.exception(e)
