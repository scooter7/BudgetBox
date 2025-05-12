# -*- coding: utf-8 -*-
import io
import re
import camelot
import pdfplumber
import fitz
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer

# --- Font Registration ---
# Register regular fonts
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))

# !!! NEW: Register bold fonts (assuming files exist in 'fonts/' directory) !!!
# Make sure you have Barlow-Bold.ttf in your fonts folder
try:
    pdfmetrics.registerFont(TTFont("Barlow-Bold", "fonts/Barlow-Bold.ttf"))
    # Link Barlow regular and bold for <b> tag support
    pdfmetrics.registerFontFamily('Barlow',
                                  normal='Barlow',
                                  bold='Barlow-Bold')
                                  # If you had italic/bold-italic, you'd add them here:
                                  # italic='Barlow-Italic',
                                  # boldItalic='Barlow-BoldItalic')
    ST_BARLOW_BOLD_LOADED = True
except Exception as e:
    # st.warning("Barlow-Bold.ttf not found or failed to register. Bold for Barlow may not work correctly.")
    # print("Warning: Barlow-Bold.ttf not found or failed to register. Bold for Barlow may not work correctly.")
    ST_BARLOW_BOLD_LOADED = False


# For DMSerif, if you have a DMSerifDisplay-Bold.ttf, you would do similarly:
# try:
#     pdfmetrics.registerFont(TTFont("DMSerif-Bold", "fonts/DMSerifDisplay-Bold.ttf"))
#     pdfmetrics.registerFontFamily('DMSerif',
#                                   normal='DMSerif',
#                                   bold='DMSerif-Bold')
# except Exception as e:
# st.warning("DMSerifDisplay-Bold.ttf not found. Bold for DMSerif may not work correctly.")
# print("Warning: DMSerifDisplay-Bold.ttf not found. Bold for DMSerif may not work correctly.")


DEFAULT_SERIF_FONT = "DMSerif"
DEFAULT_SANS_FONT = "Barlow"

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
        d = page.get_text("dict", clip=bbox) # Use clip for better accuracy within the bbox
        spans = []
        x0_bbox, y0_bbox, x1_bbox, y1_bbox = bbox

        for block in d.get("blocks", []):
            if block.get("type") != 0:  # Text blocks only
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sx0, sy0, sx1, sy1 = span["bbox"]
                    # Check for overlap between span and bbox
                    if sx0 < x1_bbox and sx1 > x0_bbox and sy0 < y1_bbox and sy1 > y0_bbox:
                        spans.append(span)

        if not spans:
            return ""

        # Group spans by their approximate baseline (y-coordinate of the origin)
        # and then sort by horizontal position.
        lines_dict = {}
        for s in spans:
            key = round(s["origin"][1], 1) # Using y-origin for line grouping
            lines_dict.setdefault(key, []).append(s)

        span_text_lines = []
        for key in sorted(lines_dict.keys()):
            row_spans = sorted(lines_dict[key], key=lambda s_item: s_item["origin"][0]) # Sort by x-origin
            line_pieces = []
            last_x1_of_span = x0_bbox # Initialize for checking horizontal spacing

            for span_idx, span_item in enumerate(row_spans):
                # Add a space if spans are not contiguous (simple heuristic)
                if span_idx > 0:
                    prev_span_x1 = row_spans[span_idx-1]["bbox"][2]
                    current_span_x0 = span_item["bbox"][0]
                    # If there's a gap (e.g., more than 1-2 points), add a space
                    if current_span_x0 > prev_span_x1 + 1.5 : # Tolerance for space detection
                        line_pieces.append(" ")

                t = span_item["text"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                is_bold = span_item["flags"] & 2  # Standard flag for bold

                if is_bold:
                    # Check if the font being used (Barlow or DMSerif) has bold loaded
                    # This is a runtime check, assumes DEFAULT_SANS_FONT or DEFAULT_SERIF_FONT is active
                    # For Paragraphs, ReportLab handles this via font family registration.
                    # This explicit check here is more for understanding.
                    font_name_from_span = span_item.get("font", DEFAULT_SANS_FONT) # Get actual font from span if possible
                    can_render_bold = False
                    if "barlow" in font_name_from_span.lower() and ST_BARLOW_BOLD_LOADED:
                        can_render_bold = True
                    # Add similar check for DMSerif if its bold variant is loaded

                    if can_render_bold or "barlow" not in font_name_from_span.lower(): # Attempt bold for others too
                        line_pieces.append(f"<b>{t}</b>")
                    else:
                        line_pieces.append(t) # Fallback to non-bold if bold font specifically missing
                else:
                    line_pieces.append(t)
                last_x1_of_span = span_item["bbox"][2]

            span_text_lines.append("".join(line_pieces))
        return "<br/>".join(span_text_lines)

    except Exception as e:
        # st.warning(f"Error in extract_rich_cell for bbox {bbox} on page {page_number}: {e}")
        return ""
# ... (rest of your existing script, no changes needed to the table parsing or PDF generation logic related to Paragraphs, as they will use the <b> tags and rely on the font family registration)
# Make sure the rest of your script (from HEADERS = ... onwards) follows here.
# I will paste the rest of the script for completeness.
# Previous code ended here in the thought process, so I'm continuing from here.

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
            for l_line in text_content.splitlines():
                line_key = (actual_page_num_for_key, l_line)
                if re.search(r'\b(?<!grand\s)total\b.*?\$\s*[\d,.]+',l_line, re.I) and line_key not in used_total_lines:
                    used_total_lines.add(line_key)
                    return l_line.strip()
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
                            cell_bbox_val = row_obj.cells[desc_i]
                            x0_c, top_c, x1_c, bottom_c = cell_bbox_val
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
                            if not new_row_values[4] and (cell_idx < len(hdr) and "monthly" in hdr[cell_idx].lower()): new_row_values[4] = cell_val
                            elif not new_row_values[5] and (cell_idx < len(hdr) and "total" in hdr[cell_idx].lower()): new_row_values[5] = cell_val
                            elif not new_row_values[4]: new_row_values[4] = cell_val # Fallback for monthly
                            elif not new_row_values[5]: new_row_values[5] = cell_val # Fallback for item total
                        elif notes_i is None and desc_i == cell_idx : continue
                        elif not new_row_values[6] and len(cell_val) > 3: # Fallback notes
                             if not any (x_char in cell_val.lower() for x_char in ["date", "term", "$", "month"]):
                                new_row_values[6] = cell_val.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")


                    if any(new_row_values[i_val].strip().replace('\n',' ') == HEADERS[i_val] for i_val in range(len(HEADERS)) if new_row_values[i_val]):
                        continue
                    table_rows_data.append(new_row_values)
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
except pdfplumber.PDFSyntaxError as e_pdfsyn:
    st.error(f"PDFPlumber Error: Failed to process PDF. It might be corrupted or password-protected. Error: {e_pdfsyn}")
    st.stop()
except Exception as e_proc:
    st.error(f"An unexpected error occurred during PDF processing: {e_proc}")
    st.exception(e_proc) # Shows full traceback in Streamlit for debugging
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
    current_col_widths = col_widths[:n_cols] if len(col_widths) >= n_cols else [table_width / n_cols] * n_cols

    header_row_styled = [Paragraph(h_text, hs) for h_text in current_headers]
    table_data_styled = [header_row_styled]

    for i, row_data_list in enumerate(current_rows):
        styled_row_elements = []
        for j, cell_text_val in enumerate(row_data_list):
            cell_style_to_use = bs
            # Apply alignment styles based on column index
            # Body style (bs) default is TA_LEFT (for Description, Notes)
            if j == 1 or j == 2 or j == 3: # Start Date, End Date, Term (Months)
                cell_style_to_use = bs_center
            elif j == 4 or j == 5: # Monthly Amount, Item Total
                cell_style_to_use = bs_right
            # Note: Column 0 (Description) and 6 (Notes) use default 'bs' style (TA_LEFT)

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
             if total_label_text == "Total": # More specific search if first pass is too generic
                 total_label_text = next((c for c in current_total_info if c and ('total' in c.lower() or 'subtotal' in c.lower())), "Total").strip()

             total_value_text = next((c for c in reversed(current_total_info) if "$" in c), "")
        elif isinstance(current_total_info, str):
            m_total = re.match(r'(.*?)\s*(\$\s*[\d,]+\.\d{2})', current_total_info)
            if m_total:
                total_label_text, total_value_text = m_total.group(1).strip(), m_total.group(2).strip()
            else:
                total_label_text = re.sub(r'\$\s*[\d,.]+', '', current_total_info).strip() or "Total"
                val_match = re.search(r'(\$\s*[\d,.]+\.\d{2})', current_total_info)
                if val_match: total_value_text = val_match.group(1)
        if not total_label_text: total_label_text = "Total" # Ensure label is not empty


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
        # Column specific ALIGNMENTS for data rows (row 1 to -1, or 1 to -2 if total row exists)
        # These apply to data rows. Header alignment is TA_CENTER by default via 'hs' style.
        ("ALIGN", (1, 1), (1, -1 if not current_total_info else -2), "CENTER"), # Start Date
        ("ALIGN", (2, 1), (2, -1 if not current_total_info else -2), "CENTER"), # End Date
        ("ALIGN", (3, 1), (3, -1 if not current_total_info else -2), "CENTER"), # Term
        ("ALIGN", (4, 1), (4, -1 if not current_total_info else -2), "RIGHT"),  # Monthly Amount
        ("ALIGN", (5, 1), (5, -1 if not current_total_info else -2), "RIGHT"),  # Item Total
    ]

    if current_total_info:
        style_cmds_list.extend([
            ("SPAN", (0, -1), (-2, -1)),
            ("ALIGN", (0, -1), (-2, -1), "RIGHT"), # Total label align (within the span)
            ("ALIGN", (-1, -1), (-1, -1), "RIGHT"), # Total value align (in the last cell)
            ("VALIGN", (0, -1), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAEAEA")),
        ])

    tbl_reportlab.setStyle(TableStyle(style_cmds_list))
    story.extend([tbl_reportlab, Spacer(1, 24)])

if grand_total:
    grand_total_row_styled = [Paragraph("Grand Total", bs)] + \
                             [Paragraph("", bs)] * (len(HEADERS) - 2) + \
                             [Paragraph(grand_total, bs_right)]

    gt_table = LongTable([grand_total_row_styled], colWidths=col_widths)
    gt_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D0D0D0")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("SPAN", (0, 0), (-2, 0)),
        ("ALIGN", (0, 0), (-2, 0), "RIGHT"),
        ("ALIGN", (-1, 0), (-1, 0), "RIGHT"),
        ("TEXTCOLOR", (0,0), (-1,-1), colors.black),
        ("FONTNAME", (0,0), (-1,-1), DEFAULT_SANS_FONT), # Use consistent font
        ("FONTSIZE", (0,0), (-1,-1), 10),
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
except Exception as e_build:
    st.error(f"Error building final PDF with ReportLab: {e_build}")
    st.exception(e_build)
