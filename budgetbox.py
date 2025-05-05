# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
import docx # Make sure docx is imported
from docx.shared import Inches, Pt, RGBColor # Import RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE # Import WD_STYLE_TYPE
# Import WD_CELL_VERTICAL_ALIGNMENT explicitly if needed elsewhere, or use docx.enum.table.*
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls # Added nsdecls just in case, qn is essential
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
import html # Import the html module for escaping

# Register fonts
# Ensure these font files exist in a 'fonts' subdirectory or provide correct paths
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    # Define default fonts in case registration fails or for Word
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except Exception as e:
    st.warning(f"Could not load custom fonts: {e}. Using default fonts.")
    # Fallback fonts (ensure they are available on the system where Streamlit runs)
    DEFAULT_SERIF_FONT = "Times New Roman"
    DEFAULT_SANS_FONT = "Arial"


# Streamlit setup
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download both PDF and Word outputs.")

uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# Split first line = Strategy, rest = Description
def split_cell_text(raw: str):
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    if not lines:
        return "", ""
    # Basic whitespace normalization within the joined description
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return lines[0], description

# --- REVISED Word hyperlink helper (Unchanged from previous working version) ---
def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    """
    Add a hyperlink to a paragraph. NOW USED ONLY FOR "- link" text.

    :param paragraph: The paragraph to add the hyperlink to.
    :param url: The URL for the hyperlink.
    :param text: The text to display for the hyperlink.
    :param font_name: Optional font name for the hyperlink text.
    :param font_size: Optional font size (in Pt) for the hyperlink text.
    :param bold: Optional boolean to make the hyperlink text bold.
    :return: The run object containing the hyperlink.
    """
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True)
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1); style.font.underline = True
        style.priority = 9; style.unhide_when_used = True
    style_element = OxmlElement('w:rStyle')
    style_element.set(qn('w:val'), 'Hyperlink')
    rPr.append(style_element)
    # Apply specific font/size if provided for the "- link" text
    if font_name:
        run_font = OxmlElement('w:rFonts')
        run_font.set(qn('w:ascii'), font_name); run_font.set(qn('w:hAnsi'), font_name)
        rPr.append(run_font)
    if font_size:
        size = OxmlElement('w:sz'); size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs'); size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size); rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b'); rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t'); t.set(qn('xml:space'), 'preserve'); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)
# --- End of Word hyperlink helper ---


# === START: REVISED PDF TABLE EXTRACTION AND PROCESSING LOGIC ===
tables_info = []
grand_total = None
proposal_title = "Untitled Proposal" # Default title

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]

        first_page_lines = page_texts[0].splitlines() if page_texts else []
        potential_title = next((line.strip() for line in first_page_lines if "proposal" in line.lower() and len(line.strip()) > 5), None)
        if potential_title: proposal_title = potential_title
        elif len(first_page_lines) > 0: proposal_title = first_page_lines[0].strip()

        used_totals = set()
        def find_total(pi):
            # (find_total function unchanged)
            if pi >= len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            links = page.hyperlinks

            page_tables = page.find_tables()
            if not page_tables: continue

            for tbl_idx, tbl in enumerate(page_tables): # Added index for logging
                data = tbl.extract(x_tolerance=1, y_tolerance=1)
                # Need at least a header and one data row conceptually
                if not data or len(data) < 2:
                    # st.info(f"Skipping table {tbl_idx+1} on page {pi+1}: Not enough data rows ({len(data)} rows found).")
                    continue

                # --- Find Original Header and Description Column ---
                # Ensure header cells are strings and stripped
                original_hdr = [(str(h).strip() if h is not None else "") for h in data[0]]
                if not any(original_hdr):
                    # st.info(f"Skipping table {tbl_idx+1} on page {pi+1}: Header is empty or invalid.")
                    continue # Skip if header is effectively empty

                # Find the index of the description column in the *original* header
                original_desc_idx = -1
                # First, try finding "description" explicitly
                for i, h in enumerate(original_hdr):
                    if h and "description" in h.lower():
                        original_desc_idx = i
                        break
                # Fallback: Find the first reasonably wide column if "description" not found
                if original_desc_idx == -1:
                     for i, h in enumerate(original_hdr):
                        # Adjust length check if needed based on typical description headers
                        if h and len(h) > 8:
                           original_desc_idx = i
                           break

                # If no suitable description column found, skip this table
                if original_desc_idx == -1:
                    # st.warning(f"Skipping table {tbl_idx+1} on page {pi+1}: Could not reliably identify a description column. Header: {original_hdr}")
                    continue

                # --- Link Finding (Associates link URI with ROW INDEX) ---
                # Use the original_desc_idx to check for links in the correct column's bounding box
                desc_links_uri = {} # Stores {row_index_in_pdf_table: uri}
                if hasattr(tbl, 'rows'):
                    for r, row_obj in enumerate(tbl.rows):
                        # r=0 is header, r=1 is first data row in pdfplumber table object
                        if r == 0: continue
                        if hasattr(row_obj, 'cells') and original_desc_idx < len(row_obj.cells):
                            cell_bbox = row_obj.cells[original_desc_idx]
                            if not cell_bbox: continue
                            # pdfplumber cell coordinates are (x0, top, x1, bottom)
                            cell_x0, cell_top, cell_x1, cell_bottom = cell_bbox

                            for link in links:
                                if not all(k in link for k in ['x0', 'x1', 'top', 'bottom', 'uri']): continue
                                link_x0, link_top, link_x1, link_bottom = link['x0'], link['top'], link['x1'], link['bottom']
                                # Simple overlap check: link bbox intersects cell bbox
                                x_overlap = (link_x0 < cell_x1) and (link_x1 > cell_x0)
                                y_overlap = (link_top < cell_bottom) and (link_bottom > cell_top)

                                if x_overlap and y_overlap:
                                    # Found an overlapping link, store its URI for this row index 'r'
                                    # Note: r is 1-based relative to tbl.rows (where 0 is header)
                                    desc_links_uri[r] = link.get("uri")
                                    break # Stop checking other links for *this* cell

                # --- CORRECTED Header and Row Processing ---
                # Build the NEW header by replacing the original description header
                # with "Strategy" and "Description"
                new_hdr = []
                for i, h in enumerate(original_hdr):
                    if i == original_desc_idx:
                        new_hdr.extend(["Strategy", "Description"])
                    else:
                        # Only include non-empty original headers
                        if h:
                            new_hdr.append(h)

                # Ensure we don't end up with an empty header if original was bad
                if not new_hdr:
                    # st.warning(f"Skipping table {tbl_idx+1} on page {pi+1}: Resulting header 'new_hdr' is empty after processing.")
                    continue

                rows_data = [] # Stores processed rows matching new_hdr structure
                row_links_uri_list = [] # Stores URI or None for each row in rows_data
                table_total_info = None # Stores total row info (list or string)

                # Iterate through data rows from the PDF table (skip header data[0])
                # ridx_pdf is the 1-based index within the PDF table (incl. header)
                for ridx_pdf, row_content in enumerate(data[1:], start=1):
                    # Get raw string values for the current row, handling potential None values
                    row_str_list = [(str(cell).strip() if cell is not None else "") for cell in row_content]
                    if all(not cell_val for cell_val in row_str_list): continue # Skip entirely empty rows

                    # Check for total row (using original row structure for detection)
                    first_cell_lower = row_str_list[0].lower() if row_str_list else ""
                    # More robust check: look for total/subtotal AND a currency symbol/amount anywhere
                    is_total_row = (("total" in first_cell_lower or "subtotal" in first_cell_lower) and \
                                   any(re.search(r'\$|€|£|¥', str(cell_val)) for cell_val in row_str_list)) or \
                                   (len(row_str_list)>1 and "total" in row_str_list[-1].lower()) # Also check last cell

                    if is_total_row:
                        if table_total_info is None: table_total_info = row_str_list
                        continue # Skip adding total rows to rows_data

                    # --- Process NORMAL data row ---
                    # Extract the text from the original description column
                    desc_text_from_pdf = ""
                    if original_desc_idx < len(row_str_list):
                         desc_text_from_pdf = row_str_list[original_desc_idx]

                    strat, desc = split_cell_text(desc_text_from_pdf)

                    # Build the NEW row structure based on original headers (before filtering)
                    new_row_content = []
                    col_idx_map = 0 # Tracks index in the original non-empty headers
                    for i, original_cell_val in enumerate(row_str_list):
                        # Make sure we don't go out of bounds for original_hdr
                        if i >= len(original_hdr): break

                        original_header_val = original_hdr[i]

                        if i == original_desc_idx:
                            # Insert the split Strategy and Description values
                            new_row_content.extend([strat, desc])
                            if original_header_val: # Count if it was a non-empty header
                                col_idx_map += 1
                        else:
                            # Keep the original cell value *if* its header wasn't empty
                            if original_header_val:
                                new_row_content.append(original_cell_val)
                                col_idx_map += 1

                    # Ensure the new row has the same number of columns as the new header
                    expected_cols = len(new_hdr)
                    current_cols = len(new_row_content)
                    if current_cols < expected_cols:
                       new_row_content.extend([""] * (expected_cols - current_cols))
                    elif current_cols > expected_cols:
                       new_row_content = new_row_content[:expected_cols] # Truncate if too long

                    rows_data.append(new_row_content) # Add the processed row data

                    # Get the link URI for this specific row using the map created earlier
                    # Use ridx_pdf because desc_links_uri keys are based on PDF table row index
                    row_links_uri_list.append(desc_links_uri.get(ridx_pdf)) # Append URI or None

                # --- Fallback for table total if not found in rows ---
                if table_total_info is None:
                    table_total_info = find_total(pi)

                # --- Store the processed table info ---
                if rows_data: # Only store table if it has valid data rows
                    tables_info.append((new_hdr, rows_data, row_links_uri_list, table_total_info))
                # else:
                #     st.info(f"Table {tbl_idx+1} on page {pi+1} processed but resulted in no data rows.")


        # Find Grand total robustly (remains unchanged)
        for tx in reversed(page_texts):
            # Look for "Grand Total" followed optionally by whitespace/newlines and then the amount
            m = re.search(r'Grand\s+Total.*?(?<!Subtotal\s)(?<!Sub Total\s)(\$\s*[\d,]+\.\d{2})', tx, re.I | re.S)
            if m:
                grand_total_candidate = m.group(1).replace(" ", "")
                # Avoid capturing a subtotal if "Grand Total" happens to be near it
                if "subtotal" not in m.group(0).lower():
                    grand_total = grand_total_candidate
                    break # Found the most likely Grand Total

except Exception as e:
    st.error(f"Error processing PDF: {e}")
    import traceback
    st.error(traceback.format_exc())
    st.stop()
# === END: REVISED PDF TABLE EXTRACTION AND PROCESSING LOGIC ===


# ─── Build PDF ────────────────────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf, pagesize=landscape((17*inch, 11*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch
)
# Styles (unchanged)
title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11)
link_style   = ParagraphStyle("LinkStyle", parent=body_style, textColor=colors.blue) # May not be needed now
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, textColor=colors.black, spaceBefore=6)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, textColor=colors.black, spaceBefore=6)

elements = []
logo = None
# Add Logo (unchanged)
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    response = requests.get(logo_url, timeout=10); response.raise_for_status()
    logo = response.content
    img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width
    img_width = min(5*inch, doc.width); img_height = img_width * ratio
    elements.append(RLImage(io.BytesIO(logo), width=img_width, height=img_height))
except Exception as e: st.warning(f"Could not load or process logo: {e}")

elements += [Spacer(1, 12), Paragraph(html.escape(proposal_title), title_style), Spacer(1, 24)]
total_page_width = doc.width

# --- Iterate through processed table data ---
for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info): # Use correct variable names
    num_cols = len(hdr)
    if num_cols == 0:
        # st.warning(f"Skipping PDF generation for table index {table_index}: Header is empty.")
        continue

    # --- Column widths calculation ---
    col_widths = []
    try:
        # Find the index of 'Description' in the *new* header
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_col_width = total_page_width * 0.45 # Assign fixed proportion for Description
        other_cols_count = num_cols - 1 # Description is one column
        if other_cols_count > 0:
            other_total_width = total_page_width - desc_col_width
            # Distribute remaining width among other columns
            # Consider giving 'Strategy' slightly more width if it exists next to Desc
            strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1

            if strategy_idx != -1:
                 # Example: Give Strategy 15% and distribute rest among others
                 strat_width = total_page_width * 0.15
                 remaining_width = other_total_width - strat_width
                 remaining_cols = other_cols_count - 1
                 other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0
                 col_widths = []
                 for i in range(num_cols):
                     if i == desc_actual_idx_in_hdr: col_widths.append(desc_col_width)
                     elif i == strategy_idx: col_widths.append(strat_width)
                     else: col_widths.append(max(0.1*inch, other_indiv_width)) # Ensure min width
            else:
                 # Just distribute evenly if no Strategy column found next to Description
                 other_col_width = other_total_width / other_cols_count
                 col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]

        elif num_cols == 1 and desc_actual_idx_in_hdr == 0: # Only description column
             col_widths = [total_page_width]
        else: # Only one non-description column? Fallback to equal width
             col_widths = [total_page_width / num_cols] * num_cols

    except ValueError:
        # 'Description' not found in header, fallback to equal widths
        # st.warning(f"PDF Gen: 'Description' column not found in header for table {table_index}. Using equal widths. Header: {hdr}")
        desc_actual_idx_in_hdr = -1 # Indicate description column not found for link logic
        if num_cols > 0:
            col_widths = [total_page_width / num_cols] * num_cols
        else:
             continue # Cannot create table with 0 columns


    # Wrap header text
    wrapped_header = [Paragraph(html.escape(str(h)), header_style) for h in hdr]
    wrapped_data = [wrapped_header]

    # --- START: Process Data Rows for PDF Output (Simpler "- link" approach) ---
    for ridx, row in enumerate(rows_data):
        line = []
        # Ensure row has expected number of cells, pad if necessary
        current_cells = len(row)
        if current_cells < num_cols:
            row = list(row) + [""] * (num_cols - current_cells)
        elif current_cells > num_cols:
            row = row[:num_cols]

        for cidx, cell_content in enumerate(row):
            cell_str = str(cell_content)
            escaped_cell_text = html.escape(cell_str)

            # Check if this is the description column (use generated header index)
            # AND if a link URI exists for this row
            link_applied = False
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                if link_uri: # Ensure URI is not None or empty
                    # Append "- link" and make only that part the link
                    paragraph_text = f"{escaped_cell_text} <link href='{html.escape(link_uri)}' color='blue'>- link</link>"
                    p = Paragraph(paragraph_text, body_style) # Apply base style
                    link_applied = True

            if not link_applied:
                # No link for this cell, or not the description column, or link was empty
                p = Paragraph(escaped_cell_text, body_style)

            line.append(p)
        wrapped_data.append(line)
    # --- END: Process Data Rows for PDF Output ---

    # Add Table Total Row (logic seems okay, ensure indices are safe)
    has_total_row = False
    if table_total_info:
        label = "Total"; value = ""
        if isinstance(table_total_info, list):
             # Safely access elements, provide defaults
             label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
             value = next((val.strip() for val in reversed(table_total_info) if val and '$' in str(val)), "")
             # Fallback to last element if no $ found and list has > 1 element
             if not value and len(table_total_info) > 1: value = table_total_info[-1].strip() if table_total_info[-1] else ""

        elif isinstance(table_total_info, str):
             # Try to parse Label and Value from the string
             total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
             if total_match:
                 label_parsed, value = total_match.groups()
                 label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total"
                 value = value.strip() if value else ""
             else:
                 # If direct match fails, try finding just the amount at the end
                 amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                 if amount_match:
                     value = amount_match.group(1).strip() if amount_match.group(1) else ""
                     potential_label = table_total_info[:amount_match.start()].strip()
                     label = potential_label if potential_label else "Total"
                 else: # Cannot parse reliably, treat whole string as value maybe? Or just label?
                     value = table_total_info # Assign whole string as value as fallback
                     label = "Total" # Default label

        if num_cols > 0:
            # Create the total row elements, ensuring indices are valid
            total_row_elements = [Paragraph(html.escape(label), bl_style)]
            if num_cols > 2:
                 total_row_elements.extend([Paragraph("", body_style)] * (num_cols - 2))
            if num_cols > 1:
                 total_row_elements.append(Paragraph(html.escape(value), br_style))
            elif num_cols == 1:
                 # If only one column, value might have been put in label, adjust if needed
                 # Or perhaps display value in the first column if label was generic
                 if label == "Total": total_row_elements[0] = Paragraph(html.escape(value), bl_style)
                 # else keep the specific label. This case is ambiguous.

            # Pad if somehow elements are still fewer than num_cols (shouldn't happen)
            total_row_elements += [Paragraph("", body_style)] * (num_cols - len(total_row_elements))

            wrapped_data.append(total_row_elements)
            has_total_row = True

    # Create and style PDF table (unchanged styling logic)
    if wrapped_data and col_widths and len(wrapped_data) > 1: # Ensure there's data beyond header
        tbl = LongTable(wrapped_data, colWidths=col_widths, repeatRows=1)
        style_commands = [ # Base styles
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")), # Header background
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey), # Grid for all cells
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"), # Header vertical align middle
            ("VALIGN", (0, 1), (-1, -1), "TOP"), # Body vertical align top
        ]
        # Total row styling (only if total row was added)
        if has_total_row:
             # (-1, -1) refers to the last row
             if num_cols > 1:
                 # Span label across columns except the last, align label left, value right
                 style_commands.extend([
                     ('SPAN', (0, -1), (-2, -1)),
                     ('ALIGN', (0, -1), (-2, -1), 'LEFT'),
                     ('ALIGN', (-1, -1), (-1, -1), 'RIGHT'),
                     ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
                 ])
             elif num_cols == 1:
                  # Only one column, align left, valign middle
                  style_commands.extend([
                      ('ALIGN', (0, -1), (0, -1), 'LEFT'),
                      ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
                 ])
        tbl.setStyle(TableStyle(style_commands))
        elements += [tbl, Spacer(1, 24)] # Add table and spacer

# Add Grand Total row (unchanged logic, just ensure indices safe)
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1] # Get header from the last processed table
    num_cols = len(last_hdr)
    if num_cols > 0:
        # Recalculate widths based on the last table's structure for consistency
        gt_col_widths = []
        try:
            desc_actual_idx_in_hdr = last_hdr.index("Description")
            desc_col_width = total_page_width * 0.45
            other_cols_count = num_cols - 1
            if other_cols_count > 0:
                 other_total_width = total_page_width - desc_col_width
                 strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and last_hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1
                 if strategy_idx != -1:
                     strat_width = total_page_width * 0.15
                     remaining_width = other_total_width - strat_width
                     remaining_cols = other_cols_count - 1
                     other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0
                     gt_col_widths = []
                     for i in range(num_cols):
                         if i == desc_actual_idx_in_hdr: gt_col_widths.append(desc_col_width)
                         elif i == strategy_idx: gt_col_widths.append(strat_width)
                         else: gt_col_widths.append(max(0.1*inch, other_indiv_width))
                 else:
                     other_col_width = other_total_width / other_cols_count
                     gt_col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
            elif num_cols == 1:
                 gt_col_widths = [total_page_width]
            else: # Should not happen if num_cols > 0
                 gt_col_widths = [total_page_width / num_cols] * num_cols

        except ValueError: # Fallback if 'Description' not in last header
            gt_col_widths = [total_page_width / num_cols] * num_cols if num_cols > 0 else []

        if gt_col_widths:
            # Create Grand Total row data
            gt_row_data = [ Paragraph("Grand Total", bl_style) ]
            if num_cols > 2:
                gt_row_data.extend([ Paragraph("", body_style) for _ in range(num_cols - 2) ])
            if num_cols > 1:
                gt_row_data.append(Paragraph(html.escape(grand_total), br_style))
            elif num_cols == 1: # Only one column, put GT value in the first cell
                 gt_row_data = [Paragraph(f"Grand Total: {html.escape(grand_total)}", bl_style)]

            # Pad if necessary (shouldn't be needed with above logic)
            gt_row_data += [Paragraph("", body_style)] * (num_cols - len(gt_row_data))

            gt_table = LongTable([gt_row_data], colWidths=gt_col_widths)
            gt_style_cmds = [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")) # Background for GT row
            ]
            if num_cols > 1:
                # Span label, align right for value
                gt_style_cmds.extend([
                    ('SPAN', (0, 0), (-2, 0)),
                    ('ALIGN', (0, 0), (-2, 0), 'LEFT'), # Align "Grand Total" text left
                    ('ALIGN', (-1, 0), (-1, 0), 'RIGHT') # Align the value text right
                 ])
            else: # num_cols == 1
                gt_style_cmds.append(('ALIGN', (0,0), (0,0), 'LEFT')) # Align left in single column

            gt_table.setStyle(TableStyle(gt_style_cmds))
            elements.append(gt_table)

# Build PDF (unchanged)
try:
    doc.build(elements); pdf_buf.seek(0)
except Exception as e:
    st.error(f"Error building PDF: {e}");
    import traceback; st.error(traceback.format_exc());
    pdf_buf = None


# ─── Build Word ───────────────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document()
# Page Setup (unchanged)
sec = docx_doc.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_height = Inches(11); sec.page_width = Inches(17)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5); sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
# Logo and Title (unchanged)
if logo:
    try:
        p_logo = docx_doc.add_paragraph(); r_logo = p_logo.add_run(); img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width; img_width_in = 5; img_height_in = img_width_in * ratio; r_logo.add_picture(io.BytesIO(logo), width=Inches(img_width_in)); p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
    except Exception as e: st.warning(f"Could not add logo to Word: {e}")
p_title = docx_doc.add_paragraph(); p_title.alignment = WD_TABLE_ALIGNMENT.CENTER; r_title = p_title.add_run(proposal_title); r_title.font.name = DEFAULT_SERIF_FONT; r_title.font.size = Pt(18); r_title.bold = True; docx_doc.add_paragraph()

TOTAL_W_INCHES = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

# --- Iterate through processed table data ---
for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info): # Use correct variable names
    n = len(hdr)
    if n == 0:
         # st.warning(f"Skipping Word generation for table index {table_index}: Header is empty.")
         continue

    # --- Column widths calculation for Word ---
    # Similar logic to PDF for consistency
    desc_actual_idx_in_hdr = -1
    desc_w_in = 0
    other_w_in = 0

    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_w_in = 0.45 * TOTAL_W_INCHES # Description width
        other_cols_count = n - 1
        if other_cols_count > 0:
            other_total_w_in = TOTAL_W_INCHES - desc_w_in
            strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1
            if strategy_idx != -1:
                strat_w_in = 0.15 * TOTAL_W_INCHES
                remaining_w_in = other_total_w_in - strat_w_in
                remaining_cols = other_cols_count - 1
                other_w_in = remaining_w_in / remaining_cols if remaining_cols > 0 else 0
                # We'll set individual widths below, store strat_w_in for later use
            else:
                other_w_in = other_total_w_in / other_cols_count
        elif n == 1: # Only description
             desc_w_in = TOTAL_W_INCHES
             other_w_in = 0
        else: # Only one non-description column?
             other_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES
             desc_w_in = other_w_in # Fallback


    except ValueError:
        # Fallback if 'Description' not found
        # st.warning(f"Word Gen: 'Description' column not found in header for table {table_index}. Using equal widths. Header: {hdr}")
        desc_actual_idx_in_hdr = -1 # Mark as not found
        desc_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES
        other_w_in = desc_w_in
        strategy_idx = -1 # Cannot exist if description doesn't


    # Create table
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid");
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER;
    tbl.autofit = False; # Required to set manual widths
    tbl.allow_autofit = False; # Prevent Word from overriding

    # --- START: Set Preferred Table Width (Seems okay) ---
    tblPr_list = tbl._element.xpath('./w:tblPr')
    if not tblPr_list:
        tblPr = OxmlElement('w:tblPr')
        tbl._element.insert(0, tblPr)
    else:
        tblPr = tblPr_list[0]
    tblW = OxmlElement('w:tblW')
    # Set table width to 100% of the text area (5000 = 100% * 50)
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    existing_tblW = tblPr.xpath('./w:tblW')
    if existing_tblW: tblPr.remove(existing_tblW[0])
    tblPr.append(tblW)
    # --- END: Set Preferred Table Width ---

    # Set Column Widths
    for idx, col in enumerate(tbl.columns):
        width_val = 0
        if idx == desc_actual_idx_in_hdr:
             width_val = desc_w_in
        elif strategy_idx != -1 and idx == strategy_idx:
             width_val = strat_w_in # Use specific strategy width if calculated
        else:
             width_val = other_w_in # Use general other width

        # Ensure a minimum width to avoid errors
        col.width = Inches(max(0.2, width_val))


    # Populate Header Row (unchanged styling logic)
    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr):
        # Make sure we don't try to access cell index beyond actual table columns
        if i >= len(hdr_cells): break

        cell = hdr_cells[i]
        # Set cell background color using Oxml
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'F2F2F2') # Light grey fill
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        tcPr.append(shd)

        # Add text and format
        p = cell.paragraphs[0]
        # It's better to clear existing content (if any) before adding run
        p.text = ""
        run = p.add_run(str(col_name))
        run.font.name = DEFAULT_SERIF_FONT
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # --- START: Populate Data Rows for Word Output (Simpler "- link" approach) ---
    for ridx, row in enumerate(rows_data):
        # Ensure row has expected number of cells for safety
        current_cells_count = len(row)
        if current_cells_count < n:
            row = list(row) + [""] * (n - current_cells_count)
        elif current_cells_count > n:
            row = row[:n]

        row_cells = tbl.add_row().cells
        for cidx, cell_content in enumerate(row):
            # Ensure we don't index past the number of cells created
            if cidx >= len(row_cells): break

            cell = row_cells[cidx]
            p = cell.paragraphs[0]
            p.text = "" # Clear paragraph content first
            cell_str = str(cell_content)

            # Add the main text content as a normal run first
            run_text = p.add_run(cell_str)
            run_text.font.name = DEFAULT_SANS_FONT
            run_text.font.size = Pt(9)

            # Check if this is the description column (use generated index) AND if a link URI exists
            link_applied = False
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                if link_uri: # Check if URI is valid
                    # Add a space before the link text for separation only if main text exists
                    if cell_str:
                        space_run = p.add_run(" ")
                        space_run.font.name = DEFAULT_SANS_FONT # Match font if needed
                        space_run.font.size = Pt(9)

                    # Use the helper function to add ONLY the "- link" hyperlinked text
                    try:
                        add_hyperlink(p, link_uri, "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
                        link_applied = True
                    except Exception as link_e:
                        # st.warning(f"Failed to add hyperlink '{link_uri}' in Word table {table_index}, row {ridx}: {link_e}")
                        # Add non-linked text as fallback
                         failed_link_run = p.add_run("- link (error)")
                         failed_link_run.font.name = DEFAULT_SANS_FONT
                         failed_link_run.font.size = Pt(9)
                         failed_link_run.font.color.rgb = RGBColor(255, 0, 0) # Red color for error


            # Set alignment and vertical alignment for the cell
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    # --- END: Populate Data Rows for Word Output ---

    # --- START: Add Table Total Row (Corrected NameError Fix & Logic) ---
    if table_total_info:
        label = "Total"; amount = ""
        # Use same parsing logic as PDF for consistency
        if isinstance(table_total_info, list):
             label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
             amount = next((val.strip() for val in reversed(table_total_info) if val and '$' in str(val)), "")
             if not amount and len(table_total_info) > 1: amount = table_total_info[-1].strip() if table_total_info[-1] else ""
        elif isinstance(table_total_info, str):
            try: # Wrap regex in try-except
                total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
                if total_match:
                    label_parsed, amount_parsed = total_match.groups()
                    label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total"
                    amount = amount_parsed.strip() if amount_parsed else ""
                else:
                    amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                    if amount_match:
                         amount = amount_match.group(1).strip() if amount_match.group(1) else ""
                         potential_label = table_total_info[:amount_match.start()].strip()
                         label = potential_label if potential_label else "Total"
                    else:
                         amount = table_total_info # Fallback amount
                         label = "Total" # Fallback label
            except Exception as e:
                # st.warning(f"Regex error parsing total string '{table_total_info}' for Word: {e}")
                amount = table_total_info # Fallback safely
                label = "Total"

        # Populate Word row using parsed label and amount
        total_cells = tbl.add_row().cells
        if n > 0: # Proceed only if columns exist
            # Merge cells for label if more than 1 column
            label_cell = total_cells[0]
            if n > 1:
                try:
                     label_cell.merge(total_cells[n-2]) # Merge from first to second-to-last
                except IndexError:
                     pass # Ignore merge error if n=1 (should be handled by condition)
                except Exception as merge_e:
                    # st.warning(f"Error merging total row cells for Word table {table_index}: {merge_e}")
                    pass


            # Format Label cell
            p_label = label_cell.paragraphs[0]; p_label.text = ""; run_label = p_label.add_run(label);
            run_label.font.name = DEFAULT_SERIF_FONT; run_label.font.size = Pt(10); run_label.bold = True;
            p_label.alignment = WD_TABLE_ALIGNMENT.LEFT;
            label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;

            # Format Amount cell (last cell)
            if n > 1: # Amount cell only exists if n > 1
                 amount_cell = total_cells[n-1]
                 p_amount = amount_cell.paragraphs[0]; p_amount.text = ""; run_amount = p_amount.add_run(amount);
                 run_amount.font.name = DEFAULT_SERIF_FONT; run_amount.font.size = Pt(10); run_amount.bold = True;
                 p_amount.alignment = WD_TABLE_ALIGNMENT.RIGHT;
                 amount_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
            elif n == 1:
                # If only one column, put amount in the first cell (potentially overwriting label if generic)
                if label == "Total": p_label.text = amount; run_label.text = amount # Update run too
                # else keep specific label from PDF? Ambiguous. Best to show amount somehow.
                # Maybe append amount to label?
                # run_label.text = f"{label}: {amount}"


    # --- END Table Total Row ---

    docx_doc.add_paragraph() # Spacer between tables

# --- Add Grand Total row (Corrected tblPr Access & Logic) ---
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1] # Use last table's header info
    n = len(last_hdr)
    if n > 0:
        # --- Recalculate widths for GT table based on last table's structure ---
        # (Using the same width logic as the main table loop for consistency)
        gt_desc_idx = -1
        gt_desc_w = 0
        gt_other_w = 0
        gt_strat_w = 0 # Initialize strategy width
        gt_strat_idx = -1

        try:
             gt_desc_idx = last_hdr.index("Description")
             gt_desc_w = 0.45 * TOTAL_W_INCHES
             gt_other_count = n - 1
             if gt_other_count > 0:
                 gt_other_total_w = TOTAL_W_INCHES - gt_desc_w
                 gt_strat_idx = gt_desc_idx - 1 if gt_desc_idx > 0 and last_hdr[gt_desc_idx - 1] == "Strategy" else -1
                 if gt_strat_idx != -1:
                     gt_strat_w = 0.15 * TOTAL_W_INCHES
                     gt_remain_w = gt_other_total_w - gt_strat_w
                     gt_remain_cols = gt_other_count - 1
                     gt_other_w = gt_remain_w / gt_remain_cols if gt_remain_cols > 0 else 0
                 else:
                     gt_other_w = gt_other_total_w / gt_other_count
             elif n == 1: gt_desc_w = TOTAL_W_INCHES; gt_other_w = 0
             else: gt_other_w = TOTAL_W_INCHES / n; gt_desc_w = gt_other_w
        except ValueError:
            gt_desc_idx = -1
            gt_desc_w = TOTAL_W_INCHES / n
            gt_other_w = gt_desc_w
            gt_strat_idx = -1

        # --- Create GT table ---
        tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid");
        tblg.alignment = WD_TABLE_ALIGNMENT.CENTER;
        tblg.autofit = False; tblg.allow_autofit = False;

        # --- Set Preferred Table Width for Grand Total Table ---
        tblgPr_list = tblg._element.xpath('./w:tblPr')
        if not tblgPr_list: tblgPr = OxmlElement('w:tblPr'); tblg._element.insert(0, tblgPr)
        else: tblgPr = tblgPr_list[0]
        tblgW = OxmlElement('w:tblW'); tblgW.set(qn('w:w'), '5000'); tblgW.set(qn('w:type'), 'pct');
        existing_tblgW_gt = tblgPr.xpath('./w:tblW')
        if existing_tblgW_gt: tblgPr.remove(existing_tblgW_gt[0])
        tblgPr.append(tblgW)
        # --- END: Set Preferred Table Width ---

        # Set column widths for GT table
        for idx, col in enumerate(tblg.columns):
            width_val_gt = 0
            if idx == gt_desc_idx: width_val_gt = gt_desc_w
            elif gt_strat_idx != -1 and idx == gt_strat_idx: width_val_gt = gt_strat_w
            else: width_val_gt = gt_other_w
            col.width = Inches(max(0.2, width_val_gt))

        # Populate GT row
        gt_cells = tblg.rows[0].cells
        if n > 0: # Ensure cells exist
             gt_label_cell = gt_cells[0]
             # Merge label cell if n > 1
             if n > 1:
                 try:
                      gt_label_cell.merge(gt_cells[n-2])
                 except IndexError: pass # n=1 case
                 except Exception as merge_e:
                    # st.warning(f"Error merging Grand Total row cells: {merge_e}")
                    pass

             # Format label cell (with background)
             tc_label = gt_label_cell._tc; tcPr_label = tc_label.get_or_add_tcPr();
             shd_label = OxmlElement('w:shd'); shd_label.set(qn('w:fill'), 'E0E0E0'); shd_label.set(qn('w:val'), 'clear'); shd_label.set(qn('w:color'), 'auto'); tcPr_label.append(shd_label); # Grey Background
             p_gt_label = gt_label_cell.paragraphs[0]; p_gt_label.text = ""; run_gt_label = p_gt_label.add_run("Grand Total");
             run_gt_label.font.name = DEFAULT_SERIF_FONT; run_gt_label.font.size = Pt(10); run_gt_label.bold = True;
             p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT; gt_label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;

             # Format value cell (last cell, only if n > 1)
             if n > 1:
                 gt_value_cell = gt_cells[n-1];
                 tc_val = gt_value_cell._tc; tcPr_val = tc_val.get_or_add_tcPr();
                 shd_val = OxmlElement('w:shd'); shd_val.set(qn('w:fill'), 'E0E0E0'); shd_val.set(qn('w:val'), 'clear'); shd_val.set(qn('w:color'), 'auto'); tcPr_val.append(shd_val); # Grey Background
                 p_gt_val = gt_value_cell.paragraphs[0]; p_gt_val.text = ""; run_gt_val = p_gt_val.add_run(grand_total);
                 run_gt_val.font.name = DEFAULT_SERIF_FONT; run_gt_val.font.size = Pt(10); run_gt_val.bold = True;
                 p_gt_val.alignment = WD_TABLE_ALIGNMENT.RIGHT; gt_value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
             elif n==1:
                 # If only one column, put GT value in the first cell
                  run_gt_label.text = f"Grand Total: {grand_total}"
                  p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT


# === End Word Build ===

# --- Save and Download Buttons (unchanged) ---
try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except Exception as e:
    st.error(f"Error building Word document: {e}")
    import traceback
    st.error(traceback.format_exc())
    docx_buf = None # Indicate failure

c1, c2 = st.columns(2)
if pdf_buf:
    with c1: st.download_button("📥 Download deliverable PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf", use_container_width=True)
else:
     with c1: st.error("PDF generation failed.")
if docx_buf:
    with c2: st.download_button("📥 Download deliverable DOCX", data=docx_buf, file_name="proposal_deliverable.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    with c2: st.error("Word document generation failed.")
