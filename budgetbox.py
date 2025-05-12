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
from reportlab.lib.enums import TA_CENTER, TA_LEFT
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
        words = page.get_text("words", clip=bbox) # Get words within the bbox

        if not words:
            return ""

        lines = {}
        # Group words by baseline (y1)
        for w in words:
            # x0, y0, x1, y1, word, block_no, line_no, word_no
            key = round(w[3], 1) # Use y1 as the key for line grouping
            lines.setdefault(key, []).append(w)

        text_lines = []
        # Sort lines by vertical position, then words by horizontal position
        for key in sorted(lines.keys()):
            # Sort words in the line by their x0 coordinate
            row = sorted(lines[key], key=lambda w: w[0])
            pieces = []
            for word_info in row:
                t = word_info[4].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                # Check flags for bold (flag 2) - need font info for more styles
                # Get span info to check flags requires "dict" or "rawdict"
                # This simplified word approach won't easily get bold.
                # Reverting to get_text("dict") approach but clipped.
                # Let's stick to the original 'extract_rich_cell' for simplicity for now,
                # accepting its limitations, but apply it to the notes column.
                # The original function is actually better suited if spans cross bbox.
                # Re-implementing the original logic slightly cleaned up:

                # Re-fetch with dict for span info, using clip
                d = page.get_text("dict", clip=bbox)
                spans = []
                x0,y0,x1,y1 = bbox
                for block in d["blocks"]:
                    if block.get("type")!=0: # Text blocks only
                        continue
                    for line in block["lines"]:
                        for span in line["spans"]:
                            # Check if span intersects bbox (more robust than simple contains)
                            sx0,sy0,sx1,sy1 = span["bbox"]
                            # Check for overlap
                            if sx0 < x1 and sx1 > x0 and sy0 < y1 and sy1 > y0:
                                spans.append(span)

                lines_dict = {}
                for s in spans:
                    # Group by baseline y1, rounded
                    key = round(s["bbox"][3], 1)
                    lines_dict.setdefault(key, []).append(s)

                span_text_lines = []
                for key in sorted(lines_dict.keys()):
                    # Sort spans in line by x0
                    row_spans = sorted(lines_dict[key], key=lambda s: s["bbox"][0])
                    line_pieces = []
                    last_x1 = bbox[0] # Start from the left edge of the bbox
                    for span in row_spans:
                        # Add space if there's a gap between spans
                        if span["bbox"][0] > last_x1 + 1: # Add tolerance for space
                             line_pieces.append(" ")
                        t = span["text"].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                        # Check flags: 2 is bold, 4 is italic (need font check really)
                        is_bold = span["flags"] & 2
                        # is_italic = span["flags"] & 4 # Example if needed
                        if is_bold:
                            line_pieces.append(f"<b>{t}</b>")
                        # elif is_italic:
                        #     line_pieces.append(f"<i>{t}</i>") # Example
                        else:
                            line_pieces.append(t)
                        last_x1 = span["bbox"][2]
                    span_text_lines.append("".join(line_pieces))
                # Handle bullet points: If a line starts with common bullet characters, add <ul><li> tags?
                # For simplicity, just preserve bullet characters and rely on <br/>.
                # ReportLab Paragraph handles basic HTML like <br/> and <b>.
                return "<br/>".join(span_text_lines)

    except Exception as e:
        # st.warning(f"Error in extract_rich_cell: {e}") # Optional: for debugging
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
    tables = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n")
    if tables:
        raw = tables[0].df.values.tolist()
        # Basic validation: check if table has >1 row and enough columns
        if len(raw) > 1 and len(raw[0]) >= 6: # Check against a reasonable number of expected cols
             # Heuristic: Check if first row looks like headers (non-numeric, contains keywords)
             header_row = [str(h).lower() for h in raw[0]]
             if any(kw in ''.join(header_row) for kw in ['description', 'date', 'term', 'amount', 'total', 'notes']):
                first_table = raw
except Exception as e:
    # st.warning(f"Camelot failed on first page: {e}") # Optional: for debugging
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        texts = [(p.page_number, p.extract_text(x_tolerance=1, y_tolerance=1) or "") for p in pdf.pages]
        first_page_num, first_text = texts[0] if texts else (0, "")
        first_lines = first_text.splitlines()

        # Extract Proposal Title
        pot = next((l.strip() for l in first_lines if "proposal" in l.lower() and len(l.strip())>10), None) # More specific title search
        if pot:
            proposal_title = pot
        elif first_lines:
            # Take the first non-empty line as a fallback title
            proposal_title = next((l.strip() for l in first_lines if l.strip()), "Untitled Proposal")

        used_total_lines = set()
        def find_total(page_num_idx):
            """Finds a 'Total' line on a specific page."""
            if page_num_idx >= len(texts):
                return None
            page_num, text_content = texts[page_num_idx]
            for l in text_content.splitlines():
                line_key = (page_num, l) # Use page number and line content as unique key
                # Regex: Looks for 'total' (not preceded by 'grand'), followed by '$' and digits/commas/dots.
                if re.search(r'\b(?<!grand\s)total\b.*?\$\s*[\d,.]+',l, re.I) and line_key not in used_total_lines:
                    used_total_lines.add(line_key)
                    return l.strip()
            return None

        # Process tables page by page
        for pi, page in enumerate(pdf.pages):
            page_num = page.page_number # pdfplumber page num is 1-based usually
            page_content = next((text for num, text in texts if num == page_num), "") # Get pre-extracted text

            # Use Camelot data if it's the first page and valid table was found
            if pi == 0 and first_table:
                # Simulate the structure expected by the loop
                # Camelot doesn't provide bbox or rows_obj easily, so set to None
                found = [("camelot", first_table, None, None)]
                links = [] # No link info from Camelot
            else:
                # Use pdfplumber for other pages or if Camelot failed
                # find_tables settings can be tuned (e.g., table_settings={"vertical_strategy": "lines"})
                found_tables = page.find_tables(table_settings={"vertical_strategy": "lines", "horizontal_strategy": "lines"})
                found = [(tbl, tbl.extract(x_tolerance=1, y_tolerance=1), tbl.bbox, tbl.rows) for tbl in found_tables]
                links = page.hyperlinks

            for tbl_obj, data, bbox, rows_obj in found:
                if not data or len(data) < 2: # Need header + at least one data row
                    continue

                # Clean and identify header columns
                hdr = [str(h).strip().replace('\n', ' ') for h in data[0]] # Clean headers
                # Find column indices (case-insensitive)
                desc_i = next((i for i, h in enumerate(hdr) if "description" in h.lower()), None)
                notes_i = next((i for i, h in enumerate(hdr) if any(x in h.lower() for x in ["note", "comment"])), None)
                start_date_i = next((i for i, h in enumerate(hdr) if "start date" in h.lower()), None)
                end_date_i = next((i for i, h in enumerate(hdr) if "end date" in h.lower()), None)
                term_i = next((i for i, h in enumerate(hdr) if "term" in h.lower() or "month" in h.lower()), None)
                monthly_i = next((i for i, h in enumerate(hdr) if "monthly" in h.lower()), None)
                total_i = next((i for i, h in enumerate(hdr) if "total" in h.lower() and "grand" not in h.lower() and "month" not in h.lower()), None) # Avoid month/grand


                # If no description column found, try a heuristic (e.g., first column) or skip
                if desc_i is None:
                     desc_i = 0 # Assume first column is description if specific header not found
                     # Add more heuristics if needed, e.g., check for longest text column
                     # If still unreliable, could skip: continue

                # Extract Links associated with description cells (only possible with pdfplumber)
                desc_links = {}
                if tbl_obj != "camelot" and rows_obj and desc_i is not None:
                    for r, row_obj in enumerate(rows_obj):
                        if r == 0: continue # Skip header row_obj
                        if desc_i < len(row_obj.cells):
                            cell_bbox = row_obj.cells[desc_i]
                            if cell_bbox: # Ensure bbox exists
                                x0, top, x1, bottom = cell_bbox
                                for link in links:
                                    # Check if link object has necessary keys and overlaps with cell
                                    if all(k in link for k in ("x0", "x1", "top", "bottom", "uri")):
                                        if not (link["x1"] < x0 or link["x0"] > x1 or link["bottom"] < top or link["top"] > bottom):
                                            desc_links[r] = link["uri"] # Map row index (1-based) to link URI
                                            break # Assume one link per cell max

                table_rows = []
                row_links_in_order = []
                table_total_row = None

                # Process data rows
                for ridx, row in enumerate(data[1:], start=1): # Start from 1 for data rows
                    cells = [str(c).strip() if c is not None else "" for c in row] # Ensure strings and handle None
                    if not any(cells): # Skip empty rows
                        continue

                    # Check for table total/subtotal rows based on content
                    first_cell_lower = cells[0].lower() if cells else ""
                    # More robust check for total rows
                    is_total_row = ("total" in first_cell_lower or "subtotal" in first_cell_lower) and any("$" in c for c in cells)
                    if is_total_row:
                        if table_total_row is None: # Capture the first total row found within the table
                            table_total_row = cells
                        continue # Don't process total rows as data

                    # Initialize the standardized row based on HEADERS
                    new_row = [""] * len(HEADERS)

                    # --- Assign data to standardized columns ---

                    # 0: Description
                    if desc_i is not None and desc_i < len(cells):
                        raw_desc = cells[desc_i]
                        # Use rich text extraction if possible (pdfplumber obj and valid bbox)
                        if tbl_obj != "camelot" and rows_obj and desc_i < len(rows_obj[ridx].cells) and rows_obj[ridx].cells[desc_i]:
                             cell_bbox = rows_obj[ridx].cells[desc_i]
                             rich_desc = extract_rich_cell(page_num - 1, cell_bbox) # Fitz pages are 0-indexed
                             new_row[0] = rich_desc or raw_desc.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")
                        else: # Fallback for Camelot or missing bbox
                            new_row[0] = raw_desc.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")

                    # 6: Notes (Handle similarly to Description with rich text)
                    if notes_i is not None and notes_i < len(cells):
                        raw_notes = cells[notes_i]
                        if tbl_obj != "camelot" and rows_obj and notes_i < len(rows_obj[ridx].cells) and rows_obj[ridx].cells[notes_i]:
                            cell_bbox = rows_obj[ridx].cells[notes_i]
                            rich_notes = extract_rich_cell(page_num - 1, cell_bbox) # Fitz pages are 0-indexed
                            # Preserve bullet points and structure from rich extraction or basic replace
                            new_row[6] = rich_notes or raw_notes.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")
                        else:
                             new_row[6] = raw_notes.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace("\n","<br/>")


                    # Attempt assignment based on identified header indices first
                    if start_date_i is not None and start_date_i < len(cells): new_row[1] = cells[start_date_i]
                    if end_date_i is not None and end_date_i < len(cells):   new_row[2] = cells[end_date_i]
                    if term_i is not None and term_i < len(cells):       new_row[3] = cells[term_i]
                    if monthly_i is not None and monthly_i < len(cells):    new_row[4] = cells[monthly_i]
                    if total_i is not None and total_i < len(cells):      new_row[5] = cells[total_i]

                    # Fallback: Iterate through remaining cells and guess based on content
                    # This is less reliable and should ideally be avoided if headers are consistent
                    for i, v in enumerate(cells):
                        if not v: continue # Skip empty cells
                        # Skip if column index corresponds to already assigned Description or Notes
                        if (desc_i is not None and i == desc_i) or \
                           (notes_i is not None and i == notes_i):
                            continue
                        # Skip if column index corresponds to an already assigned column via header index
                        if (start_date_i is not None and i == start_date_i) or \
                           (end_date_i is not None and i == end_date_i) or \
                           (term_i is not None and i == term_i) or \
                           (monthly_i is not None and i == monthly_i) or \
                           (total_i is not None and i == total_i):
                            continue

                        # Guessing based on content (use sparingly)
                        if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', v): # Date format
                            if not new_row[1]: new_row[1] = v
                            elif not new_row[2]: new_row[2] = v
                        elif re.fullmatch(r"\d{1,3}", v) or "month" in v.lower() or "mo" in v.lower(): # Term
                            if not new_row[3]: new_row[3] = v
                        elif "$" in v: # Currency
                            if not new_row[4]: new_row[4] = v # Assume monthly first
                            elif not new_row[5]: new_row[5] = v # Then item total

                    # Final check: Don't add rows that just repeat the main headers
                    if any(new_row[i].strip().replace('\n',' ') == HEADERS[i] for i in range(len(HEADERS)) if new_row[i]):
                        continue

                    table_rows.append(new_row)
                    row_links_in_order.append(desc_links.get(ridx)) # Append link for this row (or None)

                # If no total row found within table, try searching page text
                if table_total_row is None:
                    table_total_row = find_total(pi) # pi is the 0-based index for 'texts' list

                # Add extracted table data if any rows were processed
                if table_rows:
                    tables_info.append((HEADERS, table_rows, row_links_in_order, table_total_row))

        # Find Grand Total (Search from the end of the document)
        for page_num, blk in reversed(texts):
            # Regex: 'Grand Total' possibly followed by anything (incl newlines) then '$' and amount
            m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', blk, re.I | re.S)
            if m:
                grand_total = m.group(1).replace(" ", "") # Remove spaces from amount
                break
except pdfplumber.PDFSyntaxError as e:
    st.error(f"PDFPlumber Error: Failed to process PDF. It might be corrupted or password-protected. Error: {e}")
    st.stop()
except Exception as e:
    st.error(f"An unexpected error occurred during PDF processing: {e}")
    # Optionally add more detailed logging here
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
hs = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black, spaceAfter=6) # Header style
bs = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=12) # Body text style with leading
# Style for right-aligned currency
bs_right = ParagraphStyle("BodyRight", parent=bs, alignment=TA_RIGHT)
# Style for centered term
bs_center = ParagraphStyle("BodyCenter", parent=bs, alignment=TA_CENTER)

story = [Spacer(1, 12), Paragraph(proposal_title, ts), Spacer(1, 24)]
table_width = doc.width

# Column widths (adjust percentages as needed)
col_widths = [
    table_width * 0.30, # Description
    table_width * 0.08, # Start Date
    table_width * 0.08, # End Date
    table_width * 0.06, # Term (Months) - Centered
    table_width * 0.10, # Monthly Amount - Right aligned
    table_width * 0.10, # Item Total - Right aligned
    table_width * 0.28  # Notes
]

for headers, rows, links, total_info in tables_info:
    n_cols = len(headers)
    # Ensure col_widths matches header length if dynamic headers were ever introduced
    if len(col_widths) != n_cols:
        # Simple fallback: distribute width equally if lengths don't match
        col_widths = [table_width / n_cols] * n_cols

    # Header row with Paragraph styles
    header_row_styled = [Paragraph(h, hs) for h in headers]
    table_data = [header_row_styled]

    # Process data rows
    for i, row_data in enumerate(rows):
        styled_row = []
        for j, cell_text in enumerate(row_data):
            cell_style = bs # Default body style
            # Apply specific styles based on column index
            if j == 3: # Term (Months)
                cell_style = bs_center
            elif j in [4, 5]: # Monthly Amount, Item Total
                cell_style = bs_right

            # Add hyperlink if available for the description column (index 0)
            if j == 0 and i < len(links) and links[i]:
                 # Append link HTML to the cell text
                 linked_text = cell_text + f" <link href='{links[i]}' color='blue'>[link]</link>"
                 styled_row.append(Paragraph(linked_text, cell_style))
            else:
                 styled_row.append(Paragraph(cell_text, cell_style))
        table_data.append(styled_row)

    # Add total row if found
    if total_info:
        total_label = "Total"
        total_value = ""
        if isinstance(total_info, list): # If it was a row from the table
            # Try to find label (non-empty cell before potential '$' value) and value
             total_label = next((c for c in total_info if c and '$' not in c), "Total") # Find first non-empty, non-$ cell
             total_value = next((c for c in reversed(total_info) if "$" in c), "") # Find last '$' value
        elif isinstance(total_info, str): # If it was found via regex search
            m = re.match(r'(.*?)\s*(\$\s*[\d,]+\.\d{2})', total_info) # Extract label and value
            if m:
                total_label, total_value = m.group(1).strip(), m.group(2).strip()
            else: # Fallback if regex match fails but string exists
                total_label = total_info.replace("$","").strip() # Basic label guess
                if "$" in total_info: total_value = "$" + total_info.split("$")[-1] # Basic value guess


        # Create styled total row (spans first n-2 columns)
        total_row_styled = [Paragraph(total_label, bs)] + \
                           [Paragraph("", bs)] * (n_cols - 2) + \
                           [Paragraph(total_value, bs_right)] # Right-align total value
        table_data.append(total_row_styled)

    # Create LongTable
    tbl = LongTable(table_data, colWidths=col_widths, repeatRows=1) # Repeat header row

    # Table styling commands
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")), # Header background
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),            # Grid lines
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),                    # Header vertical align
        ("VALIGN", (0, 1), (-1, -1), "TOP"),                      # Body cells vertical align top
        # Alignment commands based on column content type
        ("ALIGN", (1, 1), (2, -1), "CENTER"), # Dates (Start, End) centered
        ("ALIGN", (3, 1), (3, -1), "CENTER"), # Term centered
        ("ALIGN", (4, 1), (5, -1), "RIGHT"),  # Amounts (Monthly, Total) right aligned
    ]

    # Add specific styling for the total row if it exists
    if total_info:
        style_cmds.extend([
            ("SPAN", (0, -1), (-2, -1)),             # Span label across columns 0 to n-2
            ("ALIGN", (0, -1), (-2, -1), "RIGHT"),   # Align label right within the span (optional, LEFT is default)
            ("ALIGN", (-1, -1), (-1, -1), "RIGHT"),  # Align value in the last column right
            ("VALIGN", (0, -1), (-1, -1), "MIDDLE"), # Middle align total row vertically
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAEAEA")), # Optional: Background for total row
        ])

    tbl.setStyle(TableStyle(style_cmds))
    story.extend([tbl, Spacer(1, 24)]) # Add table and spacer to story

# Add Grand Total row if found
if grand_total:
    # Styled grand total row (similar structure to table total row)
    grand_total_row_styled = [Paragraph("Grand Total", bs)] + \
                             [Paragraph("", bs)] * (len(HEADERS) - 2) + \
                             [Paragraph(grand_total, bs_right)] # Right-align grand total value

    gt_table = LongTable([grand_total_row_styled], colWidths=col_widths)
    gt_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")), # Background color
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),           # Grid
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),                  # Vertical alignment
        ("SPAN", (0, 0), (-2, 0)),                              # Span label
        ("ALIGN", (0, 0), (-2, 0), "RIGHT"),                    # Align label (optional)
        ("ALIGN", (-1, 0), (-1, 0), "RIGHT")                     # Align value right
    ]))
    story.append(gt_table)

# Build the PDF document
try:
    doc.build(story)
    pdf_buf.seek(0)
    # Provide download button
    st.download_button(
        "📥 Download Transformed PDF",
        data=pdf_buf,
        file_name="transformed_proposal.pdf",
        mime="application/pdf",
        use_container_width=True
        )
except Exception as e:
    st.error(f"Error building final PDF with ReportLab: {e}")
