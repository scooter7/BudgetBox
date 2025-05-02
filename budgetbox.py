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
    return lines[0], " ".join(lines[1:])

# --- REVISED Word hyperlink helper ---
def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
    """
    Add a hyperlink to a paragraph.

    :param paragraph: The paragraph to add the hyperlink to.
    :param url: The URL for the hyperlink.
    :param text: The text to display for the hyperlink.
    :param font_name: Optional font name for the hyperlink text.
    :param font_size: Optional font size (in Pt) for the hyperlink text.
    :param bold: Optional boolean to make the hyperlink text bold.
    :return: The run object containing the hyperlink.
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a w:r element for the hyperlink text
    new_run = OxmlElement('w:r')

    # Create a w:rPr element for run properties (styling)
    rPr = OxmlElement('w:rPr')

    # Apply the standard Hyperlink character style
    # Check if the style exists, add if necessary (optional, Word usually has it)
    styles = paragraph.part.document.styles
    if "Hyperlink" not in styles:
         # Basic definition if style is missing, might need refinement
        style = styles.add_style("Hyperlink", WD_STYLE_TYPE.CHARACTER, True) # True for built-in=False
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1) # Standard blue
        style.font.underline = True
        style.priority = 9 # Default priority
        style.unhide_when_used = True

    # Add reference to the Hyperlink style
    style_element = OxmlElement('w:rStyle')
    style_element.set(qn('w:val'), 'Hyperlink')
    rPr.append(style_element)

    # Apply optional direct formatting (overrides style if needed)
    if font_name:
        run_font = OxmlElement('w:rFonts')
        run_font.set(qn('w:ascii'), font_name)
        run_font.set(qn('w:hAnsi'), font_name) # Also set hAnsi for compatibility
        # Consider adding w:cs for complex scripts and w:eastAsia if needed
        rPr.append(run_font)
    if font_size:
        size = OxmlElement('w:sz')
        # Ensure font_size is treated as points and converted to half-points
        size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs') # Also set complex script size
        size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size)
        rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)

    # Add run properties to the run
    new_run.append(rPr)

    # Add the text preserving whitespace according to XML spec
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_run.append(t)

    # Append the run to the hyperlink element
    hyperlink.append(new_run)

    # Append the hyperlink element to the paragraph's XML element (_p)
    paragraph._p.append(hyperlink)

    # Return a proxy Run object wrapping the new w:r element
    return docx.text.run.Run(new_run, paragraph)
# --- End of REVISED Word hyperlink helper ---


# Extract tables & totals, capturing hyperlink per row
tables_info = []
grand_total = None
proposal_title = "Untitled Proposal" # Default title

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages] # Adjust tolerances if needed

        # Try to find a better proposal title (example logic)
        first_page_lines = page_texts[0].splitlines() if page_texts else []
        potential_title = next((line.strip() for line in first_page_lines if "proposal" in line.lower() and len(line.strip()) > 5), None)
        if potential_title:
            proposal_title = potential_title
        elif len(first_page_lines) > 0:
             proposal_title = first_page_lines[0].strip() # Fallback to first line

        used_totals = set()
        def find_total(pi):
            if pi >= len(page_texts): return None # Boundary check
            for ln in page_texts[pi].splitlines():
                # Improved regex to find "Total" (case-insensitive) followed by a dollar amount
                # Avoid matching "Grand Total" here if possible
                if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            # Extract hyperlinks using page.hyperlinks property
            links = page.hyperlinks # list of dicts: {'x0', 'top', 'x1', 'bottom', 'uri', 'object_id', 'page_number'}

            page_tables = page.find_tables()
            if not page_tables: continue

            for tbl in page_tables:
                data = tbl.extract(x_tolerance=1, y_tolerance=1) # Adjust tolerances if needed
                if not data or len(data) < 2: # Need header and at least one row
                    continue

                hdr = data[0]
                # Ensure header cells are strings and handle None
                hdr = [(str(h).strip() if h else "") for h in hdr]

                # Find 'description' column index robustly
                desc_i = next((i for i, h in enumerate(hdr) if h and "description" in h.lower()), None)
                if desc_i is None:
                    # Try finding a column with significant text if 'description' is missing
                    desc_i = next((i for i, h in enumerate(hdr) if len(h) > 10), None) # Example fallback
                    if desc_i is None or len(hdr) <= 1 : continue # Skip if no suitable description column found

                # Build map row_index -> URL for Description column using link bounding boxes
                desc_links = {}
                column_coords = []
                if hasattr(tbl, 'rows'): # Check if tbl object has rows
                    for r, row_obj in enumerate(tbl.rows):
                        if r == 0: continue # Skip header row
                        if hasattr(row_obj, 'cells') and desc_i is not None and desc_i < len(row_obj.cells):
                            column_coords.append((r, desc_i))

                # Iterate through the valid coordinates found
                for row_idx_rel, cell_coord in enumerate(column_coords):
                    row_tbl_idx = cell_coord[0]
                    col_tbl_idx = cell_coord[1]

                    if row_tbl_idx < len(tbl.rows) and \
                       hasattr(tbl.rows[row_tbl_idx], 'cells') and \
                       col_tbl_idx < len(tbl.rows[row_tbl_idx].cells):
                        cell_bbox = tbl.rows[row_tbl_idx].cells[col_tbl_idx]
                    else: continue

                    if not cell_bbox: continue
                    x0_cell, top_cell, x1_cell, bottom_cell = cell_bbox

                    temp_links = list(links)
                    for link in temp_links:
                        if not all(k in link for k in ['x0', 'x1', 'top', 'bottom', 'uri']): continue

                        link_center_x = (link.get('x0', 0) + link.get('x1', 0)) / 2
                        link_overlaps_vertically = not (link.get('bottom', 0) < top_cell or link.get('top', 0) > bottom_cell)

                        if (link_center_x >= x0_cell and link_center_x <= x1_cell and link_overlaps_vertically):
                            desc_links[row_tbl_idx] = link.get("uri")
                            # Optional removal:
                            # try: links.remove(link)
                            # except ValueError: pass
                            break

                # --- START OF MODIFIED SECTION TO CAPTURE TABLE TOTAL ---
                new_hdr = ["Strategy", "Description"] + [h for i, h in enumerate(hdr) if i != desc_i and h]
                rows = []
                row_links_list = []
                table_total_info = None # Initialize as None

                for ridx_data, row in enumerate(data[1:], start=1): # ridx_data is 1-based table row index
                    row_str = [(str(cell).strip() if cell else "") for cell in row]
                    if all(not cell_val for cell_val in row_str): continue # Skip empty

                    # Check if this row looks like a total row based on content
                    first_cell_lower = row_str[0].lower() if row_str else ""
                    # More specific check: "total" or "subtotal" in first cell AND a '$' sign somewhere in the row
                    if ("total" in first_cell_lower or "subtotal" in first_cell_lower) and \
                       any("$" in str(cell_val) for cell_val in row_str):
                        table_total_info = row_str # Capture the row list
                        # Optional: break here if you are sure the total is the absolute last row
                        # break
                        continue # Skip adding this row to the main 'rows' list

                    # --- If not a total row, process normally ---
                    desc_text = row_str[desc_i] if desc_i < len(row_str) else ""
                    strat, desc = split_cell_text(desc_text)
                    rest = [row_str[i] for i, h in enumerate(hdr) if i != desc_i and h and i < len(row_str)]

                    rows.append([strat, desc] + rest)
                    row_links_list.append(desc_links.get(ridx_data))

                # Fallback: If no total row found in tbl.extract(), search page text
                if table_total_info is None:
                    table_total_info = find_total(pi) # Result is a string or None

                # --- END OF MODIFIED SECTION ---

                if rows: # Only add table if it has valid data rows
                     # Store the captured info (list, string, or None)
                     tables_info.append((new_hdr, rows, row_links_list, table_total_info))


        # Find Grand total robustly
        for tx in reversed(page_texts):
            m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I | re.S)
            if m:
                grand_total = m.group(1).replace(" ", "")
                break

except Exception as e:
    st.error(f"Error processing PDF: {e}")
    # import traceback
    # st.error(traceback.format_exc())
    st.stop()


# ─── Build PDF ────────────────────────────────────────────────────────────────
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf,
    pagesize=landscape((17*inch, 11*inch)), # Swapped width/height for landscape
    leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch # Standard margins
)

# Define Paragraph Styles using registered/fallback fonts
title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11) # Added leading
link_style   = ParagraphStyle("Link", parent=body_style, textColor=colors.blue, underline=False) # Separate style for links
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, textColor=colors.black, spaceBefore=6)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, textColor=colors.black, spaceBefore=6)

elements = []
logo = None # Initialize logo
try:
    # Consider making logo uploadable or configurable
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    response = requests.get(logo_url, timeout=10)
    response.raise_for_status() # Check if request was successful
    logo = response.content
    # Calculate width/height maintaining aspect ratio, limit max width
    img = Image.open(io.BytesIO(logo))
    ratio = img.height / img.width
    img_width = min(5*inch, doc.width) # Limit width to 5 inches or page width
    img_height = img_width * ratio
    elements.append(RLImage(io.BytesIO(logo), width=img_width, height=img_height))
except requests.exceptions.RequestException as e:
    st.warning(f"Could not download logo: {e}")
except Exception as e:
     st.warning(f"Could not process logo image: {e}")

elements += [Spacer(1, 12), Paragraph(html.escape(proposal_title), title_style), Spacer(1, 24)]

total_page_width = doc.width # Usable width within margins

for hdr, rows, row_links_list, table_total_info in tables_info:
    num_cols = len(hdr)
    # Define column widths
    desc_col_width = total_page_width * 0.45
    other_col_width = (total_page_width - desc_col_width) / (num_cols - 1) if num_cols > 1 else 0
    col_widths = [other_col_width] * num_cols
    try:
        desc_actual_idx = hdr.index("Description")
        col_widths[desc_actual_idx] = desc_col_width
        other_total_width = total_page_width - desc_col_width
        other_cols_count = num_cols - 1
        if other_cols_count > 0:
            new_other_width = other_total_width / other_cols_count
            col_widths = [new_other_width if i != desc_actual_idx else desc_col_width for i in range(num_cols)]
    except ValueError:
         desc_actual_idx = 1
         col_widths = [total_page_width / num_cols] * num_cols

    # Wrap header text
    wrapped_header = [Paragraph(html.escape(str(h)), header_style) for h in hdr]
    wrapped_data = [wrapped_header] # Start table data with header

    # Process data rows
    for ridx, row in enumerate(rows):
        line = []
        for cidx, cell in enumerate(row):
            cell_str = str(cell)
            escaped_cell_text = html.escape(cell_str)
            if cidx == desc_actual_idx and ridx < len(row_links_list) and row_links_list[ridx]:
                escaped_url = html.escape(row_links_list[ridx])
                link_text = f'<link href="{escaped_url}" color="blue">{escaped_cell_text}</link>'
                p = Paragraph(link_text, body_style)
            else:
                p = Paragraph(escaped_cell_text, body_style)
            line.append(p)
        wrapped_data.append(line)

    # --- START MODIFIED PDF TOTAL ROW HANDLING ---
    has_total_row = False # Flag to track if total row is added
    if table_total_info:
        total_row_elements = []
        label = "Total" # Default
        value = "" # Default

        if isinstance(table_total_info, list):
            # Found total row in extracted data (list of strings)
            label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
            # Find the last element that contains a '$' as the likely value
            value = next((val.strip() for val in reversed(table_total_info) if '$' in str(val)), "")
            if not value and len(table_total_info) > 1: # Fallback to last element if no '$' found
                 value = table_total_info[-1].strip()

        elif isinstance(table_total_info, str):
            # Got total from find_total (string) - parse it
            total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
            if total_match:
                label_parsed, value = total_match.groups()
                label = label_parsed.strip() if label_parsed.strip() else "Total"
                value = value.strip()
            else: # Try parsing just value if pattern fails
                 amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                 if amount_match:
                      value = amount_match.group(1).strip()
                      potential_label = table_total_info[:amount_match.start()].strip()
                      label = potential_label if potential_label else "Total"
                 else: # Cannot reliably parse
                      value = table_total_info # Display raw string as value

        # Create the formatted row for ReportLab
        if num_cols > 0:
            total_row_elements = [Paragraph(html.escape(label), bl_style)] + \
                                 [Paragraph("", body_style)] * (num_cols - 2) + \
                                 [Paragraph(html.escape(value), br_style) if num_cols > 1 else Paragraph(html.escape(value), bl_style)]
            # Ensure correct column count
            total_row_elements += [Paragraph("", body_style)] * (num_cols - len(total_row_elements))
            wrapped_data.append(total_row_elements) # Add the total row to the table data
            has_total_row = True
    # --- END MODIFIED PDF TOTAL ROW HANDLING ---

    # Create and style the table
    if wrapped_data and col_widths:
        tbl = LongTable(wrapped_data, colWidths=col_widths, repeatRows=1)
        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ]
        # Add styles specific to the total row only if it exists
        if has_total_row and num_cols > 1:
             style_commands.extend([
                ('SPAN', (0, -1), (-2, -1)),
                ('ALIGN', (0, -1), (-2, -1), 'LEFT'),
                ('ALIGN', (-1, -1), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
             ])
        elif has_total_row and num_cols == 1:
              style_commands.extend([
                 ('ALIGN', (0, -1), (0, -1), 'LEFT'),
                 ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
              ])

        tbl.setStyle(TableStyle(style_commands))
        elements += [tbl, Spacer(1, 24)]

# Add Grand Total row if present (No change needed here)
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]
    num_cols = len(last_hdr)
    desc_col_width = total_page_width * 0.45
    try:
        desc_actual_idx = last_hdr.index("Description")
        if num_cols > 1:
            other_total_width = total_page_width - desc_col_width
            other_cols_count = num_cols - 1
            new_other_width = other_total_width / other_cols_count
            gt_col_widths = [new_other_width if i != desc_actual_idx else desc_col_width for i in range(num_cols)]
        else: gt_col_widths = [total_page_width]
    except ValueError:
        gt_col_widths = [total_page_width / num_cols] * num_cols if num_cols > 0 else []

    if gt_col_widths:
        gt_row_data = [ Paragraph("Grand Total", bl_style) ] + \
                      [ Paragraph("", body_style) for _ in range(num_cols - 2) ] + \
                      [ Paragraph(html.escape(grand_total), br_style) if num_cols > 1 else Paragraph(html.escape(grand_total), bl_style) ]
        if num_cols == 1: gt_row_data = [Paragraph(f"Grand Total: {html.escape(grand_total)}", bl_style)]
        gt_row_data += [Paragraph("", body_style)] * (num_cols - len(gt_row_data))

        gt_table = LongTable([gt_row_data], colWidths=gt_col_widths)
        gt_style_cmds = [
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0")),
        ]
        if num_cols > 1:
             gt_style_cmds.extend([('SPAN', (0, 0), (-2, 0)), ('ALIGN', (-1, 0), (-1, 0), 'RIGHT')])
        else:
             gt_style_cmds.append(('ALIGN', (0,0), (0,0), 'LEFT'))
        gt_table.setStyle(TableStyle(gt_style_cmds))
        elements.append(gt_table)

try:
    doc.build(elements)
    pdf_buf.seek(0)
except Exception as e:
    st.error(f"Error building PDF: {e}")
    # import traceback
    # st.error(traceback.format_exc())
    pdf_buf = None # Indicate failure

# ─── Build Word ───────────────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document() # Renamed from 'docx' to avoid conflict with import
sec = docx_doc.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
# Set page size explicitly for landscape 11x17
sec.page_height = Inches(11) # Height for landscape
sec.page_width = Inches(17)  # Width for landscape
sec.left_margin = Inches(0.5)
sec.right_margin = Inches(0.5)
sec.top_margin = Inches(0.5)
sec.bottom_margin = Inches(0.5)

# Add Logo
if logo:
    try:
        p_logo = docx_doc.add_paragraph()
        r_logo = p_logo.add_run()
        # Calculate width maintaining aspect ratio for Word, similar to PDF
        img = Image.open(io.BytesIO(logo))
        ratio = img.height / img.width
        img_width_in = 5 # Max width 5 inches
        img_height_in = img_width_in * ratio
        r_logo.add_picture(io.BytesIO(logo), width=Inches(img_width_in))
        p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER # Use enum for alignment
    except Exception as e:
        st.warning(f"Could not add logo to Word: {e}")

# Add Title
p_title = docx_doc.add_paragraph()
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r_title = p_title.add_run(proposal_title)
r_title.font.name = DEFAULT_SERIF_FONT
r_title.font.size = Pt(18)
r_title.bold = True
docx_doc.add_paragraph() # Spacer paragraph

# Define total usable width for Word table calculations
TOTAL_W_INCHES = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for hdr, rows, row_links_list, table_total_info in tables_info: # Use table_total_info
    n = len(hdr)
    if n == 0: continue # Skip if header is empty

    # Recalculate widths for Word
    try:
        desc_actual_idx = hdr.index("Description")
        desc_w_in = 0.45 * TOTAL_W_INCHES
        other_cols_count = n - 1
        other_w_in = (TOTAL_W_INCHES - desc_w_in) / other_cols_count if other_cols_count > 0 else 0
    except ValueError:
        desc_actual_idx = 1 if n > 1 else 0
        desc_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES
        other_w_in = desc_w_in

    # Create table
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.allow_autofit = False

    # Set Preferred Table Width
    if not hasattr(tbl, '_tblPr'): tbl._element.append(OxmlElement('w:tblPr'))
    tblPr = tbl._tblPr
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    # Remove existing tblW if it exists before appending
    for existing_tblW in tblPr.xpath('w:tblW'):
         tblPr.remove(existing_tblW)
    tblPr.append(tblW)

    # Set column widths
    for idx, col in enumerate(tbl.columns):
        width_val = desc_w_in if idx == desc_actual_idx else other_w_in
        col.width = Inches(max(0.1, width_val))

    # Populate header row
    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr):
        cell = hdr_cells[i]
        tc = cell._tc; tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), 'F2F2F2'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
        tcPr.append(shd)
        p = cell.paragraphs[0]; p.text = ""
        run = p.add_run(str(col_name))
        run.font.name = DEFAULT_SERIF_FONT; run.font.size = Pt(10); run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Populate data rows
    for ridx, row in enumerate(rows):
        row_cells = tbl.add_row().cells
        for cidx, val in enumerate(row):
            cell = row_cells[cidx]
            p = cell.paragraphs[0]; p.text = ""
            cell_str = str(val)
            if cidx == desc_actual_idx and ridx < len(row_links_list) and row_links_list[ridx]:
                add_hyperlink(p, row_links_list[ridx], cell_str, font_name=DEFAULT_SANS_FONT, font_size=9)
            else:
                run = p.add_run(cell_str)
                run.font.name = DEFAULT_SANS_FONT
                run.font.size = Pt(9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    # --- START MODIFIED WORD TOTAL ROW HANDLING ---
    if table_total_info:
        label = "Total" # Defaults
        amount = ""

        if isinstance(table_total_info, list):
            # Found total row in extracted data (list of strings)
            label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
            # Find the last element that contains a '$' as the likely value
            amount = next((val.strip() for val in reversed(table_total_info) if '$' in str(val)), "")
            if not amount and len(table_total_info) > 1: # Fallback to last element if no '$' found
                 amount = table_total_info[-1].strip()

        elif isinstance(table_total_info, str):
            # Got total from find_total (string) - parse it
            total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
            if total_match:
                label_parsed, amount = total_match.groups()
                label = label_parsed.strip() if label_parsed.strip() else "Total" # Use parsed label only if it's not empty
                amount = amount.strip()
            else: # Try parsing just value if pattern fails
                 amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                 if amount_match:
                      amount = amount_match.group(1).strip()
                      potential_label = table_total_info[:amount_match.start()].strip()
                      label = potential_label if potential_label else "Total"
                 else: # Cannot reliably parse
                      amount = table_total_info # Display raw string as value

        # Add the total row to the Word table
        total_cells = tbl.add_row().cells
        label_cell = total_cells[0]
        if n > 1: # Merge if more than one column
             label_cell.merge(total_cells[n-2])
        p_label = label_cell.paragraphs[0]
        p_label.text = "" # Clear text
        run_label = p_label.add_run(label)
        run_label.font.name = DEFAULT_SERIF_FONT
        run_label.font.size = Pt(10)
        run_label.bold = True
        p_label.alignment = WD_TABLE_ALIGNMENT.LEFT
        label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        # Add the amount in the last cell
        if n > 0:
            amount_cell = total_cells[n-1]
            p_amount = amount_cell.paragraphs[0]
            p_amount.text = "" # Clear text
            run_amount = p_amount.add_run(amount)
            run_amount.font.name = DEFAULT_SERIF_FONT
            run_amount.font.size = Pt(10)
            run_amount.bold = True
            p_amount.alignment = WD_TABLE_ALIGNMENT.RIGHT
            amount_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
     # --- END MODIFIED WORD TOTAL ROW HANDLING ---

    docx_doc.add_paragraph() # Spacer after table

# Add Grand Total row if present (No change needed here)
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]
    n = len(last_hdr)
    if n > 0:
        try:
            desc_actual_idx = last_hdr.index("Description")
            desc_w_in = 0.45 * TOTAL_W_INCHES
            other_cols_count = n - 1
            other_w_in = (TOTAL_W_INCHES - desc_w_in) / other_cols_count if other_cols_count > 0 else 0
        except ValueError:
            desc_actual_idx = 1 if n > 1 else 0
            desc_w_in = TOTAL_W_INCHES / n
            other_w_in = desc_w_in

        tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
        tblg.alignment = WD_TABLE_ALIGNMENT.CENTER
        tblg.autofit = False
        tblg.allow_autofit = False

        # Set Preferred Table Width for Grand Total Table
        if not hasattr(tblg, '_tblPr'): tblg._element.append(OxmlElement('w:tblPr'))
        tblgPr = tblg._tblPr
        tblgW = OxmlElement('w:tblW')
        tblgW.set(qn('w:w'), '5000')
        tblgW.set(qn('w:type'), 'pct')
        # Remove existing tblW if it exists before appending
        for existing_tblW in tblgPr.xpath('w:tblW'):
             tblgPr.remove(existing_tblW)
        tblgPr.append(tblgW)


        # Set column widths for GT table
        for idx, col in enumerate(tblg.columns):
            width_val = desc_w_in if idx == desc_actual_idx else other_w_in
            col.width = Inches(max(0.1, width_val))

        gt_cells = tblg.rows[0].cells
        gt_label_cell = gt_cells[0]
        if n > 1: gt_label_cell.merge(gt_cells[n-2])

        tc_label = gt_label_cell._tc; tcPr_label = tc_label.get_or_add_tcPr()
        shd_label = OxmlElement('w:shd'); shd_label.set(qn('w:fill'), 'E0E0E0'); tcPr_label.append(shd_label)
        p_gt_label = gt_label_cell.paragraphs[0]; p_gt_label.text = ""
        run_gt_label = p_gt_label.add_run("Grand Total")
        run_gt_label.font.name = DEFAULT_SERIF_FONT; run_gt_label.font.size = Pt(10); run_gt_label.bold = True
        p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT
        gt_label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        if n > 0:
            gt_value_cell = gt_cells[n-1]
            tc_val = gt_value_cell._tc; tcPr_val = tc_val.get_or_add_tcPr()
            shd_val = OxmlElement('w:shd'); shd_val.set(qn('w:fill'), 'E0E0E0'); tcPr_val.append(shd_val)
            p_gt_val = gt_value_cell.paragraphs[0]; p_gt_val.text = ""
            run_gt_val = p_gt_val.add_run(grand_total)
            run_gt_val.font.name = DEFAULT_SERIF_FONT; run_gt_val.font.size = Pt(10); run_gt_val.bold = True
            p_gt_val.alignment = WD_TABLE_ALIGNMENT.RIGHT
            gt_value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except Exception as e:
    st.error(f"Error building Word document: {e}")
    # import traceback
    # st.error(traceback.format_exc())
    docx_buf = None # Indicate failure

# Download buttons (only show if buffer exists)
c1, c2 = st.columns(2)
if pdf_buf:
    with c1:
        st.download_button(
            "📥 Download deliverable PDF",
            data=pdf_buf,
            file_name="proposal_deliverable.pdf",
            mime="application/pdf",
            use_container_width=True
        )
else:
     with c1:
         st.error("PDF generation failed.")

if docx_buf:
    with c2:
        st.download_button(
            "📥 Download deliverable DOCX",
            data=docx_buf,
            file_name="proposal_deliverable.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
else:
    with c2:
        st.error("Word document generation failed.")
