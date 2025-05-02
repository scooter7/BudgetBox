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


# Extract tables & totals, capturing hyperlink per row
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

            for tbl in page_tables:
                data = tbl.extract(x_tolerance=1, y_tolerance=1)
                if not data or len(data) < 2: continue

                hdr = [(str(h).strip() if h else "") for h in data[0]]
                desc_i = next((i for i, h in enumerate(hdr) if h and "description" in h.lower()), None)
                if desc_i is None:
                    desc_i = next((i for i, h in enumerate(hdr) if len(h) > 10), None)
                    if desc_i is None or len(hdr) <= 1 : continue

                # --- Simplified Link Finding: Just check if *any* link overlaps description cell ---
                desc_links_uri = {} # Stores {row_index: uri}
                if hasattr(tbl, 'rows'):
                    for r, row_obj in enumerate(tbl.rows):
                        if r == 0: continue # Skip header
                        if hasattr(row_obj, 'cells') and desc_i is not None and desc_i < len(row_obj.cells):
                             cell_bbox = row_obj.cells[desc_i]
                             if not cell_bbox: continue
                             cell_x0, cell_top, cell_x1, cell_bottom = cell_bbox

                             for link in links:
                                if not all(k in link for k in ['x0', 'x1', 'top', 'bottom', 'uri']): continue
                                link_x0, link_top, link_x1, link_bottom = link['x0'], link['top'], link['x1'], link['bottom']
                                # Simple overlap check
                                if not (link_x1 < cell_x0 or link_x0 > cell_x1 or link_bottom < cell_top or link_top > cell_bottom):
                                     # Found an overlapping link, store its URI for this row index 'r'
                                     desc_links_uri[r] = link.get("uri")
                                     break # Stop checking other links for *this* cell

                # --- Row Processing & Total Finding (Reverted to previous working logic) ---
                new_hdr = ["Strategy", "Description"] + [h for i, h in enumerate(hdr) if i != desc_i and h]
                rows_data = [] # Stores processed rows [['strat', 'desc', ...], ...]
                row_links_uri_list = [] # Stores URI or None for each row in rows_data
                table_total_info = None # Stores total row info (list or string)

                for ridx_data, row_content in enumerate(data[1:], start=1): # ridx_data is 1-based index within PDF table (incl. header)
                    row_str_list = [(str(cell).strip() if cell else "") for cell in row_content]
                    if all(not cell_val for cell_val in row_str_list): continue # Skip empty rows

                    # Check for total row
                    first_cell_lower = row_str_list[0].lower() if row_str_list else ""
                    if ("total" in first_cell_lower or "subtotal" in first_cell_lower) and \
                       any("$" in str(cell_val) for cell_val in row_str_list):
                        if table_total_info is None: table_total_info = row_str_list
                        continue # Skip adding total rows to rows_data

                    # Process normal data row
                    desc_text_from_data = row_str_list[desc_i] if desc_i < len(row_str_list) else ""
                    strat, desc = split_cell_text(desc_text_from_data)
                    rest_cols = [row_str_list[i] for i, h in enumerate(hdr) if i != desc_i and h and i < len(row_str_list)]
                    rows_data.append([strat, desc] + rest_cols) # Add the processed row data

                    # Get the link URI for this specific row using the map created earlier
                    row_links_uri_list.append(desc_links_uri.get(ridx_data)) # Append URI or None

                # Fallback for table total
                if table_total_info is None:
                    table_total_info = find_total(pi)

                if rows_data: # Only store table if it has valid data rows
                     tables_info.append((new_hdr, rows_data, row_links_uri_list, table_total_info))
                 # --- End Row Processing ---

        # Find Grand total robustly
        for tx in reversed(page_texts):
            m = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', tx, re.I | re.S)
            if m: grand_total = m.group(1).replace(" ", ""); break

except Exception as e:
    st.error(f"Error processing PDF: {e}")
    import traceback
    st.error(traceback.format_exc())
    st.stop()


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
for hdr, rows_data, row_links_uri_list, table_total_info in tables_info: # Use correct variable names
    num_cols = len(hdr)
    # Column widths (unchanged)
    desc_col_width = total_page_width * 0.45
    other_col_width = (total_page_width - desc_col_width) / (num_cols - 1) if num_cols > 1 else 0
    col_widths = [other_col_width] * num_cols
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        col_widths[desc_actual_idx_in_hdr] = desc_col_width
        other_total_width = total_page_width - desc_col_width; other_cols_count = num_cols - 1
        if other_cols_count > 0:
            new_other_width = other_total_width / other_cols_count
            col_widths = [new_other_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
    except ValueError:
         desc_actual_idx_in_hdr = 1
         col_widths = [total_page_width / num_cols] * num_cols

    # Wrap header text
    wrapped_header = [Paragraph(html.escape(str(h)), header_style) for h in hdr]
    wrapped_data = [wrapped_header]

    # --- START: Process Data Rows for PDF Output (Simpler "- link" approach) ---
    for ridx, row in enumerate(rows_data):
        line = []
        for cidx, cell_content in enumerate(row):
            cell_str = str(cell_content)
            escaped_cell_text = html.escape(cell_str)

            # Check if this is the description column (index 1 in rows_data)
            # AND if a link URI exists for this row
            # Use generated header index desc_actual_idx_in_hdr for check
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                # Append "- link" and make only that part the link
                paragraph_text = f"{escaped_cell_text} <link href='{html.escape(link_uri)}' color='blue'>- link</link>"
                p = Paragraph(paragraph_text, body_style) # Apply base style
            else:
                # No link for this cell, or not the description column
                p = Paragraph(escaped_cell_text, body_style)

            line.append(p)
        wrapped_data.append(line)
    # --- END: Process Data Rows for PDF Output ---

    # Add Table Total Row (unchanged from previous working version)
    has_total_row = False
    if table_total_info:
        label = "Total"; value = ""
        if isinstance(table_total_info, list):
             label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
             value = next((val.strip() for val in reversed(table_total_info) if '$' in str(val)), "")
             if not value and len(table_total_info) > 1: value = table_total_info[-1].strip()
        elif isinstance(table_total_info, str):
             total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
             if total_match: label_parsed, value = total_match.groups(); label = label_parsed.strip() if label_parsed.strip() else "Total"; value = value.strip()
             else: amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info);
             if amount_match: value = amount_match.group(1).strip(); potential_label = table_total_info[:amount_match.start()].strip(); label = potential_label if potential_label else "Total"
             else: value = table_total_info
        if num_cols > 0:
            total_row_elements = [Paragraph(html.escape(label), bl_style)] + \
                                 [Paragraph("", body_style)] * (num_cols - 2) + \
                                 [Paragraph(html.escape(value), br_style) if num_cols > 1 else Paragraph(html.escape(value), bl_style)]
            total_row_elements += [Paragraph("", body_style)] * (num_cols - len(total_row_elements))
            wrapped_data.append(total_row_elements)
            has_total_row = True

    # Create and style PDF table (unchanged)
    if wrapped_data and col_widths:
        tbl = LongTable(wrapped_data, colWidths=col_widths, repeatRows=1)
        style_commands = [ # Base styles
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ]
        if has_total_row and num_cols > 1: # Styles for total row
             style_commands.extend([('SPAN', (0, -1), (-2, -1)),('ALIGN', (0, -1), (-2, -1), 'LEFT'),('ALIGN', (-1, -1), (-1, -1), 'RIGHT'),('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
        elif has_total_row and num_cols == 1:
              style_commands.extend([('ALIGN', (0, -1), (0, -1), 'LEFT'),('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
        tbl.setStyle(TableStyle(style_commands))
        elements += [tbl, Spacer(1, 24)]

# Add Grand Total row (unchanged)
if grand_total and tables_info:
    # ... (Grand total logic remains the same) ...
    last_hdr, _, _, _ = tables_info[-1]; num_cols = len(last_hdr)
    desc_col_width = total_page_width * 0.45
    try:
        desc_actual_idx_in_hdr = last_hdr.index("Description")
        if num_cols > 1: other_total_width = total_page_width - desc_col_width; other_cols_count = num_cols - 1; new_other_width = other_total_width / other_cols_count; gt_col_widths = [new_other_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
        else: gt_col_widths = [total_page_width]
    except ValueError: gt_col_widths = [total_page_width / num_cols] * num_cols if num_cols > 0 else []
    if gt_col_widths:
        gt_row_data = [ Paragraph("Grand Total", bl_style) ] + [ Paragraph("", body_style) for _ in range(num_cols - 2) ] + [ Paragraph(html.escape(grand_total), br_style) if num_cols > 1 else Paragraph(html.escape(grand_total), bl_style) ]
        if num_cols == 1: gt_row_data = [Paragraph(f"Grand Total: {html.escape(grand_total)}", bl_style)]
        gt_row_data += [Paragraph("", body_style)] * (num_cols - len(gt_row_data))
        gt_table = LongTable([gt_row_data], colWidths=gt_col_widths)
        gt_style_cmds = [("GRID", (0, 0), (-1, -1), 0.25, colors.grey), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0"))]
        if num_cols > 1: gt_style_cmds.extend([('SPAN', (0, 0), (-2, 0)), ('ALIGN', (-1, 0), (-1, 0), 'RIGHT')])
        else: gt_style_cmds.append(('ALIGN', (0,0), (0,0), 'LEFT'))
        gt_table.setStyle(TableStyle(gt_style_cmds)); elements.append(gt_table)

# Build PDF (unchanged)
try:
    doc.build(elements); pdf_buf.seek(0)
except Exception as e: st.error(f"Error building PDF: {e}"); import traceback; st.error(traceback.format_exc()); pdf_buf = None


# ─── Build Word ───────────────────────────────────────────────────────────────
docx_buf = io.BytesIO()
docx_doc = Document()
# Page Setup (unchanged)
sec = docx_doc.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_height = Inches(11); sec.page_width = Inches(17)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5); sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
# Logo and Title (unchanged)
if logo:
    try: p_logo = docx_doc.add_paragraph(); r_logo = p_logo.add_run(); img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width; img_width_in = 5; img_height_in = img_width_in * ratio; r_logo.add_picture(io.BytesIO(logo), width=Inches(img_width_in)); p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
    except Exception as e: st.warning(f"Could not add logo to Word: {e}")
p_title = docx_doc.add_paragraph(); p_title.alignment = WD_TABLE_ALIGNMENT.CENTER; r_title = p_title.add_run(proposal_title); r_title.font.name = DEFAULT_SERIF_FONT; r_title.font.size = Pt(18); r_title.bold = True; docx_doc.add_paragraph()

TOTAL_W_INCHES = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

# --- Iterate through processed table data ---
for hdr, rows_data, row_links_uri_list, table_total_info in tables_info: # Use correct variable names
    n = len(hdr)
    if n == 0: continue
    # Column widths (unchanged)
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_w_in = 0.45 * TOTAL_W_INCHES; other_cols_count = n - 1; other_w_in = (TOTAL_W_INCHES - desc_w_in) / other_cols_count if other_cols_count > 0 else 0
    except ValueError: desc_actual_idx_in_hdr = 1 if n > 1 else 0; desc_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES; other_w_in = desc_w_in
    # Create table
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid"); tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False; tbl.allow_autofit = False

    # --- START: Set Preferred Table Width (Revised Access) ---
    tblPr = tbl._element.get_or_add_tblPr() # Use get_or_add_tblPr()
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '5000')
    tblW.set(qn('w:type'), 'pct')
    # Remove existing tblW before appending the new one
    existing_tblW = tblPr.xpath('./w:tblW') # Use XPath relative to tblPr
    if existing_tblW:
        tblPr.remove(existing_tblW[0])
    tblPr.append(tblW)
    # --- END: Set Preferred Table Width ---

    # Set Column Widths (unchanged)
    for idx, col in enumerate(tbl.columns): width_val = desc_w_in if idx == desc_actual_idx_in_hdr else other_w_in; col.width = Inches(max(0.1, width_val))
    # Populate Header Row (unchanged)
    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr): cell = hdr_cells[i]; tc = cell._tc; tcPr = tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'F2F2F2'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); tcPr.append(shd); p = cell.paragraphs[0]; p.text = ""; run = p.add_run(str(col_name)); run.font.name = DEFAULT_SERIF_FONT; run.font.size = Pt(10); run.bold = True; p.alignment = WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # --- START: Populate Data Rows for Word Output (Simpler "- link" approach) ---
    for ridx, row in enumerate(rows_data):
        row_cells = tbl.add_row().cells
        for cidx, cell_content in enumerate(row):
            cell = row_cells[cidx]
            p = cell.paragraphs[0]
            p.text = "" # Clear paragraph content first
            cell_str = str(cell_content)

            # Add the main text content as a normal run first
            run_text = p.add_run(cell_str)
            run_text.font.name = DEFAULT_SANS_FONT
            run_text.font.size = Pt(9)

            # Check if this is the description column (use generated index) AND if a link URI exists
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                # Add a space before the link text for separation
                space_run = p.add_run(" ")
                space_run.font.name = DEFAULT_SANS_FONT # Match font if needed
                space_run.font.size = Pt(9)
                # Use the helper function to add ONLY the "- link" hyperlinked text
                add_hyperlink(p, link_uri, "- link", font_name=DEFAULT_SANS_FONT, font_size=9)
            # Else: No link, the run_text added above is sufficient

            # Set alignment and vertical alignment for the cell
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    # --- END: Populate Data Rows for Word Output ---

    # --- Add Table Total Row (logic unchanged) ---
    if table_total_info:
        label = "Total"; amount = ""
        if isinstance(table_total_info, list): label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"; amount = next((val.strip() for val in reversed(table_total_info) if '$' in str(val)), "");
        if not amount and len(table_total_info) > 1: amount = table_total_info[-1].strip()
        elif isinstance(table_total_info, str): total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info);
        if total_match: label_parsed, amount = total_match.groups(); label = label_parsed.strip() if label_parsed.strip() else "Total"; amount = amount.strip();
        else: amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info);
        if amount_match: amount = amount_match.group(1).strip(); potential_label = table_total_info[:amount_match.start()].strip(); label = potential_label if potential_label else "Total";
        else: amount = table_total_info;
        total_cells = tbl.add_row().cells; label_cell = total_cells[0];
        if n > 1: label_cell.merge(total_cells[n-2]);
        p_label = label_cell.paragraphs[0]; p_label.text = ""; run_label = p_label.add_run(label); run_label.font.name = DEFAULT_SERIF_FONT; run_label.font.size = Pt(10); run_label.bold = True; p_label.alignment = WD_TABLE_ALIGNMENT.LEFT; label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
        if n > 0: amount_cell = total_cells[n-1]; p_amount = amount_cell.paragraphs[0]; p_amount.text = ""; run_amount = p_amount.add_run(amount); run_amount.font.name = DEFAULT_SERIF_FONT; run_amount.font.size = Pt(10); run_amount.bold = True; p_amount.alignment = WD_TABLE_ALIGNMENT.RIGHT; amount_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
    # --- End Table Total Row ---

    docx_doc.add_paragraph() # Spacer

# --- Add Grand Total row (logic mostly unchanged) ---
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]; n = len(last_hdr)
    if n > 0:
        try: desc_actual_idx_in_hdr = last_hdr.index("Description"); desc_w_in = 0.45 * TOTAL_W_INCHES; other_cols_count = n - 1; other_w_in = (TOTAL_W_INCHES - desc_w_in) / other_cols_count if other_cols_count > 0 else 0;
        except ValueError: desc_actual_idx_in_hdr = 1 if n > 1 else 0; desc_w_in = TOTAL_W_INCHES / n; other_w_in = desc_w_in;
        tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid"); tblg.alignment = WD_TABLE_ALIGNMENT.CENTER; tblg.autofit = False; tblg.allow_autofit = False;

        # --- START: Set Preferred Table Width for Grand Total Table (Revised Access) ---
        tblgPr = tblg._element.get_or_add_tblPr() # Use get_or_add_tblPr()
        tblgW = OxmlElement('w:tblW')
        tblgW.set(qn('w:w'), '5000')
        tblgW.set(qn('w:type'), 'pct')
        # Remove existing tblW before appending the new one
        existing_tblgW_gt = tblgPr.xpath('./w:tblW') # Use XPath relative to tblgPr
        if existing_tblgW_gt:
            tblgPr.remove(existing_tblgW_gt[0])
        tblgPr.append(tblgW)
        # --- END: Set Preferred Table Width ---

        # Set column widths for GT table (unchanged)
        for idx, col in enumerate(tblg.columns): width_val = desc_w_in if idx == desc_actual_idx_in_hdr else other_w_in; col.width = Inches(max(0.1, width_val));
        # Populate GT row (unchanged)
        gt_cells = tblg.rows[0].cells; gt_label_cell = gt_cells[0];
        if n > 1: gt_label_cell.merge(gt_cells[n-2]);
        tc_label = gt_label_cell._tc; tcPr_label = tc_label.get_or_add_tcPr(); shd_label = OxmlElement('w:shd'); shd_label.set(qn('w:fill'), 'E0E0E0'); tcPr_label.append(shd_label); p_gt_label = gt_label_cell.paragraphs[0]; p_gt_label.text = ""; run_gt_label = p_gt_label.add_run("Grand Total"); run_gt_label.font.name = DEFAULT_SERIF_FONT; run_gt_label.font.size = Pt(10); run_gt_label.bold = True; p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT; gt_label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
        if n > 0: gt_value_cell = gt_cells[n-1]; tc_val = gt_value_cell._tc; tcPr_val = tc_val.get_or_add_tcPr(); shd_val = OxmlElement('w:shd'); shd_val.set(qn('w:fill'), 'E0E0E0'); tcPr_val.append(shd_val); p_gt_val = gt_value_cell.paragraphs[0]; p_gt_val.text = ""; run_gt_val = p_gt_val.add_run(grand_total); run_gt_val.font.name = DEFAULT_SERIF_FONT; run_gt_val.font.size = Pt(10); run_gt_val.bold = True; p_gt_val.alignment = WD_TABLE_ALIGNMENT.RIGHT; gt_value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
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
