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
    # Adjust paths if your fonts are located elsewhere
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except Exception as e:
    st.warning(f"Could not load custom fonts: {e}. Using default system fonts.")
    DEFAULT_SERIF_FONT = "Times New Roman" # Fallback
    DEFAULT_SANS_FONT = "Arial" # Fallback


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
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    return lines[0], description

# --- Word hyperlink helper (Unchanged) ---
def add_hyperlink(paragraph, url, text, font_name=None, font_size=None, bold=None):
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


# === START: PDF TABLE EXTRACTION AND PROCESSING LOGIC ===
# --- Added Table Settings for Extraction ---
TABLE_EXTRACTION_SETTINGS = {
    "vertical_strategy": "text",    # Try 'text' alignment first
    "horizontal_strategy": "text",  # Try 'text' alignment first
    "intersection_x_tolerance": 3,  # Default is 3
    "intersection_y_tolerance": 3,  # Default is 3
    "snap_x_tolerance": 3,          # Default is 3
    "snap_y_tolerance": 3,          # Default is 3
    "join_x_tolerance": 3,          # Default is 3
    "join_y_tolerance": 3,          # Default is 3
    "edge_min_length": 3,           # Default is 3
    "min_words_vertical": 3,        # Default is 3
    "min_words_horizontal": 1,      # Default is 1
    "text_x_tolerance": 3,          # Increased tolerance might help merge text
    "text_y_tolerance": 3,
}
# --- Fallback settings if 'text' strategy fails ---
TABLE_EXTRACTION_SETTINGS_FALLBACK = {
    "vertical_strategy": "lines",    # Fallback to lines if text doesn't work
    "horizontal_strategy": "lines",  # Fallback to lines
    "intersection_x_tolerance": 5,   # Slightly increase tolerance for lines
    "snap_x_tolerance": 5,
    "join_x_tolerance": 5,
    # Keep others potentially default or as needed
}


tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]

        first_page_lines = page_texts[0].splitlines() if page_texts else []
        potential_title = next((line.strip() for line in first_page_lines if "proposal" in line.lower() and len(line.strip()) > 5), None)
        if potential_title: proposal_title = potential_title
        elif len(first_page_lines) > 0: proposal_title = first_page_lines[0].strip()

        used_totals = set()
        def find_total(pi):
            if pi >= len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                # Improved regex: looks for total/subtotal NOT preceded by 'grand', then optional chars, then a $ amount
                if re.search(r'\b(?<!Grand\s)(?:Total|Subtotal)\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            links = page.hyperlinks

            # --- Use Defined Table Settings ---
            page_tables = page.find_tables(table_settings=TABLE_EXTRACTION_SETTINGS)
            # --- If first strategy finds nothing, try fallback ---
            if not page_tables:
                page_tables = page.find_tables(table_settings=TABLE_EXTRACTION_SETTINGS_FALLBACK)


            if not page_tables: continue

            for tbl_idx, tbl in enumerate(page_tables):
                 # --- Use tighter tolerances for data extraction within found table ---
                data = tbl.extract(x_tolerance=1, y_tolerance=1)

                if not data or len(data) < 2: continue

                original_hdr = [(str(h).strip() if h is not None else "") for h in data[0]]
                if not any(original_hdr): continue

                original_desc_idx = -1
                # Prioritize "Description" header name
                for i, h in enumerate(original_hdr):
                    if h and "description" in h.lower():
                        original_desc_idx = i
                        break
                # Fallback: Find the first reasonably wide column OR common alternatives
                if original_desc_idx == -1:
                     common_desc_headers = ["details", "summary", "notes", "content"]
                     found_common = False
                     for i, h in enumerate(original_hdr):
                          if h:
                              for common_hdr in common_desc_headers:
                                   if common_hdr in h.lower():
                                       original_desc_idx = i
                                       found_common = True
                                       break
                          if found_common: break
                     # If still not found, use width heuristic (e.g., first col wider than 10 chars)
                     if original_desc_idx == -1:
                          for i, h in enumerate(original_hdr):
                             if h and len(h) > 10:
                                 original_desc_idx = i
                                 break

                if original_desc_idx == -1:
                    # st.warning(f"Skipping table {tbl_idx+1} on page {pi+1}: Could not identify a description column. Header: {original_hdr}")
                    continue # Cannot proceed without knowing which column to split

                # --- Link Finding ---
                desc_links_uri = {}
                if hasattr(tbl, 'rows'):
                    for r, row_obj in enumerate(tbl.rows):
                        if r == 0: continue
                        if hasattr(row_obj, 'cells') and original_desc_idx < len(row_obj.cells):
                            cell_bbox = row_obj.cells[original_desc_idx]
                            if not cell_bbox: continue
                            cell_x0, cell_top, cell_x1, cell_bottom = cell_bbox

                            for link in links:
                                if not all(k in link for k in ['x0', 'x1', 'top', 'bottom', 'uri']): continue
                                link_x0, link_top, link_x1, link_bottom = link['x0'], link['top'], link['x1'], link['bottom']
                                x_overlap = (link_x0 < cell_x1) and (link_x1 > cell_x0)
                                y_overlap = (link_top < cell_bottom) and (link_bottom > cell_top)
                                if x_overlap and y_overlap:
                                    desc_links_uri[r] = link.get("uri")
                                    break

                # --- Header and Row Processing (LOGIC UNCHANGED from previous correct version) ---
                new_hdr = []
                valid_original_indices = [] # Keep track of which original columns are kept
                for i, h in enumerate(original_hdr):
                    if i == original_desc_idx:
                        new_hdr.extend(["Strategy", "Description"])
                        # Mark original desc index as processed, but don't add to valid_original_indices
                    elif h: # Only keep non-empty original headers
                        new_hdr.append(h)
                        valid_original_indices.append(i) # Store index of kept original column

                if not new_hdr: continue # Skip if header becomes empty

                rows_data = []
                row_links_uri_list = []
                table_total_info = None

                for ridx_pdf, row_content in enumerate(data[1:], start=1):
                    row_str_list = [(str(cell).strip() if cell is not None else "") for cell in row_content]
                    if all(not cell_val for cell_val in row_str_list): continue

                    first_cell_lower = row_str_list[0].lower() if row_str_list else ""
                    is_total_row = (("total" in first_cell_lower or "subtotal" in first_cell_lower) and \
                                   any(re.search(r'\$|€|£|¥', str(cell_val)) for cell_val in row_str_list if cell_val)) or \
                                   (len(row_str_list)>1 and "total" in row_str_list[-1].lower())

                    if is_total_row:
                        if table_total_info is None: table_total_info = row_str_list
                        continue

                    desc_text_from_pdf = row_str_list[original_desc_idx] if original_desc_idx < len(row_str_list) else ""
                    strat, desc = split_cell_text(desc_text_from_pdf)

                    new_row_content = []
                    original_cell_idx = 0
                    processed_original = False # Flag to check if desc column was handled
                    for i in range(len(original_hdr)): # Iterate based on original header structure
                        if i == original_desc_idx:
                            new_row_content.extend([strat, desc])
                            processed_original = True
                        elif original_hdr[i]: # Was this a valid original header?
                            # Get corresponding cell value if it exists
                            cell_val = row_str_list[i] if i < len(row_str_list) else ""
                            new_row_content.append(cell_val)


                    # Pad/truncate to match new_hdr length if necessary
                    expected_cols = len(new_hdr)
                    current_cols = len(new_row_content)
                    if current_cols < expected_cols:
                       new_row_content.extend([""] * (expected_cols - current_cols))
                    elif current_cols > expected_cols:
                       new_row_content = new_row_content[:expected_cols]

                    rows_data.append(new_row_content)
                    row_links_uri_list.append(desc_links_uri.get(ridx_pdf))

                if table_total_info is None:
                    table_total_info = find_total(pi)

                if rows_data:
                    tables_info.append((new_hdr, rows_data, row_links_uri_list, table_total_info))


        # Find Grand total robustly
        for tx in reversed(page_texts):
            m = re.search(r'Grand\s+Total.*?(?<!Subtotal\s)(?<!Sub Total\s)(\$\s*[\d,]+\.\d{2})', tx, re.I | re.S)
            if m:
                grand_total_candidate = m.group(1).replace(" ", "")
                if "subtotal" not in m.group(0).lower():
                    grand_total = grand_total_candidate
                    break

except Exception as e:
    st.error(f"Error processing PDF: {e}")
    import traceback
    st.error(traceback.format_exc())
    st.stop()
# === END: PDF TABLE EXTRACTION AND PROCESSING LOGIC ===


# === PDF Building Section (No changes needed here, uses tables_info) ===
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(
    pdf_buf, pagesize=landscape((17*inch, 11*inch)),
    leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch
)
title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11)
link_style   = ParagraphStyle("LinkStyle", parent=body_style, textColor=colors.blue)
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, textColor=colors.black, spaceBefore=6)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, textColor=colors.black, spaceBefore=6)

elements = []
logo = None
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

for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info):
    num_cols = len(hdr)
    if num_cols == 0: continue

    col_widths = []
    desc_actual_idx_in_hdr = -1
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_col_width = total_page_width * 0.45
        other_cols_count = num_cols - 1
        if other_cols_count > 0:
            other_total_width = total_page_width - desc_col_width
            strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1
            if strategy_idx != -1:
                 strat_width = total_page_width * 0.15
                 remaining_width = other_total_width - strat_width
                 remaining_cols = other_cols_count - 1
                 other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0
                 col_widths = []
                 for i in range(num_cols):
                     if i == desc_actual_idx_in_hdr: col_widths.append(desc_col_width)
                     elif i == strategy_idx: col_widths.append(strat_width)
                     else: col_widths.append(max(0.1*inch, other_indiv_width))
            else:
                 other_col_width = other_total_width / other_cols_count
                 col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
        elif num_cols == 1: col_widths = [total_page_width]
        else: col_widths = [total_page_width / num_cols] * num_cols
    except ValueError:
        desc_actual_idx_in_hdr = -1
        if num_cols > 0: col_widths = [total_page_width / num_cols] * num_cols
        else: continue

    wrapped_header = [Paragraph(html.escape(str(h)), header_style) for h in hdr]
    wrapped_data = [wrapped_header]

    for ridx, row in enumerate(rows_data):
        line = []
        current_cells = len(row)
        if current_cells < num_cols: row = list(row) + [""] * (num_cols - current_cells)
        elif current_cells > num_cols: row = row[:num_cols]

        for cidx, cell_content in enumerate(row):
            cell_str = str(cell_content)
            escaped_cell_text = html.escape(cell_str)
            link_applied = False
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                if link_uri:
                    paragraph_text = f"{escaped_cell_text} <link href='{html.escape(link_uri)}' color='blue'>- link</link>"
                    p = Paragraph(paragraph_text, body_style)
                    link_applied = True
            if not link_applied:
                p = Paragraph(escaped_cell_text, body_style)
            line.append(p)
        wrapped_data.append(line)

    has_total_row = False
    if table_total_info:
        label = "Total"; value = ""
        if isinstance(table_total_info, list):
             label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"
             value = next((val.strip() for val in reversed(table_total_info) if val and '$' in str(val)), "")
             if not value and len(table_total_info) > 1: value = table_total_info[-1].strip() if table_total_info[-1] else ""
        elif isinstance(table_total_info, str):
             total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
             if total_match: label_parsed, value = total_match.groups(); label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total"; value = value.strip() if value else ""
             else:
                 amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                 if amount_match: value = amount_match.group(1).strip() if amount_match.group(1) else ""; potential_label = table_total_info[:amount_match.start()].strip(); label = potential_label if potential_label else "Total"
                 else: value = table_total_info; label = "Total"
        if num_cols > 0:
            total_row_elements = [Paragraph(html.escape(label), bl_style)]
            if num_cols > 2: total_row_elements.extend([Paragraph("", body_style)] * (num_cols - 2))
            if num_cols > 1: total_row_elements.append(Paragraph(html.escape(value), br_style))
            elif num_cols == 1:
                 if label == "Total": total_row_elements[0] = Paragraph(html.escape(value), bl_style)
            total_row_elements += [Paragraph("", body_style)] * (num_cols - len(total_row_elements))
            wrapped_data.append(total_row_elements)
            has_total_row = True

    if wrapped_data and col_widths and len(wrapped_data) > 1:
        tbl = LongTable(wrapped_data, colWidths=col_widths, repeatRows=1)
        style_commands = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("VALIGN", (0, 1), (-1, -1), "TOP"),
        ]
        if has_total_row:
             if num_cols > 1: style_commands.extend([('SPAN', (0, -1), (-2, -1)), ('ALIGN', (0, -1), (-2, -1), 'LEFT'), ('ALIGN', (-1, -1), (-1, -1), 'RIGHT'), ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
             elif num_cols == 1: style_commands.extend([('ALIGN', (0, -1), (0, -1), 'LEFT'), ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
        tbl.setStyle(TableStyle(style_commands))
        elements += [tbl, Spacer(1, 24)]

if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]; num_cols = len(last_hdr)
    if num_cols > 0:
        gt_col_widths = []
        try:
            desc_actual_idx_in_hdr = last_hdr.index("Description")
            desc_col_width = total_page_width * 0.45; other_cols_count = num_cols - 1
            if other_cols_count > 0:
                 other_total_width = total_page_width - desc_col_width
                 strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and last_hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1
                 if strategy_idx != -1:
                     strat_width = total_page_width * 0.15; remaining_width = other_total_width - strat_width; remaining_cols = other_cols_count - 1
                     other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0
                     gt_col_widths = [max(0.1*inch, other_indiv_width) if i != desc_actual_idx_in_hdr and i != strategy_idx else (desc_col_width if i == desc_actual_idx_in_hdr else strat_width) for i in range(num_cols)]
                 else: other_col_width = other_total_width / other_cols_count; gt_col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
            elif num_cols == 1: gt_col_widths = [total_page_width]
            else: gt_col_widths = [total_page_width / num_cols] * num_cols
        except ValueError: gt_col_widths = [total_page_width / num_cols] * num_cols if num_cols > 0 else []

        if gt_col_widths:
            gt_row_data = [ Paragraph("Grand Total", bl_style) ]
            if num_cols > 2: gt_row_data.extend([ Paragraph("", body_style) for _ in range(num_cols - 2) ])
            if num_cols > 1: gt_row_data.append(Paragraph(html.escape(grand_total), br_style))
            elif num_cols == 1: gt_row_data = [Paragraph(f"Grand Total: {html.escape(grand_total)}", bl_style)]
            gt_row_data += [Paragraph("", body_style)] * (num_cols - len(gt_row_data))
            gt_table = LongTable([gt_row_data], colWidths=gt_col_widths)
            gt_style_cmds = [("GRID", (0, 0), (-1, -1), 0.25, colors.grey), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0"))]
            if num_cols > 1: gt_style_cmds.extend([('SPAN', (0, 0), (-2, 0)), ('ALIGN', (0, 0), (-2, 0), 'LEFT'), ('ALIGN', (-1, 0), (-1, 0), 'RIGHT')])
            else: gt_style_cmds.append(('ALIGN', (0,0), (0,0), 'LEFT'))
            gt_table.setStyle(TableStyle(gt_style_cmds)); elements.append(gt_table)
try:
    doc.build(elements); pdf_buf.seek(0)
except Exception as e: st.error(f"Error building PDF: {e}"); import traceback; st.error(traceback.format_exc()); pdf_buf = None


# === Word Building Section (No changes needed here, uses tables_info) ===
docx_buf = io.BytesIO()
docx_doc = Document()
sec = docx_doc.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_height = Inches(11); sec.page_width = Inches(17)
sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5); sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
if logo:
    try: p_logo = docx_doc.add_paragraph(); r_logo = p_logo.add_run(); img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width; img_width_in = 5; img_height_in = img_width_in * ratio; r_logo.add_picture(io.BytesIO(logo), width=Inches(img_width_in)); p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
    except Exception as e: st.warning(f"Could not add logo to Word: {e}")
p_title = docx_doc.add_paragraph(); p_title.alignment = WD_TABLE_ALIGNMENT.CENTER; r_title = p_title.add_run(proposal_title); r_title.font.name = DEFAULT_SERIF_FONT; r_title.font.size = Pt(18); r_title.bold = True; docx_doc.add_paragraph()
TOTAL_W_INCHES = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info):
    n = len(hdr)
    if n == 0: continue

    desc_actual_idx_in_hdr = -1; desc_w_in = 0; other_w_in = 0; strat_w_in = 0; strategy_idx = -1
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_w_in = 0.45 * TOTAL_W_INCHES; other_cols_count = n - 1
        if other_cols_count > 0:
            other_total_w_in = TOTAL_W_INCHES - desc_w_in
            strategy_idx = desc_actual_idx_in_hdr - 1 if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy" else -1
            if strategy_idx != -1: strat_w_in = 0.15 * TOTAL_W_INCHES; remaining_w_in = other_total_w_in - strat_w_in; remaining_cols = other_cols_count - 1; other_w_in = remaining_w_in / remaining_cols if remaining_cols > 0 else 0
            else: other_w_in = other_total_w_in / other_cols_count
        elif n == 1: desc_w_in = TOTAL_W_INCHES; other_w_in = 0
        else: other_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES; desc_w_in = other_w_in
    except ValueError: desc_actual_idx_in_hdr = -1; desc_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES; other_w_in = desc_w_in; strategy_idx = -1

    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid");
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER; tbl.autofit = False; tbl.allow_autofit = False;
    tblPr_list = tbl._element.xpath('./w:tblPr')
    if not tblPr_list: tblPr = OxmlElement('w:tblPr'); tbl._element.insert(0, tblPr)
    else: tblPr = tblPr_list[0]
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct');
    existing_tblW = tblPr.xpath('./w:tblW');
    if existing_tblW: tblPr.remove(existing_tblW[0])
    tblPr.append(tblW)

    for idx, col in enumerate(tbl.columns):
        width_val = 0
        if idx == desc_actual_idx_in_hdr: width_val = desc_w_in
        elif strategy_idx != -1 and idx == strategy_idx: width_val = strat_w_in
        else: width_val = other_w_in
        col.width = Inches(max(0.2, width_val))

    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr):
        if i >= len(hdr_cells): break
        cell = hdr_cells[i]; tc = cell._tc; tcPr = tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'F2F2F2'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); tcPr.append(shd);
        p = cell.paragraphs[0]; p.text = ""; run = p.add_run(str(col_name)); run.font.name = DEFAULT_SERIF_FONT; run.font.size = Pt(10); run.bold = True; p.alignment = WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for ridx, row in enumerate(rows_data):
        current_cells_count = len(row)
        if current_cells_count < n: row = list(row) + [""] * (n - current_cells_count)
        elif current_cells_count > n: row = row[:n]
        row_cells = tbl.add_row().cells
        for cidx, cell_content in enumerate(row):
            if cidx >= len(row_cells): break
            cell = row_cells[cidx]; p = cell.paragraphs[0]; p.text = ""; cell_str = str(cell_content);
            run_text = p.add_run(cell_str); run_text.font.name = DEFAULT_SANS_FONT; run_text.font.size = Pt(9);
            link_applied = False
            if cidx == desc_actual_idx_in_hdr and ridx < len(row_links_uri_list) and row_links_uri_list[ridx]:
                link_uri = row_links_uri_list[ridx]
                if link_uri:
                    if cell_str: space_run = p.add_run(" "); space_run.font.name = DEFAULT_SANS_FONT; space_run.font.size = Pt(9);
                    try: add_hyperlink(p, link_uri, "- link", font_name=DEFAULT_SANS_FONT, font_size=9); link_applied = True
                    except Exception as link_e: failed_link_run = p.add_run("- link (error)"); failed_link_run.font.name = DEFAULT_SANS_FONT; failed_link_run.font.size = Pt(9); failed_link_run.font.color.rgb = RGBColor(255, 0, 0);
            p.alignment = WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    if table_total_info:
        label = "Total"; amount = ""
        if isinstance(table_total_info, list): label = table_total_info[0].strip() if table_total_info and table_total_info[0] else "Total"; amount = next((val.strip() for val in reversed(table_total_info) if val and '$' in str(val)), "");
        if not amount and len(table_total_info) > 1: amount = table_total_info[-1].strip() if table_total_info[-1] else ""
        elif isinstance(table_total_info, str):
            try:
                total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
                if total_match: label_parsed, amount_parsed = total_match.groups(); label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total"; amount = amount_parsed.strip() if amount_parsed else ""
                else:
                    amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                    if amount_match: amount = amount_match.group(1).strip() if amount_match.group(1) else ""; potential_label = table_total_info[:amount_match.start()].strip(); label = potential_label if potential_label else "Total"
                    else: amount = table_total_info; label = "Total"
            except Exception as e: amount = table_total_info; label = "Total"
        total_cells = tbl.add_row().cells;
        if n > 0:
            label_cell = total_cells[0];
            if n > 1:
                try: label_cell.merge(total_cells[n-2])
                except Exception as merge_e: pass
            p_label = label_cell.paragraphs[0]; p_label.text = ""; run_label = p_label.add_run(label); run_label.font.name = DEFAULT_SERIF_FONT; run_label.font.size = Pt(10); run_label.bold = True; p_label.alignment = WD_TABLE_ALIGNMENT.LEFT; label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
            if n > 1:
                 amount_cell = total_cells[n-1]; p_amount = amount_cell.paragraphs[0]; p_amount.text = ""; run_amount = p_amount.add_run(amount); run_amount.font.name = DEFAULT_SERIF_FONT; run_amount.font.size = Pt(10); run_amount.bold = True; p_amount.alignment = WD_TABLE_ALIGNMENT.RIGHT; amount_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
            elif n == 1:
                if label == "Total": p_label.text = amount; run_label.text = amount

    docx_doc.add_paragraph()

if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]; n = len(last_hdr)
    if n > 0:
        gt_desc_idx = -1; gt_desc_w = 0; gt_other_w = 0; gt_strat_w = 0; gt_strat_idx = -1
        try:
             gt_desc_idx = last_hdr.index("Description"); gt_desc_w = 0.45 * TOTAL_W_INCHES; gt_other_count = n - 1
             if gt_other_count > 0:
                 gt_other_total_w = TOTAL_W_INCHES - gt_desc_w
                 gt_strat_idx = gt_desc_idx - 1 if gt_desc_idx > 0 and last_hdr[gt_desc_idx - 1] == "Strategy" else -1
                 if gt_strat_idx != -1: gt_strat_w = 0.15 * TOTAL_W_INCHES; gt_remain_w = gt_other_total_w - gt_strat_w; gt_remain_cols = gt_other_count - 1; gt_other_w = gt_remain_w / gt_remain_cols if gt_remain_cols > 0 else 0
                 else: gt_other_w = gt_other_total_w / gt_other_count
             elif n == 1: gt_desc_w = TOTAL_W_INCHES; gt_other_w = 0
             else: gt_other_w = TOTAL_W_INCHES / n; gt_desc_w = gt_other_w
        except ValueError: gt_desc_idx = -1; gt_desc_w = TOTAL_W_INCHES / n; gt_other_w = gt_desc_w; gt_strat_idx = -1

        tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid"); tblg.alignment = WD_TABLE_ALIGNMENT.CENTER; tblg.autofit = False; tblg.allow_autofit = False;
        tblgPr_list = tblg._element.xpath('./w:tblPr')
        if not tblgPr_list: tblgPr = OxmlElement('w:tblPr'); tblg._element.insert(0, tblgPr)
        else: tblgPr = tblgPr_list[0]
        tblgW = OxmlElement('w:tblW'); tblgW.set(qn('w:w'), '5000'); tblgW.set(qn('w:type'), 'pct');
        existing_tblgW_gt = tblgPr.xpath('./w:tblW');
        if existing_tblgW_gt: tblgPr.remove(existing_tblgW_gt[0])
        tblgPr.append(tblgW)

        for idx, col in enumerate(tblg.columns):
            width_val_gt = 0
            if idx == gt_desc_idx: width_val_gt = gt_desc_w
            elif gt_strat_idx != -1 and idx == gt_strat_idx: width_val_gt = gt_strat_w
            else: width_val_gt = gt_other_w
            col.width = Inches(max(0.2, width_val_gt));
        gt_cells = tblg.rows[0].cells;
        if n > 0:
             gt_label_cell = gt_cells[0];
             if n > 1:
                 try: gt_label_cell.merge(gt_cells[n-2])
                 except Exception as merge_e: pass
             tc_label = gt_label_cell._tc; tcPr_label = tc_label.get_or_add_tcPr(); shd_label = OxmlElement('w:shd'); shd_label.set(qn('w:fill'), 'E0E0E0'); tcPr_label.append(shd_label); p_gt_label = gt_label_cell.paragraphs[0]; p_gt_label.text = ""; run_gt_label = p_gt_label.add_run("Grand Total"); run_gt_label.font.name = DEFAULT_SERIF_FONT; run_gt_label.font.size = Pt(10); run_gt_label.bold = True; p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT; gt_label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
             if n > 1:
                 gt_value_cell = gt_cells[n-1]; tc_val = gt_value_cell._tc; tcPr_val = tc_val.get_or_add_tcPr(); shd_val = OxmlElement('w:shd'); shd_val.set(qn('w:fill'), 'E0E0E0'); tcPr_val.append(shd_val); p_gt_val = gt_value_cell.paragraphs[0]; p_gt_val.text = ""; run_gt_val = p_gt_val.add_run(grand_total); run_gt_val.font.name = DEFAULT_SERIF_FONT; run_gt_val.font.size = Pt(10); run_gt_val.bold = True; p_gt_val.alignment = WD_TABLE_ALIGNMENT.RIGHT; gt_value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
             elif n==1: run_gt_label.text = f"Grand Total: {grand_total}"; p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT

# === Save and Download Buttons (Unchanged) ===
try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except Exception as e:
    st.error(f"Error building Word document: {e}")
    import traceback
    st.error(traceback.format_exc())
    docx_buf = None

c1, c2 = st.columns(2)
if pdf_buf:
    with c1: st.download_button("📥 Download deliverable PDF", data=pdf_buf, file_name="proposal_deliverable.pdf", mime="application/pdf", use_container_width=True)
else:
     with c1: st.error("PDF generation failed.")
if docx_buf:
    with c2: st.download_button("📥 Download deliverable DOCX", data=docx_buf, file_name="proposal_deliverable.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
else:
    with c2: st.error("Word document generation failed.")
