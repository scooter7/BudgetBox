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
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
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
import html
from collections import defaultdict

# Register fonts
try:
    pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT = "DMSerif"
    DEFAULT_SANS_FONT = "Barlow"
except Exception as e:
    st.warning(f"Could not load custom fonts: {e}. Using default system fonts.")
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
    if not lines: return "", ""
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    if len(lines) == 1: return lines[0], ""
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
        style.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
        style.font.underline = True
        style.priority = 9
        style.unhide_when_used = True
    style_element = OxmlElement('w:rStyle')
    style_element.set(qn('w:val'), 'Hyperlink')
    rPr.append(style_element)
    if font_name:
        run_font = OxmlElement('w:rFonts')
        run_font.set(qn('w:ascii'), font_name)
        run_font.set(qn('w:hAnsi'), font_name)
        rPr.append(run_font)
    if font_size:
        size = OxmlElement('w:sz')
        size.set(qn('w:val'), str(int(font_size * 2)))
        size_cs = OxmlElement('w:szCs')
        size_cs.set(qn('w:val'), str(int(font_size * 2)))
        rPr.append(size)
        rPr.append(size_cs)
    if bold:
        b = OxmlElement('w:b')
        rPr.append(b)
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return docx.text.run.Run(new_run, paragraph)
# --- End of Word hyperlink helper ---

# --- START: Manual Table Reconstruction Function ---
def reconstruct_table_from_words(table_obj, page_height, x_tolerance=3, y_tolerance=3):
    """
    Attempts to reconstruct table data using word positions.
    Returns a list-of-lists representing rows and cells, or None if failed.
    """
    bbox = table_obj.bbox
    crop_bbox = (max(0, bbox[0] - 5), max(0, bbox[1] - 5), min(table_obj.page.width, bbox[2] + 5), min(table_obj.page.height, bbox[3] + 5))
    table_page = table_obj.page.crop(crop_bbox)
    words = table_page.extract_words(x_tolerance=x_tolerance, y_tolerance=y_tolerance, keep_blank_chars=True, use_text_flow=True, horizontal_ltr=True)
    if not words: return None

    words.sort(key=lambda w: (w['top'], w['x0']))
    row_lines = []
    if words:
        current_line_top = words[0]['top']
        current_line_bottom = words[0]['bottom']
        for i in range(1, len(words)):
            line_height_guess = words[i]['bottom'] - words[i]['top']
            line_height_guess = max(1, line_height_guess)
            if words[i]['top'] > current_line_bottom + (line_height_guess * 0.5):
                row_lines.append((current_line_top, current_line_bottom))
                current_line_top = words[i]['top']
                current_line_bottom = words[i]['bottom']
            else:
                current_line_bottom = max(current_line_bottom, words[i]['bottom'])
        row_lines.append((current_line_top, current_line_bottom))
    if not row_lines: return None

    header_words = [w for w in words if abs(w['top'] - row_lines[0][0]) < y_tolerance]
    header_words.sort(key=lambda w: w['x0'])
    if not header_words: return None

    col_boundaries = []
    current_col_start = -1
    current_col_end = -1
    if header_words:
        current_col_start = header_words[0]['x0']
        current_col_end = header_words[0]['x1']
        for i in range(1, len(header_words)):
            space_guess = header_words[i]['x0'] - current_col_end
            if space_guess > 5: # Gap threshold
                col_boundaries.append((current_col_start, current_col_end))
                current_col_start = header_words[i]['x0']
                current_col_end = header_words[i]['x1']
            else:
                current_col_end = max(current_col_end, header_words[i]['x1'])
        col_boundaries.append((current_col_start, current_col_end)) # Add last column

    num_cols = len(col_boundaries)
    if num_cols == 0: return None

    table_data = [["" for _ in range(num_cols)] for _ in range(len(row_lines))]
    for word in words:
        word_mid_y = (word['top'] + word['bottom']) / 2
        word_mid_x = (word['x0'] + word['x1']) / 2
        row_idx = -1
        for idx, (r_top, r_bottom) in enumerate(row_lines):
            if word_mid_y >= r_top - y_tolerance and word_mid_y <= r_bottom + y_tolerance:
                row_idx = idx
                break
        if row_idx == -1: continue
        col_idx = -1
        for idx, (c_start, c_end) in enumerate(col_boundaries):
            if max(word['x0'], c_start) < min(word['x1'], c_end):
                 col_idx = idx
                 break
            elif word_mid_x >= c_start - x_tolerance and word_mid_x <= c_end + x_tolerance:
                 col_idx = idx
                 break

        if col_idx != -1:
            if table_data[row_idx][col_idx]:
                table_data[row_idx][col_idx] += " " + word['text']
            else:
                table_data[row_idx][col_idx] = word['text']
    for r in range(len(table_data)):
        for c in range(len(table_data[r])):
            table_data[r][c] = re.sub(r'\s+', ' ', table_data[r][c]).strip()
    return table_data
# --- END: Manual Table Reconstruction Function ---

# === START: PDF TABLE EXTRACTION AND PROCESSING LOGIC ===
tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_texts = [p.extract_text(x_tolerance=1, y_tolerance=1) or "" for p in pdf.pages]
        first_page_lines = page_texts[0].splitlines() if page_texts else []
        potential_title = next((line.strip() for line in first_page_lines if "proposal" in line.lower() and len(line.strip()) > 5), None)
        if potential_title:
             proposal_title = potential_title
        elif len(first_page_lines) > 0:
             proposal_title = first_page_lines[0].strip()

        used_totals = set()
        def find_total(pi):
            if pi >= len(page_texts): return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?<!Grand\s)(?:Total|Subtotal)\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            links = page.hyperlinks
            page_tables = page.find_tables()
            if not page_tables: continue

            for tbl_idx, tbl in enumerate(page_tables):
                data = reconstruct_table_from_words(tbl, page.height, x_tolerance=2, y_tolerance=2)
                if data is None or len(data) < 2 or not any(data[0]):
                    data = tbl.extract(x_tolerance=3, y_tolerance=3)
                if not data or len(data) < 2: continue

                original_hdr_raw = data[0]
                original_hdr = [(str(h).strip() if h is not None else "") for h in original_hdr_raw]
                if not any(original_hdr): continue

                original_desc_idx = -1
                for i, h in enumerate(original_hdr):
                    if h and "description" in h.lower(): original_desc_idx = i; break
                if original_desc_idx == -1:
                     common_desc_headers = ["details", "summary", "notes", "content"]
                     found_common = False
                     for i, h in enumerate(original_hdr):
                          if h:
                              for common_hdr in common_desc_headers:
                                   if common_hdr in h.lower(): original_desc_idx = i; found_common = True; break
                          if found_common: break
                     if original_desc_idx == -1:
                          for i, h in enumerate(original_hdr):
                             if h and len(h) > 8 : original_desc_idx = i; break
                if original_desc_idx == -1: continue

                table_links = []

                new_hdr = []; processed_desc_in_new = False
                for i, h in enumerate(original_hdr):
                    if i == original_desc_idx: new_hdr.extend(["Strategy", "Description"]); processed_desc_in_new = True
                    elif h: new_hdr.append(h)
                if not processed_desc_in_new or not new_hdr: continue

                rows_data = []; row_links_uri_list = []; table_total_info = None
                for ridx_data, row_content in enumerate(data[1:], start=1):
                    full_row_content = list(row_content) + [""] * (len(original_hdr) - len(row_content))
                    row_str_list = [(str(cell).strip() if cell is not None else "") for cell in full_row_content]
                    if all(not cell_val for cell_val in row_str_list): continue
                    first_cell_lower = row_str_list[0].lower() if row_str_list else ""
                    is_total_row = (("total" in first_cell_lower or "subtotal" in first_cell_lower) and \
                                   any(re.search(r'\$|€|£|¥', str(cell_val)) for cell_val in row_str_list if cell_val))
                    if is_total_row:
                        if table_total_info is None: table_total_info = row_str_list
                        continue
                    desc_text_from_pdf = row_str_list[original_desc_idx] if original_desc_idx < len(row_str_list) else ""
                    strat, desc = split_cell_text(desc_text_from_pdf)
                    new_row_content = []
                    for i, h in enumerate(original_hdr):
                        if i == original_desc_idx: new_row_content.extend([strat, desc])
                        elif h: new_row_content.append(row_str_list[i] if i < len(row_str_list) else "")
                    expected_cols = len(new_hdr); current_cols = len(new_row_content)
                    if current_cols < expected_cols: new_row_content.extend([""] * (expected_cols - current_cols))
                    elif current_cols > expected_cols: new_row_content = new_row_content[:expected_cols]
                    rows_data.append(new_row_content); row_links_uri_list.append(None)

                if table_total_info is None: table_total_info = find_total(pi)
                if rows_data: tables_info.append((new_hdr, rows_data, row_links_uri_list, table_total_info))

        for tx in reversed(page_texts):
            m = re.search(r'Grand\s+Total.*?(?<!Subtotal\s)(?<!Sub Total\s)(\$\s*[\d,]+\.\d{2})', tx, re.I | re.S)
            if m: grand_total_candidate = m.group(1).replace(" ", "");
            if "subtotal" not in m.group(0).lower(): grand_total = grand_total_candidate; break
except Exception as e:
    st.error(f"Error processing PDF: {e}"); import traceback; st.error(traceback.format_exc()); st.stop()
# === END: PDF TABLE EXTRACTION AND PROCESSING LOGIC ===


# === PDF Building Section ===
pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17*inch, 11*inch)), leftMargin=0.5*inch, rightMargin=0.5*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
title_style  = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
header_style = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black)
body_style   = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=11)
link_style   = ParagraphStyle("LinkStyle", parent=body_style, textColor=colors.blue)
bl_style     = ParagraphStyle("BL", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_LEFT, textColor=colors.black, spaceBefore=6)
br_style     = ParagraphStyle("BR", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_RIGHT, textColor=colors.black, spaceBefore=6)
elements = []
logo = None
try: # Logo handling
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"; response = requests.get(logo_url, timeout=10); response.raise_for_status(); logo = response.content; img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width; img_width = min(5*inch, doc.width); img_height = img_width * ratio; elements.append(RLImage(io.BytesIO(logo), width=img_width, height=img_height))
except Exception as e: st.warning(f"Could not load or process logo: {e}")
elements += [Spacer(1, 12), Paragraph(html.escape(proposal_title), title_style), Spacer(1, 24)]
total_page_width = doc.width

# --- Loop through processed tables ---
for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info):
    num_cols = len(hdr)
    if num_cols == 0: continue

    col_widths = []
    desc_actual_idx_in_hdr = -1

    # --- Calculate Column Widths ---
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_col_width = total_page_width * 0.45
        other_cols_count = num_cols - 1
        if other_cols_count > 0:
            other_total_width = total_page_width - desc_col_width
            strategy_idx = -1
            if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy":
                strategy_idx = desc_actual_idx_in_hdr - 1
            if strategy_idx != -1:
                strat_width = total_page_width * 0.15
                remaining_width = other_total_width - strat_width
                remaining_cols = other_cols_count - 1
                other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0
                col_widths = [max(0.1*inch, other_indiv_width) if i != desc_actual_idx_in_hdr and i != strategy_idx else (desc_col_width if i == desc_actual_idx_in_hdr else strat_width) for i in range(num_cols)]
            else:
                other_col_width = other_total_width / other_cols_count
                col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
        elif num_cols == 1:
            col_widths = [total_page_width]
        else:
             col_widths = []

    except ValueError:
        desc_actual_idx_in_hdr = -1
        if num_cols > 0:
             col_widths = [total_page_width / num_cols] * num_cols
        else:
             continue

    if not col_widths:
        continue

    # --- Build Table Content ---
    wrapped_header = [Paragraph(html.escape(str(h)), header_style) for h in hdr]; wrapped_data = [wrapped_header]
    for ridx, row in enumerate(rows_data):
        line = []; current_cells = len(row)
        if current_cells < num_cols: row = list(row) + [""] * (num_cols - current_cells)
        elif current_cells > num_cols: row = row[:num_cols]
        for cidx, cell_content in enumerate(row):
            cell_str = str(cell_content); escaped_cell_text = html.escape(cell_str); link_applied = False
            # Link logic disabled
            p = Paragraph(escaped_cell_text, body_style)
            line.append(p)
        wrapped_data.append(line)

    # --- Add Total Row ---
    has_total_row = False
    if table_total_info:
        label = "Total"; value = ""
        if isinstance(table_total_info, list):
            # Use original_hdr length for padding total row consistently
            original_total_row = list(table_total_info) + [""] * (len(original_hdr) - len(table_total_info))
            label = original_total_row[0].strip() if original_total_row[0] else "Total"
            value = original_total_row[-1].strip()
            if '$' not in value: value = next((val.strip() for val in reversed(original_total_row) if val and '$' in str(val)), value)
        elif isinstance(table_total_info, str):
             # FIXED SEMICOLONS HERE
             total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
             if total_match:
                  label_parsed, value_parsed = total_match.groups() # Renamed value variable
                  label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total"
                  value = value_parsed.strip() if value_parsed else "" # Use value_parsed
             else:
                  amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                  if amount_match:
                       value = amount_match.group(1).strip() if amount_match.group(1) else "" # Use value variable
                       potential_label = table_total_info[:amount_match.start()].strip()
                       label = potential_label if potential_label else "Total"
                  else:
                       value = table_total_info # Use value variable
                       label = "Total"
        # Build total row elements (same as before)
        if num_cols > 0:
            total_row_elements = [Paragraph(html.escape(label), bl_style)]
            if num_cols > 2: total_row_elements.extend([Paragraph("", body_style)] * (num_cols - 2))
            if num_cols > 1: total_row_elements.append(Paragraph(html.escape(value), br_style))
            elif num_cols == 1:
                 if label == "Total": total_row_elements[0] = Paragraph(html.escape(value), bl_style)
            total_row_elements += [Paragraph("", body_style)] * (num_cols - len(total_row_elements))
            wrapped_data.append(total_row_elements); has_total_row = True

    # --- Create and Style PDF Table ---
    if wrapped_data and len(wrapped_data) > 1:
        tbl = LongTable(wrapped_data, colWidths=col_widths, repeatRows=1); style_commands = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")), ("GRID", (0, 0), (-1, -1), 0.25, colors.grey), ("VALIGN", (0, 0), (-1, 0), "MIDDLE"), ("VALIGN", (0, 1), (-1, -1), "TOP"),];
        if has_total_row:
             if num_cols > 1: style_commands.extend([('SPAN', (0, -1), (-2, -1)), ('ALIGN', (0, -1), (-2, -1), 'LEFT'), ('ALIGN', (-1, -1), (-1, -1), 'RIGHT'), ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
             elif num_cols == 1: style_commands.extend([('ALIGN', (0, -1), (0, -1), 'LEFT'), ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),])
        tbl.setStyle(TableStyle(style_commands)); elements += [tbl, Spacer(1, 24)]

# --- Add Grand Total Row ---
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]; num_cols = len(last_hdr)
    if num_cols > 0:
        gt_col_widths = [];
        try: # Use same width logic as main loop
            desc_actual_idx_in_hdr = last_hdr.index("Description"); desc_col_width = total_page_width * 0.45; other_cols_count = num_cols - 1
            if other_cols_count > 0: other_total_width = total_page_width - desc_col_width; strategy_idx = -1;
            if desc_actual_idx_in_hdr > 0 and last_hdr[desc_actual_idx_in_hdr - 1] == "Strategy": strategy_idx = desc_actual_idx_in_hdr - 1
            if strategy_idx != -1: strat_width = total_page_width * 0.15; remaining_width = other_total_width - strat_width; remaining_cols = other_cols_count - 1; other_indiv_width = remaining_width / remaining_cols if remaining_cols > 0 else 0; gt_col_widths = [max(0.1*inch, other_indiv_width) if i != desc_actual_idx_in_hdr and i != strategy_idx else (desc_col_width if i == desc_actual_idx_in_hdr else strat_width) for i in range(num_cols)]
            else: other_col_width = other_total_width / other_cols_count; gt_col_widths = [other_col_width if i != desc_actual_idx_in_hdr else desc_col_width for i in range(num_cols)]
            elif num_cols == 1: gt_col_widths = [total_page_width]
            else: gt_col_widths = []
        except ValueError:
             if num_cols > 0: gt_col_widths = [total_page_width / num_cols] * num_cols
             else: gt_col_widths = []

        if gt_col_widths: # Build GT table (same logic)
            gt_row_data = [ Paragraph("Grand Total", bl_style) ];
            if num_cols > 2: gt_row_data.extend([ Paragraph("", body_style) for _ in range(num_cols - 2) ])
            if num_cols > 1: gt_row_data.append(Paragraph(html.escape(grand_total), br_style))
            elif num_cols == 1: gt_row_data = [Paragraph(f"Grand Total: {html.escape(grand_total)}", bl_style)]
            gt_row_data += [Paragraph("", body_style)] * (num_cols - len(gt_row_data)); gt_table = LongTable([gt_row_data], colWidths=gt_col_widths); gt_style_cmds = [("GRID", (0, 0), (-1, -1), 0.25, colors.grey), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E0E0E0"))];
            if num_cols > 1: gt_style_cmds.extend([('SPAN', (0, 0), (-2, 0)), ('ALIGN', (0, 0), (-2, 0), 'LEFT'), ('ALIGN', (-1, 0), (-1, 0), 'RIGHT')])
            else: gt_style_cmds.append(('ALIGN', (0,0), (0,0), 'LEFT'))
            gt_table.setStyle(TableStyle(gt_style_cmds)); elements.append(gt_table)

# --- Build PDF Document ---
try: doc.build(elements); pdf_buf.seek(0)
except Exception as e: st.error(f"Error building PDF: {e}"); import traceback; st.error(traceback.format_exc()); pdf_buf = None


# === Word Building Section ===
docx_buf = io.BytesIO()
docx_doc = Document()
# Page Setup, Logo, Title (same as before)
sec = docx_doc.sections[0]; sec.orientation = WD_ORIENT.LANDSCAPE; sec.page_height = Inches(11); sec.page_width = Inches(17); sec.left_margin = Inches(0.5); sec.right_margin = Inches(0.5); sec.top_margin = Inches(0.5); sec.bottom_margin = Inches(0.5)
if logo:
    try: p_logo = docx_doc.add_paragraph(); r_logo = p_logo.add_run(); img = Image.open(io.BytesIO(logo)); ratio = img.height / img.width; img_width_in = 5; img_height_in = img_width_in * ratio; r_logo.add_picture(io.BytesIO(logo), width=Inches(img_width_in)); p_logo.alignment = WD_TABLE_ALIGNMENT.CENTER
    except Exception as e: st.warning(f"Could not add logo to Word: {e}")
p_title = docx_doc.add_paragraph(); p_title.alignment = WD_TABLE_ALIGNMENT.CENTER; r_title = p_title.add_run(proposal_title); r_title.font.name = DEFAULT_SERIF_FONT; r_title.font.size = Pt(18); r_title.bold = True; docx_doc.add_paragraph()
TOTAL_W_INCHES = sec.page_width.inches - sec.left_margin.inches - sec.right_margin.inches

# Loop through tables for Word
for table_index, (hdr, rows_data, row_links_uri_list, table_total_info) in enumerate(tables_info):
    n = len(hdr);
    if n == 0: continue
    # Width calculation (same logic as PDF, semicolons fixed)
    desc_actual_idx_in_hdr = -1; desc_w_in = 0; other_w_in = 0; strat_w_in = 0; strategy_idx = -1
    try:
        desc_actual_idx_in_hdr = hdr.index("Description")
        desc_w_in = 0.45 * TOTAL_W_INCHES
        other_cols_count = n - 1
        if other_cols_count > 0:
            other_total_w_in = TOTAL_W_INCHES - desc_w_in
            strategy_idx = -1
            if desc_actual_idx_in_hdr > 0 and hdr[desc_actual_idx_in_hdr - 1] == "Strategy":
                strategy_idx = desc_actual_idx_in_hdr - 1
            if strategy_idx != -1:
                strat_w_in = 0.15 * TOTAL_W_INCHES
                remaining_w_in = other_total_w_in - strat_w_in
                remaining_cols = other_cols_count - 1
                other_w_in = remaining_w_in / remaining_cols if remaining_cols > 0 else 0
            else:
                other_w_in = other_total_w_in / other_cols_count
        elif n == 1:
            desc_w_in = TOTAL_W_INCHES
            other_w_in = 0
        # else case n=0 is skipped, case n=1 handled

    except ValueError:
        desc_actual_idx_in_hdr = -1
        desc_w_in = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES
        other_w_in = desc_w_in
        strategy_idx = -1

    # Create table
    tbl = docx_doc.add_table(rows=1, cols=n, style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.allow_autofit = False
    # Set table width
    tblPr_list = tbl._element.xpath('./w:tblPr')
    if not tblPr_list: tblPr = OxmlElement('w:tblPr'); tbl._element.insert(0, tblPr)
    else: tblPr = tblPr_list[0]
    tblW = OxmlElement('w:tblW'); tblW.set(qn('w:w'), '5000'); tblW.set(qn('w:type'), 'pct'); existing_tblW = tblPr.xpath('./w:tblW');
    if existing_tblW: tblPr.remove(existing_tblW[0])
    tblPr.append(tblW);

    # Set column widths
    for idx, col in enumerate(tbl.columns):
        width_val = 0
        if idx == desc_actual_idx_in_hdr: width_val = desc_w_in
        elif strategy_idx != -1 and idx == strategy_idx: width_val = strat_w_in
        else: width_val = other_w_in
        col.width = Inches(max(0.2, width_val));

    # Populate header
    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr):
        if i >= len(hdr_cells): break
        cell = hdr_cells[i]; tc = cell._tc; tcPr = tc.get_or_add_tcPr(); shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'F2F2F2'); shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); tcPr.append(shd); p = cell.paragraphs[0]; p.text = ""; run = p.add_run(str(col_name)); run.font.name = DEFAULT_SERIF_FONT; run.font.size = Pt(10); run.bold = True; p.alignment = WD_TABLE_ALIGNMENT.CENTER; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    # Populate data rows
    for ridx, row in enumerate(rows_data):
        current_cells_count = len(row)
        if current_cells_count < n: row = list(row) + [""] * (n - current_cells_count)
        elif current_cells_count > n: row = row[:n]
        row_cells = tbl.add_row().cells
        for cidx, cell_content in enumerate(row):
            if cidx >= len(row_cells): break
            cell = row_cells[cidx]; p = cell.paragraphs[0]; p.text = ""; cell_str = str(cell_content); run_text = p.add_run(cell_str); run_text.font.name = DEFAULT_SANS_FONT; run_text.font.size = Pt(9); link_applied = False
            # Link logic disabled
            p.alignment = WD_TABLE_ALIGNMENT.LEFT; cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    # Add total row
    if table_total_info:
        label = "Total"; amount = ""
        if isinstance(table_total_info, list):
             original_total_row = list(table_total_info) + [""] * (len(original_hdr) - len(table_total_info))
             label = original_total_row[0].strip() if original_total_row[0] else "Total"
             amount = original_total_row[-1].strip()
             if '$' not in amount: amount = next((val.strip() for val in reversed(original_total_row) if val and '$' in str(val)), amount)
        elif isinstance(table_total_info, str):
            # FIXED SEMICOLONS HERE
            try:
                total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
                if total_match:
                     label_parsed, amount_parsed = total_match.groups() # Separate line
                     label = label_parsed.strip() if label_parsed and label_parsed.strip() else "Total" # Separate line
                     amount = amount_parsed.strip() if amount_parsed else "" # Separate line
                else:
                     amount_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                     if amount_match:
                          amount = amount_match.group(1).strip() if amount_match.group(1) else "" # Separate line
                          potential_label = table_total_info[:amount_match.start()].strip() # Separate line
                          label = potential_label if potential_label else "Total" # Separate line
                     else:
                          amount = table_total_info
                          label = "Total"
            except Exception as e:
                 amount = table_total_info
                 label = "Total"
        # Populate total row cells (same logic)
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

# Add Grand Total row (same logic)
if grand_total and tables_info:
    last_hdr, _, _, _ = tables_info[-1]; n = len(last_hdr)
    if n > 0:
        gt_desc_idx = -1; gt_desc_w = 0; gt_other_w = 0; gt_strat_w = 0; gt_strat_idx = -1
        # GT width calculation (semicolons fixed)
        try:
            gt_desc_idx = last_hdr.index("Description")
            gt_desc_w = 0.45 * TOTAL_W_INCHES
            gt_other_count = n - 1
            if gt_other_count > 0:
                 gt_other_total_w = TOTAL_W_INCHES - gt_desc_w
                 gt_strat_idx = -1 # Initialize
                 if gt_desc_idx > 0 and last_hdr[gt_desc_idx - 1] == "Strategy":
                     gt_strat_idx = gt_desc_idx - 1
                 if gt_strat_idx != -1:
                     gt_strat_w = 0.15 * TOTAL_W_INCHES
                     gt_remain_w = gt_other_total_w - gt_strat_w
                     gt_remain_cols = gt_other_count - 1
                     gt_other_w = gt_remain_w / gt_remain_cols if gt_remain_cols > 0 else 0
                 else:
                     gt_other_w = gt_other_total_w / gt_other_count
            elif n == 1:
                 gt_desc_w = TOTAL_W_INCHES
                 gt_other_w = 0
            # else case for n=0 skipped, n=1 handled
        except ValueError:
            gt_desc_idx = -1
            gt_desc_w = TOTAL_W_INCHES / n if n > 0 else TOTAL_W_INCHES
            gt_other_w = gt_desc_w
            gt_strat_idx = -1
        # Create GT table (same logic)
        tblg = docx_doc.add_table(rows=1, cols=n, style="Table Grid"); tblg.alignment = WD_TABLE_ALIGNMENT.CENTER; tblg.autofit = False; tblg.allow_autofit = False; tblgPr_list = tblg._element.xpath('./w:tblPr');
        if not tblgPr_list: tblgPr = OxmlElement('w:tblPr'); tblg._element.insert(0, tblgPr)
        else: tblgPr = tblgPr_list[0]
        tblgW = OxmlElement('w:tblW'); tblgW.set(qn('w:w'), '5000'); tblgW.set(qn('w:type'), 'pct'); existing_tblgW_gt = tblgPr.xpath('./w:tblW');
        if existing_tblgW_gt: tblgPr.remove(existing_tblgW_gt[0])
        tblgPr.append(tblgW);
        # Set GT column widths (same logic)
        for idx, col in enumerate(tblg.columns):
            width_val_gt = 0
            if idx == gt_desc_idx: width_val_gt = gt_desc_w
            elif gt_strat_idx != -1 and idx == gt_strat_idx: width_val_gt = gt_strat_w
            else: width_val_gt = gt_other_w
            col.width = Inches(max(0.2, width_val_gt));
        # Populate GT row (same logic)
        gt_cells = tblg.rows[0].cells;
        if n > 0: gt_label_cell = gt_cells[0];
        if n > 1:
            try: gt_label_cell.merge(gt_cells[n-2])
            except Exception as merge_e: pass
        tc_label = gt_label_cell._tc; tcPr_label = tc_label.get_or_add_tcPr(); shd_label = OxmlElement('w:shd'); shd_label.set(qn('w:fill'), 'E0E0E0'); tcPr_label.append(shd_label); p_gt_label = gt_label_cell.paragraphs[0]; p_gt_label.text = ""; run_gt_label = p_gt_label.add_run("Grand Total"); run_gt_label.font.name = DEFAULT_SERIF_FONT; run_gt_label.font.size = Pt(10); run_gt_label.bold = True; p_gt_label.alignment = WD_TABLE_ALIGNMENT.LEFT; gt_label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
        if n > 1: gt_value_cell = gt_cells[n-1]; tc_val = gt_value_cell._tc; tcPr_val = tc_val.get_or_add_tcPr(); shd_val = OxmlElement('w:shd'); shd_val.set(qn('w:fill'), 'E0E0E0'); tcPr_val.append(shd_val); p_gt_val = gt_value_cell.paragraphs[0]; p_gt_val.text = ""; run_gt_val = p_gt_val.add_run(grand_total); run_gt_val.font.name = DEFAULT_SERIF_FONT; run_gt_val.font.size = Pt(10); run_gt_val.bold = True; p_gt_val.alignment = WD_TABLE_ALIGNMENT.RIGHT; gt_value_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER;
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
