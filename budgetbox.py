# -*- coding: utf-8 -*-
import streamlit as st
import pdfplumber
import io
import requests
from PIL import Image
from docx import Document
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
    if not lines:
        return "", ""
    description = " ".join(lines[1:])
    description = re.sub(r'\s+', ' ', description).strip()
    if len(lines) == 1:
        return lines[0], ""
    return lines[0], description

# --- Word hyperlink helper ---
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
# --- START: Manual Table Reconstruction Function ---
def reconstruct_table_from_words(table_obj, page_height, x_tolerance=3, y_tolerance=3):
    bbox = table_obj.bbox
    crop_bbox = (
        max(0, bbox[0] - 5),
        max(0, bbox[1] - 5),
        min(table_obj.page.width, bbox[2] + 5),
        min(table_obj.page.height, bbox[3] + 5)
    )
    table_page = table_obj.page.crop(crop_bbox)
    words = table_page.extract_words(
        x_tolerance=x_tolerance,
        y_tolerance=y_tolerance,
        keep_blank_chars=True,
        use_text_flow=True,
        horizontal_ltr=True
    )
    if not words:
        return None
    words.sort(key=lambda w: (w['top'], w['x0']))
    row_lines = []
    current_line_top = words[0]['top']
    current_line_bottom = words[0]['bottom']
    for i in range(1, len(words)):
        line_height_guess = max(1, words[i]['bottom'] - words[i]['top'])
        if words[i]['top'] > current_line_bottom + (line_height_guess * 0.5):
            row_lines.append((current_line_top, current_line_bottom))
            current_line_top = words[i]['top']
            current_line_bottom = words[i]['bottom']
        else:
            current_line_bottom = max(current_line_bottom, words[i]['bottom'])
    row_lines.append((current_line_top, current_line_bottom))

    header_words = [w for w in words if abs(w['top'] - row_lines[0][0]) < y_tolerance]
    header_words.sort(key=lambda w: w['x0'])
    if not header_words:
        return None

    col_boundaries = []
    current_col_start = header_words[0]['x0']
    current_col_end = header_words[0]['x1']
    for i in range(1, len(header_words)):
        space_guess = header_words[i]['x0'] - current_col_end
        if space_guess > 5:
            col_boundaries.append((current_col_start, current_col_end))
            current_col_start = header_words[i]['x0']
            current_col_end = header_words[i]['x1']
        else:
            current_col_end = max(current_col_end, header_words[i]['x1'])
    col_boundaries.append((current_col_start, current_col_end))

    num_cols = len(col_boundaries)
    if num_cols == 0:
        return None

    table_data = [["" for _ in range(num_cols)] for _ in range(len(row_lines))]
    for word in words:
        word_mid_y = (word['top'] + word['bottom']) / 2
        word_mid_x = (word['x0'] + word['x1']) / 2
        row_idx = -1
        for idx, (r_top, r_bottom) in enumerate(row_lines):
            if r_top - y_tolerance <= word_mid_y <= r_bottom + y_tolerance:
                row_idx = idx
                break
        if row_idx == -1:
            continue
        col_idx = -1
        for idx, (c_start, c_end) in enumerate(col_boundaries):
            if max(word['x0'], c_start) < min(word['x1'], c_end):
                col_idx = idx
                break
            elif c_start - x_tolerance <= word_mid_x <= c_end + x_tolerance:
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
            if pi >= len(page_texts):
                return None
            for ln in page_texts[pi].splitlines():
                if re.search(r'\b(?<!Grand\s)(?:Total|Subtotal)\b.*?\$\s*[\d,.]+', ln, re.I) and ln not in used_totals:
                    used_totals.add(ln)
                    return ln.strip()
            return None

        for pi, page in enumerate(pdf.pages):
            links = page.hyperlinks
            page_tables = page.find_tables()
            if not page_tables:
                continue
            for tbl_idx, tbl in enumerate(page_tables):
                data = reconstruct_table_from_words(tbl, page.height, x_tolerance=2, y_tolerance=2)
                if data is None or len(data) < 2 or not any(data[0]):
                    data = tbl.extract(x_tolerance=3, y_tolerance=3)
                if not data or len(data) < 2:
                    continue

                original_hdr_raw = data[0]
                original_hdr = [(str(h).strip() if h is not None else "") for h in original_hdr_raw]
                if not any(original_hdr):
                    continue

                original_desc_idx = -1
                for i, h in enumerate(original_hdr):
                    if h and "description" in h.lower():
                        original_desc_idx = i
                        break
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
                        if found_common:
                            break
                    if original_desc_idx == -1:
                        for i, h in enumerate(original_hdr):
                            if h and len(h) > 8:
                                original_desc_idx = i
                                break
                if original_desc_idx == -1:
                    continue

                table_links = []
                new_hdr = []
                processed_desc_in_new = False
                for i, h in enumerate(original_hdr):
                    if i == original_desc_idx:
                        new_hdr.extend(["Strategy", "Description"])
                        processed_desc_in_new = True
                    elif h:
                        new_hdr.append(h)
                if not processed_desc_in_new or not new_hdr:
                    continue

                rows_data = []
                row_links_uri_list = []
                table_total_info = None
                for ridx_data, row_content in enumerate(data[1:], start=1):
                    full_row_content = list(row_content) + [""] * (len(original_hdr) - len(row_content))
                    row_str_list = [(str(cell).strip() if cell is not None else "") for cell in full_row_content]
                    if all(not cell_val for cell_val in row_str_list):
                        continue
                    first_cell_lower = row_str_list[0].lower() if row_str_list else ""
                    is_total_row = (("total" in first_cell_lower or "subtotal" in first_cell_lower) and any(re.search(r'\$|€|£|¥', str(cell_val)) for cell_val in row_str_list if cell_val))
                    if is_total_row:
                        if table_total_info is None:
                            table_total_info = row_str_list
                        continue
                    desc_text_from_pdf = row_str_list[original_desc_idx] if original_desc_idx < len(row_str_list) else ""
                    strat, desc = split_cell_text(desc_text_from_pdf)
                    new_row_content = []
                    for i, h in enumerate(original_hdr):
                        if i == original_desc_idx:
                            new_row_content.extend([strat, desc])
                        elif h:
                            new_row_content.append(row_str_list[i] if i < len(row_str_list) else "")
                    expected_cols = len(new_hdr)
                    current_cols = len(new_row_content)
                    if current_cols < expected_cols:
                        new_row_content.extend([""] * (expected_cols - current_cols))
                    elif current_cols > expected_cols:
                        new_row_content = new_row_content[:expected_cols]
                    rows_data.append(new_row_content)
                    row_links_uri_list.append(None)

                if table_total_info is None:
                    table_total_info = find_total(pi)

                if rows_data:
                    tables_info.append((new_hdr, rows_data, row_links_uri_list, table_total_info))

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
# === PDF and Word download handling ===

docx_buf = io.BytesIO()
docx_doc = Document()
section = docx_doc.sections[0]
section.orientation = WD_ORIENT.LANDSCAPE
section.page_width = Inches(17)
section.page_height = Inches(11)
section.left_margin = Inches(0.5)
section.right_margin = Inches(0.5)
section.top_margin = Inches(0.5)
section.bottom_margin = Inches(0.5)

# Title block
p_title = docx_doc.add_paragraph()
p_title.alignment = WD_TABLE_ALIGNMENT.CENTER
r_title = p_title.add_run(proposal_title)
r_title.font.name = DEFAULT_SERIF_FONT
r_title.font.size = Pt(18)
r_title.bold = True
docx_doc.add_paragraph()

# Table rendering
for hdr, rows_data, _, table_total_info in tables_info:
    tbl = docx_doc.add_table(rows=1, cols=len(hdr), style="Table Grid")
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = tbl.rows[0].cells
    for i, col_name in enumerate(hdr):
        cell = hdr_cells[i]
        p = cell.paragraphs[0]
        p.text = ""
        run = p.add_run(col_name)
        run.font.name = DEFAULT_SERIF_FONT
        run.font.size = Pt(10)
        run.bold = True
        p.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

    for row in rows_data:
        row_cells = tbl.add_row().cells
        for j, cell_val in enumerate(row):
            if j >= len(row_cells): break
            cell = row_cells[j]
            p = cell.paragraphs[0]
            p.text = ""
            run = p.add_run(cell_val)
            run.font.name = DEFAULT_SANS_FONT
            run.font.size = Pt(9)
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

    if table_total_info:
        total_row = tbl.add_row().cells
        label, amount = "Total", ""
        if isinstance(table_total_info, list):
            total_row_raw = table_total_info + [""] * (len(hdr) - len(table_total_info))
            label = total_row_raw[0].strip() if total_row_raw[0] else "Total"
            amount = total_row_raw[-1].strip()
            if '$' not in amount:
                amount = next((v for v in reversed(total_row_raw) if '$' in v), amount)
        elif isinstance(table_total_info, str):
            try:
                total_match = re.match(r'(.*?)\s*(\$?[\d,.]+)$', table_total_info)
                if total_match:
                    label_parsed, amount_parsed = total_match.groups()
                    label = label_parsed.strip() if label_parsed else "Total"
                    amount = amount_parsed.strip() if amount_parsed else ""
                else:
                    amt_match = re.search(r'(\$?[\d,.]+)$', table_total_info)
                    if amt_match:
                        amount = amt_match.group(1)
                        label = table_total_info[:amt_match.start()].strip() or "Total"
                    else:
                        amount = table_total_info
                        label = "Total"
            except Exception:
                label = "Total"
                amount = table_total_info

        if len(hdr) > 0:
            label_cell = total_row[0]
            p = label_cell.paragraphs[0]
            p.text = ""
            run = p.add_run(label)
            run.font.name = DEFAULT_SERIF_FONT
            run.font.size = Pt(10)
            run.bold = True
            p.alignment = WD_TABLE_ALIGNMENT.LEFT
            label_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        if len(hdr) > 1:
            amt_cell = total_row[-1]
            p_amt = amt_cell.paragraphs[0]
            p_amt.text = ""
            run = p_amt.add_run(amount)
            run.font.name = DEFAULT_SERIF_FONT
            run.font.size = Pt(10)
            run.bold = True
            p_amt.alignment = WD_TABLE_ALIGNMENT.RIGHT
            amt_cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

# Grand Total block
if grand_total:
    p_gt = docx_doc.add_paragraph()
    p_gt.alignment = WD_TABLE_ALIGNMENT.RIGHT
    run = p_gt.add_run(f"Grand Total: {grand_total}")
    run.font.name = DEFAULT_SERIF_FONT
    run.size = Pt(12)
    run.bold = True

# Save and serve download
try:
    docx_doc.save(docx_buf)
    docx_buf.seek(0)
except Exception as e:
    st.error(f"Error generating Word file: {e}")
    import traceback
    st.error(traceback.format_exc())
    docx_buf = None

# Streamlit download UI
if docx_buf:
    st.download_button("📥 Download DOCX", data=docx_buf, file_name="proposal_output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
else:
    st.error("Word document generation failed.")
