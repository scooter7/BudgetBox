# -*- coding: utf-8 -*-
import io
import re
import html # <<<--- ADDED IMPORT
import requests # <<<--- ADDED IMPORT
import camelot
import pdfplumber
import fitz
import streamlit as st
from PIL import Image # <<<--- ADDED IMPORT
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
# Ensure RLImage is imported
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage

# --- Font Registration ---
pdfmetrics.registerFont(TTFont("DMSerif", "fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow", "fonts/Barlow-Regular.ttf"))

ST_BARLOW_BOLD_LOADED = False
ST_DMSERIF_BOLD_LOADED = False

try:
    pdfmetrics.registerFont(TTFont("Barlow-Bold", "fonts/Barlow-Bold.ttf"))
    pdfmetrics.registerFontFamily('Barlow', normal='Barlow', bold='Barlow-Bold')
    ST_BARLOW_BOLD_LOADED = True
except Exception as e:
    pass # Silently ignore if bold font not found

# DEFAULT_SERIF_FONT = "DMSerif" # Already defined below
# DEFAULT_SANS_FONT = "Barlow" # Already defined below

# --- Streamlit Setup & PDF Loading ---
st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download the re-formatted PDF output.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")

if not uploaded:
    st.stop()

pdf_bytes = uploaded.read()

try:
    doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")
except Exception as e:
    st.error(f"Error opening PDF with Fitz: {e}")
    st.stop()

# --- Helper Functions & Constants ---
DEFAULT_SERIF_FONT = "DMSerif"
DEFAULT_SANS_FONT = "Barlow"

def extract_rich_cell(page_number, bbox):
    # ...(rest of the function remains unchanged)...
    try:
        page = doc_fitz.load_page(page_number)
        d = page.get_text("dict", clip=bbox)
        spans = []
        x0_bbox, y0_bbox, x1_bbox, y1_bbox = bbox

        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    sx0, sy0, sx1, sy1 = span["bbox"]
                    if sx0 < x1_bbox and sx1 > x0_bbox and sy0 < y1_bbox and sy1 > y0_bbox:
                        spans.append(span)
        if not spans:
            return ""

        lines_dict = {}
        for s_item in spans:
            key = round(s_item["origin"][1], 1)
            lines_dict.setdefault(key, []).append(s_item)

        span_text_lines = []
        for key in sorted(lines_dict.keys()):
            row_spans = sorted(lines_dict[key], key=lambda s_item: s_item["origin"][0])
            line_pieces = []
            for span_idx, span_item in enumerate(row_spans):
                if span_idx > 0:
                    prev_span_x1 = row_spans[span_idx - 1]["bbox"][2]
                    current_span_x0 = span_item["bbox"][0]
                    if current_span_x0 > prev_span_x1 + 1.5:
                        line_pieces.append(" ")
                
                text_content = span_item["text"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                font_name_from_span = span_item.get("font", "").lower()
                is_font_explicitly_bold = any(b_str in font_name_from_span for b_str in ["bold", "demibold", "semibold", "heavy", "black"])
                is_flagged_as_bold = span_item["flags"] & 2
                
                if is_font_explicitly_bold or is_flagged_as_bold:
                    line_pieces.append(f"<b>{text_content}</b>")
                else:
                    line_pieces.append(text_content)

            span_text_lines.append("".join(line_pieces))
        return "<br/>".join(span_text_lines)
    except Exception as e:
        return ""

HEADERS = [
    "Description", "Start Date", "End Date", "Term (Months)",
    "Monthly Amount", "Item Total", "Notes"
]

# --- Table Extraction & Data Processing ---
first_table = None
try:
    tables_camelot = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n", line_scale=40)
    if tables_camelot:
        raw = tables_camelot[0].df.values.tolist()
        if len(raw) > 1 and len(raw[0]) >= 6:
            header_row_text = "".join([str(h).lower() for h in raw[0]])
            if any(kw in header_row_text for kw in ['description', 'date', 'term', 'amount', 'total', 'notes']):
                first_table = raw
except Exception:
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal" # Default title

try:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # --- Extract Proposal Title (do this first) ---
        texts_content = [(p.page_number - 1, p.extract_text(x_tolerance=1, y_tolerance=1, layout=True) or "") for p in pdf.pages]
        first_page_idx_content, first_text_content = texts_content[0] if texts_content else (0, "")
        first_lines_content = first_text_content.splitlines()
        pot_title = next((l.strip() for l in first_lines_content if "proposal" in l.lower() and len(l.strip()) > 10), None)
        if pot_title:
            proposal_title = pot_title
        elif first_lines_content:
            proposal_title = next((l.strip() for l in first_lines_content if l.strip()), "Untitled Proposal")
        # --- End Title Extraction ---

        used_total_lines = set()
        def find_total(page_idx_for_texts_list):
            # ...(function remains unchanged)...
            if page_idx_for_texts_list >= len(texts_content): return None
            actual_page_num_for_key, text_content = texts_content[page_idx_for_texts_list]
            for l_line in text_content.splitlines():
                line_key = (actual_page_num_for_key, l_line)
                if re.search(r'\b(?<!grand\s)total\b.*?\$\s*[\d,.]+', l_line, re.I) and line_key not in used_total_lines:
                    used_total_lines.add(line_key)
                    return l_line.strip()
            return None

        for pi_pdfplumber, page in enumerate(pdf.pages):
             # ...(rest of the table extraction loop remains unchanged)...
            current_page_idx_fitz = page.page_number - 1
            
            if current_page_idx_fitz == 0 and first_table:
                found_tables_data = [("camelot", first_table, None, None)]
                page_links = []
            else:
                table_settings = {
                    "vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict",
                    "explicit_vertical_lines": page.curves + page.edges,
                    "explicit_horizontal_lines": page.curves + page.edges,
                    "snap_tolerance": 5, "join_tolerance": 5,
                    "min_words_vertical": 2, "min_words_horizontal": 1
                }
                current_page_tables = page.find_tables(table_settings=table_settings)
                if not current_page_tables: current_page_tables = page.find_tables()
                found_tables_data = [(tbl, tbl.extract(x_tolerance=1, y_tolerance=1), tbl.bbox, tbl.rows) for tbl in current_page_tables]
                page_links = page.hyperlinks

            for tbl_obj, data, table_bbox, rows_obj in found_tables_data:
                if not data or len(data) < 2: continue
                hdr = [str(h).strip().replace('\n', ' ') for h in data[0]]
                
                col_indices = {
                    "desc": next((i for i, h in enumerate(hdr) if "description" in h.lower()), 0), 
                    "notes": next((i for i, h in enumerate(hdr) if any(x in h.lower() for x in ["note", "comment"])), None),
                    "start_date": next((i for i, h in enumerate(hdr) if "start date" in h.lower()), None),
                    "end_date": next((i for i, h in enumerate(hdr) if "end date" in h.lower()), None),
                    "term": next((i for i, h in enumerate(hdr) if "term" in h.lower() or ("month" in h.lower() and "amount" not in h.lower())), None),
                    "monthly": next((i for i, h in enumerate(hdr) if "monthly" in h.lower()), None),
                    "total": next((i for i, h in enumerate(hdr) if "total" in h.lower() and "grand" not in h.lower() and "month" not in h.lower()), None)
                }

                desc_links_map = {}
                if tbl_obj != "camelot" and rows_obj and col_indices["desc"] is not None:
                    for r_idx, row_obj in enumerate(rows_obj):
                        if r_idx == 0: continue
                        if col_indices["desc"] < len(row_obj.cells) and row_obj.cells[col_indices["desc"]]:
                            cell_bbox_val = row_obj.cells[col_indices["desc"]]
                            x0_c, top_c, x1_c, bottom_c = cell_bbox_val
                            for link_item in page_links:
                                if all(k in link_item for k in ("x0", "x1", "top", "bottom", "uri")):
                                    if not (link_item["x1"] < x0_c or link_item["x0"] > x1_c or link_item["bottom"] < top_c or link_item["top"] > bottom_c):
                                        desc_links_map[r_idx] = link_item["uri"]
                                        break
                
                processed_table_rows = []
                ordered_row_links = []
                current_table_total_content = None

                for ridx_data, row_content_list in enumerate(data[1:], start=1):
                    cells = [str(c).strip() if c is not None else "" for c in row_content_list]
                    if not any(cells): continue

                    first_cell_lower = cells[0].lower() if cells else ""
                    if ("total" in first_cell_lower or "subtotal" in first_cell_lower) and any("$" in c for c in cells):
                        if current_table_total_content is None: current_table_total_content = cells
                        continue

                    new_row_output = [""] * len(HEADERS)

                    def get_cell_text_with_rich_extraction(col_idx_key, cell_list_idx, target_new_row_idx):
                        if col_indices[col_idx_key] is not None and col_indices[col_idx_key] < len(cells):
                            raw_text = cells[col_indices[col_idx_key]]
                            if tbl_obj != "camelot" and rows_obj and ridx_data < len(rows_obj) and \
                               col_indices[col_idx_key] < len(rows_obj[ridx_data].cells) and rows_obj[ridx_data].cells[col_indices[col_idx_key]]:
                                cell_bbox = rows_obj[ridx_data].cells[col_indices[col_idx_key]]
                                rich_text = extract_rich_cell(current_page_idx_fitz, cell_bbox)
                                new_row_output[target_new_row_idx] = rich_text or raw_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
                            else:
                                new_row_output[target_new_row_idx] = raw_text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
                    
                    get_cell_text_with_rich_extraction("desc", col_indices["desc"], 0)
                    get_cell_text_with_rich_extraction("notes", col_indices["notes"], 6)

                    for key, std_idx in [("start_date", 1), ("end_date", 2), ("term", 3), ("monthly", 4), ("total", 5)]:
                        if col_indices[key] is not None and col_indices[key] < len(cells):
                            new_row_output[std_idx] = cells[col_indices[key]]
                    
                    # Fallback guessing
                    for cell_idx, cell_val in enumerate(cells):
                        if not cell_val: continue
                        already_mapped = False
                        for key, std_idx_map in [("desc",0), ("notes",6), ("start_date",1), ("end_date",2), ("term",3), ("monthly",4), ("total",5)]:
                            if col_indices[key] == cell_idx and new_row_output[std_idx_map]:
                                already_mapped = True
                                break
                        if already_mapped: continue
                        
                        if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}', cell_val):
                            if not new_row_output[1]: new_row_output[1] = cell_val
                            elif not new_row_output[2]: new_row_output[2] = cell_val
                        elif re.fullmatch(r"\d{1,3}", cell_val.strip()) or ("month" in cell_val.lower() and "amount" not in cell_val.lower()) or "mo" in cell_val.lower():
                            if not new_row_output[3]: new_row_output[3] = cell_val.replace("months", "").replace("month", "").strip()
                        elif "$" in cell_val:
                            is_monthly_header = cell_idx < len(hdr) and "monthly" in hdr[cell_idx].lower()
                            is_total_header = cell_idx < len(hdr) and "total" in hdr[cell_idx].lower()
                            if not new_row_output[4] and is_monthly_header: new_row_output[4] = cell_val
                            elif not new_row_output[5] and is_total_header: new_row_output[5] = cell_val
                            elif not new_row_output[4]: new_row_output[4] = cell_val
                            elif not new_row_output[5]: new_row_output[5] = cell_val
                        elif col_indices["notes"] is None and col_indices["desc"] == cell_idx: continue 
                        elif not new_row_output[6] and len(cell_val) > 3 and not any(x_char in cell_val.lower() for x_char in ["date", "term", "$", "month"]):
                            new_row_output[6] = cell_val.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")

                    if any(new_row_output[i_val].strip().replace('\n',' ') == HEADERS[i_val] for i_val in range(len(HEADERS)) if new_row_output[i_val]):
                        continue
                    processed_table_rows.append(new_row_output)
                    ordered_row_links.append(desc_links_map.get(ridx_data)) # Use 0-based r_idx link mapping

                if current_table_total_content is None: current_table_total_content = find_total(current_page_idx_fitz)
                if processed_table_rows: tables_info.append((HEADERS, processed_table_rows, ordered_row_links, current_table_total_content))

        # Find Grand Total (use texts_content)
        for page_idx_fitz_rev, blk_text_rev in reversed(texts_content):
            m_grand = re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})', blk_text_rev, re.I | re.S)
            if m_grand:
                grand_total = m_grand.group(1).replace(" ", "")
                break
except pdfplumber.PDFSyntaxError as e_pdfsyn:
    st.error(f"PDFPlumber Error: Processing PDF failed. Error: {e_pdfsyn}")
    st.stop()
except Exception as e_proc:
    st.error(f"An unexpected error during PDF processing: {e_proc}")
    st.exception(e_proc)
    st.stop()

# --- PDF Generation ---
if not tables_info and not grand_total: # Check if there's anything to display
    st.warning("No tables or grand total suitable for reformatting were found in the uploaded PDF.")
    st.stop()

pdf_buf = io.BytesIO()
doc = SimpleDocTemplate(pdf_buf, pagesize=landscape((17 * inch, 11 * inch)),
                        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                        topMargin=0.5 * inch, bottomMargin=0.5 * inch)

# Styles
ts = ParagraphStyle("Title", fontName=DEFAULT_SERIF_FONT, fontSize=18, alignment=TA_CENTER, spaceAfter=12)
hs = ParagraphStyle("Header", fontName=DEFAULT_SERIF_FONT, fontSize=10, alignment=TA_CENTER, textColor=colors.black, spaceAfter=6)
bs = ParagraphStyle("Body", fontName=DEFAULT_SANS_FONT, fontSize=9, alignment=TA_LEFT, leading=12)
bs_right = ParagraphStyle("BodyRight", parent=bs, alignment=TA_RIGHT)
bs_center = ParagraphStyle("BodyCenter", parent=bs, alignment=TA_CENTER)

# --- Initialize story list ---
story = []
logo_added = False

# --- Add Logo from URL (Integrated Snippet) ---
try:
    logo_url = "https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png"
    # st.write(f"Attempting to load logo from {logo_url}...") # Optional user feedback
    resp = requests.get(logo_url, timeout=15)
    resp.raise_for_status() 
    logo_bytes = resp.content
    
    pil_img = Image.open(io.BytesIO(logo_bytes))
    img_width_pil, img_height_pil = pil_img.size
    
    if img_width_pil > 0 and img_height_pil > 0:
        ratio = img_height_pil / img_width_pil
        # Calculate width for ReportLab Image - max 5 inches or slightly less than doc width
        max_logo_width = 5 * inch 
        # Use doc.width (needs doc object defined, defined few lines above)
        reportlab_width = min(max_logo_width, doc.width - 1*inch) 
        reportlab_height = reportlab_width * ratio
        
        logo_image_rl = RLImage(io.BytesIO(logo_bytes), width=reportlab_width, height=reportlab_height, hAlign='CENTER')
        story.append(logo_image_rl)
        logo_added = True
        # st.success("Logo loaded successfully.") # Optional user feedback
    else:
        st.warning("Logo downloaded but image dimensions are invalid.")
        
except (requests.exceptions.RequestException, IOError) as e_req:
    st.warning(f"Could not download or process logo from URL: {e_req}")
except Exception as e_logo:
    st.warning(f"An unexpected error occurred while adding the logo: {e_logo}")
# --- End of Logo Integration ---

# Add Spacer after logo only if logo was added
if logo_added:
    story.append(Spacer(1, 0.25*inch))

# Add title (use html.escape as per snippet) and spacer
story.append(Paragraph(html.escape(proposal_title), ts))
story.append(Spacer(1, 24))

# --- Define Table Width & Columns ---
table_width = doc.width # Defined here, accessible above now too
main_col_widths = [
    table_width * 0.30, table_width * 0.08, table_width * 0.08, table_width * 0.06,
    table_width * 0.10, table_width * 0.10, table_width * 0.28
]

# --- Loop through extracted tables ---
for current_headers, current_rows, current_links, current_total_info in tables_info:
    # ...(rest of the table generation loop remains unchanged)...
    n_cols = len(current_headers)
    current_col_widths_val = main_col_widths[:n_cols] if len(main_col_widths) >= n_cols else [table_width / n_cols] * n_cols
    header_row_styled = [Paragraph(h_text, hs) for h_text in current_headers]
    table_data_styled = [header_row_styled]

    for i, row_data_list in enumerate(current_rows):
        styled_row_elements = []
        for j, cell_text_val in enumerate(row_data_list):
            cell_style_to_use = bs
            if j in [1, 2, 3]: cell_style_to_use = bs_center
            elif j in [4, 5]: cell_style_to_use = bs_right
            
            text_to_render = cell_text_val
            if j == 0 and i < len(current_links) and current_links[i]:
                text_to_render += f" <link href='{current_links[i]}' color='blue'>[link]</link>"
            styled_row_elements.append(Paragraph(text_to_render, cell_style_to_use))
        table_data_styled.append(styled_row_elements)

    if current_total_info:
        total_label_text, total_value_text = "Total", ""
        if isinstance(current_total_info, list):
            total_label_text = next((c for c in current_total_info if c and '$' not in c and c.strip().lower() not in ["total", "subtotal"]), None)
            if not total_label_text or total_label_text.lower() == "total": 
                 total_label_text = next((c for c in current_total_info if c and ('total' in c.lower() or 'subtotal' in c.lower())), "Total").strip()
            else:
                total_label_text = total_label_text.strip()
            total_value_text = next((c for c in reversed(current_total_info) if "$" in c), "")
        elif isinstance(current_total_info, str):
            m_total = re.match(r'(.*?)\s*(\$\s*[\d,]+\.\d{2})', current_total_info)
            if m_total: total_label_text, total_value_text = m_total.group(1).strip(), m_total.group(2).strip()
            else:
                total_label_text = re.sub(r'\$\s*[\d,.]+', '', current_total_info).strip() or "Total"
                val_match = re.search(r'(\$\s*[\d,.]+\.\d{2})', current_total_info)
                if val_match: total_value_text = val_match.group(1)
        if not total_label_text: total_label_text = "Total"
        
        # Add Total row with bold tags
        table_data_styled.append([Paragraph(f"<b>{total_label_text}</b>", bs)] + [Paragraph("<b></b>", bs)] * (n_cols - 2) + [Paragraph(f"<b>{total_value_text}</b>", bs_right)])

    tbl_reportlab = LongTable(table_data_styled, colWidths=current_col_widths_val, repeatRows=1)
    style_cmds_list = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F2F2")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"), ("VALIGN", (0, 1), (-1, -1), "TOP"),
    ]
    data_row_end_idx = -1 if not current_total_info else -2
    if len(table_data_styled) > 2 : # Check if data rows exist before applying data-row styles
        for col_idx_align, align_type in [(1, "CENTER"), (2, "CENTER"), (3, "CENTER"), (4, "RIGHT"), (5, "RIGHT")]:
             if col_idx_align < n_cols:
                # Ensure the end row index is valid (at least 1)
                valid_end_idx = max(1, data_row_end_idx) if data_row_end_idx != -1 else -1 # Use -1 if no total row
                if valid_end_idx == -1 or valid_end_idx >= 1: # Check ensures range is valid or applies to all data rows
                     style_cmds_list.append(("ALIGN", (col_idx_align, 1), (col_idx_align, valid_end_idx), align_type))


    if current_total_info:
        style_cmds_list.extend([
            ("SPAN", (0, -1), (-2, -1)), ("ALIGN", (0, -1), (-2, -1), "RIGHT"),
            ("ALIGN", (-1, -1), (-1, -1), "RIGHT"), ("VALIGN", (0, -1), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#EAEAEA")),
        ])
    tbl_reportlab.setStyle(TableStyle(style_cmds_list))
    story.extend([tbl_reportlab, Spacer(1, 24)])

# --- Add Grand Total Row ---
if grand_total:
    story.append(
        LongTable([[Paragraph("<b>Grand Total</b>", bs)] + [Paragraph("<b></b>", bs)] * (len(HEADERS) - 2) + [Paragraph(f"<b>{grand_total}</b>", bs_right)]], # Added bold tags
                  colWidths=main_col_widths,
                  style=TableStyle([
                      ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D0D0D0")),
                      ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
                      ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("SPAN", (0, 0), (-2, 0)),
                      ("ALIGN", (0, 0), (-2, 0), "RIGHT"), ("ALIGN", (-1, 0), (-1, 0), "RIGHT"),
                      ("TEXTCOLOR", (0, 0), (-1, -1), colors.black),
                      ("FONTNAME", (0, 0), (-1, -1), DEFAULT_SANS_FONT), ("FONTSIZE", (0, 0), (-1, -1), 10),
                  ]))
    )

# --- Build PDF and Provide Download ---
try:
    doc.build(story)
    pdf_buf.seek(0)
    st.download_button(
        "📥 Download Transformed PDF", data=pdf_buf,
        file_name="transformed_proposal.pdf", mime="application/pdf", use_container_width=True
    )
except Exception as e_build:
    st.error(f"Error building final PDF with ReportLab: {e_build}")
    st.exception(e_build)
