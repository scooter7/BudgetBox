# -*- coding: utf-8 -*-
import io
import re
import html
import camelot
import pdfplumber
import fitz
import requests
import streamlit as st
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, LongTable, TableStyle, Paragraph, Spacer, Image as RLImage

pdfmetrics.registerFont(TTFont("DMSerif","fonts/DMSerifDisplay-Regular.ttf"))
pdfmetrics.registerFont(TTFont("Barlow","fonts/Barlow-Regular.ttf"))
DEFAULT_SERIF_FONT="DMSerif"
DEFAULT_SANS_FONT="Barlow"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download the re-formatted PDF output.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()
doc_fitz = fitz.open(stream=pdf_bytes, filetype="pdf")

def extract_rich_cell(page_number, bbox):
    page = doc_fitz.load_page(page_number)
    d = page.get_text("dict")
    spans = []
    x0,y0,x1,y1 = bbox
    for block in d["blocks"]:
        if block.get("type")!=0: continue
        for line in block["lines"]:
            for span in line["spans"]:
                sx0,sy0,sx1,sy1 = span["bbox"]
                if not (sx1<x0 or sx0>x1 or sy1<y0 or sy0>y1):
                    spans.append(span)
    lines = {}
    for s in spans:
        key = round(s["bbox"][1],1)
        lines.setdefault(key,[]).append(s)
    text_lines = []
    for key in sorted(lines):
        row = sorted(lines[key], key=lambda s: s["bbox"][0])
        pieces = []
        for span in row:
            t = span["text"].replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            if span["flags"] & 2:
                pieces.append(f"<b>{t}</b>")
            else:
                pieces.append(t)
        text_lines.append("".join(pieces))
    return "<br/>".join(text_lines)

def is_data_row(cells):
    txt = " ".join(cells)
    if re.search(r'\$\s*\d',txt): return True
    if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',txt): return True
    if re.search(r'\b\d+\s*(?:month|mo|yr|year)\b',txt.lower()): return True
    return False

HEADERS = [
    "Description",
    "Start Date",
    "End Date",
    "Term (Months)",
    "Monthly Amount",
    "Item Total",
    "Notes"
]

first_table = None
try:
    tables = camelot.read_pdf(io.BytesIO(pdf_bytes), pages="1", flavor="lattice", strip_text="\n")
    if tables:
        raw = tables[0].df.values.tolist()
        if len(raw)>1 and len(raw[0])>=len(HEADERS) and any(is_data_row(row) for row in raw[1:]):
            first_table = raw
except:
    first_table = None

tables_info = []
grand_total = None
proposal_title = "Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts = [p.extract_text(x_tolerance=1,y_tolerance=1) or "" for p in pdf.pages]
    first_lines = texts[0].splitlines() if texts else []
    pot = next((l.strip() for l in first_lines if "proposal" in l.lower() and len(l.strip())>5),None)
    if pot:
        proposal_title = pot
    elif first_lines:
        proposal_title = first_lines[0].strip()
    used = set()
    def find_total(pi):
        if pi>=len(texts): return None
        for l in texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+',l,re.I) and l not in used:
                used.add(l)
                return l.strip()
        return None

    for pi,page in enumerate(pdf.pages):
        if pi==0 and first_table:
            found = [("camelot", first_table, None, None)]
            links = []
        else:
            found = [(tbl, tbl.extract(x_tolerance=1,y_tolerance=1), tbl.bbox, tbl.rows) for tbl in page.find_tables()]
            links = page.hyperlinks

        for tbl_obj, data, bbox, rows_obj in found:
            if not data or len(data)<2: continue
            hdr = [str(h).strip() for h in data[0]]
            mapping = {}
            for i,h in enumerate(hdr):
                hl=h.lower()
                if "description" in hl: mapping[i]=0
                elif "start" in hl: mapping[i]=1
                elif "end" in hl: mapping[i]=2
                elif "term" in hl or "duration" in hl: mapping[i]=3
                elif any(x in hl for x in["monthly","per month","rate","recurring"]): mapping[i]=4
                elif any(x in hl for x in["item total","subtotal","line total","amount"]): mapping[i]=5
                elif "note" in hl: mapping[i]=6

            desc_links={}
            if tbl_obj!="camelot":
                for r,row_obj in enumerate(rows_obj):
                    if r==0: continue
                    for orig_i,new_i in mapping.items():
                        if new_i==0 and orig_i<len(row_obj.cells):
                            x0,top,x1,bottom = row_obj.cells[orig_i]
                            for L in links:
                                if all(k in L for k in("x0","x1","top","bottom","uri")):
                                    if not(L["x1"]<x0 or L["x0"]>x1 or L["bottom"]<top or L["top"]>bottom):
                                        desc_links[r]=L["uri"]
                                        break

            rows_clean=[]
            row_links=[]
            tbl_tot=None
            for ridx,row in enumerate(data[1:],start=1):
                cells=[str(c).strip() for c in row]
                if not any(cells): continue
                fc=cells[0].lower()
                if ("total" in fc or "subtotal" in fc) and any("$" in c for c in cells):
                    if tbl_tot is None: tbl_tot=cells
                    continue
                new=[""]*len(HEADERS)
                for orig_i,new_i in mapping.items():
                    if orig_i<len(cells):
                        val=cells[orig_i]
                        if tbl_obj!="camelot" and rows_obj:
                            bbox_i=rows_obj[ridx].cells[orig_i]
                            rich=extract_rich_cell(pi,bbox_i)
                            if rich:
                                new[new_i]=rich
                                continue
                        t=val.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                        new[new_i]=t.replace("\n","<br/>")
                for v in cells:
                    if re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}',v):
                        if not new[1]:
                            new[1]=v
                        elif not new[2]:
                            new[2]=v
                if new[4] and new[3] and not new[5]:
                    amt=float(new[4].replace("$","").replace(",",""))
                    m=re.search(r"\d+",new[3])
                    if m:
                        new[5]=f"${amt*int(m.group()):,.2f}"
                if all(new[i]==HEADERS[i] for i in range(len(HEADERS))):
                    continue
                rows_clean.append(new)
                row_links.append(desc_links.get(ridx))
            if tbl_tot is None:
                tbl_tot=find_total(pi)
            if rows_clean:
                tables_info.append((HEADERS,rows_clean,row_links,tbl_tot))

    for blk in reversed(texts):
        m=re.search(r'Grand\s+Total.*?(\$\s*[\d,]+\.\d{2})',blk,re.I|re.S)
        if m:
            grand_total=m.group(1).replace(" ","")
            break

pdf_buf=io.BytesIO()
doc=SimpleDocTemplate(pdf_buf,pagesize=landscape((17*inch,11*inch)),
                     leftMargin=0.5*inch,rightMargin=0.5*inch,
                     topMargin=0.5*inch,bottomMargin=0.5*inch)
ts=ParagraphStyle("Title",fontName=DEFAULT_SERIF_FONT,fontSize=18,alignment=TA_CENTER,spaceAfter=12)
hs=ParagraphStyle("Header",fontName=DEFAULT_SERIF_FONT,fontSize=10,alignment=TA_CENTER,textColor=colors.black)
bs=ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=12)
els=[Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
tw=doc.width

for hdr, rows, links, tot in tables_info:
    n=len(hdr)
    ws=[tw*0.25,tw*0.08,tw*0.08,tw*0.08,tw*0.10,tw*0.10,tw*0.16]
    wrapped=[[Paragraph(html.escape(h),hs) for h in hdr]]
    for i,row in enumerate(rows):
        line=[]
        for j,cell in enumerate(row):
            if j==6 and cell:
                parts = [p.strip() for p in re.split(r'<br/?>',cell) if p.strip()]
                formatted = "<br/>".join(parts)
                txt = formatted
            else:
                txt = cell
            if j==0 and i<len(links) and links[i]:
                txt += f" <link href='{html.escape(links[i])}' color='blue'>- link</link>"
            line.append(Paragraph(txt, bs))
        wrapped.append(line)
    if tot:
        lbl,val="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total"
            val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m: lbl,val=m.group(1).strip(),m.group(2)
        wrapped.append([Paragraph(lbl,bs)] + [Paragraph("",bs)]*(n-2) + [Paragraph(val,bs)])
    tbl=LongTable(wrapped,colWidths=ws,repeatRows=1)
    cmds=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
          ("GRID",(0,0),(-1,-1),0.25,colors.grey),
          ("VALIGN",(0,0),(-1,0),"MIDDLE"),
          ("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        cmds += [("SPAN",(0,-1),(-2,-1)),
                 ("ALIGN",(0,-1),(-2,-1),"LEFT"),
                 ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),
                 ("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    els += [tbl, Spacer(1,24)]

if grand_total:
    ws=[tw*0.25,tw*0.08,tw*0.08,tw*0.08,tw*0.10,tw*0.10,tw*0.16]
    row=[Paragraph("Grand Total",bs)] + [Paragraph("",bs)]*(len(HEADERS)-2) + [Paragraph(html.escape(grand_total),bs)]
    gt=LongTable([row],colWidths=ws)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
                            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("SPAN",(0,0),(-2,0)),
                            ("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    els.append(gt)

doc.build(els)
pdf_buf.seek(0)

st.download_button("📥 Download PDF",
                   data=pdf_buf,
                   file_name="transformed_proposal.pdf",
                   mime="application/pdf",
                   use_container_width=True)
