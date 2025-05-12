# -*- coding: utf-8 -*-
import io
import re
import html
import camelot
import pdfplumber
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
from reportlab.platypus import (
    SimpleDocTemplate, LongTable, TableStyle,
    Paragraph, Spacer, Image as RLImage
)

try:
    pdfmetrics.registerFont(TTFont("DMSerif","fonts/DMSerifDisplay-Regular.ttf"))
    pdfmetrics.registerFont(TTFont("Barlow","fonts/Barlow-Regular.ttf"))
    DEFAULT_SERIF_FONT="DMSerif"
    DEFAULT_SANS_FONT="Barlow"
except:
    DEFAULT_SERIF_FONT="Times New Roman"
    DEFAULT_SANS_FONT="Arial"

st.set_page_config(page_title="Proposal Transformer", layout="wide")
st.title("🔄 Proposal Layout Transformer")
st.write("Upload a vertically formatted proposal PDF and download the re-formatted PDF output.")
uploaded = st.file_uploader("Upload proposal PDF", type="pdf")
if not uploaded:
    st.stop()
pdf_bytes = uploaded.read()

# Standard headers (no Strategy)
HEADERS = ["Description","Start Date","End Date","Term (Months)",
           "Monthly Amount","Item Total","Notes"]

# Try Camelot on page 1
first_table = None
try:
    tables=camelot.read_pdf(io.BytesIO(pdf_bytes),pages="1",flavor="lattice",strip_text="\n")
    if tables:
        df=tables[0].df
        raw=df.values.tolist()
        if len(raw)>1 and len(raw[0])>=len(HEADERS):
            first_table=raw
except:
    first_table=None

tables_info=[]
grand_total=None
proposal_title="Untitled Proposal"

with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
    texts=[p.extract_text(x_tolerance=1,y_tolerance=1) or "" for p in pdf.pages]
    lines=texts[0].splitlines() if texts else []
    pt = next((l.strip() for l in lines if "proposal" in l.lower() and len(l.strip())>5),None)
    if pt: proposal_title=pt
    elif lines: proposal_title=lines[0].strip()
    used=set()
    def find_total(pi):
        if pi>=len(texts): return None
        for l in texts[pi].splitlines():
            if re.search(r'\b(?!grand\s)total\b.*?\$\s*[\d,.]+',l,re.I) and l not in used:
                used.add(l)
                return l.strip()
        return None

    for pi,page in enumerate(pdf.pages):
        if pi==0 and first_table:
            found=[("camelot",first_table,None)]
            links=[]
        else:
            found=[(tbl,tbl.extract(x_tolerance=1,y_tolerance=1),tbl.bbox) for tbl in page.find_tables()]
            links=page.hyperlinks
        for tbl_obj,data,bbox in found:
            if not data or len(data)<2: continue
            hdr=[str(h).strip() for h in data[0]]
            # find Description col
            desc_i=next((i for i,h in enumerate(hdr) if h and "description" in h.lower()),None)
            if desc_i is None:
                desc_i=next((i for i,h in enumerate(hdr) if len(h)>10),None)
                if desc_i is None: continue
            # hyperlink map
            desc_links={}
            if tbl_obj!="camelot":
                for r,row_obj in enumerate(tbl_obj.rows):
                    if r==0: continue
                    if desc_i<len(row_obj.cells):
                        x0,top,x1,bottom=row_obj.cells[desc_i]
                        for L in links:
                            if all(k in L for k in("x0","x1","top","bottom","uri")):
                                if not (L["x1"]<x0 or L["x0"]>x1 or L["bottom"]<top or L["top"]>bottom):
                                    desc_links[r]=L["uri"]
                                    break
            rows=[]; row_links=[]; tbl_tot=None
            for ridx,row in enumerate(data[1:],start=1):
                cells=[str(c).strip() for c in row]
                if not any(cells): continue
                fc=cells[0].lower()
                if ("total" in fc or "subtotal" in fc) and any("$" in c for c in cells):
                    if tbl_tot is None: tbl_tot=cells
                    continue
                # build new row
                new=[""]*len(HEADERS)
                new[0]=cells[desc_i] if desc_i<len(cells) else ""
                # fill other columns by header name match or pattern
                for i,val in enumerate(cells):
                    v=val.strip()
                    if not v or i==desc_i: continue
                    lo=v.lower()
                    if any(x in hdr[i].lower() for x in ["start"]): new[1]=v;continue
                    if any(x in hdr[i].lower() for x in ["end"]):   new[2]=v;continue
                    if re.fullmatch(r"\d{1,3}",v) or "month" in lo or "mo" in lo:
                        if not new[3]: new[3]=v; continue
                    if "$" in v:
                        if not new[4]: new[4]=v
                        elif not new[5]: new[5]=v
                        continue
                    if any(x in hdr[i].lower() for x in ["note","comment"]):
                        new[6]=v; continue
                rows.append(new)
                row_links.append(desc_links.get(ridx))
            if tbl_tot is None:
                tbl_tot=find_total(pi)
            if rows:
                tables_info.append((HEADERS,rows,row_links,tbl_tot))
    # grand total
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
bs=ParagraphStyle("Body",fontName=DEFAULT_SANS_FONT,fontSize=9,alignment=TA_LEFT,leading=11)
els=[]

logo=None
try:
    r=requests.get("https://www.carnegiehighered.com/wp-content/uploads/2021/11/Twitter-Image-2-2021.png",timeout=10);r.raise_for_status()
    logo=r.content
    img=Image.open(io.BytesIO(logo));ratio=img.height/img.width
    w=min(5*inch,doc.width);h=w*ratio
    els.append(RLImage(io.BytesIO(logo),width=w,height=h))
except:
    pass

els += [Spacer(1,12), Paragraph(html.escape(proposal_title), ts), Spacer(1,24)]
tw=doc.width

for hdr,rows,links,tot in tables_info:
    n=len(hdr)
    idx=0
    ws=[tw*0.25,tw*0.08,tw*0.08,tw*0.08,tw*0.10,tw*0.10,tw*0.16]
    wrapped=[[Paragraph(html.escape(h),hs) for h in hdr]]
    for i,row in enumerate(rows):
        line=[]
        for j,cell in enumerate(row):
            txt=html.escape(cell)
            if j==idx and links[i]:
                line.append(Paragraph(f"{txt} <link href='{html.escape(links[i])}' color='blue'>- link</link>",bs))
            else:
                line.append(Paragraph(txt,bs))
        wrapped.append(line)
    if tot:
        lbl,val="Total",""
        if isinstance(tot,list):
            lbl=tot[0] or "Total";val=next((c for c in reversed(tot) if "$" in c),"")
        else:
            m=re.match(r'(.*?)\s*(\$[\d,]+\.\d{2})',tot)
            if m:lbl,val=m.group(1).strip(),m.group(2)
        wrapped.append([Paragraph(lbl,bs)]+[Paragraph("",bs)]*(n-2)+[Paragraph(val,bs)])
    tbl=LongTable(wrapped,colWidths=ws,repeatRows=1)
    cmds=[("BACKGROUND",(0,0),(-1,0),colors.HexColor("#F2F2F2")),
          ("GRID",(0,0),(-1,-1),0.25,colors.grey),
          ("VALIGN",(0,0),(-1,0),"MIDDLE"),
          ("VALIGN",(0,1),(-1,-1),"TOP")]
    if tot:
        cmds += [("SPAN",(0,-1),(-2,-1)),("ALIGN",(0,-1),(-2,-1),"LEFT"),
                 ("ALIGN",(-1,-1),(-1,-1),"RIGHT"),("VALIGN",(0,-1),(-1,-1),"MIDDLE")]
    tbl.setStyle(TableStyle(cmds))
    els += [tbl, Spacer(1,24)]

if grand_total:
    ws=[tw*0.25,tw*0.08,tw*0.08,tw*0.08,tw*0.10,tw*0.10,tw*0.16]
    row=[Paragraph("Grand Total",bs)]+[Paragraph("",bs)]*(len(HEADERS)-2)+[Paragraph(html.escape(grand_total),bs)]
    gt=LongTable([row],colWidths=ws)
    gt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E0E0E0")),
                            ("GRID",(0,0),(-1,-1),0.25,colors.grey),
                            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                            ("SPAN",(0,0),(-2,0)),("ALIGN",(-1,0),(-1,0),"RIGHT")]))
    els.append(gt)

doc.build(els)
pdf_buf.seek(0)

st.download_button("📥 Download PDF", data=pdf_buf, file_name="transformed_proposal.pdf", mime="application/pdf", use_container_width=True)
