"""
PDF Invoice generator — matches PODDAR STORES Excel format exactly.
"""
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer, HRFlowable, Image)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from num2words import num2words
from datetime import datetime
import math

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SIGN_PATH = os.path.join(BASE_DIR, "assets", "sign.png")

# ─── colours ────────────────────────────────────────────────────────────────
BLUE_HDR   = colors.HexColor("#1F4E79")
LIGHT_BLUE = colors.HexColor("#BDD7EE")
TABLE_HDR  = colors.HexColor("#2E75B6")
ROW_ALT    = colors.HexColor("#DEEAF1")
BLACK      = colors.black
WHITE      = colors.white

def _rupees_words(amount):
    try:
        amt = int(round(amount))
        w = num2words(amt, lang='en_IN').title()
        return f"Rupees {w} Only"
    except:
        return ""

def _fmt(v, dec=2):
    if v == 0: return "0"
    if v == int(v): return str(int(v))
    return f"{v:.{dec}f}".rstrip('0').rstrip('.')

def generate_invoice_pdf(invoice: dict, output_path: str):
    """invoice keys: invoice_no, date, month, party, items[]"""
    W, H = A4
    c = canvas.Canvas(output_path, pagesize=A4)

    party  = invoice["party"]
    items  = invoice["items"]
    inv_no = invoice["invoice_no"]
    inv_dt = invoice["date"]
    month  = invoice["month"]

    # ── compute totals ──────────────────────────────────────────────────────
    for it in items:
        it["amount"]     = it["qty"] * it["rate"]
        it["gst_amount"] = round(it["amount"] * it["gst"], 4)

    subtotal   = sum(i["amount"]     for i in items)
    total_gst  = sum(i["gst_amount"] for i in items)
    grand_pre  = subtotal + total_gst
    grand      = round(grand_pre)
    round_off  = round(grand - grand_pre, 2)
    cgst = sgst = round(total_gst / 2, 2)

    # GST slab breakdown
    gst5  = sum(i["gst_amount"] for i in items if i["gst"] == 0.05)
    gst12 = sum(i["gst_amount"] for i in items if i["gst"] == 0.12)
    gst18 = sum(i["gst_amount"] for i in items if i["gst"] == 0.18)
    gst28 = sum(i["gst_amount"] for i in items if i["gst"] == 0.28)

    margin = 10*mm
    x0 = margin; x1 = W - margin
    eff_w = x1 - x0
    y = H - 8*mm

    def line(y_pos): c.line(x0, y_pos, x1, y_pos)
    def box(x,y,w,h): c.rect(x, y, w, h)

    # ── outer border ────────────────────────────────────────────────────────
    c.setLineWidth(1.5)
    c.rect(x0, 10*mm, eff_w, H-18*mm)

    # ── TAX INVOICE title ───────────────────────────────────────────────────
    c.setFont("Helvetica", 8); c.setFillColor(BLACK)
    c.drawCentredString(W/2, y, "TAX INVOICE"); y -= 5*mm

    c.setFont("Helvetica-Bold", 18); c.setFillColor(BLUE_HDR)
    c.drawCentredString(W/2, y, "PODDAR STORES"); y -= 6*mm

    c.setFont("Helvetica", 8); c.setFillColor(BLACK)
    c.drawCentredString(W/2, y, "MADHPUR, MABBI, DARBHANGA"); y -= 4*mm
    c.setLineWidth(0.5); line(y); y -= 3*mm

    # ── header info (left GSTIN/UDYAM, right Month/Invoice) ─────────────────
    lx = x0 + 2*mm; rx = W/2 + 2*mm
    c.setFont("Helvetica", 7)
    c.drawString(lx, y, "GSTIN 10BYJPP6207N1Z8")
    c.drawString(rx, y, f"Month"); c.drawString(rx+25*mm, y, month.upper()); y -= 4*mm
    c.drawString(lx, y, "UDYAM:- UDYAM-BR-6010643")
    c.drawString(rx, y, "Invoice No:"); c.drawString(rx+25*mm, y, inv_no); y -= 4*mm
    c.drawString(lx, y, "SCHOOL NAME"); c.drawString(lx+25*mm, y, party["name"])
    c.drawString(rx, y, "Date:"); c.drawString(rx+25*mm, y, inv_dt); y -= 4*mm
    c.drawString(lx, y, "ADDRESS"); c.drawString(lx+25*mm, y, party["address"])
    c.drawString(lx+55*mm, y, "BLOCK"); c.drawString(lx+70*mm, y, party.get("block",""))
    y -= 4*mm
    c.drawString(lx, y, "DISTRIC"); c.drawString(lx+25*mm, y, party["district"])
    c.drawString(lx+55*mm, y, "STATE"); c.drawString(lx+70*mm, y, party.get("state",""))
    y -= 3*mm; line(y); y -= 1*mm

    # ── items table ─────────────────────────────────────────────────────────
    col_w = [8*mm, 45*mm, 12*mm, 12*mm, 11*mm, 13*mm, 8*mm, 18*mm, 18*mm]
    hdr = ["Sr\nNo","Item","HSN","Qty","Unit","Rate","GST","GST\nAMOUNT","Amount"]

    tbl_data = [hdr]
    for idx, it in enumerate(items, 1):
        tbl_data.append([
            str(idx),
            it["name"],
            it.get("hsn",""),
            _fmt(it["qty"]),
            it.get("unit",""),
            _fmt(it["rate"]),
            f"{int(it['gst']*100)}%",
            _fmt(it["gst_amount"]),
            _fmt(it["amount"])
        ])

    # pad to 30 rows minimum
    while len(tbl_data) < 32:
        tbl_data.append(["","","","","","","","",""])

    tbl = Table(tbl_data, colWidths=col_w, repeatRows=1)
    style = TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), TABLE_HDR),
        ("TEXTCOLOR",   (0,0), (-1,0), WHITE),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,0), 7),
        ("ALIGN",       (0,0), (-1,0), "CENTER"),
        ("VALIGN",      (0,0), (-1,-1),"MIDDLE"),
        ("FONTNAME",    (0,1), (-1,-1),"Helvetica"),
        ("FONTSIZE",    (0,1), (-1,-1), 7),
        ("ALIGN",       (0,1), (0,-1), "CENTER"),
        ("ALIGN",       (3,1), (-1,-1),"RIGHT"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE, ROW_ALT]),
        ("GRID",        (0,0), (-1,-1), 0.3, colors.grey),
        ("LINEBELOW",   (0,0), (-1,0), 1, BLACK),
        ("TOPPADDING",  (0,0), (-1,-1), 1),
        ("BOTTOMPADDING",(0,0),(-1,-1), 1),
    ])
    tbl.setStyle(style)

    tbl_h = (len(tbl_data)) * 5.5*mm
    tbl.wrapOn(c, eff_w, tbl_h)
    tbl.drawOn(c, x0, y - tbl_h); y -= tbl_h; y -= 1*mm; line(y); y -= 2*mm

    # ── bottom section ───────────────────────────────────────────────────────
    bot_y = y
    lw = eff_w * 0.52   # left: bank + T&C
    mw = eff_w * 0.24   # middle: GST breakdown
    rw = eff_w * 0.24   # right: totals

    lx2 = x0; mx2 = x0 + lw; rx2 = mx2 + mw

    # Bank details (left)
    c.setFont("Helvetica-Bold", 7); c.drawString(lx2+1*mm, y, "Bank Details:")
    c.setFont("Helvetica", 7); y -= 3.5*mm
    c.drawString(lx2+1*mm, y, "BANK OF MAHARASHTRA A/C NO.  6044133923637"); y -= 3.5*mm
    c.drawString(lx2+1*mm, y, "IFSC: MAHB0001711"); y -= 3.5*mm
    c.setFont("Helvetica-Bold", 7); c.drawString(lx2+1*mm, y, "Terms & Conditions:")
    c.setFont("Helvetica", 6.5); y -= 3*mm
    terms = [
        "1. This is a computer generated invoice.",
        "2. Goods once sold will not be taken back.",
        "3. Subject to Darbhanga jurisdiction only.",
        "4. Payment should be made within due date.",
        "5. This is combined of multiple bill like monthly or weekly."
    ]
    for t in terms:
        c.drawString(lx2+1*mm, y, t); y -= 3*mm

    # GST breakdown (middle)
    gy = bot_y
    c.setFont("Helvetica-Bold", 7)
    c.drawString(mx2+1*mm, gy, "GST%"); c.drawString(mx2+12*mm, gy, "GST Total"); gy -= 3.5*mm
    c.setFont("Helvetica", 7)
    rows_gst = [("5%", gst5),("12%", gst12),("18%", gst18),("28%", gst28)]
    for label, val in rows_gst:
        c.drawString(mx2+1*mm, gy, label)
        c.drawRightString(mx2+mw-2*mm, gy, f"{val:.2f}"); gy -= 3.5*mm
    c.setFont("Helvetica-Bold", 7)
    c.drawString(mx2+1*mm, gy, "TOTAL")
    c.drawRightString(mx2+mw-2*mm, gy, f"{total_gst:.2f}")

    # Totals (right)
    ry = bot_y
    c.setFont("Helvetica-Bold", 7)
    rows_tot = [("TOTAL", f"{subtotal:.2f}"),("CSGT", f"{cgst:.2f}"),
                ("SGST", f"{sgst:.2f}"),("ROUND OFF", f"{round_off:.2f}")]
    for label, val in rows_tot:
        c.drawString(rx2+1*mm, ry, label)
        c.drawRightString(rx2+rw-2*mm, ry, val); ry -= 3.5*mm

    # amount in words
    words_y = min(y, gy, ry) - 3*mm
    c.setLineWidth(0.5); c.line(x0, words_y+3*mm, x1, words_y+3*mm)
    c.setFont("Helvetica-Bold", 7.5)
    c.drawString(x0+1*mm, words_y, _rupees_words(grand))

    # Grand total box (right bottom)
    gt_y = words_y - 1*mm
    c.setFont("Helvetica-Bold", 8)
    c.drawString(rx2+1*mm, gt_y+0.5*mm, "Grand Total")
    c.setFillColor(BLUE_HDR)
    c.rect(rx2, gt_y-3*mm, rw, 5*mm, fill=1)
    c.setFillColor(WHITE); c.setFont("Helvetica-Bold", 9)
    c.drawCentredString(rx2+rw/2, gt_y-1.5*mm, str(grand))
    c.setFillColor(BLACK)

    c.line(x0, words_y-4*mm, x1, words_y-4*mm)

    # ── signature section ────────────────────────────────────────────────────
    sig_y = words_y - 5*mm
    c.setFont("Helvetica", 7)
    c.drawString(x0+2*mm, sig_y, "Customer Signature")
    c.setFont("Helvetica-Bold", 7)
    c.drawRightString(x1-2*mm, sig_y, "Authorized Signatory")

    sig_y -= 12*mm
    if os.path.exists(SIGN_PATH):
        try:
            c.drawImage(SIGN_PATH, x1-45*mm, sig_y, width=43*mm, height=14*mm,
                       preserveAspectRatio=True, mask='auto')
        except: pass

    sig_y -= 3*mm
    c.setFont("Helvetica", 6.5)
    c.drawString(x0+2*mm, sig_y, "E. & O.E")
    c.drawCentredString(W/2, sig_y, "This is a computer generated invoice. No physical signature required.")
    c.setFont("Helvetica-Bold", 7.5)
    c.drawRightString(x1-2*mm, sig_y, "PODDAR STORES")

    c.save()
    return output_path
