import os
from io import BytesIO
from datetime import date, datetime

import pandas as pd
import streamlit as st
from sqlalchemy import create_engine, text

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.shared import RGBColor

# PDF (preview)
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# Font for Serbian diacritics in ReportLab
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# PDF->Images (preview)
try:
    import fitz  # pymupdf
    HAVE_PYMUPDF = True
except Exception:
    HAVE_PYMUPDF = False

from obracun import (
    procesiraj_obracun_iz_db,
    rules_to_df,
    RULE_TYPES,
    material_rules,
    generate_word_for_racun,  # tvoj postojeći Word za obračun
)

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Lager Lux", layout="wide")
st.title("📦 Lager Lux – Računi + Obračun (DB)")

DB_URL = os.getenv("DB_URL", "postgresql+psycopg2://postgres:postgres@localhost:5432/lager_lux")
engine = create_engine(DB_URL, pool_pre_ping=True)

# Putanja do slike zaglavlja koju koristiš i u final Word-u za obračun.
# Stavi fajl npr. u: C:\Users\hp\Desktop\lager kalkulator Branko\assets\header.png
HEADER_IMG_PATH = os.getenv("HEADER_IMG_PATH", "assets/header.png")

SERBIAN_TTF_CANDIDATES = [
    os.getenv("SERBIAN_TTF_PATH", ""),  # ako ručno setuješ env
    "C:/Windows/Fonts/DejaVuSans.ttf",
    "C:/Windows/Fonts/Arial.ttf",
    "assets/DejaVuSans.ttf",
]

def _register_serbian_font():
    # ReportLab: registruj font koji podržava č/ć/đ/š/ž
    for p in SERBIAN_TTF_CANDIDATES:
        p = (p or "").strip()
        if p and os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("SERBIAN", p))
                return "SERBIAN"
            except Exception:
                continue
    # fallback: Helvetica (ne garantuje sva slova)
    return "Helvetica"

SERB_FONT_NAME = _register_serbian_font()

# =========================
# DB HELPERS
# =========================
def db_df(sql: str, params=None) -> pd.DataFrame:
    with engine.connect() as c:
        return pd.read_sql(text(sql), c, params=params or {})

def db_exec(sql: str, params=None) -> None:
    with engine.begin() as c:
        c.execute(text(sql), params or {})

# =========================
# NUM HELPERS
# =========================
def _to_num(x, default=0.0) -> float:
    try:
        if x is None:
            return float(default)
        if pd.isna(x):
            return float(default)
        return float(x)
    except Exception:
        return float(default)

def _fmt_rs_money(x: float) -> str:
    # 12345.6 -> "12.345,60"
    s = f"{float(x):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

def _fmt_qty(x: float) -> str:
    s = f"{float(x):,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")

# =========================
# VAT / TOTALS HELPERS
# =========================
def calc_lines(df_items: pd.DataFrame, vat_rate: float) -> pd.DataFrame:
    out = df_items.copy()
    out["qty"] = out["qty"].apply(lambda v: _to_num(v, 0.0))
    out["unit_price"] = out.get("unit_price", 0).apply(lambda v: _to_num(v, 0.0))
    out["discount"] = out.get("discount", 0).apply(lambda v: _to_num(v, 0.0))

    out["line_net"] = (out["qty"] * out["unit_price"]) - out["discount"]
    out.loc[out["line_net"] < 0, "line_net"] = 0.0
    out["line_vat"] = out["line_net"] * float(vat_rate)
    out["line_gross"] = out["line_net"] + out["line_vat"]
    return out

def sum_totals(df_lines: pd.DataFrame) -> dict:
    return {
        "total_net": float(df_lines["line_net"].sum()) if not df_lines.empty else 0.0,
        "total_vat": float(df_lines["line_vat"].sum()) if not df_lines.empty else 0.0,
        "total_gross": float(df_lines["line_gross"].sum()) if not df_lines.empty else 0.0,
    }

# =========================
# WORD helpers
# =========================
def _set_cell_align(cell, h="center", v=True):
    for p in cell.paragraphs:
        if h == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif h == "right":
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    if v:
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

def _set_run_blue(run):
    run.font.color.rgb = RGBColor(0, 0, 153)  # tamno plava (možeš menjati)

def _add_header_image(doc: Document):
    if HEADER_IMG_PATH and os.path.exists(HEADER_IMG_PATH):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run()
        r.add_picture(HEADER_IMG_PATH, width=Inches(6.5))

# =========================
# WORD: "RAČUN" (Invoice form 1:1 logika, bez template-a)
# =========================
def generate_invoice_word_from_db(invoice_id: int) -> BytesIO:
    inv = db_df("""
        SELECT
            i.invoice_id, i.invoice_no, i.invoice_type, i.status,
            i.issue_date, i.due_date, i.currency, COALESCE(i.vat_rate,0) AS vat_rate,
            COALESCE(i.total_net,0) AS total_net,
            COALESCE(i.total_vat,0) AS total_vat,
            COALESCE(i.total_gross,0) AS total_gross,
            i.delivery_note_no,              -- broj otpremnice (ako postoji)
            i.place_of_issue,                -- mesto izdavanja (ako postoji)
            i.place_of_supply,               -- mesto prometa (ako postoji)
            i.supply_date,                   -- datum prometa (ako postoji)
            c.name AS client_name,
            c.address AS client_address,
            c.city AS client_city
        FROM invoices i
        JOIN clients c ON c.client_id = i.client_id
        WHERE i.invoice_id = :id
    """, {"id": invoice_id})
    if inv.empty:
        raise ValueError(f"Invoice not found: {invoice_id}")
    inv = inv.iloc[0].to_dict()

    items = db_df("""
        SELECT
            it.item_id,
            it.description,
            it.qty,
            it.uom,
            COALESCE(it.unit_price,0) AS unit_price,
            COALESCE(it.discount,0) AS discount,
            COALESCE(it.line_net,0) AS line_net
        FROM invoice_items it
        WHERE it.invoice_id = :id
        ORDER BY it.item_id
    """, {"id": invoice_id})

    # company settings (prodavac)
    cs = db_df("SELECT * FROM company_settings WHERE id=1")
    if cs.empty:
        db_exec("INSERT INTO company_settings(id, company_name) VALUES (1,'LASER - LUX d.o.o.')")
        cs = db_df("SELECT * FROM company_settings WHERE id=1")
    cs = cs.iloc[0].to_dict()

    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    # header image (isto kao u final wordu)
    _add_header_image(doc)

    # Title line (plavo)
    p = doc.add_paragraph(cs.get("company_name") or "LASER - LUX d.o.o.")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold = True
    run.font.size = Pt(16)
    _set_run_blue(run)

    doc.add_paragraph("")

    # 2 bloka: kupac (levo) + prodavac (desno)
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    left = t.cell(0, 0)
    right = t.cell(0, 1)

    # Kupac blok (1:1 logika)
    kupac_lines = []
    if inv.get("client_name"):
        kupac_lines.append(str(inv["client_name"]))
    if inv.get("client_address"):
        kupac_lines.append(str(inv["client_address"]))
    if inv.get("client_city"):
        kupac_lines.append(str(inv["client_city"]))
    left.text = "\n".join(kupac_lines) if kupac_lines else ""
    _set_cell_align(left, "left")

    # Prodavac blok (hardkodovan / company_settings)
    seller_lines = []
    if cs.get("address"): seller_lines.append(cs["address"])
    if cs.get("city"): seller_lines.append(cs["city"])
    if cs.get("country"): seller_lines.append(cs["country"])
    if cs.get("phone"): seller_lines.append(f"tel. {cs['phone']}")
    if cs.get("fax"): seller_lines.append(f"tel/fax: {cs['fax']}")
    if cs.get("email"): seller_lines.append(f"E-mail: {cs['email']}")
    if cs.get("bank_account"): seller_lines.append(f"Račun: {cs['bank_account']}")
    pib = cs.get("pib") or ""
    mb = cs.get("mb") or ""
    if pib or mb:
        seller_lines.append(f"PIB: {pib}    Matični broj: {mb}".strip())
    right.text = "\n".join(seller_lines)
    _set_cell_align(right, "left")

    doc.add_paragraph("")

    # meta info
    issue_date = inv.get("issue_date")
    due_date = inv.get("due_date")
    supply_date = inv.get("supply_date")
    place_issue = inv.get("place_of_issue") or "Beograd"
    place_supply = inv.get("place_of_supply") or "Beograd"
    delivery_note_no = inv.get("delivery_note_no")

    def dstr(d):
        if isinstance(d, (datetime, date)):
            return d.strftime("%d.%m.%Y.")
        return ""

    meta_lines = [
        f"Datum izdavanja računa: {dstr(issue_date)}",
        f"Mesto izdavanja računa: {place_issue}",
    ]
    if delivery_note_no:
        meta_lines.append(f"Broj otpremnice: {delivery_note_no}")
    if supply_date:
        meta_lines.append(f"Mesto i datum prometa dobara: {place_supply}, {dstr(supply_date)}")
    meta_lines.append(f"Rok za plaćanje: {dstr(due_date) if due_date else '/'}")

    for line in meta_lines:
        doc.add_paragraph(line)

    doc.add_paragraph("")

    # naslov RAČUN BR. (plavo)
    title = doc.add_paragraph(f"RAČUN  BR. {inv.get('invoice_no','')}")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.runs[0]
    r.bold = True
    r.font.size = Pt(14)
    _set_run_blue(r)

    doc.add_paragraph("Za nabavku materijala i to:")

    # tabela stavki (kao na slici)
    tbl = doc.add_table(rows=1, cols=6)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = tbl.rows[0].cells
    hdr[0].text = "R.br."
    hdr[1].text = "Ime proizvoda"
    hdr[2].text = "JM"
    hdr[3].text = "Količina"
    hdr[4].text = "Jed.cena"
    hdr[5].text = "Ukupno"

    for c in hdr:
        _set_cell_align(c, "center")
        for p in c.paragraphs:
            for rr in p.runs:
                rr.bold = True

    currency = inv.get("currency") or "RSD"

    for i, rrow in enumerate(items.to_dict("records"), start=1):
        row = tbl.add_row().cells
        row[0].text = str(i)
        row[1].text = str(rrow.get("description") or "")
        row[2].text = str(rrow.get("uom") or "")
        row[3].text = _fmt_qty(_to_num(rrow.get("qty"), 0.0))
        row[4].text = _fmt_rs_money(_to_num(rrow.get("unit_price"), 0.0))
        row[5].text = f"{_fmt_rs_money(_to_num(rrow.get('line_net'), 0.0))} {currency}"

        _set_cell_align(row[0], "center")
        _set_cell_align(row[1], "left")
        _set_cell_align(row[2], "center")
        _set_cell_align(row[3], "right")
        _set_cell_align(row[4], "right")
        _set_cell_align(row[5], "right")

    doc.add_paragraph("")

    total_net = _to_num(inv.get("total_net"), 0.0)
    total_vat = _to_num(inv.get("total_vat"), 0.0)
    total_gross = _to_num(inv.get("total_gross"), 0.0)
    vat_rate = _to_num(inv.get("vat_rate"), 0.0)

    p1 = doc.add_paragraph(f"UKUPNO : {_fmt_rs_money(total_net)} {currency}")
    p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p1.runs[0].bold = True

    p2 = doc.add_paragraph(f"OSNOVICA ZA OBRAČUN PDV: {_fmt_rs_money(total_net)} {currency}")
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p3 = doc.add_paragraph(f"PDV {int(vat_rate*100)}%: {_fmt_rs_money(total_vat)} {currency}")
    p3.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p4 = doc.add_paragraph(f"SVE UKUPNO ZA UPLATU: {_fmt_rs_money(total_gross)} {currency}")
    p4.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p4.runs[0].bold = True

    doc.add_paragraph("")
    doc.add_paragraph(f"Paritet isporuke: {cs.get('delivery_parity') or '/'}")
    doc.add_paragraph(f"Napomena o poreskom oslobađanju: {cs.get('tax_note') or '/'}")
    doc.add_paragraph(cs.get("payment_note") or f"Uplatu izvršiti na tekući račun br.: {cs.get('bank_account') or ''}")

    doc.add_paragraph("")
    sig = doc.add_paragraph(f"■ {cs.get('company_name') or 'LASER LUX'} ■ d.o.o.")
    sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# =========================
# PDF preview (ReportLab) + PDF->Images
# =========================
def build_invoice_pdf_preview(
    company_block: dict,
    client_block: dict,
    meta: dict,
    lines_df: pd.DataFrame,
    totals: dict,
    currency: str,
    vat_percent: float
) -> BytesIO:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setTitle("Racun Preview")

    # font
    c.setFont(SERB_FONT_NAME, 12)

    y = h - 20 * mm

    # header image (optional)
    if HEADER_IMG_PATH and os.path.exists(HEADER_IMG_PATH):
        try:
            c.drawImage(HEADER_IMG_PATH, 20 * mm, y - 18 * mm, width=170 * mm, height=18 * mm, preserveAspectRatio=True, mask='auto')
            y -= 22 * mm
        except Exception:
            pass

    # company title (blue-ish simulation: reportlab uses RGB)
    c.setFillColorRGB(0, 0, 0.6)
    c.setFont(SERB_FONT_NAME, 18)
    c.drawCentredString(w / 2, y, company_block.get("company_name", "LASER - LUX d.o.o."))
    c.setFillColorRGB(0, 0, 0)
    c.setFont(SERB_FONT_NAME, 12)
    y -= 14 * mm

    # blocks
    left_x = 20 * mm
    right_x = 110 * mm

    # client block
    client_lines = [client_block.get("name",""), client_block.get("address",""), client_block.get("city","")]
    client_lines = [x for x in client_lines if x]
    yy = y
    for line in client_lines:
        c.drawString(left_x, yy, str(line))
        yy -= 5 * mm

    # company block
    comp_lines = []
    for k in ["address","city","country"]:
        if company_block.get(k): comp_lines.append(company_block[k])
    if company_block.get("phone"): comp_lines.append(f"tel. {company_block['phone']}")
    if company_block.get("fax"): comp_lines.append(f"tel/fax: {company_block['fax']}")
    if company_block.get("email"): comp_lines.append(f"E-mail: {company_block['email']}")
    if company_block.get("bank_account"): comp_lines.append(f"Račun: {company_block['bank_account']}")
    pib = company_block.get("pib") or ""
    mb = company_block.get("mb") or ""
    if pib or mb:
        comp_lines.append(f"PIB: {pib}    Matični broj: {mb}".strip())

    yy2 = y
    for line in comp_lines:
        c.drawString(right_x, yy2, str(line))
        yy2 -= 5 * mm

    y = min(yy, yy2) - 8 * mm

    # meta
    meta_lines = [
        f"Datum izdavanja računa: {meta.get('issue_date_str','')}",
        f"Mesto izdavanja računa: {meta.get('place_of_issue','Beograd')}",
    ]
    if meta.get("delivery_note_no"):
        meta_lines.append(f"Broj otpremnice: {meta['delivery_note_no']}")
    if meta.get("supply_date_str"):
        meta_lines.append(f"Mesto i datum prometa dobara: {meta.get('place_of_supply','Beograd')}, {meta['supply_date_str']}")
    meta_lines.append(f"Rok za plaćanje: {meta.get('due_date_str','/')}")

    for line in meta_lines:
        c.drawString(left_x, y, line)
        y -= 6 * mm

    y -= 4 * mm

    # title
    c.setFillColorRGB(0, 0, 0.6)
    c.setFont(SERB_FONT_NAME, 16)
    c.drawCentredString(w / 2, y, f"RAČUN  BR. {meta.get('invoice_no','')}")
    c.setFillColorRGB(0, 0, 0)
    c.setFont(SERB_FONT_NAME, 12)
    y -= 10 * mm

    c.drawString(left_x, y, "Za nabavku materijala i to:")
    y -= 10 * mm

    # table layout
    col_w = [18*mm, 78*mm, 15*mm, 22*mm, 22*mm, 30*mm]
    headers = ["R.br.", "Ime proizvoda", "JM", "Količina", "Jed.cena", "Ukupno"]
    x0 = left_x
    row_h = 8 * mm

    def draw_row(cvs, ytop, values, bold=False):
        x = x0
        for i, val in enumerate(values):
            cvs.rect(x, ytop - row_h, col_w[i], row_h, stroke=1, fill=0)
            if bold:
                cvs.setFont(SERB_FONT_NAME, 11)
            else:
                cvs.setFont(SERB_FONT_NAME, 11)
            # alignment
            if i in (0, 2):
                cvs.drawCentredString(x + col_w[i]/2, ytop - 6*mm + 2, str(val))
            elif i in (3, 4, 5):
                cvs.drawRightString(x + col_w[i] - 2*mm, ytop - 6*mm + 2, str(val))
            else:
                cvs.drawString(x + 2*mm, ytop - 6*mm + 2, str(val))
            x += col_w[i]

    # header row
    draw_row(c, y, headers, bold=True)
    y -= row_h

    # lines
    for idx, r in enumerate(lines_df.to_dict("records"), start=1):
        uk = f"{_fmt_rs_money(_to_num(r.get('line_net',0)))} {currency}"
        draw_row(c, y, [
            idx,
            r.get("description",""),
            r.get("uom",""),
            _fmt_qty(_to_num(r.get("qty",0))),
            _fmt_rs_money(_to_num(r.get("unit_price",0))),
            uk
        ])
        y -= row_h
        if y < 40*mm:
            c.showPage()
            c.setFont(SERB_FONT_NAME, 12)
            y = h - 20*mm

    y -= 8 * mm

    # totals (right aligned)
    c.setFont(SERB_FONT_NAME, 12)
    c.drawRightString(w - 20*mm, y, f"UKUPNO : {_fmt_rs_money(totals['total_net'])} {currency}")
    y -= 7*mm
    c.drawRightString(w - 20*mm, y, f"OSNOVICA ZA OBRAČUN PDV: {_fmt_rs_money(totals['total_net'])} {currency}")
    y -= 7*mm
    c.drawRightString(w - 20*mm, y, f"PDV {int(vat_percent)}%: {_fmt_rs_money(totals['total_vat'])} {currency}")
    y -= 9*mm
    c.setFont(SERB_FONT_NAME, 12)
    c.drawRightString(w - 20*mm, y, f"SVE UKUPNO ZA UPLATU: {_fmt_rs_money(totals['total_gross'])} {currency}")
    y -= 12*mm

    # footer notes
    c.setFont(SERB_FONT_NAME, 11)
    c.drawString(left_x, y, f"Paritet isporuke: {company_block.get('delivery_parity') or '/'}")
    y -= 6*mm
    c.drawString(left_x, y, f"Napomena o poreskom oslobađanju: {company_block.get('tax_note') or '/'}")
    y -= 6*mm
    c.drawString(left_x, y, company_block.get("payment_note") or "")
    y -= 10*mm
    c.drawRightString(w - 20*mm, y, f"■ {company_block.get('company_name') or 'LASER LUX'} ■ d.o.o.")

    c.showPage()
    c.save()
    buf.seek(0)
    return buf

def pdf_to_images(pdf_bytes: bytes, zoom: float = 2.0):
    if not HAVE_PYMUPDF:
        return []
    imgs = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    mat = fitz.Matrix(zoom, zoom)
    for i in range(doc.page_count):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat)
        imgs.append(pix.tobytes("png"))
    return imgs

# =========================
# LOAD for obracun from DB
# =========================
def load_invoice_items_for_calc(invoice_id: int) -> pd.DataFrame:
    sql = """
    SELECT
      i.invoice_no AS "Broj računa",
      c.name AS "Kompanija",
      it.material_id AS "ID materijala",
      m.name AS "Materijal",
      it.qty AS "Količina za fakturisanje",
      it.uom AS "Jedinica mere za fakturisanje",
      m.uom AS "Jedinica mere za lager - skidanje količine",
      m.opis_materijala AS "Opis Materijala",
      m.tech_normative AS "Normativna potrošnja (tehnički list)",
      it.qty AS "Količina za fakturisanje (ono što piše u tabeli za račune - Normative)",
      it.uom AS "Jedinica mere za fakturisanje - u računu"
    FROM invoices i
    JOIN clients c ON c.client_id = i.client_id
    JOIN invoice_items it ON it.invoice_id = i.invoice_id
    JOIN materials m ON m.material_id = it.material_id
    WHERE i.invoice_id = :invoice_id
    ORDER BY it.item_id;
    """
    return db_df(sql, {"invoice_id": invoice_id})

# =========================
# UI
# =========================
tabs = st.tabs(["🧾 Novi račun", "📄 Računi", "🧮 Obračun + Word + Lager"])

# ======================================================
# TAB 1: Novi račun (Preview PDF + Save + Word)
# ======================================================
with tabs[0]:
    
    # =========================
    # TOP ROW: 3 columns (client edit / add client / company settings)
    # =========================
    st.markdown("### 🧩 Podešavanja (klijent / novi klijent / prodavac)")

    colL, colM, colR = st.columns(3)

    # ---------- 1) EDIT CLIENT ----------
    with colL:
        with st.expander("✏️ Uredi klijenta (iz baze)", expanded=False):

            clients = db_df("SELECT client_id, name FROM clients ORDER BY name")
            if clients.empty:
                st.warning("Nema klijenata u bazi.")
            else:
                edit_client_id = st.selectbox(
                    "Izaberi klijenta",
                    clients["client_id"].tolist(),
                    format_func=lambda x: clients.loc[clients["client_id"] == x, "name"].iloc[0],
                    key="edit_client_id_top"
                )

                crow = db_df("""
                    SELECT client_id, name, address, city, pib, mb, phone, email
                    FROM clients
                    WHERE client_id=:id
                """, {"id": int(edit_client_id)}).iloc[0].to_dict()

                with st.form("edit_client_form_top"):
                    name = st.text_input("Naziv", value=crow.get("name") or "")
                    c1, c2 = st.columns(2)
                    with c1:
                        address = st.text_input("Adresa", value=crow.get("address") or "")
                        city = st.text_input("Grad / Poštanski broj", value=crow.get("city") or "")
                        pib = st.text_input("PIB", value=crow.get("pib") or "")
                        mb = st.text_input("Matični broj", value=crow.get("mb") or "")
                    with c2:
                        phone = st.text_input("Telefon", value=crow.get("phone") or "")
                        email = st.text_input("E-mail", value=crow.get("email") or "")

                    save_edit = st.form_submit_button("💾 Sačuvaj izmene")

                if save_edit:
                    db_exec("""
                        UPDATE clients
                        SET name=:n, address=:a, city=:c, pib=:p, mb=:m, phone=:ph, email=:e
                        WHERE client_id=:id
                    """, {
                        "n": name.strip(),
                        "a": address.strip() or None,
                        "c": city.strip() or None,
                        "p": pib.strip() or None,
                        "m": mb.strip() or None,
                        "ph": phone.strip() or None,
                        "e": email.strip() or None,
                        "id": int(edit_client_id),
                    })
                    st.success("Klijent sačuvan ✅")
                    st.rerun()

    # ---------- 2) ADD NEW CLIENT ----------
    with colM:
        with st.expander("➕ Dodaj novog klijenta", expanded=False):
            with st.form("add_client_form_top"):
                n_name = st.text_input("Naziv*", value="")
                c1, c2 = st.columns(2)
                with c1:
                    n_address = st.text_input("Adresa", value="")
                    n_city = st.text_input("Grad / Poštanski broj", value="")
                    n_pib = st.text_input("PIB", value="")
                    n_mb = st.text_input("Matični broj", value="")
                with c2:
                    n_phone = st.text_input("Telefon", value="")
                    n_email = st.text_input("E-mail", value="")

                add_btn = st.form_submit_button("✅ Dodaj")

            if add_btn:
                if not n_name.strip():
                    st.error("Naziv je obavezan.")
                else:
                    db_exec("""
                        INSERT INTO clients(name, address, city, pib, mb, phone, email)
                        VALUES (:n, :a, :c, :p, :m, :ph, :e)
                    """, {
                        "n": n_name.strip(),
                        "a": n_address.strip() or None,
                        "c": n_city.strip() or None,
                        "p": n_pib.strip() or None,
                        "m": n_mb.strip() or None,
                        "ph": n_phone.strip() or None,
                        "e": n_email.strip() or None,
                    })
                    st.success("Klijent dodat ✅")
                    st.rerun()

    # ---------- 3) COMPANY SETTINGS / SELLER ----------
    with colR:
        with st.expander("🏬 Prodavac (podešavanja firme)", expanded=False):
            cs = db_df("SELECT * FROM company_settings ORDER BY id LIMIT 1")
            if cs.empty:
                db_exec("INSERT INTO company_settings(id, company_name) VALUES (1,'LASER - LUX d.o.o.')")
                cs = db_df("SELECT * FROM company_settings WHERE id=1")

            cs = cs.iloc[0].to_dict()

            with st.form("company_settings_form_top"):
                company_name = st.text_input("Naziv", value=cs.get("company_name") or "")
                address = st.text_input("Adresa", value=cs.get("address") or "")
                city = st.text_input("Grad / Poštanski broj", value=cs.get("city") or "")
                country = st.text_input("Država", value=cs.get("country") or "")
                phone = st.text_input("Telefon", value=cs.get("phone") or "")
                fax = st.text_input("Fax", value=cs.get("fax") or "")
                email = st.text_input("Email", value=cs.get("email") or "")
                bank_account = st.text_input("Tekući račun", value=cs.get("bank_account") or "")
                pib = st.text_input("PIB", value=cs.get("pib") or "")
                mb = st.text_input("Matični broj", value=cs.get("mb") or "")
                delivery_parity = st.text_input("Paritet isporuke", value=cs.get("delivery_parity") or "/")
                tax_note = st.text_input("Napomena o poreskom oslobađanju", value=cs.get("tax_note") or "/")
                payment_note = st.text_area(
                    "Napomena za uplatu",
                    value=cs.get("payment_note") or "Uplatu izvršiti na tekući račun br.: 205-508589-34 kod NLB komercijalna banka a.d.",
                    height=80
                )

                save_company = st.form_submit_button("💾 Sačuvaj firmu")

            if save_company:
                db_exec("""
                    UPDATE company_settings
                    SET company_name=:company_name,
                        address=:address,
                        city=:city,
                        country=:country,
                        phone=:phone,
                        fax=:fax,
                        email=:email,
                        bank_account=:bank_account,
                        pib=:pib,
                        mb=:mb,
                        delivery_parity=:delivery_parity,
                        tax_note=:tax_note,
                        payment_note=:payment_note
                    WHERE id=1
                """, {
                    "company_name": company_name.strip(),
                    "address": address.strip() or None,
                    "city": city.strip() or None,
                    "country": country.strip() or None,
                    "phone": phone.strip() or None,
                    "fax": fax.strip() or None,
                    "email": email.strip() or None,
                    "bank_account": bank_account.strip() or None,
                    "pib": pib.strip() or None,
                    "mb": mb.strip() or None,
                    "delivery_parity": delivery_parity.strip() or None,
                    "tax_note": tax_note.strip() or None,
                    "payment_note": payment_note.strip() or None,
                })
                st.success("Prodavac sačuvan ✅")
                st.rerun()

    st.markdown("---")

    st.subheader("🧾 Kreiraj račun (preview kao štampa + Word)")
    # ---- Company settings editor (vidljivo ovde)


    materials = db_df("SELECT material_id, name, uom FROM materials ORDER BY name")

    if clients.empty or materials.empty:
        st.warning("Nema klijenata ili materijala u bazi.")
    else:
        if "new_items" not in st.session_state:
            st.session_state["new_items"] = pd.DataFrame(
                columns=["material_id", "qty", "uom", "unit_price", "discount", "description"]
            )
        

        with st.form("new_invoice_form", clear_on_submit=False):
            col1, col2, col3, col4 = st.columns([2, 1, 1, 1])

            with col1:
                client_id = st.selectbox(
                    "Klijent",
                    clients["client_id"].tolist(),
                    format_func=lambda x: clients.loc[clients["client_id"] == x, "name"].iloc[0],
                )

                client_row = db_df(
                    "SELECT client_id, name, address, city, pib, mb, phone, email FROM clients WHERE client_id=:id",
                    {"id": int(client_id)}
                ).iloc[0].to_dict()

                

            with col2:
                inv_type = st.selectbox("Tip računa", ["ADVANCE", "FINAL"])
            with col3:
                invoice_no = st.text_input("Broj računa")
            with col4:
                issue_date = st.date_input("Datum", value=date.today())

            m1, m2, m3, m4 = st.columns([1, 1, 1, 1])
            with m1:
                vat_percent = st.number_input("PDV (%)", min_value=0.0, max_value=50.0, value=20.0, step=1.0)
            with m2:
                currency = st.selectbox("Valuta", ["RSD", "EUR"], index=0)
            with m3:
                due_date = st.date_input("Rok plaćanja", value=issue_date)
            with m4:
                delivery_note_no = st.text_input("Broj otpremnice", value="")

            m5, m6 = st.columns(2)
            with m5:
                place_of_issue = st.text_input("Mesto izdavanja", value="Beograd")
            with m6:
                place_of_supply = st.text_input("Mesto prometa", value="Beograd")

            supply_date = st.date_input("Datum prometa", value=issue_date)

            st.markdown("### Stavke")

            df_new = st.data_editor(
                st.session_state["new_items"],
                num_rows="dynamic",
                use_container_width=True,
                key="new_items_editor",
                column_config={
                    "material_id": st.column_config.SelectboxColumn(
                        "Materijal",
                        options=materials["material_id"].tolist(),
                        format_func=lambda x: materials.loc[materials["material_id"] == x, "name"].iloc[0]
                        if x in materials["material_id"].values else str(x),
                    ),
                    "qty": st.column_config.NumberColumn("Količina"),
                    "uom": st.column_config.TextColumn("JM (fakt)"),
                    "unit_price": st.column_config.NumberColumn("Jed.cena"),
                    "discount": st.column_config.NumberColumn("Popust (RSD)"),
                    "description": st.column_config.TextColumn("Opis (opciono)"),
                },
            )

            # autopopuni
            for idx, row in df_new.iterrows():
                mid = row.get("material_id")
                if pd.notna(mid) and mid in materials["material_id"].values:
                    if pd.isna(row.get("uom")) or str(row.get("uom")).strip() == "":
                        df_new.at[idx, "uom"] = materials.loc[materials["material_id"] == mid, "uom"].iloc[0]
                    if pd.isna(row.get("description")) or str(row.get("description")).strip() == "":
                        df_new.at[idx, "description"] = materials.loc[materials["material_id"] == mid, "name"].iloc[0]
                if pd.isna(row.get("discount")):
                    df_new.at[idx, "discount"] = 0

            st.session_state["new_items"] = df_new

            do_preview = st.form_submit_button("👀 Preview (kao štampa)")
            do_save = st.form_submit_button("💾 Sačuvaj u bazu")



        

        # Preview PDF (u expanderu) + download PDF
        if do_preview:
            work = df_new.copy()
            work = work[work["material_id"].notna() & work["qty"].notna()].copy()
            if work.empty:
                st.error("Dodaj bar jednu stavku (materijal + količina).")
            else:
                vat_rate = float(vat_percent) / 100.0
                work["unit_price"] = work.get("unit_price", 0).fillna(0)
                work["discount"] = work.get("discount", 0).fillna(0)
                lines = calc_lines(work, vat_rate)
                totals = sum_totals(lines)

                cs = db_df("SELECT * FROM company_settings WHERE id=1").iloc[0].to_dict()
                cli = db_df("SELECT name, address, city FROM clients WHERE client_id=:id", {"id": int(client_id)}).iloc[0].to_dict()

                meta = {
                    "invoice_no": str(invoice_no).strip(),
                    "issue_date_str": issue_date.strftime("%d.%m.%Y."),
                    "due_date_str": due_date.strftime("%d.%m.%Y.") if due_date else "/",
                    "place_of_issue": place_of_issue,
                    "delivery_note_no": delivery_note_no.strip() or None,
                    "place_of_supply": place_of_supply,
                    "supply_date_str": supply_date.strftime("%d.%m.%Y.") if supply_date else "",
                }

                pdf_buf = build_invoice_pdf_preview(
                    company_block=cs,
                    client_block=cli,
                    meta=meta,
                    lines_df=lines,
                    totals=totals,
                    currency=currency,
                    vat_percent=float(vat_percent)
                )
                st.session_state["last_preview_pdf"] = pdf_buf.getvalue()

                with st.expander("🖨️ Preview (kao štampa) — PDF", expanded=True):
                    if HAVE_PYMUPDF:
                        imgs = pdf_to_images(st.session_state["last_preview_pdf"], zoom=2.0)
                        if imgs:
                            for i, img in enumerate(imgs, start=1):
                                st.image(img, caption=f"Strana {i}", use_container_width=True)
                        else:
                            st.info("Nisam uspeo da renderujem slike iz PDF-a. Preuzmi PDF ispod.")
                    else:
                        st.warning("Za prikaz kao slike instaliraj: pip install pymupdf")

                    st.download_button(
                        "⬇️ Preuzmi PDF preview",
                        data=st.session_state["last_preview_pdf"],
                        file_name=f"preview_racun_{meta['invoice_no'] or 'draft'}.pdf",
                        mime="application/pdf",
                    )

        # Save in DB + allow Word generation
        if do_save:
            if not str(invoice_no).strip():
                st.error("Unesi broj računa.")
            else:
                work = df_new.copy()
                work = work[work["material_id"].notna() & work["qty"].notna()].copy()
                if work.empty:
                    st.error("Dodaj bar jednu stavku (materijal + količina).")
                else:
                    vat_rate = float(vat_percent) / 100.0
                    work["unit_price"] = work.get("unit_price", 0).fillna(0)
                    work["discount"] = work.get("discount", 0).fillna(0)
                    lines = calc_lines(work, vat_rate)
                    totals = sum_totals(lines)

                    with engine.begin() as c:
                        inv_id = c.execute(text("""
                            INSERT INTO invoices(
                                invoice_no, invoice_type, status, client_id,
                                issue_date, due_date, currency, vat_rate,
                                total_net, total_vat, total_gross,
                                delivery_note_no, place_of_issue, place_of_supply, supply_date
                            )
                            VALUES (
                                :no, :t, 'DRAFT', :cid,
                                :issue, :due, :cur, :vr,
                                :tn, :tv, :tg,
                                :dn, :pio, :ps, :sd
                            )
                            RETURNING invoice_id
                        """), {
                            "no": str(invoice_no).strip(),
                            "t": inv_type,
                            "cid": int(client_id),
                            "issue": issue_date,
                            "due": due_date,
                            "cur": currency,
                            "vr": vat_rate,
                            "tn": totals["total_net"],
                            "tv": totals["total_vat"],
                            "tg": totals["total_gross"],
                            "dn": delivery_note_no.strip() or None,
                            "pio": place_of_issue.strip() or None,
                            "ps": place_of_supply.strip() or None,
                            "sd": supply_date,
                        }).scalar()

                        for _, r in lines.iterrows():
                            c.execute(text("""
                                INSERT INTO invoice_items(
                                    invoice_id, material_id, description, qty, uom,
                                    unit_price, discount,
                                    line_net, line_vat, line_gross
                                )
                                VALUES (
                                    :iid, :mid, :desc, :qty, :uom,
                                    :up, :disc,
                                    :ln, :lv, :lg
                                )
                            """), {
                                "iid": int(inv_id),
                                "mid": int(r["material_id"]),
                                "desc": str(r.get("description") or ""),
                                "qty": float(_to_num(r.get("qty"), 0.0)),
                                "uom": str(r.get("uom") or ""),
                                "up": float(_to_num(r.get("unit_price"), 0.0)),
                                "disc": float(_to_num(r.get("discount"), 0.0)),
                                "ln": float(r["line_net"]),
                                "lv": float(r["line_vat"]),
                                "lg": float(r["line_gross"]),
                            })

                    st.success(f"Sačuvano ✅ (invoice_id={int(inv_id)})")
                    st.session_state["last_created_invoice_id"] = int(inv_id)

        # Word (račun) dugme u ovom tabu
        if st.session_state.get("last_created_invoice_id"):
            st.markdown("---")
            st.subheader("📄 Generiši Word račun (1:1)")

            inv_id = int(st.session_state["last_created_invoice_id"])

            c1, c2 = st.columns([1, 2])
            with c1:
                if st.button("📄 Generiši Word račun", key="gen_word_invoice_btn"):
                    try:
                        buf = generate_invoice_word_from_db(inv_id)
                        st.session_state["last_invoice_word"] = buf.getvalue()
                        st.success("Word je generisan ✅")
                    except Exception as e:
                        st.error(f"Greška pri generisanju Word-a: {e}")

            with c2:
                if st.session_state.get("last_invoice_word"):
                    st.download_button(
                        "⬇️ Preuzmi Word račun",
                        data=st.session_state["last_invoice_word"],
                        file_name=f"racun_{inv_id}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

# ======================================================
# TAB 2: Računi
# ======================================================
with tabs[1]:
    st.subheader("📄 Lista računa")

    inv = db_df("""
        SELECT i.invoice_id, i.invoice_no, i.invoice_type, i.status, i.issue_date, i.currency,
               COALESCE(i.total_net,0) AS total_net,
               COALESCE(i.total_vat,0) AS total_vat,
               COALESCE(i.total_gross,0) AS total_gross,
               c.name AS client
        FROM invoices i
        JOIN clients c ON c.client_id=i.client_id
        ORDER BY i.issue_date DESC NULLS LAST, i.invoice_id DESC
        LIMIT 1000
    """)
    st.dataframe(inv, use_container_width=True)

    pick = st.selectbox(
        "Detalji računa",
        inv["invoice_id"].tolist() if not inv.empty else [],
        format_func=lambda x: f"{inv.loc[inv.invoice_id==x,'invoice_no'].iloc[0]} — {inv.loc[inv.invoice_id==x,'client'].iloc[0]}",
    )

    if pick:
        items = db_df("""
            SELECT item_id, material_id, description, qty, uom,
                   COALESCE(unit_price,0) AS unit_price,
                   COALESCE(discount,0) AS discount,
                   COALESCE(line_net,0) AS line_net,
                   COALESCE(line_vat,0) AS line_vat,
                   COALESCE(line_gross,0) AS line_gross
            FROM invoice_items
            WHERE invoice_id = :id
            ORDER BY item_id
        """, {"id": int(pick)})

        st.markdown("### Stavke")
        st.dataframe(items, use_container_width=True)

        meta = inv.loc[inv.invoice_id == pick].iloc[0]
        st.info(
            f"NET: {meta['total_net']:.2f} {meta['currency']} | "
            f"PDV: {meta['total_vat']:.2f} {meta['currency']} | "
            f"BRUTO: {meta['total_gross']:.2f} {meta['currency']}"
        )

# ======================================================
# TAB 3: Obračun + Word + Lager (bulk u expanderu + search + rules u tab-u)
# ======================================================
with tabs[2]:
    st.subheader("🧮 Obračun + Word + Lager (DB)")

    # 2 kolone: levo obracun, desno pravila
    left_col, right_col = st.columns([2, 1])

    with right_col:
        st.markdown("### 🧮 Pravila (koeficijenti) — samo za Obračun tab")

        if "rules_df_calc" not in st.session_state:
            st.session_state["rules_df_calc"] = rules_to_df(material_rules)

        edited_rules_df = st.data_editor(
            st.session_state["rules_df_calc"],
            use_container_width=True,
            num_rows="dynamic",
            height=420,
            column_config={
                "rule_type": st.column_config.SelectboxColumn("rule_type", options=RULE_TYPES),
                "enabled": st.column_config.CheckboxColumn("enabled"),
                "factor": st.column_config.NumberColumn("factor"),
                "extra": st.column_config.NumberColumn("extra"),
            },
            key="rules_editor_calc_only",
        )

        if st.button("💾 Sačuvaj pravila (Obračun)", key="save_rules_calc_btn"):
            st.session_state["rules_df_calc"] = edited_rules_df
            st.success("Sačuvano ✅")

    with left_col:
        inv_pick = db_df("""
            SELECT i.invoice_id, i.invoice_no, i.invoice_type, i.status, i.stock_posted, c.name AS client
            FROM invoices i
            JOIN clients c ON c.client_id=i.client_id
            ORDER BY i.invoice_id DESC
            LIMIT 2000
        """)

        if inv_pick.empty:
            st.info("Nema računa u bazi.")
        else:
            with st.expander("📦 Bulk obračun (checkbox + select all) — preporučeno za DRAFT", expanded=False):
                show_draft_only = st.checkbox("Prikaži samo DRAFT", value=True, key="draft_only_bulk")
                q = st.text_input("🔎 Pretraga (broj računa ili klijent)", value="", key="bulk_search")

                view_df = inv_pick.copy()
                if show_draft_only:
                    view_df = view_df[view_df["status"] == "DRAFT"].copy()
                if q.strip():
                    qq = q.strip().lower()
                    view_df = view_df[
                        view_df["invoice_no"].astype(str).str.lower().str.contains(qq, na=False)
                        | view_df["client"].astype(str).str.lower().str.contains(qq, na=False)
                    ].copy()

                view_df = view_df.head(300)  # UI limit da ne ubije stranicu

                if "bulk_sel" not in st.session_state:
                    st.session_state["bulk_sel"] = {}

                for iid in view_df["invoice_id"].tolist():
                    st.session_state["bulk_sel"].setdefault(int(iid), False)

                b1, b2 = st.columns([1, 1])
                with b1:
                    if st.button("✅ Select all", key="bulk_all_btn"):
                        for iid in view_df["invoice_id"].tolist():
                            st.session_state["bulk_sel"][int(iid)] = True
                        st.rerun()
                with b2:
                    if st.button("⬜ Select none", key="bulk_none_btn"):
                        for iid in view_df["invoice_id"].tolist():
                            st.session_state["bulk_sel"][int(iid)] = False
                        st.rerun()

                sel_rows = []
                for _, r in view_df.iterrows():
                    iid = int(r["invoice_id"])
                    label = f'{r["invoice_no"]} — {r["client"]} ({r["invoice_type"]}, {r["status"]})'
                    checked = st.checkbox(label, value=st.session_state["bulk_sel"][iid], key=f"chk_{iid}")
                    st.session_state["bulk_sel"][iid] = bool(checked)
                    if checked:
                        sel_rows.append(iid)

                st.write(f"Izabrano računa: **{len(sel_rows)}**")

                colX, colY = st.columns([1, 1])
                with colX:
                    if st.button("🧮 Pokreni obračun za izabrane", key="bulk_run_btn"):
                        if not sel_rows:
                            st.error("Nisi izabrao nijedan račun.")
                        else:
                            results = []
                            for iid in sel_rows:
                                try:
                                    df_items = load_invoice_items_for_calc(int(iid))
                                    df_posle, _ = procesiraj_obracun_iz_db(
                                        df_items,
                                        edited_rules_df=st.session_state.get("rules_df_calc")
                                    )
                                    st.session_state.setdefault("bulk_df_posle", {})
                                    st.session_state["bulk_df_posle"][int(iid)] = df_posle
                                    results.append((iid, "OK", len(df_posle)))
                                except Exception as e:
                                    results.append((iid, f"ERR: {e}", 0))

                            st.success("Bulk obračun završen ✅")
                            st.dataframe(pd.DataFrame(results, columns=["invoice_id", "status", "rows"]), use_container_width=True)

                with colY:
                    if st.button("💾 Upiši invoice_allocations za izabrane", key="bulk_write_btn"):
                        if not sel_rows:
                            st.error("Nisi izabrao nijedan račun.")
                        else:
                            if "bulk_df_posle" not in st.session_state:
                                st.error("Prvo pokreni obračun za izabrane.")
                            else:
                                ok, err = 0, 0
                                for iid in sel_rows:
                                    df_posle = st.session_state["bulk_df_posle"].get(int(iid))
                                    if df_posle is None:
                                        err += 1
                                        continue

                                    try:
                                        db_exec("DELETE FROM invoice_allocations WHERE invoice_id=:id", {"id": int(iid)})
                                        with engine.begin() as c:
                                            for _, rr in df_posle.iterrows():
                                                mid = rr.get("ID materijala")
                                                if pd.isna(mid):
                                                    continue
                                                qtyw = rr.get("Kolicina za skidanje sa uracunatim Koef_novi za ovaj materijal")
                                                uomw = rr.get("Jedinica mere za lager - skidanje količine")
                                                note = rr.get("Napomena konverzije", "")

                                                c.execute(text("""
                                                    INSERT INTO invoice_allocations(invoice_id, material_id, qty_from_warehouse, uom, note)
                                                    VALUES (:iid, :mid, :q, :u, :n)
                                                """), {
                                                    "iid": int(iid),
                                                    "mid": int(float(mid)),
                                                    "q": float(qtyw) if pd.notna(qtyw) else 0.0,
                                                    "u": str(uomw) if pd.notna(uomw) else None,
                                                    "n": str(note)[:1000],
                                                })
                                        ok += 1
                                    except Exception:
                                        err += 1

                                st.success(f"Upisano ✅ OK={ok}, ERR={err}")

            st.markdown("---")
            st.markdown("### Pojedinačni rad (1 račun)")

            inv_id = st.selectbox(
                "Izaberi račun",
                inv_pick["invoice_id"].tolist(),
                format_func=lambda x: (
                    f'{inv_pick.loc[inv_pick["invoice_id"]==x, "invoice_no"].iloc[0]} — '
                    f'{inv_pick.loc[inv_pick["invoice_id"]==x, "client"].iloc[0]} '
                    f'({inv_pick.loc[inv_pick["invoice_id"]==x, "invoice_type"].iloc[0]}, '
                    f'{inv_pick.loc[inv_pick["invoice_id"]==x, "status"].iloc[0]})'
                ),
            )

            meta = inv_pick.loc[inv_pick["invoice_id"] == inv_id].iloc[0]
            is_final = meta["invoice_type"] == "FINAL"
            posted = bool(meta.get("stock_posted"))

            cA, cB, cC = st.columns([1, 1, 2])

            with cA:
                if st.button("🧮 Pokreni obračun (1 račun)", key="one_calc_btn"):
                    df_items = load_invoice_items_for_calc(int(inv_id))
                    df_posle, _ = procesiraj_obracun_iz_db(
                        df_items,
                        edited_rules_df=st.session_state.get("rules_df_calc")
                    )
                    st.session_state["df_posle_db"] = df_posle
                    st.success("Obračun završen ✅")

            with cB:
                if st.button("💾 Upisi obračun u bazu (invoice_allocations)", key="one_write_alloc_btn"):
                    if "df_posle_db" not in st.session_state:
                        st.error("Prvo pokreni obračun.")
                    else:
                        df_posle = st.session_state["df_posle_db"]
                        db_exec("DELETE FROM invoice_allocations WHERE invoice_id=:id", {"id": int(inv_id)})

                        with engine.begin() as c:
                            for _, rr in df_posle.iterrows():
                                mid = rr.get("ID materijala")
                                if pd.isna(mid):
                                    continue
                                qtyw = rr.get("Kolicina za skidanje sa uracunatim Koef_novi za ovaj materijal")
                                uomw = rr.get("Jedinica mere za lager - skidanje količine")
                                note = rr.get("Napomena konverzije", "")

                                c.execute(text("""
                                    INSERT INTO invoice_allocations(invoice_id, material_id, qty_from_warehouse, uom, note)
                                    VALUES (:iid, :mid, :q, :u, :n)
                                """), {
                                    "iid": int(inv_id),
                                    "mid": int(float(mid)),
                                    "q": float(qtyw) if pd.notna(qtyw) else 0.0,
                                    "u": str(uomw) if pd.notna(uomw) else None,
                                    "n": str(note)[:1000],
                                })

                        st.success("Upisano ✅")

            with cC:
                if "df_posle_db" in st.session_state:
                    df_posle = st.session_state["df_posle_db"]
                    st.dataframe(df_posle, use_container_width=True)

                    broj = str(df_posle["Broj računa"].iloc[0])
                    word_buf = generate_word_for_racun(df_posle, broj)
                    st.download_button(
                        "📄 Preuzmi Word (obračun)",
                        data=word_buf,
                        file_name=f"obracun_{broj}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )

            st.markdown("---")

            cnt = db_df("SELECT count(*) AS n FROM invoice_allocations WHERE invoice_id = :id", {"id": int(inv_id)})
            has_alloc = int(cnt.iloc[0]["n"]) > 0

            if is_final:
                if posted:
                    st.info("📦 Lager je već proknjižen za ovaj FINAL račun.")
                else:
                    if st.button("📦 Skini sa lagera (FINAL POST)", type="primary", key="final_post_btn"):
                        if not has_alloc:
                            st.error("Prvo upiši obračun u invoice_allocations.")
                        else:
                            try:
                                db_exec("SELECT post_invoice_stock(:id)", {"id": int(inv_id)})
                                st.success("Lager uspešno ažuriran ✅")
                            except Exception as e:
                                st.error(f"Greška pri knjiženju lagera: {e}")
            else:
                st.warning("Skidanje sa lagera je dozvoljeno samo za FINAL račun.")
