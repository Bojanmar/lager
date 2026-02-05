import sys
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

DB_URL = "postgresql+psycopg2://postgres:postgres@localhost:5432/lager_lux"
SHEET_NAME = "IZLAZ 2025 - Final"

def parse_date(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = s.replace("..", ".")
    s = s.rstrip(".")
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    return None

def map_type(t):
    if t is None or pd.isna(t):
        return None
    s = str(t).strip().lower()
    if "avans" in s:
        return "ADVANCE"
    if "kona" in s:
        return "FINAL"
    return None

def norm_str(x):
    if x is None or pd.isna(x):
        return None
    s = str(x).strip()
    if s.lower() == "nan" or s == "":
        return None
    return s

def main(xlsx_path: str):
    engine = create_engine(DB_URL, pool_pre_ping=True)

    df = pd.read_excel(xlsx_path, sheet_name=SHEET_NAME)
    df.columns = [str(c).strip() for c in df.columns]

    needed = [
        "Tip račuma",
        "Broj računa",
        "Datum",
        "Kompanija",
        "ID materijala",
        "Materijal",
        "Količina za fakturisanje",
        "Jedinica mere za fakturisanje",
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise SystemExit(f"Fale kolone u Excel-u: {missing}")

    # (opciono) kolone za enrichment materijala
    has_opis = "Opis Materijala" in df.columns
    has_norm = "Normativna potrošnja (tehnički list)" in df.columns

    df = df.copy()
    df["invoice_no"] = df["Broj računa"].astype(str).str.strip()
    df["client_name"] = df["Kompanija"].astype(str).str.strip()
    df["invoice_type"] = df["Tip račuma"].apply(map_type)
    df["issue_date"] = df["Datum"].apply(parse_date)

    df["material_id"] = pd.to_numeric(df["ID materijala"], errors="coerce")
    df["qty"] = pd.to_numeric(df["Količina za fakturisanje"], errors="coerce")
    df["uom"] = df["Jedinica mere za fakturisanje"].astype(str).str.strip()
    df["material_name"] = df["Materijal"].astype(str).str.strip()

    if has_opis:
        df["opis_materijala"] = df["Opis Materijala"].apply(norm_str)
    else:
        df["opis_materijala"] = None

    if has_norm:
        df["tech_normative"] = df["Normativna potrošnja (tehnički list)"].apply(norm_str)
    else:
        df["tech_normative"] = None

    df_items = df[
        df["invoice_type"].notna()
        & df["invoice_no"].notna()
        & (df["invoice_no"].str.len() > 0)
        & df["material_id"].notna()
        & df["qty"].notna()
    ].copy()

    if df_items.empty:
        raise SystemExit("Nema stavki za import (proveri ID materijala / količine).")

    invoices = (
        df_items.groupby(["invoice_no", "invoice_type", "client_name", "issue_date"], dropna=False)
        .size()
        .reset_index(name="n_items")
        .sort_values(["issue_date", "invoice_no"])
    )

    # ====== napravi dataset za materials enrichment (po material_id) ======
    # uzmi "prvu" nepraznu vrednost po koloni
    df_mat = df_items[["material_id", "material_name", "uom", "opis_materijala", "tech_normative"]].copy()
    df_mat["material_id"] = df_mat["material_id"].astype("Int64")

    def first_nonnull(series):
        for v in series:
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                s = str(v).strip()
                if s != "" and s.lower() != "nan":
                    return s
        return None

    df_mat = (
        df_mat.sort_values(["material_id"])
        .groupby("material_id", dropna=True, as_index=False)
        .agg({
            "material_name": first_nonnull,
            "uom": first_nonnull,
            "opis_materijala": first_nonnull,
            "tech_normative": first_nonnull,
        })
    )

    created_invoices = 0
    created_items = 0
    created_clients = 0
    upserted_materials = 0

    with engine.begin() as c:
        # 1) UPSERT materials (popuni opis/normativ samo ako su prazni u bazi)
        for _, r in df_mat.iterrows():
            mid = int(r["material_id"])
            name = norm_str(r["material_name"])
            uom = norm_str(r["uom"])
            opis = norm_str(r["opis_materijala"])
            tech = norm_str(r["tech_normative"])

            # ako baš nema ničega korisnog, preskoči
            if not any([name, uom, opis, tech]):
                continue

            c.execute(text("""
                INSERT INTO materials(material_id, name, uom, opis_materijala, tech_normative)
                VALUES (:mid, :name, :uom, :opis, :tech)
                ON CONFLICT (material_id) DO UPDATE
                SET
                  name = CASE
                    WHEN materials.name IS NULL OR materials.name = '' THEN COALESCE(EXCLUDED.name, materials.name)
                    ELSE materials.name
                  END,
                  uom = CASE
                    WHEN materials.uom IS NULL OR materials.uom = '' THEN COALESCE(EXCLUDED.uom, materials.uom)
                    ELSE materials.uom
                  END,
                  opis_materijala = CASE
                    WHEN materials.opis_materijala IS NULL OR materials.opis_materijala = '' THEN EXCLUDED.opis_materijala
                    ELSE materials.opis_materijala
                  END,
                  tech_normative = CASE
                    WHEN materials.tech_normative IS NULL OR materials.tech_normative = '' THEN EXCLUDED.tech_normative
                    ELSE materials.tech_normative
                  END
            """), {"mid": mid, "name": name, "uom": uom, "opis": opis, "tech": tech})

            upserted_materials += 1

        # helper: upsert client by name
        def get_or_create_client_id(name: str) -> int:
            nonlocal created_clients
            row = c.execute(text("SELECT client_id FROM clients WHERE name = :n"), {"n": name}).fetchone()
            if row:
                return int(row[0])
            new_id = c.execute(
                text("INSERT INTO clients(name) VALUES (:n) RETURNING client_id"),
                {"n": name}
            ).scalar()
            created_clients += 1
            return int(new_id)

        # helper: get or create invoice
        def get_or_create_invoice_id(no: str, inv_type: str, cid: int, d):
            nonlocal created_invoices
            row = c.execute(text("SELECT invoice_id FROM invoices WHERE invoice_no = :no"), {"no": no}).fetchone()
            if row:
                return int(row[0])

            new_id = c.execute(text("""
                INSERT INTO invoices(invoice_no, invoice_type, status, client_id, issue_date, currency, vat_rate)
                VALUES (:no, :t, 'DRAFT', :cid, :d, 'RSD', 0.000)
                RETURNING invoice_id
            """), {"no": no, "t": inv_type, "cid": cid, "d": d}).scalar()
            created_invoices += 1
            return int(new_id)

        # 2) Import invoices + items (skip duplicates)
        for _, inv in invoices.iterrows():
            invoice_no = str(inv["invoice_no"]).strip()
            inv_type = str(inv["invoice_type"]).strip()
            client_name = str(inv["client_name"]).strip()
            issue_date = inv["issue_date"]  # može None

            cid = get_or_create_client_id(client_name)
            iid = get_or_create_invoice_id(invoice_no, inv_type, cid, issue_date)

            rows = df_items[df_items["invoice_no"] == invoice_no]

            for _, r in rows.iterrows():
                mid = int(r["material_id"])
                qty = float(r["qty"])
                uom = str(r["uom"]) if r["uom"] and str(r["uom"]).lower() != "nan" else None
                desc = str(r["material_name"]) if r["material_name"] and str(r["material_name"]).lower() != "nan" else None

                exists = c.execute(text("""
                    SELECT 1
                    FROM invoice_items
                    WHERE invoice_id=:iid AND material_id=:mid
                      AND COALESCE(qty,0)=:qty
                      AND COALESCE(uom,'')=COALESCE(:uom,'')
                      AND COALESCE(description,'')=COALESCE(:desc,'')
                    LIMIT 1
                """), {"iid": iid, "mid": mid, "qty": qty, "uom": uom, "desc": desc}).fetchone()

                if exists:
                    continue

                c.execute(text("""
                    INSERT INTO invoice_items(invoice_id, material_id, description, qty, uom, unit_price, discount)
                    VALUES (:iid, :mid, :desc, :qty, :uom, NULL, 0)
                """), {"iid": iid, "mid": mid, "desc": desc, "qty": qty, "uom": uom})

                created_items += 1

    print("IMPORT GOTOV ✅")
    print(f"  materials upserted: {upserted_materials} (popunjava NULL/prazno)")
    print(f"  +clients:  {created_clients}")
    print(f"  +invoices: {created_invoices}")
    print(f"  +items:    {created_items}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Upotreba: python import_izlaz.py <putanja_do_xlsx>")
    main(sys.argv[1])
