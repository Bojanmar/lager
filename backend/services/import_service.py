"""Service for importing Excel files."""
import pandas as pd
import io
from typing import Dict, List
from sqlalchemy.orm import Session
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Material, WarehouseStock, Invoice, InvoiceItem, ConversionRule
from services.calculation_service import norm_text, norm_unit


class ImportService:
    """Service for importing data from Excel files."""
    
    def __init__(self, db: Session):
        self.db = db
    
    def import_warehouse_stock(self, file_content: bytes, stock_type: str = "lager") -> Dict:
        """Import warehouse stock from wide-format Excel (materials as columns)."""
        df = pd.read_excel(io.BytesIO(file_content))
        
        # First row: units, Second row: values
        units = df.iloc[0].to_dict()
        values = df.iloc[1].to_dict()
        
        imported_count = 0
        for material_name, stock_value in values.items():
            if pd.isna(material_name) or pd.isna(stock_value):
                continue
            
            # Find or create material
            mat_key = norm_text(material_name)
            material = self.db.query(Material).filter(
                Material.name_normalized == mat_key
            ).first()
            
            if not material:
                material = Material(
                    name=str(material_name),
                    name_normalized=mat_key,
                    unit=norm_unit(units.get(material_name, ""))
                )
                self.db.add(material)
                self.db.flush()
            
            # Create or update stock
            stock = self.db.query(WarehouseStock).filter(
                WarehouseStock.material_id == material.id,
                WarehouseStock.stock_type == stock_type
            ).first()
            
            if stock:
                stock.stock_level = float(stock_value) if pd.notna(stock_value) else 0.0
            else:
                stock = WarehouseStock(
                    material_id=material.id,
                    stock_level=float(stock_value) if pd.notna(stock_value) else 0.0,
                    unit=norm_unit(units.get(material_name, "")),
                    stock_type=stock_type
                )
                self.db.add(stock)
            
            imported_count += 1
        
        self.db.commit()
        return {"status": "success", "imported": imported_count}
    
    def import_invoices(self, file_content: bytes) -> Dict:
        """Import invoices from IZLAZ Excel file."""
        df = pd.read_excel(io.BytesIO(file_content))
        
        # Group by invoice number
        if "Broj računa" not in df.columns:
            return {"status": "error", "message": "Missing 'Broj računa' column"}
        
        imported_invoices = 0
        imported_items = 0
        
        for invoice_num, group in df.groupby("Broj računa", dropna=False):
            if pd.isna(invoice_num):
                continue
            
            # Find or create invoice
            invoice = self.db.query(Invoice).filter(
                Invoice.invoice_number == str(invoice_num)
            ).first()
            
            if not invoice:
                company = group["Kompanija"].iloc[0] if "Kompanija" in group.columns else None
                invoice = Invoice(
                    invoice_number=str(invoice_num),
                    company=str(company) if pd.notna(company) else None
                )
                self.db.add(invoice)
                self.db.flush()
                imported_invoices += 1
            
            # Import items
            for _, row in group.iterrows():
                material_name = row.get("Materijal", "")
                if pd.isna(material_name):
                    continue
                
                # Find material
                mat_key = norm_text(material_name)
                material = self.db.query(Material).filter(
                    Material.name_normalized == mat_key
                ).first()
                
                # Get units from warehouse if material exists
                unit_for_warehouse = None
                if material and material.unit:
                    unit_for_warehouse = material.unit
                
                item = InvoiceItem(
                    invoice_id=invoice.id,
                    material_id=material.id if material else None,
                    material_name=str(material_name),
                    material_id_from_file=str(row.get("ID materijala", "")) if "ID materijala" in row else None,
                    quantity_for_billing=float(row.get("Količina za fakturisanje", 0)) if pd.notna(row.get("Količina za fakturisanje")) else 0.0,
                    unit_for_billing=str(row.get("Jedinica mere za fakturisanje", "")),
                    quantity_normative=float(row.get("Količina za fakturisanje (ono što piše u tabeli za račune - Normative)", 0)) if "Količina za fakturisanje (ono što piše u tabeli za račune - Normative)" in row and pd.notna(row.get("Količina za fakturisanje (ono što piše u tabeli za račune - Normative)")) else None,
                    unit_normative=str(row.get("Jedinica mere za fakturisanje - u računu", "")) if "Jedinica mere za fakturisanje - u računu" in row else None,
                    unit_for_warehouse=unit_for_warehouse,
                    work_type=str(row.get("Pozicija za fakturisanje - tip hidroizolacije", "")) if "Pozicija za fakturisanje - tip hidroizolacije" in row else None,
                    material_description=str(row.get("Opis Materijala", "")) if "Opis Materijala" in row else None,
                    technical_spec=str(row.get("Normativna potrošnja (tehnički list)", "")) if "Normativna potrošnja (tehnički list)" in row else None,
                )
                self.db.add(item)
                imported_items += 1
            
            self.db.commit()
        
        return {
            "status": "success",
            "invoices_imported": imported_invoices,
            "items_imported": imported_items
        }