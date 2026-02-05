"""FastAPI main application."""
from fastapi import FastAPI, Depends, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy.orm import Session
from typing import List, Optional
import uvicorn

from database import get_db, init_db
from models import (
    Material, WarehouseStock, Invoice, InvoiceItem,
    ConversionRule, ManualMapping, Calculation, CalculationResult
)
from services.import_service import ImportService
from services.calculation_service import CalculationService
from schemas import (
    MaterialSchema, InvoiceSchema, ConversionRuleSchema,
    ManualMappingSchema, CalculationResultSchema
)

app = FastAPI(title="Lager Kalkulator API", version="1.0.0")

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000", "http://localhost:5173"],  # React dev servers
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.on_event("startup")
async def startup_event():
    """Initialize database on startup."""
    init_db()


@app.get("/")
def root():
    """Root endpoint - health check."""
    return {"status": "ok", "message": "Lager Kalkulator API is running"}


@app.get("/api/health")
def health_check():
    """Health check endpoint."""
    return {"status": "healthy", "service": "Lager Kalkulator API"}


# ==================== Materials ====================

@app.get("/api/materials", response_model=List[MaterialSchema])
def get_materials(skip: int = 0, limit: int = 100, db: Session = Depends(get_db)):
    """Get all materials."""
    materials = db.query(Material).offset(skip).limit(limit).all()
    return materials


@app.get("/api/materials/{material_id}", response_model=MaterialSchema)
def get_material(material_id: int, db: Session = Depends(get_db)):
    """Get a single material."""
    material = db.query(Material).filter(Material.id == material_id).first()
    if not material:
        raise HTTPException(status_code=404, detail="Material not found")
    return material


# ==================== Warehouse Stock ====================

@app.get("/api/warehouse-stock")
def get_warehouse_stock(stock_type: str = "lager", db: Session = Depends(get_db)):
    """Get warehouse stock levels."""
    stocks = db.query(WarehouseStock).filter(
        WarehouseStock.stock_type == stock_type
    ).all()
    return [
        {
            "material_id": s.material_id,
            "material_name": s.material.name,
            "stock_level": s.stock_level,
            "unit": s.unit,
            "stock_type": s.stock_type
        }
        for s in stocks
    ]


@app.post("/api/warehouse-stock/import")
def import_warehouse_stock(
    file: UploadFile = File(...),
    stock_type: str = "lager",
    db: Session = Depends(get_db)
):
    """Import warehouse stock from Excel file."""
    content = file.file.read()
    service = ImportService(db)
    result = service.import_warehouse_stock(content, stock_type)
    return result


# ==================== Invoices ====================

@app.get("/api/invoices", response_model=List[InvoiceSchema])
def get_invoices(skip: int = 0, limit: int = 100, db: Session = Depends(get_db)):
    """Get all invoices."""
    invoices = db.query(Invoice).offset(skip).limit(limit).all()
    return invoices


@app.get("/api/invoices/{invoice_id}", response_model=InvoiceSchema)
def get_invoice(invoice_id: int, db: Session = Depends(get_db)):
    """Get a single invoice with items."""
    invoice = db.query(Invoice).filter(Invoice.id == invoice_id).first()
    if not invoice:
        raise HTTPException(status_code=404, detail="Invoice not found")
    return invoice


@app.post("/api/invoices/import")
def import_invoices(
    file: UploadFile = File(...),
    db: Session = Depends(get_db)
):
    """Import invoices from Excel file."""
    content = file.file.read()
    service = ImportService(db)
    result = service.import_invoices(content)
    return result


# ==================== Conversion Rules ====================

@app.get("/api/conversion-rules", response_model=List[ConversionRuleSchema])
def get_conversion_rules(
    material_id: Optional[int] = None,
    db: Session = Depends(get_db)
):
    """Get conversion rules."""
    query = db.query(ConversionRule)
    if material_id:
        query = query.filter(ConversionRule.material_id == material_id)
    rules = query.all()
    return rules


@app.post("/api/conversion-rules", response_model=ConversionRuleSchema)
def create_conversion_rule(
    rule: ConversionRuleSchema,
    db: Session = Depends(get_db)
):
    """Create a new conversion rule."""
    db_rule = ConversionRule(**rule.dict())
    db.add(db_rule)
    db.commit()
    db.refresh(db_rule)
    return db_rule


@app.put("/api/conversion-rules/{rule_id}", response_model=ConversionRuleSchema)
def update_conversion_rule(
    rule_id: int,
    rule: ConversionRuleSchema,
    db: Session = Depends(get_db)
):
    """Update a conversion rule."""
    db_rule = db.query(ConversionRule).filter(ConversionRule.id == rule_id).first()
    if not db_rule:
        raise HTTPException(status_code=404, detail="Rule not found")
    
    for key, value in rule.dict().items():
        setattr(db_rule, key, value)
    
    db.commit()
    db.refresh(db_rule)
    return db_rule


@app.delete("/api/conversion-rules/{rule_id}")
def delete_conversion_rule(rule_id: int, db: Session = Depends(get_db)):
    """Delete a conversion rule."""
    db_rule = db.query(ConversionRule).filter(ConversionRule.id == rule_id).first()
    if not db_rule:
        raise HTTPException(status_code=404, detail="Rule not found")
    
    db.delete(db_rule)
    db.commit()
    return {"status": "deleted"}


# ==================== Manual Mappings ====================

@app.get("/api/manual-mappings", response_model=List[ManualMappingSchema])
def get_manual_mappings(db: Session = Depends(get_db)):
    """Get all manual mappings."""
    mappings = db.query(ManualMapping).all()
    return mappings


@app.post("/api/manual-mappings", response_model=ManualMappingSchema)
def create_manual_mapping(
    mapping: ManualMappingSchema,
    db: Session = Depends(get_db)
):
    """Create a manual mapping."""
    db_mapping = ManualMapping(**mapping.dict())
    db.add(db_mapping)
    db.commit()
    db.refresh(db_mapping)
    return db_mapping


@app.delete("/api/manual-mappings/{mapping_id}")
def delete_manual_mapping(mapping_id: int, db: Session = Depends(get_db)):
    """Delete a manual mapping."""
    db_mapping = db.query(ManualMapping).filter(ManualMapping.id == mapping_id).first()
    if not db_mapping:
        raise HTTPException(status_code=404, detail="Mapping not found")
    
    db.delete(db_mapping)
    db.commit()
    return {"status": "deleted"}


# ==================== Calculations ====================

@app.post("/api/calculations/run")
def run_calculation(
    invoice_ids: Optional[List[int]] = None,
    db: Session = Depends(get_db)
):
    """Run calculation for invoices."""
    service = CalculationService(db)
    
    if invoice_ids:
        for invoice_id in invoice_ids:
            service.process_invoice_items(invoice_id)
    else:
        # Process all invoices
        invoices = db.query(Invoice).all()
        for invoice in invoices:
            service.process_invoice_items(invoice.id)
    
    # Calculate comparison
    result = service.calculate_comparison(0)  # TODO: create calculation record
    
    return result


@app.get("/api/calculations/comparison")
def get_comparison(db: Session = Depends(get_db)):
    """Get comparison results."""
    service = CalculationService(db)
    result = service.calculate_comparison(0)
    return result


# ==================== Export ====================

@app.get("/api/export/excel")
def export_excel(db: Session = Depends(get_db)):
    """Export calculation results to Excel."""
    from services.export_service import ExportService
    from fastapi.responses import StreamingResponse
    
    service = ExportService(db)
    buffer = service.export_to_excel()
    
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=obracun_zaliha.xlsx"}
    )


@app.get("/api/export/word/{invoice_number}")
def export_word(invoice_number: str, db: Session = Depends(get_db)):
    """Export Word document for an invoice."""
    from services.export_service import ExportService
    from fastapi.responses import StreamingResponse
    
    service = ExportService(db)
    buffer = service.export_word_for_invoice(invoice_number)
    
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename=racun_{invoice_number}.docx"}
    )


if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)