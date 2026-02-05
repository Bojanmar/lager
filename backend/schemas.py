"""Pydantic schemas for API requests/responses."""
from pydantic import BaseModel
from typing import Optional, List
from datetime import datetime


class MaterialSchema(BaseModel):
    id: Optional[int] = None
    name: str
    name_normalized: Optional[str] = None
    unit: Optional[str] = None
    
    class Config:
        from_attributes = True


class WarehouseStockSchema(BaseModel):
    id: Optional[int] = None
    material_id: int
    stock_level: float
    unit: str
    stock_type: str
    
    class Config:
        from_attributes = True


class InvoiceItemSchema(BaseModel):
    id: Optional[int] = None
    material_id: Optional[int] = None
    material_name: str
    material_id_from_file: Optional[str] = None
    quantity_for_billing: float
    unit_for_billing: str
    quantity_normative: Optional[float] = None
    unit_normative: Optional[str] = None
    quantity_for_warehouse: Optional[float] = None
    unit_for_warehouse: Optional[str] = None
    quantity_with_coef: Optional[float] = None
    work_type: Optional[str] = None
    material_description: Optional[str] = None
    technical_spec: Optional[str] = None
    conversion_note: Optional[str] = None
    
    class Config:
        from_attributes = True


class InvoiceSchema(BaseModel):
    id: Optional[int] = None
    invoice_number: str
    company: Optional[str] = None
    items: List[InvoiceItemSchema] = []
    
    class Config:
        from_attributes = True


class ConversionRuleSchema(BaseModel):
    id: Optional[int] = None
    material_id: int
    rule_type: str
    from_unit: str
    to_unit: str
    factor: Optional[float] = None
    extra: Optional[float] = 0.0
    enabled: bool = True
    
    class Config:
        from_attributes = True


class ManualMappingSchema(BaseModel):
    id: Optional[int] = None
    material_izlaz_id: int
    material_lager_id: int
    material_izlaz_name: str
    material_lager_name: str
    
    class Config:
        from_attributes = True


class CalculationResultSchema(BaseModel):
    id: Optional[int] = None
    material_name: str
    total_consumption: Optional[float] = None
    warehouse_stock: Optional[float] = None
    difference_before: Optional[float] = None
    warehouse_actual: Optional[float] = None
    final_consumption: Optional[float] = None
    new_coefficient: Optional[float] = None
    calibration_status: Optional[str] = None
    
    class Config:
        from_attributes = True