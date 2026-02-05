"""Database models for the inventory calculation system."""
from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, ForeignKey, Text, JSON
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime

Base = declarative_base()


class Material(Base):
    """Materials master table."""
    __tablename__ = "materials"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True, index=True, nullable=False)
    name_normalized = Column(String, index=True, nullable=False)  # normalized for matching
    unit = Column(String, nullable=True)  # default unit for warehouse
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    # Relationships
    warehouse_stocks = relationship("WarehouseStock", back_populates="material")
    invoice_items = relationship("InvoiceItem", back_populates="material")
    conversion_rules = relationship("ConversionRule", back_populates="material")
    manual_mappings = relationship("ManualMapping", foreign_keys="ManualMapping.material_izlaz_id", back_populates="material_izlaz")


class WarehouseStock(Base):
    """Warehouse stock levels."""
    __tablename__ = "warehouse_stocks"
    
    id = Column(Integer, primary_key=True, index=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    stock_level = Column(Float, nullable=False, default=0.0)
    unit = Column(String, nullable=False)
    date = Column(DateTime, nullable=False, default=datetime.utcnow)
    stock_type = Column(String, nullable=False)  # 'lager' or 'magacin'
    created_at = Column(DateTime, default=datetime.utcnow)
    
    material = relationship("Material", back_populates="warehouse_stocks")


class Invoice(Base):
    """Invoices/računi."""
    __tablename__ = "invoices"
    
    id = Column(Integer, primary_key=True, index=True)
    invoice_number = Column(String, unique=True, index=True, nullable=False)
    company = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    items = relationship("InvoiceItem", back_populates="invoice", cascade="all, delete-orphan")


class InvoiceItem(Base):
    """Items on invoices."""
    __tablename__ = "invoice_items"
    
    id = Column(Integer, primary_key=True, index=True)
    invoice_id = Column(Integer, ForeignKey("invoices.id"), nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=True)
    material_name = Column(String, nullable=False)  # original name from file
    material_id_from_file = Column(String, nullable=True)  # ID from IZLAZ file
    
    # Fakturisanje (billing)
    quantity_for_billing = Column(Float, nullable=False)
    unit_for_billing = Column(String, nullable=False)
    quantity_normative = Column(Float, nullable=True)  # Normative quantity
    unit_normative = Column(String, nullable=True)  # Normative unit
    
    # Lager (warehouse)
    quantity_for_warehouse = Column(Float, nullable=True)
    unit_for_warehouse = Column(String, nullable=True)
    quantity_with_coef = Column(Float, nullable=True)  # With Koef_novi applied
    
    # Additional info
    work_type = Column(String, nullable=True)  # tip hidroizolacije
    material_description = Column(String, nullable=True)
    technical_spec = Column(String, nullable=True)
    conversion_note = Column(Text, nullable=True)
    
    created_at = Column(DateTime, default=datetime.utcnow)
    
    invoice = relationship("Invoice", back_populates="items")
    material = relationship("Material", back_populates="invoice_items")


class ConversionRule(Base):
    """Material conversion rules."""
    __tablename__ = "conversion_rules"
    
    id = Column(Integer, primary_key=True, index=True)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    rule_type = Column(String, nullable=False)  # factor_per, per_piece, m2_to_rolna, etc.
    from_unit = Column(String, nullable=False)
    to_unit = Column(String, nullable=False)
    factor = Column(Float, nullable=True)
    extra = Column(Float, nullable=True, default=0.0)  # for markup/extra percentage
    enabled = Column(Boolean, default=True, nullable=False)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    material = relationship("Material", back_populates="conversion_rules")


class ManualMapping(Base):
    """Manual mappings between IZLAZ and LAGER materials."""
    __tablename__ = "manual_mappings"
    
    id = Column(Integer, primary_key=True, index=True)
    material_izlaz_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    material_lager_id = Column(Integer, ForeignKey("materials.id"), nullable=False)
    material_izlaz_name = Column(String, nullable=False)  # original name from IZLAZ
    material_lager_name = Column(String, nullable=False)  # mapped name from LAGER
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    material_izlaz = relationship("Material", foreign_keys=[material_izlaz_id], back_populates="manual_mappings")


class Calculation(Base):
    """Calculation runs/sessions."""
    __tablename__ = "calculations"
    
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=True)
    status = Column(String, nullable=False, default="pending")  # pending, completed, error
    created_at = Column(DateTime, default=datetime.utcnow)
    completed_at = Column(DateTime, nullable=True)
    
    results = relationship("CalculationResult", back_populates="calculation", cascade="all, delete-orphan")


class CalculationResult(Base):
    """Results from a calculation run."""
    __tablename__ = "calculation_results"
    
    id = Column(Integer, primary_key=True, index=True)
    calculation_id = Column(Integer, ForeignKey("calculations.id"), nullable=False)
    material_id = Column(Integer, ForeignKey("materials.id"), nullable=True)
    material_name = Column(String, nullable=False)
    
    # Comparison data
    total_consumption = Column(Float, nullable=True)
    warehouse_stock = Column(Float, nullable=True)
    difference_before = Column(Float, nullable=True)
    warehouse_actual = Column(Float, nullable=True)
    final_consumption = Column(Float, nullable=True)
    new_coefficient = Column(Float, nullable=True)  # Koef_novi
    calibration_status = Column(String, nullable=True)  # ok, EKSTREMNO
    
    created_at = Column(DateTime, default=datetime.utcnow)
    
    calculation = relationship("Calculation", back_populates="results")