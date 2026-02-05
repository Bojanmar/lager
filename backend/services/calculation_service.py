"""Core calculation service - ported from obracun.py."""
import re
import pandas as pd
import numpy as np
from typing import Dict, List, Optional, Tuple
from sqlalchemy.orm import Session
import sys
import os

# Add parent directory to path for imports
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from models import Material, ConversionRule, WarehouseStock, Invoice, InvoiceItem, ManualMapping


def norm_text(s):
    """Normalize text for matching."""
    if pd.isna(s):
        return ""
    s = str(s).replace('"', ' ').replace("'", ' ')
    s = s.replace("\n", " ")
    return re.sub(r"\s+", " ", s.strip()).lower()


def norm_unit(u):
    """Normalize unit names."""
    u = norm_text(u)
    repl = {
        "m^2": "m2", "m²": "m2",
        "m^1": "m", "m¹": "m", "m1": "m",
        "kom (rolni)": "kom", "kom (rolna)": "kom", "kom (pak)": "kom",
        "kg.": "kg",
        "l": "lit", "litar": "lit", "litra": "lit", "litara": "lit",
        "džak": "dzak", "djak": "dzak",
        "rolni": "rolna",
        "kom rupa": "kom",
    }
    return repl.get(u, u)


# Same unit markup (from obracun.py)
SAME_UNIT_MARKUP = {
    "borner multiplex av 4": 0.20,
    "volteco, volgrip h.1.10 light": 0.15,
    "vintex mp fr 1.5mm": 0.10,
    "geotekstil 300gr-m2": 0.10,
    "poliesterska tkanina filc 100%, rolna 2x1m": 0.10,
}

# Default heuristics
PAIR_DEFAULTS = {
    ("m2", "kg"): 1.5,
    ("m", "kg"): 0.30,
    ("kom", "kg"): 0.30
}


class ConversionRuleEngine:
    """Engine for applying conversion rules."""
    
    def __init__(self, db: Session):
        self.db = db
        self._rules_cache = None
        self._load_rules()
    
    def _load_rules(self):
        """Load all conversion rules from database."""
        rules = self.db.query(ConversionRule).filter(ConversionRule.enabled == True).all()
        self._rules_cache = {}
        for rule in rules:
            mat_key = rule.material.name_normalized
            if mat_key not in self._rules_cache:
                self._rules_cache[mat_key] = []
            self._rules_cache[mat_key].append(rule)
    
    def apply_rule(self, rule: ConversionRule, quantity: float, from_unit: str, to_unit: str) -> Optional[Tuple[float, str]]:
        """Apply a single conversion rule."""
        uf = norm_unit(from_unit)
        ul = norm_unit(to_unit)
        
        rule_from = norm_unit(rule.from_unit)
        rule_to = norm_unit(rule.to_unit)
        
        if uf != rule_from or ul != rule_to:
            return None
        
        if rule.rule_type == "factor_per":
            result = quantity * rule.factor
            note = f"factor {rule.factor} {rule_to}/{rule_from}"
            return result, note
        
        elif rule.rule_type == "per_piece":
            if uf != "kom":
                return None
            result = quantity * rule.factor
            note = f"{rule.factor} {rule_to}/kom"
            return result, note
        
        elif rule.rule_type == "m2_to_rolna":
            if uf != "m2" or ul != "rolna":
                return None
            rolls = (quantity / rule.factor) * (1.0 + (rule.extra or 0.0))
            note = f"m2→rolna; {rule.factor} m2/rolna; +{int((rule.extra or 0.0) * 100)}%"
            return rolls, note
        
        elif rule.rule_type == "m_to_rolna":
            if uf != "m" or ul != "rolna":
                return None
            rolls = (quantity / rule.factor) * (1.0 + (rule.extra or 0.0))
            note = f"m→rolna; {rule.factor} m/rolna; +{int((rule.extra or 0.0) * 100)}%"
            return rolls, note
        
        elif rule.rule_type == "identity":
            if uf != rule_from or ul != rule_to:
                return None
            return quantity, f"1:1 {rule_from}->{rule_to}"
        
        elif rule.rule_type == "m2_to_lit":
            if uf != "m2" or ul != "lit":
                return None
            result = quantity * rule.factor
            note = f"{rule.factor} lit/m2"
            return result, note
        
        elif rule.rule_type == "m_to_lit":
            if uf != "m" or ul != "lit":
                return None
            result = quantity * rule.factor
            note = f"{rule.factor} lit/m"
            return result, note
        
        elif rule.rule_type == "kg_to_dzak":
            if uf != "kg" or ul != "dzak":
                return None
            result = quantity / rule.factor
            note = f"1/{rule.factor} dzak/kg"
            return result, note
        
        return None
    
    def calculate_warehouse_quantity(
        self,
        material_name: str,
        quantity: float,
        from_unit: str,
        to_unit: str
    ) -> Tuple[Optional[float], str]:
        """Calculate warehouse quantity using rules."""
        mat_key = norm_text(material_name)
        
        # Same units - direct conversion with possible markup
        if norm_unit(from_unit) == norm_unit(to_unit):
            base = quantity
            markup = SAME_UNIT_MARKUP.get(mat_key, 0.0)
            if markup:
                return base * (1 + markup), f"same_unit +{int(markup * 100)}%"
            return base, "same_unit"
        
        # Try rules
        rules = self._rules_cache.get(mat_key, [])
        for rule in rules:
            result = self.apply_rule(rule, quantity, from_unit, to_unit)
            if result is not None:
                return result[0], f"rule: {result[1]}"
        
        # Try heuristics
        uf = norm_unit(from_unit)
        ul = norm_unit(to_unit)
        if (uf, ul) in PAIR_DEFAULTS:
            factor = PAIR_DEFAULTS[(uf, ul)]
            result = quantity * factor
            return result, f"heuristic: {factor} {ul}/{uf}"
        
        return None, f"nema_pravila({from_unit}->{to_unit})"


class CalculationService:
    """Main calculation service."""
    
    def __init__(self, db: Session):
        self.db = db
        self.rule_engine = ConversionRuleEngine(db)
    
    def process_invoice_items(self, invoice_id: int) -> Dict:
        """Process all items in an invoice and calculate warehouse quantities."""
        invoice = self.db.query(Invoice).filter(Invoice.id == invoice_id).first()
        if not invoice:
            raise ValueError(f"Invoice {invoice_id} not found")
        
        items = self.db.query(InvoiceItem).filter(InvoiceItem.invoice_id == invoice_id).all()
        
        for item in items:
            if item.quantity_for_billing and item.unit_for_billing and item.unit_for_warehouse:
                qty, note = self.rule_engine.calculate_warehouse_quantity(
                    item.material_name,
                    item.quantity_for_billing,
                    item.unit_for_billing,
                    item.unit_for_warehouse
                )
                if qty is not None:
                    item.quantity_for_warehouse = qty
                    item.conversion_note = note
        
        self.db.commit()
        return {"status": "success", "items_processed": len(items)}
    
    def calculate_comparison(self, calculation_id: int) -> Dict:
        """Calculate comparison between consumption and warehouse stock."""
        # Get all invoice items
        items = self.db.query(InvoiceItem).all()
        
        # Group by material
        consumption_by_material = {}
        for item in items:
            if item.quantity_for_warehouse:
                mat_key = norm_text(item.material_name)
                if mat_key not in consumption_by_material:
                    consumption_by_material[mat_key] = 0.0
                consumption_by_material[mat_key] += item.quantity_for_warehouse
        
        # Get warehouse stocks
        stocks = self.db.query(WarehouseStock).filter(
            WarehouseStock.stock_type == "lager"
        ).all()
        
        results = []
        for stock in stocks:
            mat_key = stock.material.name_normalized
            consumption = consumption_by_material.get(mat_key, 0.0)
            difference = stock.stock_level - consumption
            results.append({
                "material_name": stock.material.name,
                "total_consumption": consumption,
                "warehouse_stock": stock.stock_level,
                "difference_before": difference,
            })
        
        return {"results": results}