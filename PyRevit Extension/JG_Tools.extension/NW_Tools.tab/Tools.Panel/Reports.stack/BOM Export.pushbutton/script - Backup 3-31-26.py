# -*- coding: utf-8 -*-
"""
NW BOM Export - Pipeline Architecture
======================================

Modern pipeline-based architecture for Revit BOM exports.
Built from scratch using proper design patterns while maintaining
100% output compatibility with the original script.

Architecture:
    INPUT → VALIDATE → EXPAND → CLASSIFY → PROCESS → ENRICH → SORT → SERIALIZE → OUTPUT

Elements flow through processors that know how to handle specific types.
Field extraction is centralized in specialized extractors.
"""

from pyrevit import revit, forms, script as pyrevit_script
from Autodesk.Revit.DB import (
    FabricationPart, FamilyInstance, FilteredElementCollector,
    BuiltInCategory, BuiltInParameter, ElementId, CategoryType, Element
)
from Autodesk.Revit.UI.Selection import ObjectType
from System.Windows.Forms import Clipboard
from System.Collections.Generic import List
import clr
import re
import os
import shutil

clr.AddReference("System.Windows.Forms")


# =============================================================================
# CONFIGURATION
# =============================================================================

class Config:
    """Centralized configuration for the BOM export system."""
    
    # Debug mode
    DEBUG = False
    
    # Parameter Names (Southland workflow)
    class Params:
        ITEM_NUMBER = "CP_Item Number"
        DESCRIPTION = "CP_BOM_Description"
        SIZE = "CP_Size"
        LENGTH = "CP_Length"
        TYPE_NAME = "CP_Type Name"
        AREA = "CP_Construction Area"
        SERVICE = "CP_Service"
        HANGER_NUMBER = "CP_Hanger Number"
        MODEL_NUMBER = "CP_Model Number"
        MATERIAL = "CP_Material & Coating"
        FAMILY_CATEGORY = "CPL_Family Category"
    
    # Business Rules
    class Rules:
        ZERO_TOLERANCE_FT = 1e-9
        MAX_NESTED_RFA_DEPTH = 3
    
    # Formatting
    class Format:
        FRAC_DENOMINATOR = 16
        LENGTH_DECIMAL_PLACES = 3
    
    # Category Identifiers
    class Categories:
        FAB_PIPEWORK_SUBSTR = "MEP Fabrication Pipework"
        FAB_DUCTWORK_SUBSTR = "MEP Fabrication Ductwork"
        
        NATIVE_MEP_IDS = set([
            -2008044,  # Pipes
            -2008049,  # Pipe Fittings
            -2008000,  # Ducts
            -2008010,  # Duct Fittings
            -2008050,  # Flex Pipes
            -2008020,  # Flex Ducts
        ])
        
        EXCLUDE_GTP_CATEGORY = "Generic Models"
        EXCLUDE_GTP_FAMILY = "GTP"
        
        EXCLUDED_TOKENS = (
            "level", "grid", "reference plane", "scope box", "lines"
        )
    
    # Sorting
    class Sorting:
        PREFERENCES = {
            "Pipe": 110,
            "Pipes": 110,
            "Pipe Fittings": 210,
            "Pipe Accessories": 220,
            "Structural Framing": 610,
            "Structural Columns": 620,
            "Structural Foundations": 630,
            "Structural Connections": 640,
            "Conduit Fittings": 650,
        }
        
        STRUCTURAL_KEYWORDS = (
            "structural", "unistrut", "strut", "steel", "channel", "trapeze"
        )
    
    # Output Columns
    class Columns:
        HEADER = [
            "Sorting", "Category", "Family", "Type", "ItemNumber", 
            "Description", "Size", "Length_ft", "Length_ft_in", "Count", 
            "Item count", "Model", "Area", "Material", "Comments", "Spool", 
            "Service", "Hanger Number", "Install Type", "ElementId"
        ]
        
        # Column indices
        IDX_SORT = 0
        IDX_CAT = 1
        IDX_FAM = 2
        IDX_TYP = 3
        IDX_ITEMNO = 4
        IDX_DESC = 5
        IDX_SIZE = 6
        IDX_LENFT = 7
        IDX_LENFR = 8
        IDX_COUNT = 9
        IDX_ITEMCT = 10
        IDX_MODEL = 11
        IDX_AREA = 12
        IDX_MATERIAL = 13
        IDX_COMMENTS = 14
        IDX_SPOOL = 15
        IDX_SERVICE = 16
        IDX_HANGER = 17
        IDX_INSTALL = 18
        IDX_ELEMID = 19


# =============================================================================
# DOMAIN MODEL
# =============================================================================

class BOMRow:
    """Represents a single row in the BOM export.
    
    This is the core data model - a structured representation of one BOM line item.
    Can represent either a primary element or a derived item (like hanger rods).
    """
    
    def __init__(self):
        self.sorting = ""
        self.category = ""
        self.family = ""
        self.type = ""
        self.item_number = ""
        self.description = ""
        self.size = ""
        self.length_ft = None
        self.length_ft_in = ""
        self.count = ""
        self.item_count = "1"
        self.model = ""
        self.area = ""
        self.material = ""
        self.comments = ""
        self.spool = ""
        self.service = ""
        self.hanger_number = ""
        self.install_type = ""
        self.element_id = ""
        
        # Internal tracking
        self._source_element = None
        self._is_derived = False  # True for rod rows derived from hangers
    
    def to_list(self):
        """Convert to list for TSV export."""
        return [
            self.sorting,
            self.category,
            self.family,
            self.type,
            self.item_number,
            self.description,
            self.size,
            self.length_ft if isinstance(self.length_ft, str) else "",
            self.length_ft_in,
            self.count,
            self.item_count,
            self.model,
            self.area,
            self.material,
            self.comments,
            self.spool,
            self.service,
            self.hanger_number,
            self.install_type,
            self.element_id,
        ]
    
    def validate(self):
        """Apply business rules to ensure valid BOM row.
        
        Ensures pivotable fields are never blank (uses "N/A" placeholder).
        """
        if not (self.type or "").strip():
            self.type = "N/A"
        if not (self.size or "").strip():
            self.size = '="N/A"'
        if not (self.material or "").strip():
            self.material = "N/A"
        if not (self.model or "").strip():
            self.model = "N/A"
    
    def copy(self):
        """Create a deep copy of this row."""
        new_row = BOMRow()
        new_row.sorting = self.sorting
        new_row.category = self.category
        new_row.family = self.family
        new_row.type = self.type
        new_row.item_number = self.item_number
        new_row.description = self.description
        new_row.size = self.size
        new_row.length_ft = self.length_ft
        new_row.length_ft_in = self.length_ft_in
        new_row.count = self.count
        new_row.item_count = self.item_count
        new_row.model = self.model
        new_row.area = self.area
        new_row.material = self.material
        new_row.comments = self.comments
        new_row.spool = self.spool
        new_row.service = self.service
        new_row.hanger_number = self.hanger_number
        new_row.install_type = self.install_type
        new_row.element_id = self.element_id
        new_row._source_element = self._source_element
        new_row._is_derived = self._is_derived
        return new_row


# =============================================================================
# SHARED CONTEXT
# =============================================================================

class Cache:
    """Performance caching for expensive lookups."""
    
    def __init__(self):
        self.type_elements = {}
        self.type_info = {}
        self.material_names = {}
        self.sort_codes_used = set()
        self.category_codes = {}
    
    def clear(self):
        """Reset all caches."""
        self.type_elements.clear()
        self.type_info.clear()
        self.material_names.clear()
        self.sort_codes_used.clear()
        self.category_codes.clear()


class ExportContext:
    """Shared execution context for the export pipeline.
    
    Contains:
    - Document reference
    - Configuration
    - Caches
    - Shared utilities (formatters, extractors)
    
    This is passed to all processors and extractors so they have
    access to shared state without globals.
    """
    
    def __init__(self, doc):
        self.doc = doc
        self.config = Config()
        self.cache = Cache()
        
        # Initialize utilities (done lazily to avoid circular dependencies)
        self._formatters = None
        self._extractors = None
    
    @property
    def formatters(self):
        """Lazy initialization of formatters."""
        if self._formatters is None:
            self._formatters = FormatterRegistry(self)
        return self._formatters
    
    @property
    def extractors(self):
        """Lazy initialization of extractors."""
        if self._extractors is None:
            self._extractors = ExtractorRegistry(self)
        return self._extractors
    
    def reset(self):
        """Clear caches for fresh run."""
        self.cache.clear()


# =============================================================================
# UTILITIES
# =============================================================================

def log_debug(msg, context=""):
    """Debug logger (only active when Config.DEBUG = True)."""
    if not Config.DEBUG:
        return
    try:
        output = "[DEBUG] {}: {}".format(context, msg) if context else "[DEBUG] {}".format(msg)
        print(output)
    except:
        pass


def safe_str(val):
    """Safely convert value to string."""
    try:
        return "" if val is None else str(val)
    except:
        return ""


def sanitize_tsv(val):
    """Remove tabs and newlines for TSV export."""
    return "" if val is None else str(val).replace("\t", " ").replace("\n", " ").replace("\r", " ")


def normalize_blank(val, placeholder="N/A"):
    """Replace empty values with placeholder."""
    try:
        s = (val or "").strip()
    except:
        s = ""
    return s if s else placeholder


# =============================================================================
# FORMATTERS
# =============================================================================

class Formatter:
    """Base class for value formatters."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config


class ExcelFormulaFormatter(Formatter):
    """Wraps text in Excel formula to prevent auto-formatting."""
    
    def format(self, text):
        """Wrap text in ="..." formula."""
        s = (text or "").strip()
        if not s:
            return ""
        s = s.replace('"', '""')
        return '="{}"'.format(s)


class LengthDecimalFormatter(Formatter):
    """Format length as decimal feet."""
    
    def format(self, val_ft):
        """Format length to 3 decimal places."""
        if val_ft is None:
            return ""
        try:
            places = self.config.Format.LENGTH_DECIMAL_PLACES
            return "{:.3f}".format(float(val_ft))
        except:
            return ""


class LengthFractionalFormatter(Formatter):
    """Format length as feet-inches with fractions."""
    
    def format(self, val_ft):
        """Format as feet-inches (e.g., 5'-3 1/2")."""
        if val_ft is None:
            return ""
        try:
            ft_in = float(val_ft)
        except:
            return ""
        
        denom = self.config.Format.FRAC_DENOMINATOR
        ft = abs(ft_in)
        sign = "-" if ft_in < 0 else ""
        
        total_inches = ft * 12.0
        feet = int(total_inches // 12)
        inches = total_inches - (feet * 12)
        
        inches = round(inches * denom) / denom
        if inches >= 12:
            feet += 1
            inches -= 12
        
        whole = int(inches)
        frac = inches - whole
        
        frac_str = ""
        if frac > 0:
            num = int(round(frac * denom))
            g = self._gcd(num, denom)
            frac_str = " {}/{}".format(num // g, denom // g)
        
        return '{}{}\'-{}{}"'.format(sign, feet, whole, frac_str)
    
    @staticmethod
    def _gcd(a, b):
        """Calculate greatest common divisor."""
        while b:
            a, b = b, a % b
        return a


class InchesFractionalFormatter(Formatter):
    """Format inches as fractional string."""
    
    def format(self, inches):
        """Format as fractional inches (e.g., 3 1/2")."""
        if inches is None:
            return ""
        try:
            x = float(inches)
        except:
            return ""
        
        denom = self.config.Format.FRAC_DENOMINATOR
        sign = "-" if x < 0 else ""
        x = abs(x)
        x = round(x * denom) / float(denom)
        whole = int(x)
        frac = x - whole
        
        if frac < 1e-12:
            return '{}{}"'.format(sign, whole)
        
        num = int(round(frac * denom))
        g = LengthFractionalFormatter._gcd(num, denom)
        num //= g
        den = denom // g
        
        if whole == 0:
            return '{}{}/{}"'.format(sign, num, den)
        return '{}{} {}/{}"'.format(sign, whole, num, den)


class SizeFormatter(Formatter):
    """Format size strings with inch marks."""
    
    def format(self, size_str):
        """Add inch marks to size dimensions."""
        s = (size_str or "").strip()
        if not s:
            return s
        if '"' in s:
            return s
        
        s_norm = s.replace("×", "x")
        s_norm = re.sub(r"\s*x\s*", " x ", s_norm)
        
        if " x " in s_norm:
            parts = [p.strip() for p in s_norm.split(" x ") if p.strip()]
            return " x ".join(self._ensure_inches(p) for p in parts)
        
        return self._ensure_inches(s_norm)
    
    def _ensure_inches(self, tok):
        """Add inch marks to numeric tokens if missing."""
        t = (tok or "").strip()
        if not t:
            return t
        if '"' in t or "'" in t:
            return t
        if re.search(r"[A-Za-z\[\]]", t):
            return t
        
        inch_like = (
            r"^\d+$"
            r"|^\d+\.\d+$"
            r"|^\d+/\d+$"
            r"|^\d+-\d+/\d+$"
            r"|^\d+\s+\d+/\d+$"
        )
        if re.match(inch_like, t):
            return t + '"'
        return t


class CountFlagFormatter(Formatter):
    """Format count flag for Excel pivot tables."""
    
    def format(self, length_ft):
        """Return '1' if no length, else blank."""
        if length_ft is None:
            return "1"
        try:
            lf = float(length_ft)
        except:
            return "1"
        if abs(lf) <= self.config.Rules.ZERO_TOLERANCE_FT:
            return "1"
        return "" if lf > 0.0 else "1"


class FormatterRegistry:
    """Container for all formatters."""
    
    def __init__(self, context):
        self.excel_formula = ExcelFormulaFormatter(context)
        self.length_decimal = LengthDecimalFormatter(context)
        self.length_fractional = LengthFractionalFormatter(context)
        self.inches_fractional = InchesFractionalFormatter(context)
        self.size = SizeFormatter(context)
        self.count_flag = CountFlagFormatter(context)


# =============================================================================
# PARAMETER ACCESS
# =============================================================================

class ParameterHelper:
    """Centralized parameter access with caching."""
    
    def __init__(self, context):
        self.context = context
        self.doc = context.doc
        self.cache = context.cache
    
    def get_param(self, el, name):
        """Get parameter by name."""
        try:
            return el.LookupParameter(name)
        except:
            return None
    
    def get_param_str(self, el, name):
        """Get parameter value as string."""
        p = self.get_param(el, name)
        if not p:
            return ""
        try:
            return (p.AsString() or "").strip()
        except:
            try:
                return (p.AsValueString() or "").strip()
            except:
                return ""
    
    def get_param_double(self, el, name):
        """Get parameter value as double."""
        p = self.get_param(el, name)
        if not p:
            return None
        try:
            return p.AsDouble()
        except:
            return None
    
    def get_bip_param(self, el, bip):
        """Get built-in parameter."""
        try:
            return el.get_Parameter(bip)
        except:
            return None
    
    def get_bip_str(self, el, bip):
        """Get built-in parameter as string."""
        p = self.get_bip_param(el, bip)
        if not p:
            return ""
        try:
            return (p.AsString() or "").strip()
        except:
            try:
                return (p.AsValueString() or "").strip()
            except:
                return ""
    
    def get_type_element(self, element):
        """Get type element from cache or document."""
        try:
            type_id = element.GetTypeId()
        except:
            return None
        
        if not type_id:
            return None
        
        try:
            tid_int = type_id.IntegerValue
        except:
            return None
        
        # Check cache
        te = self.cache.type_elements.get(tid_int)
        if te is not None:
            return te
        
        # Fetch from document
        try:
            te = self.doc.GetElement(type_id)
        except:
            te = None
        
        self.cache.type_elements[tid_int] = te
        return te
    
    def get_type_info(self, element):
        """Get cached type info dict."""
        try:
            tid = element.GetTypeId()
            tid_int = tid.IntegerValue
        except:
            return {}
        
        # Check cache
        info = self.cache.type_info.get(tid_int)
        if info is not None:
            return info
        
        # Build info
        te = self.get_type_element(element)
        
        info = {
            "type_el": te,
            "type_name": self._get_type_name(te),
            "family_name": self._get_family_name(te),
            "cpl_family_category": self.get_param_str(te, self.context.config.Params.FAMILY_CATEGORY).strip() if te else "",
            "type_comments": self.get_bip_str(te, BuiltInParameter.ALL_MODEL_TYPE_COMMENTS) if te else "",
        }
        
        self.cache.type_info[tid_int] = info
        return info
    
    def _get_type_name(self, te):
        """Extract type name from type element."""
        if not te:
            return ""
        try:
            s = (Element.Name.GetValue(te) or "").strip()
            if s:
                return s
        except:
            pass
        try:
            return (te.Name or "").strip()
        except:
            return ""
    
    def _get_family_name(self, te):
        """Extract family name from type element."""
        if not te:
            return ""
        try:
            s = (getattr(te, "FamilyName", None) or "").strip()
            if s:
                return s
        except:
            pass
        try:
            return self.get_bip_str(te, BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        except:
            return ""
    
    def get_material_name_from_param(self, p):
        """Resolve material name from parameter."""
        if not p:
            return ""
        
        try:
            mid = p.AsElementId()
            if mid and mid.IntegerValue > 0:
                mid_int = mid.IntegerValue
                
                # Check cache
                cached = self.cache.material_names.get(mid_int)
                if cached is not None:
                    return cached
                
                # Fetch from document
                m = self.doc.GetElement(mid)
                if m:
                    nm = (m.Name or "").strip()
                    self.cache.material_names[mid_int] = nm
                    return nm
        except:
            pass
        
        try:
            return (p.AsString() or "").strip()
        except:
            try:
                return (p.AsValueString() or "").strip()
            except:
                return ""


# Continuing in next file segment...


# =============================================================================
# ELEMENT TYPE DETECTION
# =============================================================================

class ElementClassifier:
    """Determines element types and routing decisions."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
    
    def is_fabrication_part(self, element):
        """Check if element is a FabricationPart."""
        return isinstance(element, FabricationPart)
    
    def is_revit_mep_curve(self, element):
        """Check if element is a native Revit MEP curve."""
        try:
            cid = element.Category.Id.IntegerValue if (element and element.Category and element.Category.Id) else None
        except:
            cid = None
        if cid is None:
            return False
        return cid in (
            int(BuiltInCategory.OST_PipeCurves),
            int(BuiltInCategory.OST_FlexPipeCurves),
            int(BuiltInCategory.OST_DuctCurves),
            int(BuiltInCategory.OST_FlexDuctCurves),
        )
    
    def is_revit_mep_fitting(self, element):
        """Check if element is a native Revit MEP fitting (Pipe/Duct Fittings only).
        
        Note: Pipe Accessories and Duct Accessories are intentionally excluded.
        Those are RFA families (e.g. Sikla beams, valves) that carry CP parameters
        and should route to RFAProcessor for proper field extraction.
        """
        try:
            if not element or not element.Category:
                return False
            cid = element.Category.Id.IntegerValue
            return cid in (
                int(BuiltInCategory.OST_PipeFitting),
                int(BuiltInCategory.OST_DuctFitting),
            )
        except:
            return False
    
    def is_native_familyinstance(self, fi):
        """Check if FamilyInstance should use native MEP rules (not CP rules)."""
        try:
            if not isinstance(fi, FamilyInstance):
                return False
            
            cat = getattr(fi, "Category", None)
            if not cat or not hasattr(cat, "Id"):
                return False
            if cat.Id.IntegerValue not in self.config.Categories.NATIVE_MEP_IDS:
                return False
            
            params = self.context.config.Params
            param_helper = ParameterHelper(self.context)
            
            # Check instance parameters
            inst_hits = [
                param_helper.get_param_str(fi, params.FAMILY_CATEGORY).strip(),
                param_helper.get_param_str(fi, params.SIZE).strip(),
                param_helper.get_param_str(fi, params.TYPE_NAME).strip(),
                param_helper.get_param_str(fi, params.SERVICE).strip(),
            ]
            if any(inst_hits):
                return False
            
            # Check type parameters
            try:
                sym = fi.Symbol
            except:
                sym = None
            if sym:
                type_hits = [
                    param_helper.get_param_str(sym, params.FAMILY_CATEGORY).strip(),
                    param_helper.get_param_str(sym, params.SIZE).strip(),
                    param_helper.get_param_str(sym, params.TYPE_NAME).strip(),
                    param_helper.get_param_str(sym, params.SERVICE).strip(),
                ]
                if any(type_hits):
                    return False
            
            return True
        except:
            return False
    
    def is_modeling_element(self, element):
        """Check if element is a valid model element."""
        if not element:
            return False
        try:
            cat = element.Category
        except:
            cat = None
        if not cat:
            return False
        try:
            if cat.CategoryType != CategoryType.Model:
                return False
        except:
            return False
        try:
            if element.ViewSpecific:
                return False
        except:
            pass
        try:
            if cat.Id.IntegerValue == int(BuiltInCategory.OST_Assemblies):
                return False
        except:
            pass
        try:
            nm = (cat.Name or "").strip().lower()
            for token in self.config.Categories.EXCLUDED_TOKENS:
                if token in nm:
                    return False
        except:
            pass
        return True
    
    def should_exclude_by_family(self, element):
        """Check if element should be excluded (GTP family)."""
        try:
            cat = element.Category.Name if element.Category else ""
        except:
            cat = ""
        
        if cat != self.config.Categories.EXCLUDE_GTP_CATEGORY:
            return False
        
        param_helper = ParameterHelper(self.context)
        info = param_helper.get_type_info(element)
        fam = (info.get("family_name") or "").strip()
        
        if not fam and info.get("type_el"):
            try:
                fam = param_helper.get_bip_str(info["type_el"], BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            except:
                fam = ""
        
        return (fam or "").strip().lower() == self.config.Categories.EXCLUDE_GTP_FAMILY.lower()


# =============================================================================
# FIELD EXTRACTORS
# =============================================================================

class FieldExtractor:
    """Base class for field extractors."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
        self.param_helper = ParameterHelper(context)


class CategoryExtractor(FieldExtractor):
    """Extract category field."""
    
    def extract(self, element):
        """Get category name (CPL_Family Category or Revit category)."""
        # Revit category (fallback)
        rev_cat = ""
        try:
            if element and element.Category:
                rev_cat = element.Category.Name or ""
        except:
            rev_cat = ""
        
        # RFA: check CPL_Family Category
        try:
            if isinstance(element, FamilyInstance):
                # Instance param
                v = self.param_helper.get_param_str(element, self.config.Params.FAMILY_CATEGORY)
                if v:
                    return v
                
                # Type param
                sym = None
                try:
                    sym = element.Symbol
                except:
                    sym = None
                if sym:
                    v = self.param_helper.get_param_str(sym, self.config.Params.FAMILY_CATEGORY)
                    if v:
                        return v
                
                return rev_cat
        except:
            return rev_cat
        
        return rev_cat


class ItemNumberExtractor(FieldExtractor):
    """Extract item number field."""
    
    def extract(self, element):
        """Get item number (CP_Item Number or fabrication Item Number)."""
        v = self.param_helper.get_param_str(element, self.config.Params.ITEM_NUMBER).strip()
        if v:
            return v
        
        # ITM fallback
        is_itm = False
        try:
            is_itm = (element is not None and element.GetType().FullName == "Autodesk.Revit.DB.FabricationPart")
        except:
            is_itm = hasattr(element, "GetDimensions") and hasattr(element, "ServiceName")
        
        if is_itm:
            try:
                bip = getattr(BuiltInParameter, "FABRICATION_PART_ITEM_NUMBER", None)
                if bip is not None:
                    p = element.get_Parameter(bip)
                    if p:
                        s = (p.AsString() or p.AsValueString() or "").strip()
                        if s:
                            return s
            except:
                pass
            
            s = self.param_helper.get_param_str(element, "Item Number").strip()
            if s:
                return s
        
        return ""


class FamilyTypeExtractor(FieldExtractor):
    """Extract family and type names."""
    
    def extract(self, element):
        """Get (family_name, type_name) tuple."""
        if isinstance(element, FamilyInstance):
            return self._extract_family_instance(element)
        elif isinstance(element, FabricationPart):
            return self._extract_fabrication_part(element)
        else:
            return self._extract_other(element)
    
    def _extract_family_instance(self, fi):
        """Extract for FamilyInstance."""
        fam = ""
        try:
            fam = (fi.Symbol.Family.Name or "").strip()
        except:
            fam = ""
        
        info = self.param_helper.get_type_info(fi)
        typ = (info.get("type_name") or "").strip()
        if not typ:
            try:
                typ = (fi.Symbol.Name or "").strip()
            except:
                typ = ""
        
        return fam, typ
    
    def _extract_fabrication_part(self, itm):
        """Extract for FabricationPart."""
        info = self.param_helper.get_type_info(itm)
        typ = (info.get("type_name") or "").strip()
        fam = (info.get("family_name") or "").strip()
        
        if not typ:
            for attr in ("ItemName", "ProductName", "Name"):
                try:
                    v = getattr(itm, attr, None)
                    if v:
                        typ = str(v).strip()
                        break
                except:
                    pass
        
        if not fam:
            # Ductwork
            if self._is_ductwork_duct(itm):
                nm = safe_str(getattr(itm, "ProductLongDescription", None))
                nm += " " + safe_str(getattr(itm, "ProductName", None))
                nm += " " + safe_str(getattr(itm, "Name", None))
                fam = "Flex Duct" if "flex" in nm.lower() else "Duct"
            # Pipework
            elif self._is_pipework_pipe(itm):
                fam = "Pipe"
            else:
                fam = (self._get_product_data_range(itm) or "").strip()
        
        return fam, typ
    
    def _extract_other(self, element):
        """Extract for other elements."""
        info = self.param_helper.get_type_info(element)
        typ = (info.get("type_name") or "").strip()
        fam = (info.get("family_name") or "").strip()
        return fam, typ
    
    def _is_ductwork_duct(self, itm):
        """Check if ITM is ductwork duct."""
        if not self._is_ductwork(itm):
            return False
        pdr_raw = (self._get_product_data_range(itm) or "").strip().lower()
        if not pdr_raw:
            nm = safe_str(getattr(itm, "ProductLongDescription", None))
            nm += " " + safe_str(getattr(itm, "ProductName", None))
            nm += " " + safe_str(getattr(itm, "Name", None))
            return "duct" in nm.lower()
        return (pdr_raw == "duct" or pdr_raw.endswith(".duct") or 
                pdr_raw.endswith(":duct") or "duct" in pdr_raw)
    
    def _is_ductwork(self, itm):
        """Check if ITM is ductwork."""
        try:
            return self.config.Categories.FAB_DUCTWORK_SUBSTR in (itm.Category.Name or "")
        except:
            return False
    
    def _is_pipework_pipe(self, itm):
        """Check if ITM is pipework pipe straight."""
        try:
            cat_ok = self.config.Categories.FAB_PIPEWORK_SUBSTR in (itm.Category.Name or "")
        except:
            cat_ok = False
        if not cat_ok:
            return False
        try:
            return bool(itm.IsAStraight())
        except:
            return False
    
    def _get_product_data_range(self, itm):
        """Get ITM product data range."""
        try:
            v = getattr(itm, "ProductDataRange", None)
            if v:
                s = str(v).strip()
                if s:
                    return s
        except:
            pass
        for nm in ("ProductDataRange", "Product Data Range"):
            s = self.param_helper.get_param_str(itm, nm).strip()
            if s:
                return s
        return ""


class CommonFieldExtractor(FieldExtractor):
    """Extract common fields across all element types."""
    
    def get_area(self, element):
        """Get construction area."""
        return self.param_helper.get_param_str(element, self.config.Params.AREA)
    
    def get_spool(self, element):
        """Get spool/assembly name."""
        try:
            aid = element.AssemblyInstanceId
            if aid and aid.IntegerValue != -1:
                asm = self.context.doc.GetElement(aid)
                if asm:
                    return (asm.Name or "").strip()
        except:
            pass
        return ""
    
    def get_comments(self, element):
        """Get comments from instance or type."""
        s = self.param_helper.get_bip_str(element, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if s:
            return s
        info = self.param_helper.get_type_info(element)
        s2 = (info.get("type_comments") or "").strip()
        if s2:
            return s2
        return self.param_helper.get_param_str(element, "Comments")
    
    def get_hanger_number(self, element):
        """Get hanger number."""
        return self.param_helper.get_param_str(element, self.config.Params.HANGER_NUMBER)
    
    def get_element_id(self, element):
        """Get Revit element ID."""
        try:
            return str(element.Id.IntegerValue)
        except:
            return ""


# Continuing with more extractors and processors...


class SizeExtractor(FieldExtractor):
    """Extract size field with proper fallback hierarchy."""
    
    def extract_itm(self, element):
        """Extract size for Fabrication Part."""
        # 1. ProductSizeDescription
        base_size = (element.ProductSizeDescription or "")
        
        # 2. Ductwork fallback
        if not base_size and self._is_ductwork(element):
            base_size = (self.param_helper.get_param_str(element, "OverallSize") or
                        self.param_helper.get_param_str(element, "Overall Size") or
                        self.param_helper.get_param_str(element, "Overall_Size") or "")
        
        # 3. Format with inch marks
        formatted = self.context.formatters.size.format(base_size)
        
        # 4. Excel formula wrap
        return self.context.formatters.excel_formula.format(formatted)
    
    def extract_native_curve(self, element):
        """Extract size for native Revit curve."""
        # 1. CP_Size override
        override = self._get_cp_override(element)
        if override:
            formatted = self.context.formatters.size.format(override)
            return self.context.formatters.excel_formula.format(formatted)
        
        # 2. RBS_CALCULATED_SIZE
        return self._get_native_size(element)
    
    def extract_native_fitting(self, element):
        """Extract size for native Revit fitting."""
        # 1. CP_Size override
        override = self._get_cp_override(element)
        if override:
            formatted = self.context.formatters.size.format(override)
            return self.context.formatters.excel_formula.format(formatted)
        
        # 2. RBS_CALCULATED_SIZE
        return self._get_native_size(element)
    
    def extract_rfa(self, element):
        """Extract size for RFA."""
        # 1. CP_Size parameter
        cp_size = self.param_helper.get_param_str(element, self.config.Params.SIZE).strip()
        if cp_size:
            formatted = self.context.formatters.size.format(cp_size)
            return self.context.formatters.excel_formula.format(formatted)
        
        # 2. Native size if available
        native = self._get_native_size(element)
        if native:
            return native
        
        return ""
    
    def _get_cp_override(self, element):
        """Get CP_Size override from instance or type."""
        v = self.param_helper.get_param_str(element, self.config.Params.SIZE).strip()
        if v:
            return v
        try:
            tid = element.GetTypeId()
            if tid and tid.IntegerValue > 0:
                t = self.context.doc.GetElement(tid)
                if t:
                    v = self.param_helper.get_param_str(t, self.config.Params.SIZE).strip()
                    if v:
                        return v
        except:
            pass
        return ""
    
    def _get_native_size(self, element):
        """Get native Revit size (RBS_CALCULATED_SIZE)."""
        s = self.param_helper.get_bip_str(element, BuiltInParameter.RBS_CALCULATED_SIZE).strip()
        if not s:
            return ""
        s = s.replace("/", " x ")
        s = self.context.formatters.size.format(s)
        return self.context.formatters.excel_formula.format(s)
    
    def _is_ductwork(self, itm):
        """Check if ITM is ductwork."""
        try:
            return self.config.Categories.FAB_DUCTWORK_SUBSTR in (itm.Category.Name or "")
        except:
            return False


class LengthExtractor(FieldExtractor):
    """Extract length field."""
    
    def extract_itm(self, element):
        """Extract length for Fabrication Part."""
        # Fishmouth override: use ProductName text, not CID.
        # CID 2875 is shared by olets, T-drills, and other taps — text is the
        # only reliable discriminator confirmed by diagnostic sweep.
        try:
            pn = (getattr(element, "ProductName", None) or "").strip().lower()
            if pn == "fishmouths":
                return float(element.CenterlineLength)
        except:
            pass

        # Straight detection
        try:
            is_straight = bool(element.IsAStraight())
        except:
            is_straight = False

        if not is_straight:
            return 0.0

        # Get centerline length
        try:
            return float(element.CenterlineLength)
        except:
            pass

        # Fallback parameters
        for nm in ("Centerline Length", "CenterlineLength", "Length"):
            v = self.param_helper.get_param_double(element, nm)
            if v is not None:
                try:
                    return float(v)
                except:
                    pass

        return 0.0
    
    def extract_native_curve(self, element):
        """Extract length for native Revit curve."""
        p = self.param_helper.get_bip_param(element, BuiltInParameter.CURVE_ELEM_LENGTH)
        try:
            return float(p.AsDouble()) if p else None
        except:
            return None
    
    def extract_rfa(self, element):
        """Extract length for RFA."""
        v = self.param_helper.get_param_double(element, self.config.Params.LENGTH)
        return float(v) if v is not None else None


class MaterialExtractor(FieldExtractor):
    """Extract material field."""
    
    def extract_itm(self, element):
        """Extract material for Fabrication Part."""
        # 1. Primary: product material
        try:
            v = getattr(element, "ProductMaterialDescription", None)
            if v:
                return str(v).strip()
        except:
            pass
        
        # 2. Fallback: parameter
        for nm in ("Product Material Description", "ProductMaterialDescription"):
            s = self.param_helper.get_param_str(element, nm)
            if s:
                return s.strip()
        
        # 3. Last resort: fabrication material
        p = self.param_helper.get_bip_param(element, BuiltInParameter.FABRICATION_PART_MATERIAL)
        s = self.param_helper.get_material_name_from_param(p)
        if s:
            return s
        
        return ""
    
    def extract_native(self, element):
        """Extract material for native Revit element."""
        # 1. Instance 'Material'
        v = self.param_helper.get_param_str(element, "Material").strip()
        if v:
            return v
        
        # 2. Type 'Material'
        try:
            info = self.param_helper.get_type_info(element)
            te = info.get("type_el")
        except:
            te = None
        return self.param_helper.get_param_str(te, "Material").strip() if te else ""
    
    def extract_rfa(self, element):
        """Extract material for RFA."""
        return self.param_helper.get_param_str(element, self.config.Params.MATERIAL).strip()


class ServiceExtractor(FieldExtractor):
    """Extract service field."""
    
    def extract_itm(self, element):
        """Extract service for Fabrication Part."""
        try:
            v = getattr(element, "ServiceName", None)
            if v:
                return str(v).strip()
        except:
            pass
        return ""
    
    def extract_native(self, element):
        """Extract service for native Revit element."""
        cls = self.param_helper.get_bip_str(element, BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).strip()
        nm = self.param_helper.get_bip_str(element, BuiltInParameter.RBS_SYSTEM_NAME_PARAM).strip()
        if cls and nm:
            return cls + " - " + nm
        return cls or nm or ""
    
    def extract_rfa(self, element):
        """Extract service for RFA."""
        return self.param_helper.get_param_str(element, self.config.Params.SERVICE).strip()


class InstallTypeExtractor(FieldExtractor):
    """Extract install type field."""
    
    def extract_itm(self, element):
        """Extract install type for Fabrication Part."""
        for nm in ("Install Type", "ITM Install Type", "Fab Install Type"):
            s = self.param_helper.get_param_str(element, nm)
            if s:
                return s
        return ""
    
    def extract_native(self, element):
        """Extract install type for native Revit element."""
        return self.param_helper.get_param_str(element, "Connection Type").strip()


class ModelExtractor(FieldExtractor):
    """Extract model field."""
    
    def extract_itm(self, element):
        """Extract model for Fabrication Part."""
        try:
            v = getattr(element, "ProductName", None)
            if v:
                return str(v).strip()
        except:
            pass
        for nm in ("Product Name", "ProductName", "Product"):
            s = self.param_helper.get_param_str(element, nm)
            if s:
                return s
        return "N/A"
    
    def extract_rfa(self, element):
        """Extract model for RFA."""
        def _is_blankish(s):
            s = (s or "").strip()
            return (not s) or (s.lower() in ("n/a", "na", "-", "none"))
        
        # Instance CP_Model Number
        s = self.param_helper.get_param_str(element, self.config.Params.MODEL_NUMBER)
        if not _is_blankish(s):
            return s.strip()
        
        # Type CP_Model Number
        sym = None
        try:
            sym = element.Symbol
        except:
            sym = None
        
        if sym:
            s2 = self.param_helper.get_param_str(sym, self.config.Params.MODEL_NUMBER)
            if not _is_blankish(s2):
                return s2.strip()
            
            # Built-in Model parameter
            s3 = self.param_helper.get_param_str(sym, "Model")
            if not _is_blankish(s3):
                return s3.strip()
        
        return "N/A"


class DescriptionExtractor(FieldExtractor):
    """Extract description field."""
    
    def extract_itm(self, element, row_data):
        """Extract description for Fabrication Part.
        
        Args:
            element: The FabricationPart
            row_data: BOMRow with family/type already populated
        """
        # 1. CP_BOM_Description
        v = self.param_helper.get_param_str(element, self.config.Params.DESCRIPTION).strip()
        if v:
            return v
        
        # 2. CP_Type Name
        v = self.param_helper.get_param_str(element, self.config.Params.TYPE_NAME).strip()
        if v:
            return v
        
        # 3. Family - Type
        _f = (row_data.family or "").strip()
        _t = (row_data.type or "").strip()
        if _f and _t and _t.lower() not in ("default", "standard", "unknown"):
            return "{} - {}".format(_f, _t)
        
        # 4. ProductLongDescription
        base_desc = (element.ProductLongDescription or "")
        if base_desc.strip():
            return base_desc.strip()
        
        # 5. Category name
        try:
            cn = element.Category.Name if element.Category else ""
            if cn:
                return cn
        except:
            pass
        
        return "N/A"
    
    def extract_rfa(self, element):
        """Extract description for RFA."""
        # 1. Instance CP_Type Name
        v = self.param_helper.get_param_str(element, self.config.Params.TYPE_NAME).strip()
        if v:
            return v
        
        # 2. Type CP_Type Name
        te = None
        try:
            te = self.context.doc.GetElement(element.GetTypeId())
        except:
            te = None
        if te:
            v = self.param_helper.get_param_str(te, self.config.Params.TYPE_NAME).strip()
            if v:
                return v
        
        # 3. Family - Type
        fam = ""
        typ = ""
        try:
            info = self.param_helper.get_type_info(element)
            fam = (info.get("family_name") or "").strip()
            typ = (info.get("type_name") or "").strip()
        except:
            fam = ""
            typ = ""
        
        if not fam or not typ:
            try:
                if not fam:
                    fam = (element.Symbol.Family.Name or "").strip()
            except:
                pass
            try:
                if not typ:
                    typ = (element.Symbol.Name or "").strip()
            except:
                pass
        
        if fam and typ:
            return "{} - {}".format(fam, typ)
        
        return "MISSING_RFA_FAMILY_OR_TYPE"
    
    def extract_native_curve(self, element, type_name, category_name):
        """Extract description for native Revit curve."""
        if (type_name or "").strip() and (type_name or "").strip() != "N/A":
            return "{} - {}".format(category_name or "MEP", (type_name or "").strip())
        else:
            return category_name or "MEP"
    
    def extract_native_fitting(self, element, family_name, type_name, category_name):
        """Extract description for native Revit fitting."""
        _t = (type_name or '').strip()
        _f = (family_name or '').strip()
        if _t and _t != 'N/A':
            return '{} - {}'.format(_f or (category_name or 'MEP'), _t)
        else:
            return _f or (category_name or 'MEP')


class ExtractorRegistry:
    """Container for all field extractors."""
    
    def __init__(self, context):
        self.category = CategoryExtractor(context)
        self.item_number = ItemNumberExtractor(context)
        self.family_type = FamilyTypeExtractor(context)
        self.common = CommonFieldExtractor(context)
        self.size = SizeExtractor(context)
        self.length = LengthExtractor(context)
        self.material = MaterialExtractor(context)
        self.service = ServiceExtractor(context)
        self.install_type = InstallTypeExtractor(context)
        self.model = ModelExtractor(context)
        self.description = DescriptionExtractor(context)


# Continuing with processors and pipeline...


# =============================================================================
# HANGER PROCESSING
# =============================================================================

class HangerProcessor:
    """Specialized processor for fabrication hangers and rod derivation."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
    
    def is_hanger(self, element, dia_ft, len_ft):
        """Check if element is a hanger based on cached rod usage."""
        if not isinstance(element, FabricationPart):
            return False
        if dia_ft is not None or len_ft is not None:
            return True
        try:
            cn = (element.Category.Name or "").lower()
        except:
            cn = ""
        return ("fabrication hangers" in cn) or ("hangers" in cn)
    
    def get_rod_usage(self, element):
        """Get rod diameter and length from ancillary usage.
        
        Returns:
            (dia_ft, len_ft) tuple or (None, None)
        """
        if not isinstance(element, FabricationPart):
            return (None, None)
        
        # Primary: ancillary usage
        try:
            anc_list = list(element.GetPartAncillaryUsage())
        except:
            anc_list = []
        
        for anc in anc_list:
            try:
                u = safe_str(getattr(anc, "UsageType", "")).lower()
                t = safe_str(getattr(anc, "Type", "")).lower()
            except:
                u, t = "", ""
            
            if u == "hanger" and t == "supportrod":
                dia_ft = getattr(anc, "AncillaryWidthOrDiameter", None)
                len_ft = getattr(anc, "Length", None)
                
                try:
                    dia_ft = float(dia_ft) if dia_ft is not None else None
                except:
                    dia_ft = None
                try:
                    len_ft = float(len_ft) if len_ft is not None else None
                except:
                    len_ft = None
                
                if dia_ft is not None and abs(dia_ft) <= self.config.Rules.ZERO_TOLERANCE_FT:
                    dia_ft = None
                if len_ft is not None and abs(len_ft) <= self.config.Rules.ZERO_TOLERANCE_FT:
                    len_ft = None
                
                return (dia_ft, len_ft)
        
        # Fallback: parse from text
        dia_ft = self._parse_rod_diameter_ft(element)
        if dia_ft is not None and abs(dia_ft) > self.config.Rules.ZERO_TOLERANCE_FT:
            return (dia_ft, None)
        
        return (None, None)
    
    def build_dimension_lookup(self, element):
        """Build dimension name -> definition lookup."""
        try:
            dims = list(element.GetDimensions())
        except:
            return {}
        
        lut = {}
        for d in dims:
            try:
                name = (d.Name or '').strip().lower()
            except:
                name = ''
            if name:
                lut[name] = d
        return lut
    
    def get_rod_lengths_ab_ft(self, element, dim_lut=None):
        """Get Length A and Length B from fabrication dimensions."""
        if not isinstance(element, FabricationPart):
            return (None, None)
        
        name_to_def = dim_lut or self.build_dimension_lookup(element)
        
        def _get_dim_ft(name_lower):
            ddef = name_to_def.get(name_lower)
            if not ddef:
                return None
            
            v = None
            try:
                v = element.GetDimensionValue(ddef)
            except:
                v = None
            
            if v is None:
                try:
                    v = element.GetCalculatedDimensionValue(ddef)
                except:
                    v = None
            
            if v is None:
                return None
            
            try:
                v = float(v)
            except:
                return None
            
            try:
                if abs(v) < self.config.Rules.ZERO_TOLERANCE_FT:
                    return None
            except:
                pass
            
            return v
        
        return (_get_dim_ft("length a"), _get_dim_ft("length b"))
    
    def get_strut_length_ft(self, element, dim_lut=None):
        """Get trapeze strut length: Width + 2 * Bearer Extn."""
        if not isinstance(element, FabricationPart):
            return None
        
        name_to_def = dim_lut or self.build_dimension_lookup(element)
        
        def _get_dim_ft(name_lower):
            ddef = name_to_def.get(name_lower)
            if not ddef:
                return None
            v = None
            try:
                v = element.GetDimensionValue(ddef)
            except:
                v = None
            if v is None:
                try:
                    v = element.GetCalculatedDimensionValue(ddef)
                except:
                    v = None
            if v is None:
                return None
            try:
                return float(v)
            except:
                return None
        
        width_ft = _get_dim_ft("width")
        bext_ft = _get_dim_ft("bearer extn")
        
        if width_ft is None or bext_ft is None:
            return None
        
        return width_ft + (2.0 * bext_ft)
    
    def derive_rod_rows(self, base_row, dia_ft, lenA_ft, lenB_ft, element):
        """Derive rod rows from hanger base row.
        
        Returns:
            List of BOMRow objects for rods
        """
        rows = []
        
        # Rod A
        if (lenA_ft is not None) and (abs(lenA_ft) > self.config.Rules.ZERO_TOLERANCE_FT):
            rod_row = self._build_rod_row(base_row, dia_ft, lenA_ft, element)
            if rod_row:
                rows.append(rod_row)
            
            # Rod B (only if non-zero)
            if (lenB_ft is not None) and (abs(lenB_ft) > self.config.Rules.ZERO_TOLERANCE_FT):
                rod_row_2 = self._build_rod_row(base_row, dia_ft, lenB_ft, element)
                if rod_row_2:
                    rows.append(rod_row_2)
        
        return rows
    
    def _build_rod_row(self, hanger_row, dia_ft, rod_len_ft, itm_hanger):
        """Build a single rod row from hanger base row."""
        if dia_ft is None and rod_len_ft is None:
            return None
        
        rod_row = hanger_row.copy()
        rod_row._is_derived = True
        
        rod_row.category = "Structural Framing"
        rod_row.family = "All-Thread Rod"
        rod_row.description = "All-Thread Rod"
        rod_row.type = "N/A"
        rod_row.model = "N/A"
        rod_row.install_type = "N/A"
        
        # Size = rod diameter
        if dia_ft is not None:
            dia_in = dia_ft * 12.0
            rod_row.size = self.context.formatters.excel_formula.format(
                self.context.formatters.inches_fractional.format(dia_in)
            )
        else:
            rod_row.size = self.context.formatters.excel_formula.format("N/A")
        
        # Length
        rod_row.length_ft = self.context.formatters.length_decimal.format(rod_len_ft)
        rod_row.length_ft_in = self.context.formatters.length_fractional.format(rod_len_ft)
        rod_row.count = self.context.formatters.count_flag.format(rod_len_ft)
        
        # Sorting will be recomputed later in enrichment stage
        
        return rod_row
    
    def _parse_rod_diameter_ft(self, itm_part):
        """Parse rod diameter from ITM text fields."""
        if not isinstance(itm_part, FabricationPart):
            return None
        
        try:
            txt = " ".join([
                safe_str(getattr(itm_part, "ProductLongDescription", None)),
                safe_str(getattr(itm_part, "ProductName", None)),
                safe_str(getattr(itm_part, "Name", None)),
            ])
        except:
            txt = ""
        
        if not txt:
            return None
        
        s = txt.lower().replace("\t", " ").replace("\r", " ").replace("\n", " ")
        
        rod_tokens = ("rod", "all thread", "all-thread", "allthread")
        has_rod_token = any(tok in s for tok in rod_tokens)
        
        is_hanger_cat = False
        try:
            cn = (itm_part.Category.Name or "").lower()
            is_hanger_cat = ("fabrication hangers" in cn) or ("hangers" in cn)
        except:
            is_hanger_cat = False
        
        if (not has_rod_token) and (not is_hanger_cat):
            return None
        
        candidates_in = []
        
        # Fractions (only with rod tokens)
        if has_rod_token:
            for mm in re.finditer(r"\b(\d+)\s*/\s*(\d+)\b", s):
                try:
                    num = float(mm.group(1))
                    den = float(mm.group(2))
                    if den != 0:
                        candidates_in.append(num / den)
                except:
                    pass
        
        # Decimals
        for mm in re.finditer(r"\b(0?\.\d{2,6})\b", s):
            try:
                candidates_in.append(float(mm.group(1)))
            except:
                pass
        
        if not candidates_in:
            return None
        
        rod_like = [v for v in candidates_in if 0.10 <= abs(v) <= 1.00]
        if not rod_like:
            return None
        
        inch = rod_like[-1]
        try:
            return float(inch) / 12.0
        except:
            return None


# =============================================================================
# SORTING
# =============================================================================

class SortingService:
    """Assigns sorting codes for Excel pivot stability."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
        self.cache = context.cache
    
    def get_sorting_value(self, element, category_name):
        """Get final sorting value for element."""
        # Special case: Fabrication Pipework
        if isinstance(element, FabricationPart) and self._is_pipework_fp(element):
            return self._pipework_sort_code(element)
        
        base = self._series_base_for_element(element, category_name)
        return str(self._assign_code(base, category_name))
    
    def _series_base_for_element(self, element, category_name):
        """Get base series number for element type."""
        if isinstance(element, FabricationPart):
            try:
                if bool(element.IsAStraight()):
                    return 100
                elif bool(element.IsATap()):
                    return 300
                else:
                    return 200
            except:
                return 200
        
        if category_name == "Conduit Fittings" or self._is_structuralish(category_name):
            return 600
        
        return 900
    
    def _assign_code(self, base, category_name):
        """Assign unique sorting code for category."""
        cat = (category_name or "").strip() or "Uncategorized"
        
        # Check cache
        cached = self.cache.category_codes.get(cat)
        if cached is not None:
            return cached
        
        # Check preferences
        if cat in self.config.Sorting.PREFERENCES:
            code = int(self.config.Sorting.PREFERENCES[cat])
            self.cache.sort_codes_used.add(code)
            self.cache.category_codes[cat] = code
            return code
        
        # Hash-based assignment
        slot = (self._stable_hash(cat) % 9) + 1
        code = base + slot * 10
        tries = 0
        while code in self.cache.sort_codes_used and tries < 50:
            slot += 1
            if slot > 9:
                slot = 1
            code = base + slot * 10
            tries += 1
        
        self.cache.sort_codes_used.add(code)
        self.cache.category_codes[cat] = code
        return code
    
    @staticmethod
    def _stable_hash(s):
        """Stable hash for category names."""
        if s is None:
            s = ""
        try:
            s = str(s)
        except:
            s = ""
        h = 2166136261
        for ch in s:
            h = h ^ ord(ch)
            h = (h * 16777619) & 0xFFFFFFFF
        return h
    
    def _is_structuralish(self, category_name):
        """Check if category is structural-related."""
        s = (category_name or "").lower()
        return any(k in s for k in self.config.Sorting.STRUCTURAL_KEYWORDS)
    
    def _is_pipework_fp(self, fp):
        """Check if FabricationPart is Pipework."""
        try:
            return fp.Category and (self.config.Categories.FAB_PIPEWORK_SUBSTR in (fp.Category.Name or ""))
        except:
            return False
    
    def _pipework_sort_code(self, fp):
        """Get sorting code for Fabrication Pipework.
        
        Sort order: pipe straight (110) → fishmouth (120) → fittings (210) →
                    taps/olets/etc (220) → welds (240) → other (230)
        """
        if self._is_pipe_weld(fp):
            return "240"
        try:
            if bool(fp.IsAStraight()):
                return "110"
        except:
            pass
        # Fishmouth: sits right after pipe, before all other fittings
        try:
            pn = (getattr(fp, "ProductName", None) or "").strip().lower()
            if pn == "fishmouths":
                return "120"
        except:
            pass
        try:
            if (not bool(fp.IsAStraight())) and (not bool(fp.IsATap())):
                return "210"
        except:
            pass
        try:
            if bool(fp.IsATap()):
                return "220"
        except:
            pass
        return "230"
    
    def _is_pipe_weld(self, fp):
        """Check if pipe part is a weld."""
        for attr in ("ItemType", "PartType", "FabricationPartType"):
            v = getattr(fp, attr, None)
            if v is not None:
                try:
                    if "weld" in str(v).lower():
                        return True
                except:
                    pass
        for attr in ("ItemName", "Name"):
            v = getattr(fp, attr, None)
            if v:
                try:
                    if "weld" in str(v).lower():
                        return True
                except:
                    pass
        fam = self._pipe_family_str(fp)
        return "weld" in (fam or "").lower()
    
    def _pipe_family_str(self, fp):
        """Get pipe family string."""
        try:
            p = fp.LookupParameter("Family")
            if p:
                try:
                    return (p.AsValueString() or "").strip()
                except:
                    try:
                        return (p.AsString() or "").strip()
                    except:
                        return ""
        except:
            pass
        return ""


# Continuing with Element Processors...


# =============================================================================
# ELEMENT PROCESSORS
# =============================================================================

class ElementProcessor:
    """Base class for element-specific processors.
    
    Each processor knows how to:
    1. Identify if it can handle an element
    2. Extract all fields for that element type
    3. Generate BOMRow(s) including any derived rows
    """
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
        self.extractors = context.extractors
        self.formatters = context.formatters
    
    def can_process(self, element):
        """Can this processor handle this element?"""
        raise NotImplementedError
    
    def process(self, element):
        """Process element into list of BOMRow objects."""
        raise NotImplementedError


class FabricationPartProcessor(ElementProcessor):
    """Processor for Fabrication Parts (ITMs)."""
    
    def can_process(self, element):
        return isinstance(element, FabricationPart)
    
    def process(self, element):
        """Process ITM into BOMRow(s) - may include derived rod rows."""
        row = BOMRow()
        row._source_element = element
        
        # Extract fields using specialized extractors
        row.category = self.extractors.category.extract(element)
        row.family, row.type = self.extractors.family_type.extract(element)
        row.item_number = self.extractors.item_number.extract(element)
        row.area = self.extractors.common.get_area(element)
        row.comments = self.extractors.common.get_comments(element)
        row.spool = self.extractors.common.get_spool(element)
        row.hanger_number = self.extractors.common.get_hanger_number(element)
        row.element_id = self.extractors.common.get_element_id(element)
        
        # Length
        length_ft = self.extractors.length.extract_itm(element)
        row.length_ft = self.formatters.length_decimal.format(length_ft)
        row.length_ft_in = self.formatters.length_fractional.format(length_ft)
        row.count = self.formatters.count_flag.format(length_ft)
        
        # Size
        row.size = self.extractors.size.extract_itm(element)
        
        # Description (needs row data for Family - Type fallback)
        row.description = self.extractors.description.extract_itm(element, row)
        
        # Ductwork override: description should be family name only
        if self._is_ductwork(element):
            fam_name = (row.family or "").strip() or (row.type or "").strip()
            if fam_name:
                row.description = fam_name
        
        # Material, Service, Install Type, Model
        row.material = self.extractors.material.extract_itm(element)
        row.service = self.extractors.service.extract_itm(element)
        row.install_type = self.extractors.install_type.extract_itm(element)
        row.model = self.extractors.model.extract_itm(element)
        
        # Validate before continuing
        row.validate()
        
        rows = [row]
        
        # Hanger processing
        hanger_proc = HangerProcessor(self.context)
        dia_ft, rod_len_ft = hanger_proc.get_rod_usage(element)
        
        if hanger_proc.is_hanger(element, dia_ft, rod_len_ft):
            # Build dimension lookup once
            dim_lut = hanger_proc.build_dimension_lookup(element)
            lenA_ft, lenB_ft = hanger_proc.get_rod_lengths_ab_ft(element, dim_lut)
            
            # Two-rod hanger: use strut length on base row
            is_two_rod = (lenB_ft is not None) and (abs(lenB_ft) > self.config.Rules.ZERO_TOLERANCE_FT)
            if is_two_rod:
                strut_len_ft = hanger_proc.get_strut_length_ft(element, dim_lut)
                if strut_len_ft is not None and abs(strut_len_ft) > self.config.Rules.ZERO_TOLERANCE_FT:
                    row.length_ft = self.formatters.length_decimal.format(strut_len_ft)
                    row.length_ft_in = self.formatters.length_fractional.format(strut_len_ft)
                    row.count = self.formatters.count_flag.format(strut_len_ft)
            
            # Generate rod rows
            rod_rows = hanger_proc.derive_rod_rows(row, dia_ft, lenA_ft, lenB_ft, element)
            rows.extend(rod_rows)
        
        return rows
    
    def _is_pipework_pipe(self, itm):
        """Check if ITM is a pipework pipe straight."""
        try:
            cat_ok = self.config.Categories.FAB_PIPEWORK_SUBSTR in (itm.Category.Name or "")
        except:
            cat_ok = False
        if not cat_ok:
            return False
        try:
            return bool(itm.IsAStraight())
        except:
            return False

    def _is_ductwork(self, itm):
        """Check if ITM is ductwork."""
        try:
            return self.config.Categories.FAB_DUCTWORK_SUBSTR in (itm.Category.Name or "")
        except:
            return False


class NativeMEPCurveProcessor(ElementProcessor):
    """Processor for native Revit MEP curves."""
    
    def can_process(self, element):
        classifier = ElementClassifier(self.context)
        return classifier.is_revit_mep_curve(element)
    
    def process(self, element):
        """Process native MEP curve into BOMRow."""
        row = BOMRow()
        row._source_element = element
        
        # Extract fields
        row.category = self.extractors.category.extract(element)
        row.family, row.type = self.extractors.family_type.extract(element)
        row.item_number = self.extractors.item_number.extract(element)
        row.area = self.extractors.common.get_area(element)
        row.comments = self.extractors.common.get_comments(element)
        row.spool = self.extractors.common.get_spool(element)
        row.hanger_number = self.extractors.common.get_hanger_number(element)
        row.element_id = self.extractors.common.get_element_id(element)
        
        # Length
        length_ft = self.extractors.length.extract_native_curve(element)
        row.length_ft = self.formatters.length_decimal.format(length_ft)
        row.length_ft_in = self.formatters.length_fractional.format(length_ft)
        row.count = self.formatters.count_flag.format(length_ft)
        
        # Size
        row.size = self.extractors.size.extract_native_curve(element)
        
        # Material, Service, Install Type
        row.material = self.extractors.material.extract_native(element)
        row.service = self.extractors.service.extract_native(element)
        row.install_type = self.extractors.install_type.extract_native(element)
        
        # Model
        row.model = "N/A"
        
        # Description
        param_helper = ParameterHelper(self.context)
        try:
            cat_name = element.Category.Name if element.Category else "MEP"
        except:
            cat_name = "MEP"
        row.description = self.extractors.description.extract_native_curve(element, row.type, cat_name)
        
        row.validate()
        return [row]


class NativeMEPFittingProcessor(ElementProcessor):
    """Processor for native Revit MEP fittings (Pipe Fittings, Duct Fittings).
    
    Handles native Revit fittings like elbows, tees, and reducers that don't
    carry CP parameters. Pipe/Duct Accessories are excluded and handled by
    RFAProcessor, which reads CP_Length and other CP fields.
    """
    
    def can_process(self, element):
        classifier = ElementClassifier(self.context)
        return classifier.is_revit_mep_fitting(element)
    
    def process(self, element):
        """Process native MEP fitting into BOMRow."""
        row = BOMRow()
        row._source_element = element
        
        # Extract fields
        row.category = self.extractors.category.extract(element)
        row.family, row.type = self.extractors.family_type.extract(element)
        row.item_number = self.extractors.item_number.extract(element)
        row.area = self.extractors.common.get_area(element)
        row.comments = self.extractors.common.get_comments(element)
        row.spool = self.extractors.common.get_spool(element)
        row.hanger_number = self.extractors.common.get_hanger_number(element)
        row.element_id = self.extractors.common.get_element_id(element)
        
        # No length for fittings
        row.length_ft = ""
        row.length_ft_in = ""
        row.count = "1"
        
        # Size
        row.size = self.extractors.size.extract_native_fitting(element)
        
        # Material, Service, Install Type
        row.material = self.extractors.material.extract_native(element)
        row.service = self.extractors.service.extract_native(element)
        row.install_type = self.extractors.install_type.extract_native(element)
        
        # Model
        row.model = "N/A"
        
        # Description
        param_helper = ParameterHelper(self.context)
        try:
            cat_name = element.Category.Name if element.Category else "MEP"
        except:
            cat_name = "MEP"
        row.description = self.extractors.description.extract_native_fitting(element, row.family, row.type, cat_name)
        
        row.validate()
        return [row]


class RFAProcessor(ElementProcessor):
    """Processor for RFA/FamilyInstance elements."""
    
    def can_process(self, element):
        return isinstance(element, FamilyInstance)
    
    def process(self, element):
        """Process RFA into BOMRow."""
        row = BOMRow()
        row._source_element = element
        
        # Extract fields
        row.category = self.extractors.category.extract(element)
        row.family, row.type = self.extractors.family_type.extract(element)
        row.item_number = self.extractors.item_number.extract(element)
        row.area = self.extractors.common.get_area(element)
        row.comments = self.extractors.common.get_comments(element)
        row.spool = self.extractors.common.get_spool(element)
        row.hanger_number = self.extractors.common.get_hanger_number(element)
        row.element_id = self.extractors.common.get_element_id(element)
        
        # Description
        row.description = self.extractors.description.extract_rfa(element)
        
        # Size
        classifier = ElementClassifier(self.context)
        if classifier.is_native_familyinstance(element):
            # Native fitting: use native size
            row.size = self.extractors.size.extract_native_fitting(element)
        else:
            # CP-driven RFA: use RFA size
            row.size = self.extractors.size.extract_rfa(element)
        
        # Length
        length_ft = self.extractors.length.extract_rfa(element)
        row.length_ft = self.formatters.length_decimal.format(length_ft)
        row.length_ft_in = self.formatters.length_fractional.format(length_ft)
        row.count = self.formatters.count_flag.format(length_ft)
        
        # Model
        row.model = self.extractors.model.extract_rfa(element)
        
        # Material
        if classifier.is_native_familyinstance(element):
            row.material = self.extractors.material.extract_native(element)
        else:
            mat = self.extractors.material.extract_rfa(element)
            if not mat:
                mat = self.extractors.material.extract_native(element)
            row.material = mat
        
        # Service
        if classifier.is_native_familyinstance(element):
            row.service = self.extractors.service.extract_native(element)
        else:
            svc = self.extractors.service.extract_rfa(element)
            if not svc:
                svc = self.extractors.service.extract_native(element)
            row.service = svc
        
        # Install Type
        if classifier.is_native_familyinstance(element):
            row.install_type = self.extractors.install_type.extract_native(element)
        else:
            row.install_type = ""
        
        row.validate()
        return [row]


class FallbackProcessor(ElementProcessor):
    """Fallback processor for any unhandled element types."""
    
    def can_process(self, element):
        return True  # Catches everything
    
    def process(self, element):
        """Process unknown element into minimal BOMRow."""
        row = BOMRow()
        row._source_element = element
        
        row.category = self.extractors.category.extract(element)
        row.family, row.type = self.extractors.family_type.extract(element)
        row.item_number = self.extractors.item_number.extract(element)
        row.area = self.extractors.common.get_area(element)
        row.comments = self.extractors.common.get_comments(element)
        row.spool = self.extractors.common.get_spool(element)
        row.hanger_number = self.extractors.common.get_hanger_number(element)
        row.element_id = self.extractors.common.get_element_id(element)
        
        # Minimal description
        param_helper = ParameterHelper(self.context)
        try:
            row.description = element.Category.Name if element.Category else ""
        except:
            row.description = ""
        
        # Defaults
        row.size = self.formatters.excel_formula.format("N/A")
        row.length_ft = ""
        row.length_ft_in = ""
        row.count = "1"
        row.model = "N/A"
        row.material = "N/A"
        row.service = ""
        row.install_type = ""
        
        row.validate()
        return [row]


# Continuing with Pipeline...


# =============================================================================
# NESTED RFA EXPANSION
# =============================================================================

class RFAExpander:
    """Expands selection to include nested family instances."""
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
    
    def expand_selection(self, elements):
        """Split elements into non-RFA and expanded RFA lists.
        
        Args:
            elements: List of selected elements
            
        Returns:
            (non_rfa_list, rfa_list) tuple
        """
        non_rfa = []
        seed_rfa_ids = set()
        
        for el in elements:
            if isinstance(el, FamilyInstance):
                try:
                    subs = el.GetSubComponentIds()
                    if subs and subs.Count > 0:
                        seed_rfa_ids.add(el.Id.IntegerValue)
                    else:
                        non_rfa.append(el)
                except:
                    non_rfa.append(el)
            else:
                non_rfa.append(el)
        
        all_rfa_ids = set(seed_rfa_ids)
        
        # Recursive expansion
        for rid in list(seed_rfa_ids):
            host = self.context.doc.GetElement(ElementId(rid))
            if isinstance(host, FamilyInstance):
                self._collect_nested(host, all_rfa_ids, depth=0)
        
        # Resolve IDs to elements
        rfas = []
        for rid in all_rfa_ids:
            fi = self.context.doc.GetElement(ElementId(rid))
            if isinstance(fi, FamilyInstance):
                rfas.append(fi)
        
        return non_rfa, rfas
    
    def _collect_nested(self, host_fi, out_ids, depth):
        """Recursively collect nested subcomponents."""
        if depth >= self.config.Rules.MAX_NESTED_RFA_DEPTH:
            return
        
        try:
            sub_ids = list(host_fi.GetSubComponentIds())
        except:
            sub_ids = []
        
        for sid in sub_ids:
            try:
                sub_el = self.context.doc.GetElement(sid)
            except:
                sub_el = None
            
            if isinstance(sub_el, FamilyInstance):
                eid = sub_el.Id.IntegerValue
                if eid not in out_ids:
                    out_ids.add(eid)
                    self._collect_nested(sub_el, out_ids, depth + 1)


# =============================================================================
# PIPELINE
# =============================================================================

class BOMExportPipeline:
    """Main orchestrator for the BOM export process.
    
    Pipeline stages:
        INPUT → VALIDATE → EXPAND → CLASSIFY → PROCESS → ENRICH → SORT → SERIALIZE → OUTPUT
    """
    
    def __init__(self, context):
        self.context = context
        self.config = context.config
        
        # Initialize processors (order matters - first match wins)
        # 1. FabricationPart (ITMs) - checked first via isinstance
        # 2. Native MEP curves (Pipes, Ducts, Flex) - by category ID
        # 3. Native MEP fittings (Pipe/Duct Fittings only, NOT accessories) - by category ID
        # 4. RFA/FamilyInstance (all remaining, including Pipe/Duct Accessories) - via isinstance
        # 5. Fallback catch-all for anything else
        self.processors = [
            FabricationPartProcessor(context),
            NativeMEPCurveProcessor(context),
            NativeMEPFittingProcessor(context),
            RFAProcessor(context),
            FallbackProcessor(context),  # Catch-all
        ]
        
        self.sorting_service = SortingService(context)
    
    def execute(self, elements):
        """Execute the full export pipeline.
        
        Args:
            elements: Selected Revit elements
            
        Returns:
            TSV-formatted string ready for clipboard
        """
        log_debug("Pipeline started", "PIPELINE")
        
        # Stage 1: Validate input
        elements = self._stage_validate(elements)
        log_debug("Validated {} elements".format(len(elements)), "VALIDATE")
        
        # Stage 2: Expand nested RFAs
        elements = self._stage_expand(elements)
        log_debug("Expanded to {} elements".format(len(elements)), "EXPAND")
        
        # Stage 3: Classify & Process
        rows = self._stage_process(elements)
        log_debug("Processed to {} rows".format(len(rows)), "PROCESS")
        
        # Stage 4: Enrich (sorting codes)
        rows = self._stage_enrich(rows)
        log_debug("Enriched rows", "ENRICH")
        
        # Stage 5: Sort
        rows = self._stage_sort(rows)
        log_debug("Sorted rows", "SORT")
        
        # Stage 6: Serialize to TSV
        output = self._stage_serialize(rows)
        log_debug("Serialized to TSV", "SERIALIZE")
        
        return output
    
    def _stage_validate(self, elements):
        """Stage 1: Filter out invalid elements."""
        classifier = ElementClassifier(self.context)
        
        valid = []
        for el in elements:
            if classifier.is_modeling_element(el):
                if not classifier.should_exclude_by_family(el):
                    valid.append(el)
        
        return valid
    
    def _stage_expand(self, elements):
        """Stage 2: Expand nested RFAs."""
        expander = RFAExpander(self.context)
        non_rfa, rfas = expander.expand_selection(elements)
        
        # Re-validate expanded elements
        classifier = ElementClassifier(self.context)
        non_rfa = [e for e in non_rfa if classifier.is_modeling_element(e) and 
                  (not classifier.should_exclude_by_family(e))]
        rfas = [fi for fi in rfas if classifier.is_modeling_element(fi) and 
               (not classifier.should_exclude_by_family(fi))]
        
        return non_rfa + rfas
    
    def _stage_process(self, elements):
        """Stage 3: Process each element into BOMRow(s)."""
        rows = []
        
        for element in elements:
            processor = self._find_processor(element)
            element_rows = processor.process(element)
            rows.extend(element_rows)
        
        return rows
    
    def _stage_enrich(self, rows):
        """Stage 4: Enrich rows with sorting codes."""
        for row in rows:
            if row._source_element:
                row.sorting = self.sorting_service.get_sorting_value(
                    row._source_element, 
                    row.category
                )
        
        return rows
    
    def _stage_sort(self, rows):
        """Stage 5: Sort rows for Excel export."""
        def _sort_int(x):
            try:
                return int(str(x).strip())
            except:
                return 999999
        
        rows.sort(key=lambda r: (
            _sort_int(r.sorting),
            (r.category or ""),
            (r.description or ""),
            (r.family or ""),
            (r.type or ""),
        ))
        
        return rows
    
    def _stage_serialize(self, rows):
        """Stage 6: Convert rows to TSV format."""
        lines = ["\t".join(self.config.Columns.HEADER)]
        
        for row in rows:
            row_list = row.to_list()
            lines.append("\t".join([sanitize_tsv(c) for c in row_list]))
        
        return "\r\n".join(lines)
    
    def _find_processor(self, element):
        """Find the appropriate processor for this element."""
        for processor in self.processors:
            if processor.can_process(element):
                return processor
        
        # Should never reach here (FallbackProcessor catches all)
        raise ValueError("No processor for element: {}".format(element.Id))


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def get_selected_elements(uidoc, doc):
    """Get elements from current selection or prompt user to select."""
    picked = []
    
    # Try current selection
    try:
        sel_ids = list(uidoc.Selection.GetElementIds())
    except:
        sel_ids = []
    
    if sel_ids:
        for sid in sel_ids:
            try:
                el = doc.GetElement(sid)
                if el:
                    picked.append(el)
            except:
                pass
    else:
        # Prompt for selection
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select elements to export (ITMs + RFAs, nested included)."
            )
        except:
            forms.alert("Selection cancelled.", exitscript=True)
        
        picked = [doc.GetElement(r) for r in refs if r]
        picked = [e for e in picked if e is not None]
    
    return picked


def main():
    """Main export function."""
    # Setup
    doc = revit.doc
    uidoc = revit.uidoc
    
    # Create context
    context = ExportContext(doc)
    context.reset()
    
    # Get selection
    elements = get_selected_elements(uidoc, doc)
    if not elements:
        forms.alert("Nothing selected.", exitscript=True)
    
    # Execute pipeline
    pipeline = BOMExportPipeline(context)
    
    try:
        output_text = pipeline.execute(elements)
    except Exception as ex:
        forms.alert("Export failed: {}".format(ex), exitscript=True)
    
    # Count rows (excluding header)
    row_count = output_text.count('\n')
    
    if row_count == 0:
        forms.alert("No valid elements found after filtering.", exitscript=True)
    
    # Copy to clipboard
    Clipboard.SetText(output_text)
    
    # Clear selection
    try:
        uidoc.Selection.SetElementIds(List[ElementId]())
        uidoc.RefreshActiveView()
    except:
        pass
    
    # Show completion dialog with download option
    show_completion_dialog(row_count)


def show_completion_dialog(row_count):
    """Show completion dialog with clearer instructions."""
    from rpw.ui.forms import FlexForm, Label, Button
    
    # Create custom dialog with improved messaging
    components = [
        Label("Export Complete - Data Copied to Clipboard!"),
        Label(""),
        Label("Rows exported: {}".format(row_count)),
        Label(""),
        Label("NEXT STEP:"),
        Label("1. Open the NW_BOM Template V1 Excel file"),
        Label("2. Go to the 'Data' tab"),
        Label("3. Paste (Ctrl+V)"),
        Label(""),
        Label("Don't have the template? Click below to download it:"),
        Label("(Only needed once - saves to your Downloads folder)"),
        Label(""),
        Button("Download Blank Template to Downloads", on_click=download_template),
        Button("Close")
    ]
    
    form = FlexForm("BOM Export Complete", components)
    form.show()


def download_template(sender, args):
    """Download template Excel file to user's Downloads folder."""
    try:
        # Get script directory (where pushbutton is located)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Template filename
        template_name = "NW_BOM_Template V1.xlsx"
        template_path = os.path.join(script_dir, template_name)
        
        # Check if template exists
        if not os.path.exists(template_path):
            forms.alert(
                "Template file not found!\n\n"
                "Expected location:\n{}\n\n"
                "Please ensure template is in the pushbutton folder.".format(template_path),
                title="Template Not Found"
            )
            return
        
        # Get user's Downloads folder
        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        destination = os.path.join(downloads_folder, template_name)
        
        # Copy template to Downloads
        shutil.copy(template_path, destination)
        
        # File downloaded successfully (no popup)
    
    except Exception as ex:
        forms.alert(
            "Failed to download template:\n\n{}".format(str(ex)),
            title="Download Failed"
        )


# Run main
if __name__ == '__main__':
    main()