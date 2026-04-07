# -*- coding: utf-8 -*-
"""
NW BOM Export (pyRevit / IronPython)
=================================

Purpose
-------
Exports selected Revit elements to a tab-delimited table (clipboard) suitable for
Excel PivotTables + downstream BOM/cut-sheet workflows.

This script supports three data sources in a single unified export:
  1) Fabrication Parts (ITMs)
  2) CP-driven RFAs (your custom families with CP_* and CPL_* parameters)
  3) Native Revit MEP elements (pipes/ducts + their fittings)

Output Columns (fixed order)
----------------------------
Sorting, Category, Family, Type, ItemNumber, Description, Size, Length_ft, Length_ft_in,
# ---------------------------------------------------------------------
# Hanger & rod special rules: extra rod rows, strut overrides, hanger numbering
# (These rules drive downstream BOM behavior; keep aligned with shop standards.)
# ---------------------------------------------------------------------

Count, Item count, Model, Area, Material, Comments, Spool, Service, Hanger Number, Install Type

Key Rules (high level)
----------------------
* Sorting: numeric bucket so Excel pivots sort correctly (e.g., straights before fittings).
* Length:
    - Length_ft is decimal feet for totals
    - Length_ft_in is formatted feet/inches for cut sheets
    - If element has length, Count column is left blank (intentional)
    - If element has no length, Count is set to 1 (intentional)
* Item count: always 1 per exported row (for counting instances).
* Description: standardized naming for human readability; DO NOT place Size into Description.
* Native Revit Service: System Classification + " - " + System Name (pinned sources).
* Native Revit Install Type: uses "Connection Type" (pinned source).
* Native Revit Material: uses parameter named "Material" (instance first, then type).

Performance
-----------
Production logic uses pinned, exact sources for speed. Diagnostic probe scripts are separate.

Notes
-----
- This file is a comment-cleanup pass only. Logic/behavior is unchanged from the prior stabilized version.
"""

# Clipboard Export (Read-Only, ITMs + RFAs, nested included)
# Model-only filter; excludes GTP families.
# Adds a second "All-Thread Rod" row for each ITM hanger using GetPartAncillaryUsage() SupportRod/Hanger.
# Adds "Item count" column (always 1).
# Excel-safe sizing (prevents date auto-format).
#
# PERF TUNED (same outputs):
# - Caches type elements + type-derived info (family/type name, CPL category, type mark fallback)
# - Caches material ElementId -> name
# - Avoids double GetPartAncillaryUsage() call per hanger
# - Skips full-model SuperComponent scan if no seed RFAs
# - Uses a few local bindings in hot loops

# =====================================================================
# Imports / Revit API bindings
# =====================================================================

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    FabricationPart, FamilyInstance, FilteredElementCollector,
    BuiltInCategory, BuiltInParameter, ElementId, CategoryType, Element
)
from Autodesk.Revit.UI.Selection import ObjectType

import clr
import re

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Clipboard

from System.Collections.Generic import List



# ---------------- Diagnostics ----------------
DEBUG = False

# =====================================================================
# Helper functions (parameter access, formatting, categorization)
# =====================================================================

def log_exception(ex, context=""):
    """Lightweight debug logger. Set DEBUG=True to print exceptions to pyRevit output."""
    if not DEBUG:
        return
    try:
        print("[DEBUG] {}: {}".format(context, ex))
    except Exception:
        pass

# ---------------- CONFIG ----------------
# =====================================================================
# Constants: custom parameter names, built-in parameter IDs, formatting
# =====================================================================

CP_ITEMNO_PARAM     = "CP_Item Number"
CP_DESC_PARAM       = "CP_BOM_Description"
CP_SIZE_PARAM       = "CP_Size"
CP_LENGTH_PARAM     = "CP_Length"
CP_TYPE_NAME_PARAM  = "CP_Type Name"

CPL_FAMILY_CATEGORY_PARAM = "CPL_Family Category"
CP_AREA_PARAM             = "CP_Construction Area"
CP_RFA_SERVICE_PARAM      = "CP_Service"
CP_HANGER_NUMBER_PARAM    = "CP_Hanger Number"
CP_RFA_MODELNO_PARAM      = "CP_Model Number"

CP_RFA_MATERIAL_PARAM     = "CP_Material & Coating"

FISHMOUTH_CID = 2875
FRAC_DENOM = 16
ZERO_TOL_FT = 1e-9

PIPE_STOCK_THRESHOLD_FT = 21.0
PIPE_PREFIX_SHORT = "21FT-"
PIPE_PREFIX_LONG  = "40FT-"
FAB_PIPEWORK_SUBSTR = "MEP Fabrication Pipework"
FAB_DUCTWORK_SUBSTR = "MEP Fabrication Ductwork"

# Native Revit MEP category IDs (stable across Revit versions)
# Pinned to integer Category.Id values to avoid BuiltInCategory name differences across versions.
# Pipes, Pipe Fittings, Ducts, Duct Fittings, Flex Pipes, Flex Ducts
NATIVE6_CAT_IDS = set([
    -2008044,  # Pipes (PipeCurves)
    -2008049,  # Pipe Fittings
    -2008000,  # Ducts (DuctCurves)
    -2008010,  # Duct Fittings
    -2008050,  # Flex Pipes
    -2008020,  # Flex Ducts
])

EXCLUDE_GTP_CATEGORY = "Generic Models"
EXCLUDE_GTP_FAMILY   = "GTP"


# ---------------- Sorting preferences ----------------
SORT_PREF_GLOBAL = {
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

STRUCTURAL_KEYWORDS = ("structural", "unistrut", "strut", "steel", "channel", "trapeze")

EXCLUDED_MODEL_CATEGORY_NAME_CONTAINS = (
    "level",
    "grid",
    "reference plane",
    "scope box",
    "lines",
)


# ---------------- Helpers ----------------
def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except Exception:
        return None

def get_param_str(el, name):
    p = get_param(el, name)
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except Exception:
        try:
            return (p.AsValueString() or "").strip()
        except Exception:
            return ""

def get_param_double(el, name):
    p = get_param(el, name)
    if not p:
        return None
    try:
        return p.AsDouble()
    except Exception:
        return None

def get_bip_param(el, bip):
    try:
        return el.get_Parameter(bip)
    except Exception:
        return None

def get_bip_str(el, bip):
    p = get_bip_param(el, bip)
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except Exception:
        try:
            return (p.AsValueString() or "").strip()
        except Exception:
            return ""

def get_bip_str_by_id(el, bip_int_id):
    try:
        p = el.get_Parameter(ElementId(int(bip_int_id)))
    except Exception:
        p = None
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except Exception:
        try:
            return (p.AsValueString() or "").strip()
        except Exception:
            return ""

def is_revit_duct(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves)
    except Exception:
        return False

def is_revit_mep_curve(el):
    """Covers native Revit pipe/duct/flex curves."""
    try:
        cid = el.Category.Id.IntegerValue if (el and el.Category and el.Category.Id) else None
    except Exception:
        cid = None
    if cid is None:
        return False
    return cid in (
        int(BuiltInCategory.OST_PipeCurves),
        int(BuiltInCategory.OST_FlexPipeCurves),
        int(BuiltInCategory.OST_DuctCurves),
        int(BuiltInCategory.OST_FlexDuctCurves),
    )

def get_revit_curve_length_ft(el):
    p = get_bip_param(el, BuiltInParameter.CURVE_ELEM_LENGTH)
    try:
        return float(p.AsDouble()) if p else None
    except Exception:
        return None

def get_revit_curve_size_text(el):
    """Pinned native curve size (single source): RBS_CALCULATED_SIZE"""
    return get_native_revit_size_text(el) or ""


def get_revit_category_name(el):
    try:
        return el.Category.Name if el.Category else ""
    except Exception:
        return ""


def _native_cat_id(el):
    try:
        return el.Category.Id.IntegerValue if (el and el.Category and el.Category.Id) else None
    except Exception:
        return None

def get_native_revit_size_text(el):
    """
    Native Revit size for BOTH curves and fittings.
    SINGLE SOURCE (per probe): BuiltInParameter.RBS_CALCULATED_SIZE (value string).
    Returns Excel-safe text formula or "".
    """
    s = get_bip_str(el, BuiltInParameter.RBS_CALCULATED_SIZE).strip()
    if not s:
        return ""
    # Normalize common native formats to your Size column expectations
    # e.g. 12"/12" -> 12" x 12"
    s = s.replace("/", " x ")
    s = format_size_inch_marks(s)
    return excel_text_formula(s)

def get_native_revit_service(el):
    """Native Revit service/system for BOM.
    PINNED SOURCE (per SystemsProbe):
      Service = RBS_SYSTEM_CLASSIFICATION_PARAM + ' - ' + RBS_SYSTEM_NAME_PARAM
    Works for pipes/ducts and their fittings.
    """
    cls = get_bip_str(el, BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).strip()
    nm  = get_bip_str(el, BuiltInParameter.RBS_SYSTEM_NAME_PARAM).strip()
    if cls and nm:
        return cls + " - " + nm
    return cls or nm or ""

def get_native_revit_material(doc, el):
    """Native Revit material.
    PINNED SOURCE (per SystemsProbe): parameter named 'Material'.
      - Pipes often have instance Material populated
      - Fittings may have type Material populated
    Exact name only; no ambiguous lookup chains.
    """
    # 1) Instance 'Material'
    v = get_param_str(el, "Material").strip()
    if v:
        return v
    # 2) Type 'Material'
    try:
        info = _get_type_info(doc, el)
        te = info.get("type_el")
    except Exception:
        te = None
    return get_param_str(te, "Material").strip() if te else ""


def get_native_revit_install_type(el):
    """Native Revit Install Type equivalency (pinned).
    Exact source: instance parameter named 'Connection Type'.
    """
    return get_param_str(el, "Connection Type").strip()


def is_native_familyinstance(fi):
    """Gate for treating a FamilyInstance as a *native* Revit fitting (Pipe/Duct Fittings only),
    rather than routing it into the CP-driven RFA rule path.

    Rules:
      - Only applies inside the 6 native Revit MEP categories:
        Pipes, Pipe Fittings, Ducts, Duct Fittings, Flex Pipes, Flex Ducts.
      - For FamilyInstances in those categories (typically fittings),
        default to NATIVE_RULES unless the instance/type has CP/CPL fields with a VALUE.
        (Blank CP fields do *not* force RFA routing.)
    """
    try:
        if not isinstance(fi, FamilyInstance):
            return False

        # Must be one of the six native MEP categories (by BuiltInCategory int id)
        cat = getattr(fi, "Category", None)
        if not cat or not hasattr(cat, "Id"):
            return False
        if cat.Id.IntegerValue not in NATIVE6_CAT_IDS:
            return False

        # Any CP/CPL values present => intentional RFA workflow, do NOT treat as native
        inst_hits = [
            get_param_str(fi, CPL_FAMILY_CATEGORY_PARAM).strip(),
            get_param_str(fi, CP_SIZE_PARAM).strip(),
            get_param_str(fi, CP_TYPE_NAME_PARAM).strip(),
            get_param_str(fi, CP_RFA_SERVICE_PARAM).strip(),
        ]
        if any(inst_hits):
            return False

        # Check type/symbol too
        try:
            sym = fi.Symbol
        except Exception:
            sym = None
        if sym:
            type_hits = [
                get_param_str(sym, CPL_FAMILY_CATEGORY_PARAM).strip(),
                get_param_str(sym, CP_SIZE_PARAM).strip(),
                get_param_str(sym, CP_TYPE_NAME_PARAM).strip(),
                get_param_str(sym, CP_RFA_SERVICE_PARAM).strip(),
            ]
            if any(type_hits):
                return False

        return True
    except Exception:
        return False



def fmt_len_decimal_ft(val):
    if val is None:
        return ""
    try:
        return "{:.3f}".format(float(val))
    except Exception:
        return ""

def _gcd(a, b):
    while b:
        a, b = b, a % b
    return a

def fmt_len_ft_in_frac(val_ft, denom=16):
    if val_ft is None:
        return ""
    try:
        ft_in = float(val_ft)
    except Exception:
        return ""
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
        g = _gcd(num, denom)
        frac_str = " {}/{}".format(num // g, denom // g)

    return '{}{}\'-{}{}"'.format(sign, feet, whole, frac_str)

def fmt_inches_frac(inches, denom=16):
    if inches is None:
        return ""
    try:
        x = float(inches)
    except Exception:
        return ""
    sign = "-" if x < 0 else ""
    x = abs(x)

    x = round(x * denom) / float(denom)
    whole = int(x)
    frac = x - whole

    if frac < 1e-12:
        return '{}{}"'.format(sign, whole)

    num = int(round(frac * denom))
    g = _gcd(num, denom)
    num //= g
    den = denom // g

    if whole == 0:
        return '{}{}/{}"'.format(sign, num, den)
    return '{}{} {}/{}"'.format(sign, whole, num, den)

def count_flag(length_ft):
    if length_ft is None:
        return "1"
    try:
        lf = float(length_ft)
    except Exception:
        return "1"
    if abs(lf) <= ZERO_TOL_FT:
        return "1"
    return "" if lf > 0.0 else "1"

def sanitize(val):
    return "" if val is None else str(val).replace("\t", " ").replace("\n", " ").replace("\r", " ")

def excel_text_formula(s):
    s = (s or "").strip()
    if not s:
        return ""
    s = s.replace('"', '""')
    return '="{}"'.format(s)


def normalize_for_pivot(val, placeholder="N/A"):
    """Avoid PivotTable '(blank)' buckets by replacing empty values."""
    try:
        s = (val or "").strip()
    except Exception:
        s = ""
    return s if s else placeholder


# ---------------- Model-only selection filter ----------------
def is_modeling_element(el):
    if el is None:
        return False
    try:
        cat = el.Category
    except Exception:
        cat = None
    if cat is None:
        return False
    try:
        if cat.CategoryType != CategoryType.Model:
            return False
    except Exception:
        return False
    try:
        if el.ViewSpecific:
            return False
    except Exception:
        pass
    try:
        if cat.Id.IntegerValue == int(BuiltInCategory.OST_Assemblies):
            return False
    except Exception:
        pass
    try:
        nm = (cat.Name or "").strip().lower()
        for token in EXCLUDED_MODEL_CATEGORY_NAME_CONTAINS:
            if token in nm:
                return False
    except Exception:
        pass
    return True


# ---------------- Size formatting ----------------
def _ensure_inches(tok):
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

def format_size_inch_marks(size_str):
    s = (size_str or "").strip()
    if not s:
        return s
    if '"' in s:
        return s
    s_norm = s.replace("×", "x")
    s_norm = re.sub(r"\s*x\s*", " x ", s_norm)
    if " x " in s_norm:
        parts = [p.strip() for p in s_norm.split(" x ") if p.strip()]
        return " x ".join(_ensure_inches(p) for p in parts)
    return _ensure_inches(s_norm)


# ---------------- Type + material caches ----------------
_uidoc = revit.uidoc
_doc = revit.doc

_type_el_cache = {}       # typeIdInt -> type element
_type_info_cache = {}     # typeIdInt -> dict of derived info
_material_name_cache = {} # materialIdInt -> name


def reset_state():
    """Reset module-level caches/state for stability across repeated runs in the same session."""
    try:
        _type_el_cache.clear()
        _type_info_cache.clear()
        _material_name_cache.clear()
    except Exception as ex:
        log_exception(ex, "reset_state: caches")
    try:
        _USED_CODES.clear()
        _CAT_CODE_CACHE.clear()
    except Exception as ex:
        # these may not be defined yet in some versions
        log_exception(ex, "reset_state: code caches")


def _get_type_el(doc, type_id):
    if not type_id:
        return None
    try:
        tid = type_id.IntegerValue
    except Exception:
        return None
    te = _type_el_cache.get(tid)
    if te is not None:
        return te
    try:
        te = doc.GetElement(type_id)
    except Exception:
        te = None
    _type_el_cache[tid] = te
    return te

def _get_type_name(te):
    if te is None:
        return ""
    try:
        s = (Element.Name.GetValue(te) or "").strip()
        if s:
            return s
    except Exception:
        pass
    try:
        return (te.Name or "").strip()
    except Exception:
        return ""

def _get_type_family_name(te):
    if te is None:
        return ""
    try:
        s = (getattr(te, "FamilyName", None) or "").strip()
        if s:
            return s
    except Exception:
        pass
    try:
        return get_bip_str(te, BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    except Exception:
        return ""

def _get_type_info(doc, el):
    """Returns cached dict for el.GetTypeId()"""
    try:
        tid = el.GetTypeId()
    except Exception:
        return {}
    try:
        tid_int = tid.IntegerValue
    except Exception:
        return {}
    info = _type_info_cache.get(tid_int)
    if info is not None:
        return info

    te = _get_type_el(doc, tid)

    info = {
        "type_el": te,
        "type_name": _get_type_name(te),
        "family_name": _get_type_family_name(te),
        "cpl_family_category": get_param_str(te, CPL_FAMILY_CATEGORY_PARAM).strip() if te else "",
        "type_comments": get_bip_str(te, BuiltInParameter.ALL_MODEL_TYPE_COMMENTS) if te else "",
    }
    _type_info_cache[tid_int] = info
    return info


# ---------------- Nested RFA expansion ----------------
MAX_NESTED_RFA_DEPTH = 3  # safety cap; typical nesting is 1-2 levels
def collect_nested_subcomponents(doc, host_fi, out_ids, depth=0):
    """Selection-bound nested expansion.
    Uses GetSubComponentIds() recursively to collect nested FamilyInstances.
    No whole-model scans; depth-capped for safety.
    """
    if depth >= MAX_NESTED_RFA_DEPTH:
        return
    try:
        sub_ids = list(host_fi.GetSubComponentIds())
    except Exception:
        sub_ids = []
    for sid in sub_ids:
        try:
            sub_el = doc.GetElement(sid)
        except Exception:
            sub_el = None
        if isinstance(sub_el, FamilyInstance):
            eid = sub_el.Id.IntegerValue
            if eid not in out_ids:
                out_ids.add(eid)
                collect_nested_subcomponents(doc, sub_el, out_ids, depth + 1)


def expand_selection_with_nested_rfas(doc, picked):
    non_rfa = []
    seed_rfa_ids = set()

    for el in picked:
        if isinstance(el, FamilyInstance):
            seed_rfa_ids.add(el.Id.IntegerValue)
        else:
            non_rfa.append(el)

    all_rfa_ids = set(seed_rfa_ids)

    # Subcomponent recursion (fast, local graph)
    for rid in list(seed_rfa_ids):
        host = doc.GetElement(ElementId(rid))
        if isinstance(host, FamilyInstance):
            collect_nested_subcomponents(doc, host, all_rfa_ids)

    rfas = []
    for rid in all_rfa_ids:
        fi = doc.GetElement(ElementId(rid))
        if isinstance(fi, FamilyInstance):
            rfas.append(fi)

    return non_rfa, rfas


# ---------------- ITM Pipe prefix logic ----------------
def get_itm_product_data_range(itm):
    try:
        v = getattr(itm, "ProductDataRange", None)
        if v is not None:
            s = str(v).strip()
            if s:
                return s
    except Exception:
        pass
    for nm in ("ProductDataRange", "Product Data Range"):
        s = get_param_str(itm, nm).strip()
        if s:
            return s
    return ""

def is_pipework_pipe_itm(itm):
    try:
        cat_ok = FAB_PIPEWORK_SUBSTR in (itm.Category.Name or "")
    except Exception:
        cat_ok = False
    pdr_raw = (get_itm_product_data_range(itm) or "").strip().lower()
    pdr_ok = (pdr_raw == "pipe") or pdr_raw.endswith(".pipe") or pdr_raw.endswith(":pipe")
    return cat_ok and pdr_ok

def is_ductwork_itm(itm):
    """True if FabricationPart belongs to MEP Fabrication Ductwork (any data range)."""
    try:
        return FAB_DUCTWORK_SUBSTR in (itm.Category.Name or "")
    except Exception:
        return False

def is_ductwork_duct_itm(itm):
    """True if FabricationPart is duct-like (duct/flex duct) in MEP Fabrication Ductwork."""
    if not is_ductwork_itm(itm):
        return False
    pdr_raw = (get_itm_product_data_range(itm) or "").strip().lower()
    if not pdr_raw:
        # Some ITMs don't expose ProductDataRange cleanly; fall back to description/name.
        nm = ((getattr(itm, "ProductLongDescription", None) or "") + " " + (getattr(itm, "ProductName", None) or "") + " " + (getattr(itm, "Name", None) or "")).lower()
        return ("duct" in nm)
    return (pdr_raw == "duct") or pdr_raw.endswith(".duct") or pdr_raw.endswith(":duct") or ("duct" in pdr_raw)

def get_itm_product_short_description(itm):
    try:
        v = getattr(itm, "ProductShortDescription", None)
        if v:
            return str(v).strip()
    except Exception:
        pass
    for nm in ("Product Short Description", "ProductShortDescription"):
        s = get_param_str(itm, nm).strip()
        if s:
            return s
    return ""

def stock_prefix_from_short_desc_or_length(itm, length_ft):
    psd = (get_itm_product_short_description(itm) or "").lower()
    nums = re.findall(r"\d+", psd)
    nums = [int(n) for n in nums] if nums else []
    if 40 in nums:
        return PIPE_PREFIX_LONG
    if 21 in nums or 20 in nums:
        return PIPE_PREFIX_SHORT
    try:
        lf = float(length_ft) if length_ft is not None else 0.0
    except Exception:
        lf = 0.0
    return PIPE_PREFIX_SHORT if lf <= PIPE_STOCK_THRESHOLD_FT else PIPE_PREFIX_LONG

def apply_pipe_prefix_to_description(itm, base_desc, length_ft):
    if not base_desc:
        base_desc = ""
    if is_pipework_pipe_itm(itm):
        pref = stock_prefix_from_short_desc_or_length(itm, length_ft)
        if base_desc.startswith(PIPE_PREFIX_SHORT) or base_desc.startswith(PIPE_PREFIX_LONG):
            return base_desc
        return pref + base_desc
    return base_desc


# ---------------- Extraction ----------------
def get_item_number(el):
    """Pinned ItemNumber logic:
    1) CP_Item Number (Southland shared workflow) for everything.
    2) If blank AND element is a FabricationPart (ITM), fall back to Revit's native fabrication Item Number.
       (This supports users exporting ITMs where only the built-in Item Number was filled.)
    No other fallbacks.
    """
    # 1) Authoritative workflow field
    v = get_param_str(el, CP_ITEMNO_PARAM).strip()
    if v:
        return v

    # 2) ITM-only fallback (native Revit fabrication item number)
    is_itm = False
    try:
        is_itm = (el is not None and el.GetType().FullName == "Autodesk.Revit.DB.FabricationPart")
    except:
        # duck-typing fallback
        is_itm = hasattr(el, "GetDimensions") and hasattr(el, "ServiceName")

    if is_itm:
        try:
            bip = getattr(BuiltInParameter, "FABRICATION_PART_ITEM_NUMBER", None)
            if bip is not None:
                p = el.get_Parameter(bip)
                if p:
                    s = (p.AsString() or p.AsValueString() or "").strip()
                    if s:
                        return s
        except:
            pass

        # Last-resort: single pinned display name (still native Revit, not 3rd-party)
        s = get_param_str(el, "Item Number").strip()
        if s:
            return s

    return ""

def get_itm_length_ft(itm):
    """Fabrication centerline length for straights (pipe/duct/flex).

    Notes:
      - Fittings/valves/etc. intentionally return 0.0.
      - Ductwork was previously excluded because we only checked "Pipe...".
    """

    # Always honor fishmouth
    try:
        if itm.ItemCustomId == FISHMOUTH_CID:
            return float(itm.CenterlineLength)
    except Exception:
        pass

    # Prefer robust straight-curve detection
    try:
        is_straight = bool(itm.IsAStraight())
    except Exception:
        is_straight = False

    if not is_straight:
        # Non-straights generally shouldn't get a length
        return 0.0

    # Pipework + Ductwork straights (including flex)
    try:
        return float(itm.CenterlineLength)
    except Exception:
        pass

    # Fallback (older builds): sometimes CenterlineLength isn't exposed; try a param
    for nm in ("Centerline Length", "CenterlineLength", "Length"):
        v = get_param_double(itm, nm)
        if v is not None:
            try:
                return float(v)
            except Exception:
                pass

    return 0.0

def get_rfa_length_ft(fi):
    v = get_param_double(fi, CP_LENGTH_PARAM)
    return float(v) if v is not None else None

def get_rfa_description(doc, fi):
    # Pinned description rules:
    # 1) CP_BOM_Description (instance)
    # 2) CP_Type Name (instance)
    # 3) Deterministic fallback: "{Family} - {Type}" (avoids "Standard"/"Unknown")
    v = get_param_str(fi, CP_DESC_PARAM).strip()
    if v:
        return v
    v = get_param_str(fi, CP_TYPE_NAME_PARAM).strip()
    if v:
        return v
    fam = ""
    try:
        fam = (fi.Symbol.Family.Name or "").strip()
    except Exception:
        fam = ""
    info = _get_type_info(doc, fi)
    typ = (info.get("type_name") or "").strip()
    if not typ:
        try:
            typ = (fi.Symbol.Name or "").strip()
        except Exception:
            typ = ""
    if fam and typ:
        return "{} - {}".format(fam, typ)
    return fam or typ or "N/A"


def get_category_for_export(doc, el):
    """Pinned Category rules:
    - ITM / FabricationPart: Revit Category name
    - RFA / FamilyInstance: CPL_Family Category (INSTANCE first, then TYPE). If missing/blank -> Revit Category name

    Important: CPL_Family Category may be bound as a TYPE parameter in some families.
    """
    # Revit Category Name (fallback and ITM default)
    rev_cat = ""
    try:
        if el is not None and el.Category is not None:
            rev_cat = el.Category.Name or ""
    except Exception:
        rev_cat = ""

    # RFA rule
    try:
        if isinstance(el, FamilyInstance):
            # 1) INSTANCE param
            v = get_param_str(el, "CPL_Family Category")
            if v:
                return v

            # 2) TYPE param
            sym = None
            try:
                sym = el.Symbol
            except Exception:
                sym = None
            if sym:
                v = get_param_str(sym, "CPL_Family Category")
                if v:
                    return v

            # 3) fallback
            return rev_cat
    except Exception:
        return rev_cat

    # ITM + everything else
    return rev_cat


def get_area(el):
    return get_param_str(el, CP_AREA_PARAM)

def get_spool_assembly_name(doc, el):
    try:
        aid = el.AssemblyInstanceId
        if aid and aid.IntegerValue != -1:
            asm = doc.GetElement(aid)
            if asm:
                return (asm.Name or "").strip()
    except Exception:
        pass
    return ""

def get_comments(doc, el):
    s = get_bip_str(el, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
    if s:
        return s
    info = _get_type_info(doc, el)
    s2 = (info.get("type_comments") or "").strip()
    if s2:
        return s2
    return get_param_str(el, "Comments")

def get_family_type(doc, el):
    """
    Returns (RevitFamilyName, RevitTypeName).
    Cached via _get_type_info for type-derived values.
    """
    if isinstance(el, FamilyInstance):
        fam = ""
        try:
            fam = (el.Symbol.Family.Name or "").strip()
        except Exception:
            fam = ""
        info = _get_type_info(doc, el)
        typ = (info.get("type_name") or "").strip()
        if not typ:
            try:
                typ = (el.Symbol.Name or "").strip()
            except Exception:
                typ = ""
        return fam, typ

    if isinstance(el, FabricationPart):
        info = _get_type_info(doc, el)
        typ = (info.get("type_name") or "").strip()
        fam = (info.get("family_name") or "").strip()
        # Fill gaps for Fabrication Ductwork/Pipework where type/...
        if not typ:
            for attr in ("ItemName", "ProductName", "Name"):
                try:
                    v = getattr(el, attr, None)
                    if v:
                        typ = str(v).strip()
                        break
                except Exception:
                    pass
        if not fam:
            # Prefer cleaner family buckets for ductwork
            if is_ductwork_duct_itm(el):
                nm = ((getattr(el, "ProductLongDescription", None) or "") + " " + (getattr(el, "ProductName", None) or "") + " " + (getattr(el, "Name", None) or "")).lower()
                fam = "Flex Duct" if "flex" in nm else "Duct"
            elif is_pipework_pipe_itm(el):
                fam = "Pipe"
            else:
                fam = (get_itm_product_data_range(el) or "").strip()
        return fam, typ

    # Other typed elements (e.g., native Revit pipes/ducts/...) can still resolve a type name
    info = _get_type_info(doc, el)
    typ = (info.get("type_name") or "").strip()
    fam = (info.get("family_name") or "").strip()
    return fam, typ

def _material_name_from_param(doc, p):
    if not p:
        return ""
    try:
        mid = p.AsElementId()
        if mid and mid.IntegerValue > 0:
            mid_int = mid.IntegerValue
            cached = _material_name_cache.get(mid_int)
            if cached is not None:
                return cached
            m = doc.GetElement(mid)
            if m:
                nm = (m.Name or "").strip()
                _material_name_cache[mid_int] = nm
                return nm
    except Exception:
        pass
    try:
        return (p.AsString() or "").strip()
    except Exception:
        try:
            return (p.AsValueString() or "").strip()
        except Exception:
            return ""

def get_itm_material(doc, itm):
    """
    ITM Material (authoritative):
    Use Fabrication Product Material Description (grade-level, e.g. 304L),
    not the fabrication category material.
    """

    # 1) Primary: Fabrication database product material (correct source)
    try:
        v = getattr(itm, "ProductMaterialDescription", None)
        if v:
            return str(v).strip()
    except Exception:
        pass

    # 2) Fallback: same data exposed as a parameter
    for nm in (
        "Product Material Description",
        "ProductMaterialDescription",
    ):
        s = get_param_str(itm, nm)
        if s:
            return s.strip()

    # 3) LAST resort only: fabrication category material
    p = get_bip_param(itm, BuiltInParameter.FABRICATION_PART_MATERIAL)
    s = _material_name_from_param(doc, p)
    if s:
        return s

    # Optional: if you ever need a secondary built-in fallback
    # p = get_bip_param(itm, BuiltInParameter.MATERIAL_ID_PARAM)
    # s = _material_name_from_param(doc, p)
    # if s:
    #     return s

    return ""
def get_rfa_material(fi):
    """Pinned RFA material source: exact shared parameter name only."""
    return get_param_str(fi, CP_RFA_MATERIAL_PARAM).strip()


def get_itm_service(itm):
    try:
        v = getattr(itm, "ServiceName", None)
        if v:
            return str(v).strip()
    except Exception:
        pass
    return ""

def get_install_type_itm(itm):
    for nm in ("Install Type", "ITM Install Type", "Fab Install Type"):
        s = get_param_str(itm, nm)
        if s:
            return s
    return ""

def get_model_field_for_itm(itm):
    try:
        v = getattr(itm, "ProductName", None)
        if v:
            return str(v).strip()
    except Exception:
        pass
    for nm in ("Product Name", "ProductName", "Product"):
        s = get_param_str(itm, nm)
        if s:
            return s
    return "N/A"

def get_model_field_for_rfa(doc, fi):
    """
    RFA Model (pinned, Southland workflow):
      1) CP_Model Number (INSTANCE)
      2) CP_Model Number (TYPE)
      3) Revit built-in Identity Data 'Model' (TYPE param named 'Model')
      4) N/A

    Notes:
      - Treat empty, '-', and 'N/A' (any case) as blank.
      - Do NOT fall back to Type Name or any 3rd-party parameters.
      - In your content, 'Model' is often a TYPE parameter (fi.Symbol), not an instance parameter.
    """

    def _is_blankish(s):
        s = (s or "").strip()
        return (not s) or (s.lower() in ("n/a", "na", "-", "none"))

    # 1) CP_Model Number on INSTANCE
    s = get_param_str(fi, CP_RFA_MODELNO_PARAM)
    if not _is_blankish(s):
        return s.strip()

    # 2) CP_Model Number on TYPE
    sym = None
    try:
        sym = fi.Symbol
    except:
        sym = None

    if sym:
        s2 = get_param_str(sym, CP_RFA_MODELNO_PARAM)
        if not _is_blankish(s2):
            return s2.strip()

        # 3) Revit built-in 'Model' on TYPE (Identity Data)
        s3 = get_param_str(sym, "Model")
        if not _is_blankish(s3):
            return s3.strip()

    return "N/A"


def should_exclude_by_family(doc, el):
    """Exclude exactly one known helper family.
    Pinned rule:
      - Category == 'Generic Models'
      - Family   == 'GTP'
    No substring matching.
    """
    try:
        cat = el.Category.Name if el.Category else ""
    except Exception:
        cat = ""
    if cat != EXCLUDE_GTP_CATEGORY:
        return False
    info = _get_type_info(doc, el)
    fam = (info.get("family_name") or "").strip()
    if not fam and info.get("type_el") is not None:
        try:
            fam = get_bip_str(info["type_el"], BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        except Exception:
            fam = ""
    return (fam or "").strip().lower() == EXCLUDE_GTP_FAMILY.lower()



# ---------------- TRUE rod extraction ----------------
def safe_str(x):
    try:
        return "" if x is None else str(x)
    except Exception:
        return ""

def try_get_support_rod_usage(itm_part):
    try:
        anc_list = list(itm_part.GetPartAncillaryUsage())
    except Exception:
        return (None, None)
    for anc in anc_list:
        try:
            u = safe_str(getattr(anc, "UsageType", "")).lower()
            t = safe_str(getattr(anc, "Type", "")).lower()
        except Exception:
            u, t = "", ""
        if u == "hanger" and t == "supportrod":
            dia_ft = getattr(anc, "AncillaryWidthOrDiameter", None)
            len_ft = getattr(anc, "Length", None)
            try:
                dia_ft = float(dia_ft) if dia_ft is not None else None
            except Exception:
                dia_ft = None
            try:
                len_ft = float(len_ft) if len_ft is not None else None
            except Exception:
                len_ft = None
            return (dia_ft, len_ft)

    # Native fallback (still ITM database): parse rod diameter from fabrication product text.
    # This avoids any 3rd-party shared parameters (BIMrx/Naviate/etc.) while remaining
    # deterministic and fast (no wide-net parameter scanning).
    #
    # Why parsing?
    # - Some hanger ITMs do not expose a supportrod ancillary usage object.
    # - The Fabrication dimension name "Diameter" is frequently the *pipe size* (as you saw),
    #   not the rod diameter, so we do NOT use it.
    #
    # We instead parse the fabrication item's own descriptions/names, which are ITM-backed.
    dia_ft = try_parse_itm_rod_diameter_ft(itm_part)
    if dia_ft is not None and abs(dia_ft) >= ZERO_TOL_FT:
        return (dia_ft, None)

    return (None, None)


def try_parse_itm_rod_diameter_ft(itm_part):
    """Return rod diameter in feet parsed from ITM-backed text fields.

    Inputs (concatenated in priority order):
      - ProductLongDescription
      - ProductName
      - Name

    Strategy:
      - Prefer explicit fractions like "3/8".
      - Otherwise, accept decimals like "0.375".
      - Choose the last "rod-like" token in the text (most hanger naming conventions place
        rod size near the end, e.g. "... 0.375").

    Guardrails to avoid pipe sizes:
      - Only accept inch values in a sane rod range: 0.10" to 1.00".

    Returns None if no suitable value found.
    """
    if not isinstance(itm_part, FabricationPart):
        return None

    try:
        txt = " ".join([
            safe_str(getattr(itm_part, "ProductLongDescription", None)),
            safe_str(getattr(itm_part, "ProductName", None)),
            safe_str(getattr(itm_part, "Name", None)),
        ])
    except Exception:
        txt = ""

    if not txt:
        return None

    s = txt.lower()
    # Normalize a few separators
    s = s.replace("\t", " ").replace("\r", " ").replace("\n", " ")

    candidates_in = []

    # 1) Fractions like 3/8, 1/2, 5/16
    for m in re.finditer(r"\b(\d+)\s*/\s*(\d+)\b", s):
        try:
            num = float(m.group(1))
            den = float(m.group(2))
            if den != 0:
                inch = num / den
                candidates_in.append(inch)
        except Exception:
            pass

    # 2) Decimals like 0.375, 0.3125
    for m in re.finditer(r"\b(0?\.\d{2,6})\b", s):
        try:
            inch = float(m.group(1))
            candidates_in.append(inch)
        except Exception:
            pass

    if not candidates_in:
        return None

    # Filter to plausible rod diameters (inches)
    rod_like = [v for v in candidates_in if 0.10 <= abs(v) <= 1.00]
    if not rod_like:
        return None

    # Choose last plausible candidate (matches typical naming: "... 0.375")
    inch = rod_like[-1]
    try:
        return float(inch) / 12.0
    except Exception:
        return None

def is_itm_hanger_with_cached_usage(itm, dia_ft, len_ft):
    if not isinstance(itm, FabricationPart):
        return False
    if dia_ft is not None or len_ft is not None:
        return True
    try:
        cn = (itm.Category.Name or "").lower()
    except Exception:
        cn = ""
    return ("fabrication hangers" in cn) or ("hangers" in cn)



def parse_len_value_to_ft(valstr):
    """Parse a Revit length value string like 0' 6 5/8" or 2'-2 9/16" into feet (float). Returns None if unparseable."""
    if not valstr:
        return None
    s = str(valstr).strip()
    if not s:
        return None
    # Normalize separators
    s = s.replace("’", "'").replace("′","'").replace("”", '"').replace("″", '"')
    s = s.replace("-", " ").replace("  ", " ")
    feet = 0.0
    inches = 0.0

    # Feet part: 2' or 2 '
    m = re.search(r"(-?\d+)\s*'\s*", s)
    if m:
        try:
            feet = float(m.group(1))
        except Exception:
            feet = 0.0

    # Inch part: may include whole inches and fraction, e.g. 2 9/16" or 6 5/8"
    # Grab substring after feet mark if present
    if "'" in s:
        s2 = s.split("'",1)[1]
    else:
        s2 = s
    # Extract the first quoted inches segment if any
    # Remove trailing quote
    s2 = s2.replace('"'," ").strip()
    # Now s2 might be like: 2 9/16 or 6 5/8 or 6 or 9/16
    if s2:
        parts = s2.split()
        whole = 0.0
        frac = 0.0
        if len(parts) >= 1:
            if "/" in parts[0]:
                # fraction only
                num, den = parts[0].split("/",1)
                try:
                    frac = float(num)/float(den)
                except Exception:
                    frac = 0.0
            else:
                try:
                    whole = float(parts[0])
                except Exception:
                    whole = 0.0
            if len(parts) >= 2 and "/" in parts[1]:
                num, den = parts[1].split("/",1)
                try:
                    frac = float(num)/float(den)
                except Exception:
                    frac = frac
        inches = whole + frac

    # If no explicit inches but string contains a decimal inches value like 6.5"
    if inches == 0.0:
        m2 = re.search(r"(-?\d+(?:\.\d+)?)\s*\"", s)
        if m2:
            try:
                inches = float(m2.group(1))
            except Exception:
                pass

    return feet + (inches / 12.0)

def try_get_rod_lengths_ab_ft(itm_part):
    """Return (lenA_ft, lenB_ft) from the ITM (fabrication database) dimensions.

    IMPORTANT:
      - Do NOT use shared parameters from 3rd-party add-ins (e.g. BIMrx_*).
      - Do NOT scan / fuzzy-match parameter names (performance + ambiguity).
      - Use ONLY the FabricationPart dimension names proven by probe:
          * "Length A"
          * "Length B"

    Values are returned in internal feet. Missing / zero values return None.
    """
    if not isinstance(itm_part, FabricationPart):
        return (None, None)

    lenA_ft = None
    lenB_ft = None

    try:
        dims = list(itm_part.GetDimensions())
    except Exception:
        return (None, None)

    # Build exact lookup table (case-insensitive) for the two pinned dimension names.
    name_to_def = {}
    for ddef in dims:
        try:
            nm = (ddef.Name or "").strip().lower()
        except Exception:
            nm = ""
        if nm:
            name_to_def[nm] = ddef

    def _get_dim_ft(name_lower):
        ddef = name_to_def.get(name_lower)
        if not ddef:
            return None

        v = None
        # Preferred: direct dimension value
        try:
            v = itm_part.GetDimensionValue(ddef)
        except Exception:
            v = None

        # Some ITM dims only expose calculated value in certain builds
        if v is None:
            try:
                v = itm_part.GetCalculatedDimensionValue(ddef)
            except Exception:
                v = None

        if v is None:
            return None

        try:
            v = float(v)
        except Exception:
            return None

        # Treat near-zero as "not present"
        try:
            if abs(v) < ZERO_TOL_FT:
                return None
        except Exception:
            pass
        return v

    lenA_ft = _get_dim_ft("length a")
    lenB_ft = _get_dim_ft("length b")

    return (lenA_ft, lenB_ft)

def try_get_hanger_strut_length_ft(itm_part):
    """Return trapeze strut length (ft) for Fabrication hangers, **without** relying on 3rd-party parameters.

    Based on Fabrication dimensions observed in your dumps:
        Strut Length = Width + 2 * Bearer Extn

    If either dimension is missing, returns None.
    """
    if itm_part is None:
        return None

    def _get_dim_ft(target_name):
        try:
            dims = itm_part.GetDimensions()
        except Exception as ex:
            log_exception(ex, "try_get_hanger_strut_length_ft:GetDimensions")
            return None

        if not dims:
            return None

        # Find exact-name match (stable; avoids accidental 3rd-party params)
        try:
            for d in dims:
                try:
                    if (d.Name or "").strip() == target_name:
                        try:
                            return itm_part.GetDimensionValue(d)
                        except Exception as ex:
                            log_exception(ex, "try_get_hanger_strut_length_ft:GetDimensionValue:%s" % target_name)
                            return None
                except Exception:
                    continue
        except Exception:
            return None

        return None

    width_ft = _get_dim_ft("Width")
    bext_ft  = _get_dim_ft("Bearer Extn")

    if width_ft is None or bext_ft is None:
        return None

    return width_ft + (2.0 * bext_ft)

def build_rod_row_from_hanger_cached(hanger_row, dia_ft, rod_len_ft, itm_hanger):
    if dia_ft is None and rod_len_ft is None:
        return None

    r = list(hanger_row)

    r[IDX_CAT] = "Structural Framing"
    r[IDX_FAM] = "All-Thread Rod"
    r[IDX_DESC] = "All-Thread Rod"

    # Put rod diameter in SIZE (not Type)
    if dia_ft is not None:
        dia_in = dia_ft * 12.0
        r[IDX_SIZE] = excel_text_formula(fmt_inches_frac(dia_in, denom=16))
    else:
        r[IDX_SIZE] = excel_text_formula("N/A")

    # Type for rod doesn't carry meaning in our BOM; keep it clean
    r[IDX_TYP] = "N/A"

    r[IDX_LENFT] = fmt_len_decimal_ft(rod_len_ft)
    r[IDX_LENFR] = fmt_len_ft_in_frac(rod_len_ft, denom=FRAC_DENOM)
    r[IDX_COUNT] = count_flag(rod_len_ft)

    # Item count stays "1"
    r[IDX_ITEMCT] = "1"

    # Model/Install Type not applicable
    r[IDX_MODEL] = "N/A"
    r[IDX_INSTALL] = "N/A"

    r[IDX_SORT] = get_sorting_value(itm_hanger, r[IDX_CAT])
    return r

# ---------------- Pipework sort helpers (Fabrication Pipework) ----------------
PIPEWORK_CATEGORY_TOKEN = "Fabrication Pipework"
PIPEWORK_FAMILY_PARAM = "Family"
PIPEWORK_DIAM_PARAM = "Main Primary Diameter"

def _pipe_family_str(fp):
    # Match SI Tools behavior (Family compared via AsValueString where possible)
    try:
        p = fp.LookupParameter(PIPEWORK_FAMILY_PARAM)
        if p:
            try:
                return (p.AsValueString() or "").strip()
            except Exception:
                try:
                    return (p.AsString() or "").strip()
                except Exception:
                    return ""
    except Exception:
        pass
    return ""

def _is_pipe_weld(fp):
    # Best-effort weld detection (ported from SI Tools renumber logic)
    for attr in ("ItemType", "PartType", "FabricationPartType"):
        v = getattr(fp, attr, None)
        if v is not None:
            try:
                if "weld" in str(v).lower():
                    return True
            except Exception:
                pass
    for attr in ("ItemName", "Name"):
        v = getattr(fp, attr, None)
        if v:
            try:
                if "weld" in str(v).lower():
                    return True
            except Exception:
                pass
    fam = _pipe_family_str(fp)
    return "weld" in (fam or "").lower()

def _is_pipework_fp(fp):
    try:
        return fp.Category and (PIPEWORK_CATEGORY_TOKEN in (fp.Category.Name or ""))
    except Exception:
        return False

def _pipework_sort_code(fp):
    """Return fixed sorting code strings for Fabrication Pipework.
    Mirrors SI Tools ordering buckets:
      1) non-weld straights
      2) non-weld non-straight non-tap
      3) non-weld taps
      4) remaining non-weld
      5) welds
    """
    if _is_pipe_weld(fp):
        return "240"
    try:
        if bool(fp.IsAStraight()):
            return "110"
    except Exception:
        pass
    try:
        if (not bool(fp.IsAStraight())) and (not bool(fp.IsATap())):
            return "210"
    except Exception:
        pass
    try:
        if bool(fp.IsATap()):
            return "220"
    except Exception:
        pass
    return "230"
# ---------------- Sorting ----------------
def stable_hash(s):
    if s is None:
        s = ""
    try:
        s = str(s)
    except Exception:
        s = ""
    h = 2166136261
    for ch in s:
        h = h ^ ord(ch)
        h = (h * 16777619) & 0xFFFFFFFF
    return h

_USED_CODES = set()
_CAT_CODE_CACHE = {}  # PERF: category -> assigned code (keeps first-assigned behavior)

def is_structuralish(category_name):
    s = (category_name or "").lower()
    return any(k in s for k in STRUCTURAL_KEYWORDS)

def series_base_for_element(el, exported_category_name):
    if isinstance(el, FabricationPart):
        try:
            if bool(el.IsAStraight()):
                return 100
            elif bool(el.IsATap()):
                return 300
            else:
                return 200
        except Exception:
            return 200
    if exported_category_name == "Conduit Fittings" or is_structuralish(exported_category_name):
        return 600
    return 900

def assign_code(base, category_name):
    cat = (category_name or "").strip() or "Uncategorized"
    cached = _CAT_CODE_CACHE.get(cat)
    if cached is not None:
        return cached

    if cat in SORT_PREF_GLOBAL:
        code = int(SORT_PREF_GLOBAL[cat])
        _USED_CODES.add(code)
        _CAT_CODE_CACHE[cat] = code
        return code

    slot = (stable_hash(cat) % 9) + 1
    code = base + slot * 10
    tries = 0
    while code in _USED_CODES and tries < 50:
        slot += 1
        if slot > 9:
            slot = 1
        code = base + slot * 10
        tries += 1

    _USED_CODES.add(code)
    _CAT_CODE_CACHE[cat] = code
    return code

def get_sorting_value(el, cat_export):
    # Special-case: Fabrication Pipework needs deterministic sub-ordering (straights, fittings, taps, welds)
    if isinstance(el, FabricationPart) and _is_pipework_fp(el):
        return _pipework_sort_code(el)

    base = series_base_for_element(el, cat_export)
    return str(assign_code(base, cat_export))


# ---------------- Selection ----------------
uidoc = _uidoc
doc = _doc

reset_state()

picked = []
try:
    sel_ids = list(uidoc.Selection.GetElementIds())
except Exception:
    sel_ids = []

if sel_ids:
    for sid in sel_ids:
        try:
            el = doc.GetElement(sid)
            if el:
                picked.append(el)
        except Exception:
            pass
else:
    try:
# =====================================================================
# Main: selection, expansion (nested RFAs), row build, clipboard export
# =====================================================================

        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select scope to export (ITMs + RFAs, nested included).")
    except Exception:
        forms.alert("Selection cancelled.", exitscript=True)
    picked = [doc.GetElement(r) for r in refs if r]
    picked = [e for e in picked if e is not None]

picked = [e for e in picked if is_modeling_element(e)]
picked = [e for e in picked if not should_exclude_by_family(doc, e)]

if not picked:
    forms.alert("Nothing selected (after filtering).", exitscript=True)


# ---------------- Expand nested RFAs ----------------
non_rfa, rfas_expanded = expand_selection_with_nested_rfas(doc, picked)
non_rfa = [e for e in non_rfa if is_modeling_element(e) and (not should_exclude_by_family(doc, e))]
# ---------------------------------------------------------------------
# RFAs (CP-driven families): use CP_* and CPL_* rules; may apply pinned native
# fallbacks ONLY for native-style fittings (strict gate)
# ---------------------------------------------------------------------

rfas_expanded = [fi for fi in rfas_expanded if is_modeling_element(fi) and (not should_exclude_by_family(doc, fi))]


# ---------------- Build rows ----------------
rows = []

header = [
    "Sorting","Category","Family","Type","ItemNumber","Description","Size",
    "Length_ft","Length_ft_in","Count","Item count",
    "Model","Area","Material","Comments","Spool","Service","Hanger Number","Install Type",
]

# Column indices (keep in sync with header)
IDX_SORT    = 0
IDX_CAT     = 1
IDX_FAM     = 2
IDX_TYP     = 3
IDX_ITEMNO  = 4
IDX_DESC    = 5
IDX_SIZE    = 6
IDX_LENFT   = 7
IDX_LENFR   = 8
IDX_COUNT   = 9
IDX_ITEMCT  = 10
IDX_MODEL   = 11
IDX_AREA    = 12
IDX_MATERIAL= 13
IDX_COMMENTS= 14
IDX_SPOOL   = 15
IDX_SERVICE = 16
IDX_HANGER  = 17
IDX_INSTALL = 18

# Hot-loop local bindings (small but free speed)
rows_append = rows.append
sanitize_local = sanitize
fmt_len_dec = fmt_len_decimal_ft
fmt_len_frac = fmt_len_ft_in_frac
count_flag_local = count_flag
excel_formula = excel_text_formula
format_size = format_size_inch_marks

# ITMs / non-RFAs
# ---------------------------------------------------------------------
# Non-RFA elements (native Revit MEP curves + other categories):
# - Native pipes/ducts use pinned sources for service/material/install type
# - Other categories follow existing baseline rules
# ---------------------------------------------------------------------

for el in non_rfa:
    cat_export = get_category_for_export(doc, el)
    sorting_val = get_sorting_value(el, cat_export)

    fam, typ = get_family_type(doc, el)
    item_no = get_item_number(el)

    area_val = get_area(el)
    comments_val = get_comments(doc, el)
    spool_val = get_spool_assembly_name(doc, el)
    hanger_val = get_param_str(el, CP_HANGER_NUMBER_PARAM)

    desc_val = ""
    size_val = ""
    length_ft = None
    material_val = ""
    service_val = ""
    install_type_val = get_native_revit_install_type(el) if is_native_familyinstance(el) else ""
    model_val = "N/A"
    length_for_count = None

    dia_ft = None
    rod_len_ft = None
    lenA_ft = None  # 2-rod hangers: Length A
    lenB_ft = None  # 2-rod hangers: Length B
    strut_len_ft = None  # 2-rod hangers: strut/bearer length
    is_hanger = False

    if isinstance(el, FabricationPart):
        length_ft = get_itm_length_ft(el)
        length_for_count = length_ft

        # Size (used in Size column)
        base_size = (el.ProductSizeDescription or "")
        # Fabrication ductwork often has blank ProductSizeDescription; fall back to OverallSize.
        if (not base_size) and is_ductwork_itm(el):
            base_size = (get_param_str(el, "OverallSize") or get_param_str(el, "Overall Size") or get_param_str(el, "Overall_Size") or "")
        _size_fmt = format_size(base_size)
        size_val = excel_formula(_size_fmt)

        # Description:
        # - Pipework keeps existing behavior (prefix on ProductLongDescription)
        # - Ductwork should mirror pipework style, but Description (col F) = Fabrication family name + ": " + size
        base_desc = (el.ProductLongDescription or "")
        desc_val = apply_pipe_prefix_to_description(el, base_desc, length_ft)

        # Fabrication ductwork: Description (col F) should match fabrication pipe behavior:
        # use the fabrication family name (e.g., "Straight - Square") only.
        if is_ductwork_itm(el):
            fam_name = (fam or "").strip() or (typ or "").strip()
            if fam_name:
                desc_val = fam_name

        material_val = get_itm_material(doc, el)
        service_val = get_itm_service(el)
        install_type_val = get_install_type_itm(el)
        model_val = get_model_field_for_itm(el)

        # PERF: compute rod usage once
        dia_ft, rod_len_ft = try_get_support_rod_usage(el)
        is_hanger = is_itm_hanger_with_cached_usage(el, dia_ft, rod_len_ft)
        # If this is a hanger, pull 2-rod lengths once (used for rod lines + strut length on hanger row)
        if is_hanger:
            lenA_ft, lenB_ft = try_get_rod_lengths_ab_ft(el)
            if (lenA_ft is not None) and (lenB_ft is not None):
                # 2-rod hanger: use strut/bearer length on the hanger row (requested for BOM)
                strut_len_ft = try_get_hanger_strut_length_ft(el)
                if strut_len_ft is not None and abs(strut_len_ft) > ZERO_TOL_FT:
                    length_ft = strut_len_ft
                    length_for_count = strut_len_ft


    elif is_revit_mep_curve(el):
        length_ft = get_revit_curve_length_ft(el)
        length_for_count = length_ft
        size_val = get_revit_curve_size_text(el)
        # Native Revit curves: pinned sources for material/service
        material_val = get_native_revit_material(doc, el)
        service_val = get_native_revit_service(el)
        install_type_val = get_native_revit_install_type(el)
        # Prefer type name when available; otherwise fall back to category
        if (typ or "").strip() and (typ or "").strip() != "N/A":
            desc_val = "{} - {}".format(get_revit_category_name(el) or "MEP", (typ or "").strip())
        else:
            desc_val = get_revit_category_name(el) or "MEP"

    else:
        desc_val = get_revit_category_name(el) or ""
    # --- Normalize blanks to avoid Excel PivotTable '(blank)' ---
    if not (size_val or "").strip():
        size_val = excel_formula("N/A")
    if not (material_val or "").strip():
        material_val = "N/A"
    if not (typ or "").strip():
        typ = "N/A"
    if not (model_val or "").strip():
        model_val = "N/A"


    base_row = [
        sorting_val, cat_export, fam, typ, item_no, desc_val, size_val,
        fmt_len_dec(length_ft), fmt_len_frac(length_ft, denom=FRAC_DENOM),
        count_flag_local(length_for_count),
        "1",  # Item count
        model_val, area_val, material_val, comments_val, spool_val, service_val,
        hanger_val, install_type_val,
    ]
    rows_append(base_row)

    if isinstance(el, FabricationPart) and is_hanger:
        # Some hangers have two independent rod lengths (Length A / Length B). If present, output two rod lines.
        if (lenA_ft is not None) and (lenB_ft is not None):
            rod_row_1 = build_rod_row_from_hanger_cached(base_row, dia_ft, lenA_ft, el)
            if rod_row_1:
                rows_append(rod_row_1)
            rod_row_2 = build_rod_row_from_hanger_cached(base_row, dia_ft, lenB_ft, el)
            if rod_row_2:
                rows_append(rod_row_2)
        else:
            # Default: one rod derived from ancillary usage (or whatever the hanger reports)
            # If the hanger exposes a single rod length via fabrication dimensions, prefer it over ancillary-derived length
            use_len_ft = lenA_ft if (lenA_ft is not None) else rod_len_ft
            rod_row = build_rod_row_from_hanger_cached(base_row, dia_ft, use_len_ft, el)
            if rod_row:
                rows_append(rod_row)

# RFAs
for fi in rfas_expanded:
    cat_export = get_category_for_export(doc, fi)
    sorting_val = get_sorting_value(fi, cat_export)

    fam, typ = get_family_type(doc, fi)
    item_no = get_item_number(fi)
    desc_val = get_rfa_description(doc, fi)

    _cp_size = get_param_str(fi, CP_SIZE_PARAM).strip()
    size_val = excel_formula(format_size(_cp_size)) if _cp_size else (get_native_revit_size_text(fi) or "")
    length_ft = get_rfa_length_ft(fi)

    model_val = get_model_field_for_rfa(doc, fi)
    area_val = get_area(fi)
    material_val = get_rfa_material(fi).strip()
    if not material_val:
        material_val = get_native_revit_material(doc, fi)
    comments_val = get_comments(doc, fi)
    spool_val = get_spool_assembly_name(doc, fi)
    service_val = get_param_str(fi, CP_RFA_SERVICE_PARAM).strip()
    if not service_val:
        service_val = get_native_revit_service(fi)


    # Native Revit fittings (FamilyInstance) — apply pinned native sources, keep Description rules unchanged
    if is_native_familyinstance(fi):
        _nsz = get_native_revit_size_text(fi)
        if _nsz:
            size_val = _nsz
        service_val = get_native_revit_service(fi)
        install_type_val = get_native_revit_install_type(fi)
        material_val = get_native_revit_material(doc, fi) or material_val
    hanger_val = get_param_str(fi, CP_HANGER_NUMBER_PARAM)
    install_type_val = get_native_revit_install_type(fi) if is_native_familyinstance(fi) else ""

    # --- Normalize blanks to avoid Excel PivotTable '(blank)' ---
    if not (size_val or "").strip():
        size_val = excel_formula("N/A")
    if not (material_val or "").strip():
        material_val = "N/A"
    if not (typ or "").strip():
        typ = "N/A"
    if not (model_val or "").strip():
        model_val = "N/A"

    rows_append([
        sorting_val, cat_export, fam, typ, item_no, desc_val, size_val,
        fmt_len_dec(length_ft), fmt_len_frac(length_ft, denom=FRAC_DENOM),
        count_flag_local(length_ft),
        "1",  # Item count
        model_val, area_val, material_val, comments_val, spool_val, service_val,
        hanger_val, install_type_val,
    ])

if not rows:
    forms.alert("No valid elements found after filtering.", exitscript=True)
# Pre-sort rows so Excel/PivotTables (even in Data Source Order) come in correctly ordered
def _sort_int(x):
    try:
        return int(str(x).strip())
    except Exception:
        return 999999

rows.sort(key=lambda r: (
    _sort_int(r[0]),   # Sorting
    (r[1] or ""),      # Category
    (r[5] or ""),      # Description
    (r[2] or ""),      # Family
    (r[3] or ""),      # Type
))



# Build clipboard text (same output)
# Final normalization for pivot stability
# - Keep core fields non-empty where pivot tables tend to break
# - Standardize blanks to "N/A" for a few key columns
IDX_TYPE = 3
IDX_SIZE = 6
IDX_MODEL = 11
IDX_MATERIAL = 13
for _r in rows:
    try:
        _r[IDX_TYPE] = normalize_for_pivot(_r[IDX_TYPE])
        _r[IDX_SIZE] = normalize_for_pivot(_r[IDX_SIZE])
        _r[IDX_MODEL] = normalize_for_pivot(_r[IDX_MODEL])
        _r[IDX_MATERIAL] = normalize_for_pivot(_r[IDX_MATERIAL])
    except Exception as ex:
        log_exception(ex, "finalize row normalization")

lines = ["\t".join(header)]
for row in rows:
    lines.append("\t".join([sanitize_local(c) for c in row]))

Clipboard.SetText("\r\n".join(lines))

# Clear selection so elements don't stay highlighted (blue) after export
try:
    uidoc.Selection.SetElementIds(List[ElementId]())
    uidoc.RefreshActiveView()
except Exception:
    pass

forms.alert(
    "Clipboard Export Complete\n\n"
    "Rows exported: {rows}\n\n"
    "Data copied to Windows clipbord\n"
    "Paste into Data tab of NW_BOM Template V1".format(rows=len(rows))
)
