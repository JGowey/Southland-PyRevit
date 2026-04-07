# -*- coding: utf-8 -*-
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


# ---------------- CONFIG ----------------
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

CP_RFA_MATERIAL_PARAM_1   = "CP_Meterial & Coating"
CP_RFA_MATERIAL_PARAM_2   = "CP_Material & Coating"

FISHMOUTH_CID = 2875
FRAC_DENOM = 16
ZERO_TOL_FT = 1e-9

PIPE_STOCK_THRESHOLD_FT = 21.0
PIPE_PREFIX_SHORT = "21FT-"
PIPE_PREFIX_LONG  = "40FT-"
FAB_PIPEWORK_SUBSTR = "MEP Fabrication Pipework"

BIP_TYPE_MARK_ID_GUESS = -1002002  # safe built-in id guess; plus name fallback

EXCLUDE_FAMILY_SUBSTR = "gtp"


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
    except:
        return None

def get_param_str(el, name):
    p = get_param(el, name)
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except:
        try:
            return (p.AsValueString() or "").strip()
        except:
            return ""

def get_param_double(el, name):
    p = get_param(el, name)
    if not p:
        return None
    try:
        return p.AsDouble()
    except:
        return None

def get_bip_param(el, bip):
    try:
        return el.get_Parameter(bip)
    except:
        return None

def get_bip_str(el, bip):
    p = get_bip_param(el, bip)
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except:
        try:
            return (p.AsValueString() or "").strip()
        except:
            return ""

def get_bip_str_by_id(el, bip_int_id):
    try:
        p = el.get_Parameter(ElementId(int(bip_int_id)))
    except:
        p = None
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except:
        try:
            return (p.AsValueString() or "").strip()
        except:
            return ""

def is_revit_duct(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves)
    except:
        return False

def get_revit_category_name(el):
    try:
        return el.Category.Name if el.Category else ""
    except:
        return ""

def fmt_len_decimal_ft(val):
    if val is None:
        return ""
    return str(float(val))

def _gcd(a, b):
    while b:
        a, b = b, a % b
    return a

def fmt_len_ft_in_frac(val_ft, denom=16):
    if val_ft is None:
        return ""
    try:
        ft_in = float(val_ft)
    except:
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
    except:
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
    except:
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


# ---------------- Model-only selection filter ----------------
def is_modeling_element(el):
    if el is None:
        return False
    try:
        cat = el.Category
    except:
        cat = None
    if cat is None:
        return False
    try:
        if cat.CategoryType != CategoryType.Model:
            return False
    except:
        return False
    try:
        if el.ViewSpecific:
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
        for token in EXCLUDED_MODEL_CATEGORY_NAME_CONTAINS:
            if token in nm:
                return False
    except:
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

def _get_type_el(doc, type_id):
    if not type_id:
        return None
    try:
        tid = type_id.IntegerValue
    except:
        return None
    te = _type_el_cache.get(tid)
    if te is not None:
        return te
    try:
        te = doc.GetElement(type_id)
    except:
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
    except:
        pass
    try:
        return (te.Name or "").strip()
    except:
        return ""

def _get_type_family_name(te):
    if te is None:
        return ""
    try:
        s = (getattr(te, "FamilyName", None) or "").strip()
        if s:
            return s
    except:
        pass
    try:
        return get_bip_str(te, BuiltInParameter.ALL_MODEL_FAMILY_NAME)
    except:
        return ""

def _get_type_info(doc, el):
    """Returns cached dict for el.GetTypeId()"""
    try:
        tid = el.GetTypeId()
    except:
        return {}
    try:
        tid_int = tid.IntegerValue
    except:
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
        "type_mark_guess": get_bip_str_by_id(te, BIP_TYPE_MARK_ID_GUESS) if te else "",
        "type_comments": get_bip_str(te, BuiltInParameter.ALL_MODEL_TYPE_COMMENTS) if te else "",
    }
    _type_info_cache[tid_int] = info
    return info


# ---------------- Nested RFA expansion ----------------
def collect_nested_subcomponents(doc, host_fi, out_ids):
    try:
        sub_ids = list(host_fi.GetSubComponentIds())
    except:
        sub_ids = []
    for sid in sub_ids:
        try:
            sub_el = doc.GetElement(sid)
        except:
            sub_el = None
        if isinstance(sub_el, FamilyInstance):
            eid = sub_el.Id.IntegerValue
            if eid not in out_ids:
                out_ids.add(eid)
                collect_nested_subcomponents(doc, sub_el, out_ids)

def collect_children_by_supercomponent(doc, host_ids_int, out_ids):
    # NOTE: This is an unavoidable full-model scan if you want SuperComponent children reliably.
    # PERF: only called if there are seed RFAs (guarded in expand_selection_with_nested_rfas).
    host_ids = set(host_ids_int)
    for fi in FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements():
        try:
            sc = fi.SuperComponent
        except:
            sc = None
        if sc and sc.Id and sc.Id.IntegerValue in host_ids:
            out_ids.add(fi.Id.IntegerValue)

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

    # SuperComponent children scan (global scan, only if needed)
    if seed_rfa_ids:
        collect_children_by_supercomponent(doc, list(seed_rfa_ids), all_rfa_ids)

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
    except:
        pass
    for nm in ("ProductDataRange", "Product Data Range"):
        s = get_param_str(itm, nm).strip()
        if s:
            return s
    return ""

def is_pipework_pipe_itm(itm):
    try:
        cat_ok = FAB_PIPEWORK_SUBSTR in (itm.Category.Name or "")
    except:
        cat_ok = False
    pdr_raw = (get_itm_product_data_range(itm) or "").strip().lower()
    pdr_ok = (pdr_raw == "pipe") or pdr_raw.endswith(".pipe") or pdr_raw.endswith(":pipe")
    return cat_ok and pdr_ok

def get_itm_product_short_description(itm):
    try:
        v = getattr(itm, "ProductShortDescription", None)
        if v:
            return str(v).strip()
    except:
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
    except:
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
    v = get_param_str(el, CP_ITEMNO_PARAM).strip()
    if v:
        return v
    for nm in ("ItemNumber", "Item Number", "Item #"):
        s = get_param_str(el, nm).strip()
        if s:
            return s
    return ""

def get_itm_length_ft(itm):
    try:
        if itm.ItemCustomId == FISHMOUTH_CID:
            return float(itm.CenterlineLength)
    except:
        pass
    try:
        if (itm.ProductLongDescription or "").startswith("Pipe"):
            return float(itm.CenterlineLength)
    except:
        pass
    return 0.0

def get_rfa_length_ft(fi):
    v = get_param_double(fi, CP_LENGTH_PARAM)
    return float(v) if v is not None else None

def get_rfa_description(doc, fi):
    v = get_param_str(fi, CP_DESC_PARAM).strip()
    if v:
        return v
    v = get_param_str(fi, CP_TYPE_NAME_PARAM).strip()
    if v:
        return v
    info = _get_type_info(doc, fi)
    tn = (info.get("type_name") or "").strip()
    return tn if tn else "Unknown"

def get_category_for_export(doc, el):
    if not isinstance(el, FamilyInstance):
        return get_revit_category_name(el)
    s = get_param_str(el, CPL_FAMILY_CATEGORY_PARAM).strip()
    if s:
        return s
    info = _get_type_info(doc, el)
    s2 = (info.get("cpl_family_category") or "").strip()
    if s2:
        return s2
    return get_revit_category_name(el)

def get_area(el):
    return get_param_str(el, CP_AREA_PARAM)

def get_spool_assembly_name(doc, el):
    try:
        aid = el.AssemblyInstanceId
        if aid and aid.IntegerValue != -1:
            asm = doc.GetElement(aid)
            if asm:
                return (asm.Name or "").strip()
    except:
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
        except:
            fam = ""
        info = _get_type_info(doc, el)
        typ = (info.get("type_name") or "").strip()
        if not typ:
            try:
                typ = (el.Symbol.Name or "").strip()
            except:
                typ = ""
        return fam, typ

    if isinstance(el, FabricationPart):
        info = _get_type_info(doc, el)
        typ = (info.get("type_name") or "").strip()
        fam = (info.get("family_name") or "").strip()
        if not typ:
            try:
                typ = (el.Name or "").strip()
            except:
                typ = ""
        return fam, typ

    return "", ""

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
    except:
        pass
    try:
        return (p.AsString() or "").strip()
    except:
        try:
            return (p.AsValueString() or "").strip()
        except:
            return ""

def get_itm_material(doc, itm):
    p = get_bip_param(itm, BuiltInParameter.FABRICATION_PART_MATERIAL)
    s = _material_name_from_param(doc, p)
    if s:
        return s
    p = get_bip_param(itm, BuiltInParameter.MATERIAL_ID_PARAM)
    s = _material_name_from_param(doc, p)
    if s:
        return s
    return ""

def get_rfa_material(fi):
    s = get_param_str(fi, CP_RFA_MATERIAL_PARAM_1)
    if s:
        return s
    return get_param_str(fi, CP_RFA_MATERIAL_PARAM_2)

def get_itm_service(itm):
    try:
        v = getattr(itm, "ServiceName", None)
        if v:
            return str(v).strip()
    except:
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
    except:
        pass
    for nm in ("Product Name", "ProductName", "Product"):
        s = get_param_str(itm, nm)
        if s:
            return s
    return "N/A"

def get_model_field_for_rfa(doc, fi):
    s = get_param_str(fi, CP_RFA_MODELNO_PARAM)
    if s:
        return s
    s = get_bip_str_by_id(fi, BIP_TYPE_MARK_ID_GUESS)
    if s:
        return s
    info = _get_type_info(doc, fi)
    s2 = (info.get("type_mark_guess") or "").strip()
    if s2:
        return s2
    tn = (info.get("type_name") or "").strip()
    return tn if tn else "N/A"

def should_exclude_by_family(doc, el):
    info = _get_type_info(doc, el)
    fam = (info.get("family_name") or "").strip()
    if not fam and info.get("type_el") is not None:
        try:
            fam = get_bip_str(info["type_el"], BuiltInParameter.ALL_MODEL_FAMILY_NAME)
        except:
            fam = ""
    return (EXCLUDE_FAMILY_SUBSTR in (fam or "").lower())


# ---------------- TRUE rod extraction ----------------
def safe_str(x):
    try:
        return "" if x is None else str(x)
    except:
        return ""

def try_get_support_rod_usage(itm_part):
    try:
        anc_list = list(itm_part.GetPartAncillaryUsage())
    except:
        return (None, None)
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
            return (dia_ft, len_ft)
    return (None, None)

def is_itm_hanger_with_cached_usage(itm, dia_ft, len_ft):
    if not isinstance(itm, FabricationPart):
        return False
    if dia_ft is not None or len_ft is not None:
        return True
    try:
        cn = (itm.Category.Name or "").lower()
    except:
        cn = ""
    return ("fabrication hangers" in cn) or ("hangers" in cn)

# Row indices (kept identical to your original)
IDX_SORT     = 0
IDX_CAT      = 1
IDX_FAM      = 2
IDX_TYP      = 3
IDX_DESC     = 5
IDX_SIZE     = 6
IDX_LENFT    = 7
IDX_LENFR    = 8
IDX_COUNT    = 9
IDX_ITEMCT   = 10
IDX_MODEL    = 11
IDX_INSTALL  = 18

def build_rod_row_from_hanger_cached(hanger_row, dia_ft, rod_len_ft, itm_hanger):
    if dia_ft is None and rod_len_ft is None:
        return None

    r = list(hanger_row)

    r[IDX_CAT] = "Structural Framing"
    r[IDX_FAM] = "All-Thread Rod"
    r[IDX_DESC] = "All-Thread Rod"

    if dia_ft is not None:
        dia_in = dia_ft * 12.0
        r[IDX_TYP] = excel_text_formula(fmt_inches_frac(dia_in, denom=16))
    else:
        r[IDX_TYP] = ""

    r[IDX_SIZE] = ""

    r[IDX_LENFT] = fmt_len_decimal_ft(rod_len_ft)
    r[IDX_LENFR] = fmt_len_ft_in_frac(rod_len_ft, denom=FRAC_DENOM)
    r[IDX_COUNT] = count_flag(rod_len_ft)

    # Item count stays "1"
    r[IDX_ITEMCT] = "1"

    r[IDX_MODEL] = ""
    r[IDX_INSTALL] = ""

    r[IDX_SORT] = get_sorting_value(itm_hanger, r[IDX_CAT])
    return r


# ---------------- Sorting ----------------
def stable_hash(s):
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
        except:
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
    base = series_base_for_element(el, cat_export)
    return str(assign_code(base, cat_export))


# ---------------- Selection ----------------
uidoc = _uidoc
doc = _doc

picked = []
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
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select scope to export (ITMs + RFAs, nested included).")
    except:
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
rfas_expanded = [fi for fi in rfas_expanded if is_modeling_element(fi) and (not should_exclude_by_family(doc, fi))]


# ---------------- Build rows ----------------
rows = []

header = [
    "Sorting","Category","Family","Type","ItemNumber","Description","Size",
    "Length_ft","Length_ft_in","Count","Item count",
    "Model","Area","Material","Comments","Spool","Service","Hanger Number","Install Type",
]

# Hot-loop local bindings (small but free speed)
rows_append = rows.append
sanitize_local = sanitize
fmt_len_dec = fmt_len_decimal_ft
fmt_len_frac = fmt_len_ft_in_frac
count_flag_local = count_flag
excel_formula = excel_text_formula
format_size = format_size_inch_marks

# ITMs / non-RFAs
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
    install_type_val = ""
    model_val = "N/A"
    length_for_count = None

    dia_ft = None
    rod_len_ft = None
    is_hanger = False

    if isinstance(el, FabricationPart):
        length_ft = get_itm_length_ft(el)
        length_for_count = length_ft

        base_desc = (el.ProductLongDescription or "")
        desc_val = apply_pipe_prefix_to_description(el, base_desc, length_ft)

        base_size = (el.ProductSizeDescription or "")
        size_val = excel_formula(format_size(base_size))

        material_val = get_itm_material(doc, el)
        service_val = get_itm_service(el)
        install_type_val = get_install_type_itm(el)
        model_val = get_model_field_for_itm(el)

        # PERF: compute rod usage once
        dia_ft, rod_len_ft = try_get_support_rod_usage(el)
        is_hanger = is_itm_hanger_with_cached_usage(el, dia_ft, rod_len_ft)

    elif is_revit_duct(el):
        p = get_bip_param(el, BuiltInParameter.CURVE_ELEM_LENGTH)
        length_ft = float(p.AsDouble()) if p else None
        length_for_count = length_ft
        desc_val = "Duct"
        size_val = ""

    else:
        desc_val = get_revit_category_name(el) or ""

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
        rod_row = build_rod_row_from_hanger_cached(base_row, dia_ft, rod_len_ft, el)
        if rod_row:
            rows_append(rod_row)

# RFAs
for fi in rfas_expanded:
    cat_export = get_category_for_export(doc, fi)
    sorting_val = get_sorting_value(fi, cat_export)

    fam, typ = get_family_type(doc, fi)
    item_no = get_item_number(fi)
    desc_val = get_rfa_description(doc, fi)

    size_val = excel_formula(format_size(get_param_str(fi, CP_SIZE_PARAM)))
    length_ft = get_rfa_length_ft(fi)

    model_val = get_model_field_for_rfa(doc, fi)
    area_val = get_area(fi)
    material_val = get_rfa_material(fi)
    comments_val = get_comments(doc, fi)
    spool_val = get_spool_assembly_name(doc, fi)
    service_val = get_param_str(fi, CP_RFA_SERVICE_PARAM)
    hanger_val = get_param_str(fi, CP_HANGER_NUMBER_PARAM)
    install_type_val = ""

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

# Build clipboard text (same output)
lines = ["\t".join(header)]
for row in rows:
    lines.append("\t".join([sanitize_local(c) for c in row]))

Clipboard.SetText("\r\n".join(lines))

forms.alert(
    "Clipboard export complete (read-only).\n\n"
    "Rows exported: {rows}\n\n"
    "Data copied to Windows clipboard.\n"
    "Paste directly into Excel.".format(rows=len(rows))
)
