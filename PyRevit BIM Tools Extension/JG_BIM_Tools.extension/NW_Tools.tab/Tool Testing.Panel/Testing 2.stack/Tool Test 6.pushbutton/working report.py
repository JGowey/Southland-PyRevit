# -*- coding: utf-8 -*-
# Clipboard Export (Read-Only, Instance Itemized + Nested RFAs)
# Preselection supported; otherwise prompts.
# Exports one row per instance to Windows clipboard (Excel-ready).
# Pipe tweak:
#   For ITM in "MEP Fabrication Pipework" AND ProductDataRange == Pipe:
#     Prefix DESCRIPTION with 21FT-/40FT- (based on Product Short Description, fallback to placed length).
# Size tweak:
#   Restore inch symbols in ITM sizes (e.g. 3 -> 3", 3 x 1/2 -> 3" x 1/2").
# NO writes.

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    FabricationPart, FamilyInstance, FilteredElementCollector,
    BuiltInCategory, BuiltInParameter, ElementId
)
from Autodesk.Revit.UI.Selection import ObjectType

import clr
import math
import re

clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Clipboard


# ---------------- CONFIG ----------------
CP_ITEMNO_PARAM     = "CP_Item Number"
CP_DESC_PARAM       = "CP_BOM_Description"
CP_SIZE_PARAM       = "CP_Size"
CP_LENGTH_PARAM     = "CP_Length"
CP_TYPE_NAME_PARAM  = "CP_Type Name"

FISHMOUTH_CID = 2875
FRAC_DENOM = 16
ZERO_TOL_FT = 1e-9

PIPE_STOCK_THRESHOLD_FT = 21.0
PIPE_PREFIX_SHORT = "21FT-"
PIPE_PREFIX_LONG  = "40FT-"
FAB_PIPEWORK_SUBSTR = "MEP Fabrication Pipework"


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
        return p.AsString() or ""
    except:
        try:
            return p.AsValueString() or ""
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

def is_revit_duct(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves)
    except:
        return False

def get_category_name(el):
    try:
        return el.Category.Name if el.Category else ""
    except:
        return ""

def rfa_family_name(fi):
    try:
        return fi.Symbol.Family.Name
    except:
        return ""

def rfa_type_name(fi):
    try:
        return fi.Symbol.Name
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

    for rid in list(seed_rfa_ids):
        host = doc.GetElement(ElementId(rid))
        if isinstance(host, FamilyInstance):
            collect_nested_subcomponents(doc, host, all_rfa_ids)

    collect_children_by_supercomponent(doc, list(seed_rfa_ids), all_rfa_ids)

    rfas = []
    for rid in all_rfa_ids:
        fi = doc.GetElement(ElementId(rid))
        if isinstance(fi, FamilyInstance):
            rfas.append(fi)

    return non_rfa, rfas


# ---------------- ITM Pipe detection + description prefix ----------------
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
    # property first
    try:
        v = getattr(itm, "ProductShortDescription", None)
        if v:
            return str(v).strip()
    except:
        pass
    # parameter fallback (matches your UI label)
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

    # fallback to placed length
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
        # avoid double-prefix if rerun/pasted twice
        if base_desc.startswith(PIPE_PREFIX_SHORT) or base_desc.startswith(PIPE_PREFIX_LONG):
            return base_desc
        return pref + base_desc
    return base_desc


# ---------------- Size formatting (restore inch marks) ----------------
def _ensure_quotes_on_token(tok):
    """
    For a token like '3' or '1/2' return '3"' or '1/2"'.
    If it already contains " or ' or ft/in symbols, leave it.
    """
    t = tok.strip()
    if not t:
        return t
    if ('"' in t) or ("'" in t):
        return t
    # if it has digits or fractions, treat as inches
    if re.match(r"^[0-9]+(\s*[0-9]+/[0-9]+)?$", t) or re.match(r"^[0-9]+/[0-9]+$", t):
        return t + '"'
    return t

def format_itm_size_inch_marks(size_str):
    """
    Examples:
      '3' -> '3"'
      '3 x 1/2' -> '3" x 1/2"'
      '3x1/2' -> '3" x 1/2"'
      '6 x 3' -> '6" x 3"'
    Leaves anything that already has quotes alone.
    """
    s = (size_str or "").strip()
    if not s:
        return s
    if '"' in s:
        return s  # already good

    # normalize separators to " x "
    s_norm = s.replace("×", "x")
    s_norm = re.sub(r"\s*x\s*", " x ", s_norm)

    if " x " in s_norm:
        parts = [p.strip() for p in s_norm.split(" x ") if p.strip() != ""]
        parts_q = [_ensure_quotes_on_token(p) for p in parts]
        return " x ".join(parts_q)

    # single token
    return _ensure_quotes_on_token(s_norm)


# ---------------- Value extraction ----------------
def get_item_number(el):
    v = get_param_str(el, "CP_Item Number").strip()
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
        # keep your existing length logic
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
    try:
        te = doc.GetElement(fi.GetTypeId())
        if te:
            p = te.LookupParameter(CP_TYPE_NAME_PARAM)
            if p and p.HasValue:
                return p.AsString() or ""
    except:
        pass
    return rfa_family_name(fi) or "Unknown RFA Family"


# ---------------- Selection: pre or post ----------------
uidoc = revit.uidoc
doc = revit.doc

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

if not picked:
    forms.alert("Nothing selected.", exitscript=True)


# ---------------- Build rows ----------------
non_rfa, rfas_expanded = expand_selection_with_nested_rfas(doc, picked)

rows = []
itm_count = 0
rfa_count = 0
duct_count = 0

for el in non_rfa:
    eid = str(el.Id.IntegerValue)
    cat = get_category_name(el)

    if isinstance(el, FabricationPart):
        itm_count += 1
        length_ft = get_itm_length_ft(el)

        base_desc = (el.ProductLongDescription or "")
        desc_val = apply_pipe_prefix_to_description(el, base_desc, length_ft)

        base_size = (el.ProductSizeDescription or "")
        size_val = format_itm_size_inch_marks(base_size)

        rows.append([
            eid, cat, "ITM", "", "",
            get_item_number(el),
            desc_val,
            size_val,
            fmt_len_decimal_ft(length_ft),
            fmt_len_ft_in_frac(length_ft, denom=FRAC_DENOM),
            count_flag(length_ft),
        ])
        continue

    if is_revit_duct(el):
        duct_count += 1
        p = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        length_ft = float(p.AsDouble()) if p else None
        rows.append([
            eid, cat, "Duct", "", "",
            get_item_number(el),
            "Duct",
            "",
            fmt_len_decimal_ft(length_ft),
            fmt_len_ft_in_frac(length_ft, denom=FRAC_DENOM),
            count_flag(length_ft),
        ])
        continue

for fi in rfas_expanded:
    rfa_count += 1
    eid = str(fi.Id.IntegerValue)
    cat = get_category_name(fi)
    length_ft = get_rfa_length_ft(fi)

    rows.append([
        eid, cat, "RFA",
        rfa_family_name(fi),
        rfa_type_name(fi),
        get_item_number(fi),
        get_rfa_description(doc, fi),
        get_param_str(fi, CP_SIZE_PARAM),
        fmt_len_decimal_ft(length_ft),
        fmt_len_ft_in_frac(length_ft, denom=FRAC_DENOM),
        count_flag(length_ft),
    ])

if not rows:
    forms.alert("No valid elements found.", exitscript=True)

header = [
    "ElementId", "Category", "Kind", "Family", "Type",
    "ItemNumber", "Description", "Size",
    "Length_ft", "Length_ft_in", "Count"
]

lines = ["\t".join(header)]
for row in rows:
    lines.append("\t".join([sanitize(c) for c in row]))

Clipboard.SetText("\r\n".join(lines))

forms.alert(
    "Clipboard export complete (read-only).\n\n"
    "Rows exported: {rows}\n"
    "  ITMs: {itms}\n"
    "  RFAs (incl. nested): {rfas}\n"
    "  Ducts: {ducts}\n\n"
    "Data copied to Windows clipboard.\n"
    "Paste directly into Excel.".format(
        rows=len(rows), itms=itm_count, rfas=rfa_count, ducts=duct_count
    )
)
