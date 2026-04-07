# -*- coding: utf-8 -*-
# SI Tools Number - Stable Renumber (Fab Pipe/Duct + RFAs w/ nested)
# Stable reruns by: (1) clear, (2) doc.Regenerate(), (3) overwrite numbering (no blank-gating)
#
# RFA updates:
# - Sort RFAs: Category > Family > Type > CP_Length
# - Group RFAs: Family + TypeId + CP_Length(normalized numeric) + CP_Width/Depth/Diameter/Material Thickness

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    Transaction, FabricationPart, FamilyInstance, StorageType,
    BuiltInCategory, FilteredElementCollector, ElementId
)
from Autodesk.Revit.UI.Selection import ObjectType


# ---------------- CONFIG ----------------
TARGET_PARAM = "CP_Item Number"
DIAM_PARAM   = "Main Primary Diameter"
FAMILY_PARAM = "Family"

NESTED_FLAG_PARAM = "CP_Nested Family"
SKIP_GTP_FAMILY_NAME = "GTP"

ALLOWED_FAB_CAT_SUBSTRINGS = (
    "Fabrication Pipework",
    "Fabrication Ductwork",
)

# Revit internal units = feet. ~1/16" = 0.005208333 ft
LEN_TOL_FT = 0.005208333


# ---------------- Param helpers ----------------
def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None

def get_param_str(el, name):
    """Best for true string params. For numeric params, prefer AsValueString or AsDouble."""
    p = get_param(el, name)
    if not p:
        return ""
    try:
        v = p.AsString()
        return v if v else ""
    except:
        return ""

def get_param_valstr(el, name):
    """Human-readable value; works for many numeric params but may vary by units formatting."""
    p = get_param(el, name)
    if not p:
        return ""
    try:
        v = p.AsValueString()
        return v if v else ""
    except:
        return ""

def get_param_double(el, name, default=0.0):
    p = get_param(el, name)
    if not p:
        return default
    try:
        return p.AsDouble()
    except:
        return default

def norm_len_ft(x):
    try:
        xf = float(x)
        return round(xf / LEN_TOL_FT) * LEN_TOL_FT
    except:
        return None

def get_param_length_double(el, name):
    """
    Safely read a Length/Double parameter for grouping.
    Returns a rounded float in feet, or None.
    """
    p = get_param(el, name)
    if not p:
        return None
    try:
        if p.StorageType == StorageType.Double:
            val = p.AsDouble()
            if val is None:
                return None
            return norm_len_ft(val)
    except:
        pass
    return None

def set_param_value(p, value_str):
    """Write strings/ints safely. Returns True if written."""
    if (not p) or p.IsReadOnly:
        return False
    st = p.StorageType
    if st == StorageType.String:
        p.Set(value_str if value_str is not None else "")
        return True
    if st == StorageType.Integer:
        if value_str in (None, ""):
            p.Set(0)
            return True
        try:
            p.Set(int(value_str))
            return True
        except:
            return False
    return False

def is_param_true(p):
    if not p:
        return False
    try:
        if p.StorageType == StorageType.Integer:
            return p.AsInteger() == 1
    except:
        pass
    try:
        vs = (p.AsValueString() or "").strip().lower()
        return vs in ("yes", "true", "1")
    except:
        return False


# ---------------- Fab helpers ----------------
def is_allowed_fab_category(fp):
    try:
        if not fp.Category:
            return False
        cname = fp.Category.Name or ""
        for sub in ALLOWED_FAB_CAT_SUBSTRINGS:
            if sub in cname:
                return True
        return False
    except:
        return False

def fab_display_name(fp):
    try:
        nm = getattr(fp, "ItemName", None)
        if nm:
            return str(nm)
    except:
        pass
    fam = get_param_valstr(fp, FAMILY_PARAM) or get_param_str(fp, FAMILY_PARAM)
    if fam:
        return fam
    try:
        return str(fp.Name)
    except:
        return ""

def is_flex_fab(fp):
    s = (fab_display_name(fp) or "").lower()
    c = ""
    try:
        c = (fp.Category.Name or "").lower()
    except:
        pass
    return ("flex" in s) or ("flex" in c)

def is_pipe_weld(fp):
    # NOTE: This is legacy "weld" detection by substring match only.
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
    fam = get_param_valstr(fp, FAMILY_PARAM) or get_param_str(fp, FAMILY_PARAM) or ""
    return "weld" in fam.lower()

def fab_group_key(fp):
    fam = get_param_valstr(fp, FAMILY_PARAM) or get_param_str(fp, FAMILY_PARAM)
    ln  = None
    try:
        ln = fp.CenterlineLength
    except:
        ln = None
    ln_n = norm_len_ft(ln) if ln is not None else None

    return (
        fam,
        getattr(fp, "Size", ""),
        getattr(fp, "Material", ""),
        getattr(fp, "Specification", ""),
        ln_n
    )

def sort_key_fab(fp):
    is_straight = bool(fp.IsAStraight())
    not_tap = bool(not fp.IsATap())
    diam = float(get_param_double(fp, DIAM_PARAM, 0.0))
    length = 0.0
    try:
        length = float(fp.CenterlineLength)
    except:
        length = 0.0
    return (is_straight, not_tap, -diam, -length, fp.Id.IntegerValue)


# ---------------- RFA helpers (selection-based + nested) ----------------
def rfa_family_name(fi):
    try:
        sym = fi.Symbol
        return (sym.Family.Name if sym and sym.Family else "") or ""
    except:
        return ""

def rfa_type_name(fi):
    try:
        sym = fi.Symbol
        return (sym.Name if sym else "") or ""
    except:
        return ""

def rfa_type_id_int(fi):
    try:
        tid = fi.GetTypeId()
        if tid:
            return tid.IntegerValue
    except:
        pass
    return -1

def rfa_category_name(fi):
    try:
        if fi.Category:
            return fi.Category.Name or ""
    except:
        pass
    return ""

def is_gtp_rfa(fi):
    try:
        is_generic = fi.Category and fi.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel)
    except:
        is_generic = False
    return is_generic and (rfa_family_name(fi).strip().upper() == SKIP_GTP_FAMILY_NAME)

def should_skip_host_but_include_nested(fi):
    return is_param_true(get_param(fi, NESTED_FLAG_PARAM))

def collect_nested_subcomponents(doc, host_fi, out_ids):
    try:
        sub_ids = list(host_fi.GetSubComponentIds())
    except:
        sub_ids = []
    for sid in sub_ids:
        sub_el = doc.GetElement(sid)
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

def rfa_group_key(fi):
    """
    Group RFAs by Family + TypeId + normalized CP_Length (numeric) + other CP descriptors.
    """
    return (
        rfa_family_name(fi),
        rfa_type_id_int(fi),
        get_param_length_double(fi, "CP_Length"),
        get_param_str(fi, "CP_Width"),
        get_param_str(fi, "CP_Depth"),
        get_param_str(fi, "CP_Diameter"),
        get_param_str(fi, "CP_Material Thickness"),
    )

def sort_key_rfa(fi):
    """
    RFA ordering:
    1) Revit Category
    2) Family
    3) Type Name
    4) TypeId (tie-breaker for same name)
    5) CP_Length (numeric normalized; None last)
    6) ElementId (final tie-breaker)
    """
    cat = rfa_category_name(fi)
    fam = rfa_family_name(fi)
    typ = rfa_type_name(fi)
    tid = rfa_type_id_int(fi)

    cp_len = get_param_length_double(fi, "CP_Length")
    cp_len_sort = cp_len if cp_len is not None else 1e99  # push None to the end

    return (cat, fam, typ, tid, cp_len_sort, fi.Id.IntegerValue)

def build_rfa_list(doc, picked):
    rfa_ids = set()
    host_ids = set()

    for el in picked:
        if not isinstance(el, FamilyInstance):
            continue

        host_ids.add(el.Id.IntegerValue)

        nested = set()
        collect_nested_subcomponents(doc, el, nested)
        for nid in nested:
            rfa_ids.add(nid)

        if is_gtp_rfa(el):
            continue
        if should_skip_host_but_include_nested(el):
            continue
        if get_param(el, TARGET_PARAM) is None:
            continue
        rfa_ids.add(el.Id.IntegerValue)

    collect_children_by_supercomponent(doc, host_ids, rfa_ids)

    rfas = []
    for eid in rfa_ids:
        fi = doc.GetElement(ElementId(eid))
        if isinstance(fi, FamilyInstance) and (not is_gtp_rfa(fi)) and get_param(fi, TARGET_PARAM) is not None:
            rfas.append(fi)

    return sorted(rfas, key=sort_key_rfa)


# ---------------- Grouping / numbering (OVERWRITE) ----------------
def build_groups_in_order(items_sorted, key_func, filter_func=None):
    groups = {}
    ordered = []
    for it in items_sorted:
        if filter_func and (not filter_func(it)):
            continue
        k = key_func(it)
        if k not in groups:
            groups[k] = []
            ordered.append(k)
        groups[k].append(it)
    return ordered, groups

def assign_numbers_overwrite(ordered_keys, groups, start_n):
    n = start_n
    for k in ordered_keys:
        items = groups.get(k, [])
        for it in items:
            set_param_value(get_param(it, TARGET_PARAM), str(n))
        n += 1
    return n

def renumber_fab_all_overwrite(fab_sorted, start_n):
    n = start_n

    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and q.IsAStraight()
    )
    n = assign_numbers_overwrite(keys, groups, n)

    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and (not q.IsAStraight()) and (not q.IsATap())
    )
    n = assign_numbers_overwrite(keys, groups, n)

    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and q.IsATap()
    )
    n = assign_numbers_overwrite(keys, groups, n)

    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and is_pipe_weld(q)
    )
    n = assign_numbers_overwrite(keys, groups, n)

    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: is_flex_fab(q)
    )
    n = assign_numbers_overwrite(keys, groups, n)

    return n

def renumber_rfas_overwrite(rfas_sorted, start_n):
    keys, groups = build_groups_in_order(rfas_sorted, rfa_group_key)
    return assign_numbers_overwrite(keys, groups, start_n)


# ---------------- MAIN ----------------
uidoc = revit.uidoc
doc = revit.doc

try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select scope (Fab Pipe/Duct + RFAs).")
except:
    forms.alert("Selection cancelled.", exitscript=True)

picked = [doc.GetElement(r) for r in refs if r]
picked = [e for e in picked if e is not None]

# Collect Fab parts in allowed categories that have TARGET_PARAM
fab_parts = []
fab_missing_param = 0
fab_readonly_param = 0

for el in picked:
    if isinstance(el, FabricationPart) and is_allowed_fab_category(el):
        p = get_param(el, TARGET_PARAM)
        if p is None:
            fab_missing_param += 1
            continue
        if p.IsReadOnly:
            fab_readonly_param += 1
            continue
        fab_parts.append(el)

fab_sorted = sorted(fab_parts, key=sort_key_fab)
rfas_sorted = build_rfa_list(doc, picked)

if (not fab_sorted) and (not rfas_sorted):
    forms.alert(
        "Nothing in scope.\n\n"
        "- Select Fabrication Pipework/Ductwork parts and/or RFAs\n"
        "- Ensure '{0}' exists and is writable on those instances".format(TARGET_PARAM),
        exitscript=True
    )

with Transaction(doc, "Renumber SI Tools Stable (overwrite)") as t:
    t.Start()

    # Clear all in-scope first
    for fp in fab_sorted:
        set_param_value(get_param(fp, TARGET_PARAM), "")
    for fi in rfas_sorted:
        set_param_value(get_param(fi, TARGET_PARAM), "")

    # Critical: keep reads/writes coherent inside this transaction
    doc.Regenerate()

    n = 1
    if fab_sorted:
        n = renumber_fab_all_overwrite(fab_sorted, n)
    if rfas_sorted:
        n = renumber_rfas_overwrite(rfas_sorted, n)

    t.Commit()

fab_flex_count = len([x for x in fab_sorted if is_flex_fab(x)])
forms.alert(
    "Done (stable overwrite).\n\n"
    "Fab total (Pipe+Duct): {0}\n"
    "  Flex: {1}\n"
    "Fab skipped missing '{2}': {3}\n"
    "Fab skipped read-only '{2}': {4}\n\n"
    "RFAs (incl nested): {5}\n\n"
    "RFA sort: Category > Family > Type > CP_Length\n"
    "RFA group: Family + TypeId + CP_Length(normalized) + CP descriptors\n"
    "Behavior is idempotent: reruns should produce identical results.".format(
        len(fab_sorted),
        fab_flex_count,
        TARGET_PARAM,
        fab_missing_param,
        fab_readonly_param,
        len(rfas_sorted)
    )
)
