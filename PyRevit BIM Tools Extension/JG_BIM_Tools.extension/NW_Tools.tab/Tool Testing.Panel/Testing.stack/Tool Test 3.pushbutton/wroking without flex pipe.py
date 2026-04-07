# -*- coding: utf-8 -*-
# SI Tools Number - Stable Renumber (Fab Pipe/Duct + Flex + RFAs w/ nested)
# IronPython-safe (pyRevit)

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

# Nested-family flag used in your expanded script
NESTED_FLAG_PARAM = "CP_Nested Family"
SKIP_GTP_FAMILY_NAME = "GTP"

# Include both (fixes ducting being excluded)
ALLOWED_FAB_CAT_SUBSTRINGS = (
    "Fabrication Pipework",
    "Fabrication Ductwork",
)

# Grouping tolerance for CenterlineLength (Revit internal feet)
# ~1/16" = 0.005208333 ft, ~1/8" = 0.010416667 ft
LEN_TOL_FT = 0.005208333


# ---------------- Param helpers ----------------
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
        v = p.AsValueString()
        return v if v else ""
    except:
        try:
            v = p.AsString()
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

def set_param_value(p, value_str):
    if (not p) or p.IsReadOnly:
        return False
    st = p.StorageType
    if st == StorageType.String:
        p.Set(value_str if value_str is not None else "")
        return True
    if st == StorageType.Integer:
        # integer "blank" is 0
        if value_str in (None, ""):
            p.Set(0)
            return True
        try:
            p.Set(int(value_str))
            return True
        except:
            return False
    return False

def is_blank_param(p):
    if not p:
        return True
    st = p.StorageType
    if st == StorageType.String:
        s = p.AsString()
        return (s is None) or (s == "")
    if st == StorageType.Integer:
        try:
            return p.AsInteger() == 0
        except:
            return True
    return True

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

def is_pipe_weld(fp):
    # best-effort "SI Tools extension" replacement
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
    fam = get_param_str(fp, FAMILY_PARAM) or ""
    return "weld" in fam.lower()

def fab_display_name(fp):
    # try ItemName first; fall back to family param; then type name
    try:
        nm = getattr(fp, "ItemName", None)
        if nm:
            return str(nm)
    except:
        pass
    fam = get_param_str(fp, FAMILY_PARAM)
    if fam:
        return fam
    try:
        return str(fp.Name)
    except:
        return ""

def is_flex_fab(fp):
    # broaden to catch more flex cases (duct/pipe naming variations)
    s = (fab_display_name(fp) or "").lower()
    c = ""
    try:
        c = (fp.Category.Name or "").lower()
    except:
        pass
    return ("flex" in s) or ("flex" in c)

def norm_len_ft(x):
    try:
        xf = float(x)
        return round(xf / LEN_TOL_FT) * LEN_TOL_FT
    except:
        return None

def fab_group_key(fp):
    # IMPORTANT: normalize length so tiny float diffs don’t split groups
    fam = get_param_str(fp, FAMILY_PARAM)
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
    # SI Tools-style ordering + deterministic tie-breaker
    is_straight = bool(fp.IsAStraight())
    not_tap = bool(not fp.IsATap())
    diam = float(get_param_double(fp, DIAM_PARAM, 0.0))
    length = 0.0
    try:
        length = float(fp.CenterlineLength)
    except:
        length = 0.0
    return (is_straight, not_tap, -diam, -length, fp.Id.IntegerValue)


# ---------------- Grouping / numbering ----------------
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

def assign_numbers_from_groups(ordered_keys, groups, start_n):
    n = start_n
    for k in ordered_keys:
        items = groups.get(k, [])
        any_blank = False
        for it in items:
            if is_blank_param(get_param(it, TARGET_PARAM)):
                any_blank = True
                break
        if not any_blank:
            continue
        for it in items:
            p = get_param(it, TARGET_PARAM)
            if is_blank_param(p):
                set_param_value(p, str(n))
        n += 1
    return n

def renumber_fab_all(fab_sorted, start_n):
    n = start_n

    # 1) Non-flex, non-weld straights
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and q.IsAStraight()
    )
    n = assign_numbers_from_groups(keys, groups, n)

    # 2) Non-flex, non-weld, non-straight, non-tap
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and (not q.IsAStraight()) and (not q.IsATap())
    )
    n = assign_numbers_from_groups(keys, groups, n)

    # 3) Non-flex, non-weld taps
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and q.IsATap()
    )
    n = assign_numbers_from_groups(keys, groups, n)

    # 4) Remaining non-flex, non-weld blanks (catch-all)
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and (not is_pipe_weld(q)) and is_blank_param(get_param(q, TARGET_PARAM))
    )
    n = assign_numbers_from_groups(keys, groups, n)

    # 5) Non-flex welds (remaining blanks)
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: (not is_flex_fab(q)) and is_pipe_weld(q) and is_blank_param(get_param(q, TARGET_PARAM))
    )
    n = assign_numbers_from_groups(keys, groups, n)

    # 6) Flex last
    keys, groups = build_groups_in_order(
        fab_sorted, fab_group_key,
        filter_func=lambda q: is_flex_fab(q) and is_blank_param(get_param(q, TARGET_PARAM))
    )
    n = assign_numbers_from_groups(keys, groups, n)

    return n


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
    # uses your CP_* parameters (if they exist); still stable if blanks
    return (
        rfa_family_name(fi),
        rfa_type_name(fi),
        get_param_str(fi, "CP_Length"),
        get_param_str(fi, "CP_Width"),
        get_param_str(fi, "CP_Depth"),
        get_param_str(fi, "CP_Diameter"),
        get_param_str(fi, "CP_Material Thickness"),
    )

def sort_key_rfa(fi):
    return (rfa_family_name(fi), rfa_type_name(fi), fi.Id.IntegerValue)

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

def renumber_rfas(rfas_sorted, start_n):
    keys, groups = build_groups_in_order(
        rfas_sorted, rfa_group_key,
        filter_func=lambda fi: is_blank_param(get_param(fi, TARGET_PARAM))
    )
    return assign_numbers_from_groups(keys, groups, start_n)


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

# RFAs (selection + nested)
rfas_sorted = build_rfa_list(doc, picked)

if (not fab_sorted) and (not rfas_sorted):
    forms.alert(
        "Nothing in scope.\n\n"
        "- Did you select Fabrication Pipework/Ductwork parts?\n"
        "- Do selected items have '{0}' and is it writable?\n".format(TARGET_PARAM),
        exitscript=True
    )

with Transaction(doc, "Renumber SI Tools Stable ({0})".format(TARGET_PARAM)) as t:
    t.Start()

    # Clear ALL Fab in scope (including flex) so nothing stays null/old unintentionally
    for fp in fab_sorted:
        set_param_value(get_param(fp, TARGET_PARAM), "")

    # Clear ALL RFAs in scope
    for fi in rfas_sorted:
        set_param_value(get_param(fi, TARGET_PARAM), "")

    n = 1

    # Phase 1+2: Fab (non-flex buckets, then flex)
    if fab_sorted:
        n = renumber_fab_all(fab_sorted, n)

    # Phase 3: RFAs
    if rfas_sorted:
        n = renumber_rfas(rfas_sorted, n)

    t.Commit()

# Counts
fab_flex_count = len([x for x in fab_sorted if is_flex_fab(x)])
fab_nonflex_count = len(fab_sorted) - fab_flex_count

forms.alert(
    "Done.\n\n"
    "Fab total (Pipe+Duct): {0}\n"
    "  Non-flex: {1}\n"
    "  Flex: {2}\n"
    "Fab skipped (missing '{3}'): {4}\n"
    "Fab skipped (read-only '{3}'): {5}\n\n"
    "RFAs (incl nested): {6}\n\n"
    "Wrote '{3}' starting at 1 (stable ordering).".format(
        len(fab_sorted),
        fab_nonflex_count,
        fab_flex_count,
        TARGET_PARAM,
        fab_missing_param,
        fab_readonly_param,
        len(rfas_sorted)
    )
)
