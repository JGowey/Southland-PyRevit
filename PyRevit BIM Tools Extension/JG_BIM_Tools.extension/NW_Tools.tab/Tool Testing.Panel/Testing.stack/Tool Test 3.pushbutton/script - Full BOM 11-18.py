# -*- coding: utf-8 -*-
# SI Tools Number - Stable Renumber (Fab Pipe/Duct + RFAs w/ nested)
# + CP_Size, CP_Length, CP_BOM_Description, CP_Assembly Name (same pass / same transaction)
#
# IronPython notes:
# - Do NOT import Duct class; detect ducts by category (OST_DuctCurves)
# - Assembly name: use BuiltInParameter.ASSEMBLY_NAME first (matches Properties "Assembly Name")
# - IMPORTANT: target param name must match exactly: "CP_Assembly Name"

from pyrevit import revit, forms
from Autodesk.Revit.DB import (
    Transaction, FabricationPart, FamilyInstance, StorageType,
    BuiltInCategory, FilteredElementCollector, ElementId, BuiltInParameter
)
from Autodesk.Revit.UI.Selection import ObjectType


# ---------------- CONFIG ----------------
TARGET_PARAM = "CP_Item Number"  # always written when present

# ALSO write to ItemNumber (only for allowed MEP Fabrication Parts)
ITEMNUMBER_PARAM_NAMES = (
    "ItemNumber",
    "Item Number",
)

# Extra targets
SIZE_PARAM_TARGET      = "CP_Size"
LENGTH_PARAM_TARGET    = "CP_Length"
BOM_PARAM_TARGET       = "CP_BOM_Description"

# FIX: correct spelling to match your shared parameter name in the screenshot
ASSEMBLY_PARAM_TARGET  = "CP_Assembly Name"

# Inputs / helpers
DIAM_PARAM   = "Main Primary Diameter"
FAMILY_PARAM = "Family"

NESTED_FLAG_PARAM = "CP_Nested Family"
SKIP_GTP_FAMILY_NAME = "GTP"

ALLOWED_FAB_CAT_SUBSTRINGS = (
    "Fabrication Pipework",
    "Fabrication Ductwork",
)

PIPEWORK_SUBSTRING = "Fabrication Pipework"

# Revit internal units = feet. ~1/16" = 0.005208333 ft
LEN_TOL_FT = 0.005208333


# ---------------- Reporting / issues tracking ----------------
ISSUES = []
ISSUE_COUNTS = {
    "fab_cp_write_fail": 0,
    "fab_item_write_fail": 0,
    "rfa_cp_write_fail": 0,

    "size_write_fail": 0,
    "len_write_fail": 0,
    "bom_write_fail": 0,
    "asm_write_fail": 0,
}

WRITE_COUNTS = {
    "size_written": 0,
    "len_written": 0,
    "bom_written": 0,
    "asm_written": 0,
}

def log_issue(msg, key=None):
    if key:
        ISSUE_COUNTS[key] = ISSUE_COUNTS.get(key, 0) + 1
    if len(ISSUES) < 25:
        ISSUES.append(msg)


# ---------------- Param helpers ----------------
def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None

def get_first_param(el, names):
    for nm in names:
        p = get_param(el, nm)
        if p is not None:
            return p
    return None

def get_param_str(el, name):
    p = get_param(el, name)
    if not p:
        return ""
    try:
        v = p.AsString()
        return v if v else ""
    except:
        return ""

def get_param_valstr(el, name):
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

def set_param_value(p, value):
    """Safe write for String / Integer / Double."""
    if (not p) or p.IsReadOnly:
        return False

    st = p.StorageType

    if st == StorageType.String:
        try:
            p.Set("" if value is None else str(value))
            return True
        except:
            return False

    if st == StorageType.Integer:
        try:
            if value in (None, ""):
                p.Set(0)
                return True
            p.Set(int(value))
            return True
        except:
            return False

    if st == StorageType.Double:
        try:
            if value in (None, ""):
                p.Set(0.0)
                return True
            p.Set(float(value))
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


# ---------------- Revit Duct detection (IronPython-safe) ----------------
def is_revit_duct(el):
    try:
        return el.Category and el.Category.Id.IntegerValue == int(BuiltInCategory.OST_DuctCurves)
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

def is_pipework_fab(fp):
    try:
        return fp.Category and (PIPEWORK_SUBSTRING in (fp.Category.Name or ""))
    except:
        return False

def is_allowed_fab_part(el):
    return isinstance(el, FabricationPart) and is_allowed_fab_category(el)

def set_numbers_for_element(el, value_str):
    cp_p = get_param(el, TARGET_PARAM)
    cp_written = set_param_value(cp_p, value_str)

    if cp_p is not None and (not cp_written):
        if is_allowed_fab_part(el):
            log_issue("Fab: couldn't write '{}' on element {}".format(TARGET_PARAM, el.Id.IntegerValue),
                      key="fab_cp_write_fail")
        else:
            log_issue("RFA: couldn't write '{}' on element {}".format(TARGET_PARAM, el.Id.IntegerValue),
                      key="rfa_cp_write_fail")

    if is_allowed_fab_part(el):
        item_p = get_first_param(el, ITEMNUMBER_PARAM_NAMES)
        if item_p is not None:
            item_written = set_param_value(item_p, value_str)
            if not item_written:
                log_issue("Fab: couldn't write '{}' on element {}".format(item_p.Definition.Name, el.Id.IntegerValue),
                          key="fab_item_write_fail")

    return cp_written

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
    cat = rfa_category_name(fi)
    fam = rfa_family_name(fi)
    typ = rfa_type_name(fi)
    tid = rfa_type_id_int(fi)

    cp_len = get_param_length_double(fi, "CP_Length")
    cp_len_sort = cp_len if cp_len is not None else 1e99

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
            set_numbers_for_element(it, str(n))
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


# ---------------- Extra value calculators ----------------
def calc_cp_size(el):
    if isinstance(el, FabricationPart) and is_pipework_fab(el):
        try:
            return el.ProductSizeDescription or ""
        except:
            return ""
    return None

def calc_cp_length(doc, el):
    if isinstance(el, FabricationPart) and is_allowed_fab_category(el):
        try:
            cid = el.ItemCustomId
        except:
            cid = None

        try:
            if cid == 2875:
                return float(el.CenterlineLength)
        except:
            pass

        try:
            desc = el.ProductLongDescription or ""
            if desc.startswith("Pipe"):
                return float(el.CenterlineLength)
        except:
            pass

        return 0.0

    if is_revit_duct(el):
        try:
            p = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
            if p:
                return float(p.AsDouble())
        except:
            pass
        return 0.0

    return None

def calc_cp_bom_description(doc, el):
    if isinstance(el, FabricationPart):
        try:
            return el.ProductLongDescription or ""
        except:
            return ""

    if isinstance(el, FamilyInstance):
        inst_p = get_param(el, "CP_Type Name")
        try:
            if inst_p and inst_p.HasValue:
                v = (inst_p.AsString() or "").strip()
                if v:
                    return v
        except:
            pass

        try:
            tid = el.GetTypeId()
            te = doc.GetElement(tid) if tid else None
            if te:
                tp = te.LookupParameter("CP_Type Name")
                if tp and tp.HasValue:
                    v = (tp.AsString() or "").strip()
                    if v:
                        return v
        except:
            pass

        fam = rfa_family_name(el).strip()
        return fam if fam else "Unknown RFA Family"

    return None

def calc_cp_assembly_name(doc, el):
    """
    Prefer the built-in "Assembly Name" parameter (matches Properties palette).
    Fallback to AssemblyInstanceId lookup if needed.
    """
    # 1) Built-in "Assembly Name" on the element
    try:
        p = el.get_Parameter(BuiltInParameter.ASSEMBLY_NAME)
        if p:
            v = (p.AsString() or "").strip()
            if v:
                return v
    except:
        pass

    # 2) Fallback: AssemblyInstanceId -> AssemblyInstance.Name
    try:
        aid = el.AssemblyInstanceId
        if (aid is None) or (aid.IntegerValue == -1):
            return None
        asm = doc.GetElement(aid)
        if not asm:
            return None
        nm = getattr(asm, "Name", None)
        return nm if nm else None
    except:
        return None


def try_write_string(el, param_name, value, fail_key, count_key):
    p = get_param(el, param_name)
    if p is None:
        return
    ok = set_param_value(p, value if value is not None else "")
    if ok:
        WRITE_COUNTS[count_key] += 1
    else:
        log_issue("Couldn't write '{}' on element {}".format(param_name, el.Id.IntegerValue), key=fail_key)

def try_write_double(el, param_name, value, fail_key, count_key):
    p = get_param(el, param_name)
    if p is None:
        return
    ok = set_param_value(p, value)
    if ok:
        WRITE_COUNTS[count_key] += 1
    else:
        log_issue("Couldn't write '{}' on element {}".format(param_name, el.Id.IntegerValue), key=fail_key)


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

ducts_selected = [e for e in picked if is_revit_duct(e)]

if (not fab_sorted) and (not rfas_sorted) and (not ducts_selected):
    forms.alert(
        "Nothing in scope.\n\n"
        "- Select Fabrication Pipework/Ductwork parts and/or RFAs and/or Ducts\n"
        "- Ensure '{}' exists and is writable where expected".format(TARGET_PARAM),
        exitscript=True
    )

all_to_touch = []
all_to_touch.extend(fab_sorted)
all_to_touch.extend(rfas_sorted)
all_to_touch.extend(ducts_selected)

with Transaction(doc, "Renumber + CP Params (Stable overwrite)") as t:
    t.Start()

    # ---- Renumbering ----
    for fp in fab_sorted:
        set_numbers_for_element(fp, "")
    for fi in rfas_sorted:
        set_numbers_for_element(fi, "")

    doc.Regenerate()

    n = 1
    if fab_sorted:
        n = renumber_fab_all_overwrite(fab_sorted, n)
    if rfas_sorted:
        n = renumber_rfas_overwrite(rfas_sorted, n)

    # ---- Extra parameter writes ----
    for el in all_to_touch:
        v_size = calc_cp_size(el)
        if v_size is not None:
            try_write_string(el, SIZE_PARAM_TARGET, v_size, "size_write_fail", "size_written")

        v_len = calc_cp_length(doc, el)
        if v_len is not None:
            try_write_double(el, LENGTH_PARAM_TARGET, v_len, "len_write_fail", "len_written")

        v_bom = calc_cp_bom_description(doc, el)
        if v_bom is not None:
            try_write_string(el, BOM_PARAM_TARGET, v_bom, "bom_write_fail", "bom_written")

        v_asm = calc_cp_assembly_name(doc, el)
        if v_asm is not None:
            try_write_string(el, ASSEMBLY_PARAM_TARGET, v_asm, "asm_write_fail", "asm_written")

    t.Commit()

# ---------------- USER-FRIENDLY REPORT ----------------
fab_count = len(fab_sorted)
rfa_count = len(rfas_sorted)
duct_count = len(ducts_selected)
total = fab_count + rfa_count

msg = (
    "Renumber + parameter write complete.\n\n"
    "Renumbered:\n"
    "  Total items renumbered: {0}\n"
    "    Fabrication components: {1}\n"
    "    RFA components: {2}\n\n"
    "Also processed (for CP_Length/Assembly/etc):\n"
    "  Ducts selected: {3}\n\n"
    "Parameter writes:\n"
    "  CP_Size written: {4}\n"
    "  CP_Length written: {5}\n"
    "  CP_BOM_Description written: {6}\n"
    "  CP_Assembly Name written: {7}"
).format(
    total, fab_count, rfa_count, duct_count,
    WRITE_COUNTS["size_written"],
    WRITE_COUNTS["len_written"],
    WRITE_COUNTS["bom_written"],
    WRITE_COUNTS["asm_written"],
)

any_failures = any(v > 0 for v in ISSUE_COUNTS.values())
if any_failures:
    msg += (
        "\n\nIssues detected:\n"
        "  Fab CP write failures: {0}\n"
        "  Fab ItemNumber write failures: {1}\n"
        "  RFA CP write failures: {2}\n"
        "  CP_Size write failures: {3}\n"
        "  CP_Length write failures: {4}\n"
        "  CP_BOM_Description write failures: {5}\n"
        "  CP_Assembly Name write failures: {6}"
    ).format(
        ISSUE_COUNTS.get("fab_cp_write_fail", 0),
        ISSUE_COUNTS.get("fab_item_write_fail", 0),
        ISSUE_COUNTS.get("rfa_cp_write_fail", 0),
        ISSUE_COUNTS.get("size_write_fail", 0),
        ISSUE_COUNTS.get("len_write_fail", 0),
        ISSUE_COUNTS.get("bom_write_fail", 0),
        ISSUE_COUNTS.get("asm_write_fail", 0),
    )

    if ISSUES:
        msg += "\n\nExamples:\n  - " + "\n  - ".join(ISSUES)

forms.alert(msg)
