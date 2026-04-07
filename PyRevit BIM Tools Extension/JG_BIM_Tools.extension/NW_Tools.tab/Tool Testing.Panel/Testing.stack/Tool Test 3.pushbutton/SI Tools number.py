# -*- coding: utf-8 -*-
# pyRevit remake of SI Tools RenumberPipeCommand
# Writes to: CP_Item Number

from pyrevit import revit, forms
from Autodesk.Revit.DB import Transaction, FabricationPart, StorageType
from Autodesk.Revit.UI.Selection import ObjectType


TARGET_PARAM = "CP_Item Number"
DIAM_PARAM = "Main Primary Diameter"
FAMILY_PARAM = "Family"


def get_param(el, name):
    return el.LookupParameter(name)


def get_param_str(el, name):
    p = get_param(el, name)
    if not p:
        return None
    # Match SI Tools: Family compared via AsValueString()
    try:
        return p.AsValueString()
    except:
        try:
            return p.AsString()
        except:
            return None


def get_param_double(el, name, default=0.0):
    p = get_param(el, name)
    if not p:
        return default
    try:
        return p.AsDouble()
    except:
        return default


def set_param_value(p, value_str):
    """Set target param; supports Text (string) and Integer."""
    if not p or p.IsReadOnly:
        return False

    st = p.StorageType
    if st == StorageType.String:
        p.Set(value_str if value_str is not None else "")
        return True

    if st == StorageType.Integer:
        # If your CP_Item Number is integer, it must be pure digits.
        if value_str in (None, ""):
            p.Set(0)
            return True
        try:
            p.Set(int(value_str))
            return True
        except:
            return False

    # Not expected for item number
    return False


def is_blank_param(p):
    """Blank = empty string for Text; 0 for Integer."""
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


def is_pipe_weld(fp):
    """Replacement for SI Tools extension IsPipeWeld(). Best-effort."""
    # Try common enum-like properties (varies by Revit version / API)
    for attr in ("ItemType", "PartType", "FabricationPartType"):
        v = getattr(fp, attr, None)
        if v is not None:
            try:
                if "weld" in str(v).lower():
                    return True
            except:
                pass

    # Try item name / element name
    for attr in ("ItemName", "Name"):
        v = getattr(fp, attr, None)
        if v:
            try:
                if "weld" in str(v).lower():
                    return True
            except:
                pass

    # Fallback: check Family parameter string
    fam = get_param_str(fp, FAMILY_PARAM) or ""
    if "weld" in fam.lower():
        return True

    return False


def group_key(fp):
    """Grouping keys per SI Tools: Family (AsValueString), Size, Material, Specification, CenterlineLength"""
    fam = get_param_str(fp, FAMILY_PARAM) or ""
    return (fam, fp.Size, fp.Material, fp.Specification, fp.CenterlineLength)


def sort_key(fp):
    """Match SI Tools ordering:
       OrderBy(IsAStraight) then ThenBy(!IsATap) then diameter desc then length desc.
       NOTE: OrderBy(bool) puts False before True.
    """
    is_straight = bool(fp.IsAStraight())
    not_tap = bool(not fp.IsATap())
    diam = float(get_param_double(fp, DIAM_PARAM, 0.0))
    length = float(fp.CenterlineLength)
    return (is_straight, not_tap, -diam, -length)


def clear_target(parts):
    for fp in parts:
        tp = get_param(fp, TARGET_PARAM)
        set_param_value(tp, "")


def renumber_bucket(all_parts, bucket_filter, number_start):
    """Apply SI Tools numbering logic to one bucket."""
    number = number_start
    bucket = [p for p in all_parts if bucket_filter(p)]

    for fp in bucket:
        tp = get_param(fp, TARGET_PARAM)
        if not is_blank_param(tp):
            continue

        # Assign current number to this part
        set_param_value(tp, str(number))
        k = group_key(fp)

        # Assign same number to all identical parts still blank in this bucket
        for other in bucket:
            otp = get_param(other, TARGET_PARAM)
            if is_blank_param(otp) and group_key(other) == k:
                set_param_value(otp, str(number))

        number += 1

    return number


# -------------------- MAIN --------------------
uidoc = revit.uidoc
doc = revit.doc

try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select Fabrication Pipework to renumber")
except:
    forms.alert("Selection cancelled.", exitscript=True)

# Pre-filter matches SI Tools except for parameter name:
# - Category contains "Fabrication Pipework"
# - Must have TARGET_PARAM (CP_Item Number)
# - Must be FabricationPart
parts = []
for r in refs:
    el = doc.GetElement(r)
    if isinstance(el, FabricationPart):
        if el.Category and ("Fabrication Pipework" in el.Category.Name):
            if get_param(el, TARGET_PARAM) is not None:
                parts.append(el)

if not parts:
    forms.alert("No Fabrication Pipework parts found with parameter '{}'.".format(TARGET_PARAM), exitscript=True)

# Sort per SI Tools
parts = sorted(parts, key=sort_key)

with Transaction(doc, "Renumber Pipes ({})".format(TARGET_PARAM)) as t:
    t.Start()

    # Clear all target values (SI Tools clears ItemNumber first)
    clear_target(parts)

    n = 1

    # Pass 1: non-weld straights
    n = renumber_bucket(parts, lambda q: (not is_pipe_weld(q)) and q.IsAStraight(), n)

    # Pass 2: non-weld, non-straight, non-tap
    n = renumber_bucket(parts, lambda q: (not is_pipe_weld(q)) and (not q.IsAStraight()) and (not q.IsATap()), n)

    # Pass 3: non-weld taps
    n = renumber_bucket(parts, lambda q: (not is_pipe_weld(q)) and q.IsATap(), n)

    # Pass 4: remaining non-welds that are still blank (catch-all)
    n = renumber_bucket(parts, lambda q: (not is_pipe_weld(q)) and is_blank_param(get_param(q, TARGET_PARAM)), n)

    # Pass 5: welds
    n = renumber_bucket(parts, lambda q: is_pipe_weld(q), n)

    t.Commit()

forms.alert("Renumbered {} parts into '{}'.".format(len(parts), TARGET_PARAM))
