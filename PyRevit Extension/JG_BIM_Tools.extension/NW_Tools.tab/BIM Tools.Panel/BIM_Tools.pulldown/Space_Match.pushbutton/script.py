# -*- coding: utf-8 -*-
from pyrevit import revit, forms
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, StorageType, Transaction
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType


doc = revit.doc
uidoc = revit.uidoc


# ----------------------------
# Safe BuiltInParameter getter
# ----------------------------
def bip(name):
    """Return BuiltInParameter.<name> if it exists in this Revit version, else None."""
    return getattr(BuiltInParameter, name, None)

def bips(*names):
    """Return list of existing BIPs from the given enum names."""
    out = []
    for n in names:
        v = bip(n)
        if v is not None:
            out.append(v)
    return out


# ----------------------------
# Helpers
# ----------------------------
def _try_get_param(elem, bip_candidates=None, name_candidates=None):
    """Return first found Parameter (not value) from BIPs or by name."""
    if bip_candidates:
        for bp in bip_candidates:
            try:
                p = elem.get_Parameter(bp)
                if p:
                    return p
            except:
                pass
    if name_candidates:
        for n in name_candidates:
            try:
                p = elem.LookupParameter(n)
                if p:
                    return p
            except:
                pass
    return None


def _get_param_value(elem, bip_candidates=None, name_candidates=None):
    p = _try_get_param(elem, bip_candidates=bip_candidates, name_candidates=name_candidates)
    if not p:
        return None
    try:
        st = p.StorageType
        if st == StorageType.String:
            return p.AsString()
        elif st == StorageType.Double:
            return p.AsDouble()
        elif st == StorageType.Integer:
            return p.AsInteger()
        elif st == StorageType.ElementId:
            return p.AsElementId()
    except:
        return None
    return None


def _set_param_value(elem, value, bip_candidates=None, name_candidates=None):
    p = _try_get_param(elem, bip_candidates=bip_candidates, name_candidates=name_candidates)
    if not p or p.IsReadOnly:
        return False
    try:
        st = p.StorageType
        if st == StorageType.String:
            p.Set("" if value is None else str(value))
            return True
        elif st == StorageType.Double:
            if value is None:
                return False
            p.Set(float(value))
            return True
        elif st == StorageType.Integer:
            if value is None:
                return False
            p.Set(int(value))
            return True
        elif st == StorageType.ElementId:
            if value is None:
                return False
            p.Set(value)
            return True
    except:
        return False
    return False


def _pick_number(prompt, default=0.0):
    """Ask for a numeric value. Interpreted as Revit internal units (feet)."""
    s = forms.ask_for_string(default=str(default), prompt=prompt, title="Input")
    if s is None:
        return None
    try:
        return float(s)
    except:
        forms.alert("Invalid number: {}".format(s), exitscript=True)


# ----------------------------
# Selection Filters
# ----------------------------
class LevelOrSpaceFilter(ISelectionFilter):
    def AllowElement(self, e):
        if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_Levels):
            return True
        if e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces):
            return True
        return False
    def AllowReference(self, ref, pt):
        return True


class SpaceOnlyFilter(ISelectionFilter):
    def AllowElement(self, e):
        return e.Category and e.Category.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces)
    def AllowReference(self, ref, pt):
        return True


# ----------------------------
# Parameter mappings
# (Prefer ROOM_* because SPACE_* may not exist in your API)
# ----------------------------

# Identity: Name
NAME_BIPS = bips("ROOM_NAME", "SPACE_NAME")  # SPACE_NAME might not exist
NAME_NAMES = ["Name"]  # fallback

# Constraints: Associated/Base Level
BASE_LEVEL_BIPS = bips("ROOM_LEVEL_ID", "SPACE_LEVEL_ID")
BASE_LEVEL_NAMES = ["Level", "Base Level"]  # fallback labels vary

# Constraints: Upper Limit Level
UPPER_LEVEL_BIPS = bips("ROOM_UPPER_LEVEL", "SPACE_UPPER_LEVEL")
UPPER_LEVEL_NAMES = ["Upper Limit"]

# Constraints: Offsets
BASE_OFFSET_BIPS = bips("ROOM_LOWER_OFFSET", "SPACE_LOWER_OFFSET")
BASE_OFFSET_NAMES = ["Base Offset"]

LIMIT_OFFSET_BIPS = bips("ROOM_UPPER_OFFSET", "SPACE_UPPER_OFFSET")
LIMIT_OFFSET_NAMES = ["Limit Offset"]

# Shared parameter candidates for "Base Name"
BASE_NAME_NAMES = [
    "Base Name", "BASE NAME", "BaseName", "BASE_NAME",
    "BASE NAME - IDENTITY", "Base Name (Identity Data)"
]


# ----------------------------
# Main
# ----------------------------
template_ref = uidoc.Selection.PickObject(
    ObjectType.Element,
    LevelOrSpaceFilter(),
    "Pick a TEMPLATE: a Level or a Space"
)
template_elem = doc.GetElement(template_ref.ElementId)

template_is_level = template_elem.Category and template_elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_Levels)
template_is_space = template_elem.Category and template_elem.Category.Id.IntegerValue == int(BuiltInCategory.OST_MEPSpaces)

if not (template_is_level or template_is_space):
    forms.alert("Template must be a Level or a Space.", exitscript=True)

# Defaults
t_name = None
t_base_name = None
t_base_level_id = None
t_upper_level_id = None
t_base_offset = None
t_limit_offset = None

if template_is_space:
    t_name = _get_param_value(template_elem, bip_candidates=NAME_BIPS, name_candidates=NAME_NAMES)
    t_base_name = _get_param_value(template_elem, name_candidates=BASE_NAME_NAMES)

    t_base_level_id = _get_param_value(template_elem, bip_candidates=BASE_LEVEL_BIPS, name_candidates=BASE_LEVEL_NAMES)
    t_upper_level_id = _get_param_value(template_elem, bip_candidates=UPPER_LEVEL_BIPS, name_candidates=UPPER_LEVEL_NAMES)

    t_base_offset = _get_param_value(template_elem, bip_candidates=BASE_OFFSET_BIPS, name_candidates=BASE_OFFSET_NAMES)
    t_limit_offset = _get_param_value(template_elem, bip_candidates=LIMIT_OFFSET_BIPS, name_candidates=LIMIT_OFFSET_NAMES)

else:
    # Template is a Level: prompt for everything
    levels = sorted(
        [lvl for lvl in revit.query.get_elements_by_category(BuiltInCategory.OST_Levels) if lvl],
        key=lambda x: x.Elevation
    )
    level_names = ["{} (elev {:.3f})".format(l.Name, l.Elevation) for l in levels]

    pick = forms.SelectFromList.show(level_names, title="Pick Upper Limit Level", multiselect=False)
    if pick is None:
        forms.alert("Cancelled.", exitscript=True)

    upper_level = levels[level_names.index(pick)]

    t_base_level_id = template_elem.Id
    t_upper_level_id = upper_level.Id

    t_base_offset = _pick_number("Base Offset (internal units = feet). Example: 0, 1.5, -0.25", default=0.0)
    if t_base_offset is None:
        forms.alert("Cancelled.", exitscript=True)

    t_limit_offset = _pick_number("Limit Offset (internal units = feet). Example: 0, 1.5, -0.25", default=0.0)
    if t_limit_offset is None:
        forms.alert("Cancelled.", exitscript=True)

    t_name = forms.ask_for_string(default="", prompt="Space Name to apply (Identity Data → Name)", title="Input")
    if t_name is None:
        forms.alert("Cancelled.", exitscript=True)

    t_base_name = forms.ask_for_string(default="", prompt='Base Name to apply (parameter named "Base Name")', title="Input")
    if t_base_name is None:
        forms.alert("Cancelled.", exitscript=True)

# Pick target spaces
target_refs = uidoc.Selection.PickObjects(
    ObjectType.Element,
    SpaceOnlyFilter(),
    "Pick TARGET Spaces to match the template"
)
targets = [doc.GetElement(r.ElementId) for r in target_refs]

if not targets:
    forms.alert("No target spaces selected.", exitscript=True)

changed = 0
failed = 0

with Transaction(doc, "Match Space Parameters") as t:
    t.Start()
    for sp in targets:
        ok_any = False

        # Identity: Name
        if t_name is not None:
            ok_any = _set_param_value(sp, t_name, bip_candidates=NAME_BIPS, name_candidates=NAME_NAMES) or ok_any

        # Identity: Base Name (shared)
        if t_base_name is not None:
            ok_any = _set_param_value(sp, t_base_name, name_candidates=BASE_NAME_NAMES) or ok_any

        # Constraints: Base/Associated Level
        if t_base_level_id is not None:
            ok_any = _set_param_value(sp, t_base_level_id, bip_candidates=BASE_LEVEL_BIPS, name_candidates=BASE_LEVEL_NAMES) or ok_any

        # Constraints: Upper Limit Level
        if t_upper_level_id is not None:
            ok_any = _set_param_value(sp, t_upper_level_id, bip_candidates=UPPER_LEVEL_BIPS, name_candidates=UPPER_LEVEL_NAMES) or ok_any

        # Constraints: Base Offset
        if t_base_offset is not None:
            ok_any = _set_param_value(sp, t_base_offset, bip_candidates=BASE_OFFSET_BIPS, name_candidates=BASE_OFFSET_NAMES) or ok_any

        # Constraints: Limit Offset
        if t_limit_offset is not None:
            ok_any = _set_param_value(sp, t_limit_offset, bip_candidates=LIMIT_OFFSET_BIPS, name_candidates=LIMIT_OFFSET_NAMES) or ok_any

        if ok_any:
            changed += 1
        else:
            failed += 1

    t.Commit()

forms.alert("Done.\nChanged: {}\nNo writable matches: {}".format(changed, failed))
