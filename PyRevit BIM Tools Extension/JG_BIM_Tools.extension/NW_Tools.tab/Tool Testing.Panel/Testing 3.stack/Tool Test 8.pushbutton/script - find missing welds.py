# -*- coding: utf-8 -*-
from __future__ import print_function

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import FabricationPart, ElementId
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc

# ---------------- CONFIG ----------------
LEN_TOL_FT = 0.005208333  # ~1/16" clustering
PARAM_PRODUCT_NAME = "Product Name"
PRODUCT_WELDS_VALUE = "Welds"
# ---------------------------------------


def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None

def get_param_str(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return ""
    try:
        return p.AsString() or ""
    except:
        return ""

def get_param_valstr(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return ""
    try:
        return p.AsValueString() or ""
    except:
        return ""

def display_name(el):
    if isinstance(el, FabricationPart):
        try:
            nm = getattr(el, "ItemName", None)
            if nm:
                return str(nm)
        except:
            pass
        pn = get_param_valstr(el, PARAM_PRODUCT_NAME) or get_param_str(el, PARAM_PRODUCT_NAME)
        if pn:
            return pn
    try:
        return el.Name or el.GetType().Name
    except:
        return el.GetType().Name

def is_weld_itm(fp):
    pn = (get_param_valstr(fp, PARAM_PRODUCT_NAME) or get_param_str(fp, PARAM_PRODUCT_NAME) or "").strip()
    return pn.lower() == PRODUCT_WELDS_VALUE.lower()

def get_connectors(el):
    try:
        cm = el.ConnectorManager
        if cm:
            return list(cm.Connectors)
    except:
        pass
    try:
        mep = el.MEPModel
        if mep and mep.ConnectorManager:
            return list(mep.ConnectorManager.Connectors)
    except:
        pass
    return []

def connected_owner_ids(conn):
    ids = set()
    try:
        for c in conn.AllRefs:
            try:
                o = c.Owner
                if o and o.Id != conn.Owner.Id:
                    ids.add(o.Id.IntegerValue)
            except:
                continue
    except:
        pass
    return ids

def round_key(xyz):
    rx = int(round(xyz.X / LEN_TOL_FT))
    ry = int(round(xyz.Y / LEN_TOL_FT))
    rz = int(round(xyz.Z / LEN_TOL_FT))
    return (rx, ry, rz)

def xyz_str(xyz):
    try:
        return "({:.4f}, {:.4f}, {:.4f})".format(xyz.X, xyz.Y, xyz.Z)
    except:
        return "(?)"


# ---------------- MAIN ----------------
try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select scope (ITMs + RFAs).")
except:
    forms.alert("Selection cancelled.", exitscript=True)

picked = [doc.GetElement(r) for r in refs if r]
picked = [e for e in picked if e is not None]
picked_ids = set([e.Id.IntegerValue for e in picked])

# joints[key] = {"pt": XYZ, "elem_ids": set(int)}
joints = {}

def joint_add(key, pt, eid_int):
    if key not in joints:
        joints[key] = {"pt": pt, "elem_ids": set()}
    joints[key]["elem_ids"].add(eid_int)

# Build joints from connector origins and AllRefs (within picked scope)
for el in picked:
    conns = get_connectors(el)
    if not conns:
        continue

    el_id = el.Id.IntegerValue
    for c in conns:
        other_ids = [oid for oid in connected_owner_ids(c) if oid in picked_ids]
        if not other_ids:
            continue

        try:
            pt = c.Origin
        except:
            continue

        key = round_key(pt)
        joint_add(key, pt, el_id)
        for oid in other_ids:
            joint_add(key, pt, oid)

if not joints:
    forms.alert("No connected joints found inside your selection.", exitscript=True)

rows = []
missing_joint_pts = []   # <--- you’ll use these later for placing RFAs
missing_elem_ids = set()

for _, data in joints.items():
    pt = data["pt"]
    elem_ids = sorted(list(data["elem_ids"]))
    elems = [doc.GetElement(ElementId(i)) for i in elem_ids]
    elems = [e for e in elems if e is not None]

    weld_ids = []
    for e in elems:
        if isinstance(e, FabricationPart) and is_weld_itm(e):
            weld_ids.append(e.Id.IntegerValue)
    has_weld = len(weld_ids) > 0

    participants = []
    for e in elems:
        if isinstance(e, FabricationPart) and is_weld_itm(e):
            continue
        participants.append(e.Id.IntegerValue)
    participants = sorted(set(participants))

    missing = (len(participants) >= 2) and (not has_weld)
    if missing:
        missing_joint_pts.append(pt)
        for pid in participants:
            if pid in picked_ids:
                missing_elem_ids.add(pid)

    part_str = ", ".join(["{}({})".format(display_name(doc.GetElement(ElementId(pid))), pid) for pid in participants[:10]])
    if len(participants) > 10:
        part_str += " ...(+{})".format(len(participants) - 10)

    rows.append([
        xyz_str(pt),
        "YES" if has_weld else "NO",
        ", ".join([str(x) for x in sorted(set(weld_ids))]),
        str(len(participants)),
        "YES" if missing else "NO",
        part_str
    ])

output.print_md("## Joint-Based Weld Identification (Product Name = 'Welds')")
output.print_md("- Elements in scope: **{}**".format(len(picked)))
output.print_md("- Physical joints found: **{}**".format(len(joints)))
output.print_md("- Missing-weld joints: **{}**".format(len(missing_joint_pts)))

output.print_table(
    table_data=rows,
    columns=["JointXYZ", "HasWeldITM", "WeldITM_IDs", "Participants", "MissingWeld", "Participants (name+id)"],
    title="Joints"
)

if missing_elem_ids:
    id_list = List[ElementId]()
    for i in sorted(missing_elem_ids):
        id_list.Add(ElementId(i))
    uidoc.Selection.SetElementIds(id_list)
    forms.alert("Missing-weld joints: {}\nSelected elements involved: {}"
                .format(len(missing_joint_pts), len(missing_elem_ids)),
                warn_icon=True)
else:
    forms.alert("No missing-weld joints found.", warn_icon=False)
