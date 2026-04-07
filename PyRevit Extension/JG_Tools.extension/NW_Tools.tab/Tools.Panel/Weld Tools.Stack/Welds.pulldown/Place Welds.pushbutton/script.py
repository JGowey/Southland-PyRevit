# ===========================
# SCRIPT 1/2 PLACE MISSING WELD MARKERS (IDEMPOTENT + HOST IDS)
# ===========================
# -*- coding: utf-8 -*-
"""
Missing Weld Marker Placer (Pipe Weld)

- Finds missing-weld joints between connected ITMs/RFAs in your selection
- Places Pipe Accessory family "Pipe Weld" at each missing weld location
- Size rules:
    * NON-TAP: use Revit "Outside Diameter" from participants (max)
    * TAP/OLET: size from CROSS-SECTION bbox OD (bbox measured perp to branch/flow axis)
- TAP/OLET placement:
    * place at run outside surface along BRANCH axis (tap meets run)
- Orientation:
    * try axis X/Y/Z + roll 0/90/180/270; keep best by bbox thickness along target axis
- NO OVER-PLACEMENT on reruns:
    * stamps stable key CP_WeldJointKey and skips if key already exists
- HOST LOCK DATA:
    * stamps CP_WeldHostIds = comma list of participant ElementIds at the joint

Requires Weld marker family instance shared params:
- CP_WeldJointKey (Text)
- CP_WeldHostIds  (Text)
"""

from __future__ import print_function

import clr

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FabricationPart, ElementId, Family, FamilySymbol, FilteredElementCollector,
    BuiltInCategory, XYZ, Line, Transaction, ElementTransformUtils, IFamilyLoadOptions, FamilySource
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ObjectType
import math
import os
from System.Collections.Generic import List

output = None  # debug output disabled
doc = revit.doc
uidoc = revit.uidoc

# ---------------- CONFIG ----------------
LEN_TOL_FT = 0.005208333          # ~1/16" joint clustering
KEY_TOL_FT = 0.001302083          # ~1/64" key rounding (prevents close-joint collisions)
DUP_TOL_FT = 0.10                 # ~1.2" duplicate check radius (secondary safety)

SIZE_OFFSET_FT = 1.0 / 96.0       # +1/8" in feet (visual brim)

PARAM_PRODUCT_NAME = "Product Name"
PRODUCT_WELDS_VALUE = "Welds"

WELD_RFA_FAMILY_NAME = "Pipe Weld"
WELD_RFA_TYPE_NAME = ""           # "" = first type found
WELD_SIZE_PARAM = "CP_Diameter"

# If the weld family is not loaded, the script will try to load it from this .RFA
# Place the RFA in the same folder as this script (the pushbutton folder).
WELD_RFA_FILENAME = "Pipe Weld.rfa"

# Must exist as INSTANCE Text params on the weld RFA
WELD_KEY_PARAM = "CP_WeldJointKey"
WELD_HOSTIDS_PARAM = "CP_WeldHostIds"

# Tap/o-let recognition
TAP_KEYWORDS = ["branch connection", "branch connections", "weldolet", "o-let", "olet", "tap", "fishmouth"]

AXES = ["X", "Y", "Z"]
ROLLS_DEG = [0.0, 90.0, 180.0, 270.0]
# --------------------------------------


def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None


def get_param_valstr(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return ""
    try:
        return p.AsValueString() or ""
    except:
        return ""


def get_param_str(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return ""
    try:
        return p.AsString() or ""
    except:
        return ""


def get_param_len_ft(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return None
    try:
        return float(p.AsDouble())
    except:
        return None


def display_name(el):
    try:
        nm = getattr(el, "ItemName", None)
        if nm:
            return str(nm)
    except:
        pass
    try:
        n = getattr(el, "Name", None)
        if n:
            return str(n)
    except:
        pass
    try:
        pn = (get_param_valstr(el, PARAM_PRODUCT_NAME) or get_param_str(el, PARAM_PRODUCT_NAME) or "").strip()
        if pn:
            return pn
    except:
        pass
    try:
        return el.GetType().Name
    except:
        return "<Unknown>"


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


def dist(a, b):
    try:
        return a.DistanceTo(b)
    except:
        return 1e9


def unitize(v):
    try:
        l = v.GetLength()
        if l < 1e-9:
            return None
        return v / l
    except:
        return None


def round_key(xyz):
    rx = int(round(xyz.X / KEY_TOL_FT))
    ry = int(round(xyz.Y / KEY_TOL_FT))
    rz = int(round(xyz.Z / LEN_TOL_FT))
    return (rx, ry, rz)


def xyz_str(xyz):
    try:
        return "({:.4f}, {:.4f}, {:.4f})".format(xyz.X, xyz.Y, xyz.Z)
    except:
        return "(?)"


def get_level_for_point(pt):
    try:
        lv = doc.ActiveView.GenLevel
        if lv:
            return lv
    except:
        pass
    return None



class _FamilyLoadOpts(IFamilyLoadOptions):
    # Always allow load; overwrite parameter values when reloading.
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        overwriteParameterValues.Value = True
        source.Value = FamilySource.Family
        return True


def _find_family_by_name(fam_name):
    try:
        for f in FilteredElementCollector(doc).OfClass(Family):
            try:
                if f and f.Name == fam_name:
                    return f
            except:
                pass
    except:
        pass
    return None


def _find_symbol_from_family(fam):
    if not fam:
        return None
    try:
        type_ids = list(fam.GetFamilySymbolIds())
    except:
        type_ids = []
    if not type_ids:
        return None

    # Prefer specified type name if provided
    if WELD_RFA_TYPE_NAME:
        for tid in type_ids:
            try:
                sym = doc.GetElement(tid)
                if sym and sym.Name == WELD_RFA_TYPE_NAME:
                    return sym
            except:
                continue

    # Otherwise first type
    try:
        return doc.GetElement(type_ids[0])
    except:
        return None


def ensure_weld_symbol_loaded():
    """
    Returns a FamilySymbol for the weld marker.
    - If already loaded: returns it.
    - Else: loads the .RFA from the pushbutton folder and returns it.
    Supports older-year RFAs by allowing Revit to upgrade on load.
    """
    # 1) already loaded?
    fam = _find_family_by_name(WELD_RFA_FAMILY_NAME)
    sym = _find_symbol_from_family(fam) if fam else None
    if sym:
        return sym

    # 2) load from script folder
    script_dir = os.path.dirname(__file__)
    rfa_path = os.path.join(script_dir, WELD_RFA_FILENAME)
    if not os.path.exists(rfa_path):
        return None

    opts = _FamilyLoadOpts()
    fam_ref = clr.Reference[Family]()
    ok = False

    t = Transaction(doc, "Load weld family")
    t.Start()
    try:
        # Prefer overload that returns the loaded Family reference (more reliable).
        try:
            ok = doc.LoadFamily(rfa_path, opts, fam_ref)
        except:
            ok = doc.LoadFamily(rfa_path)

        t.Commit()
    except:
        try:
            t.RollBack()
        except:
            pass

        # Fallback: open the family document, then load into project (handles some upgrade edge cases)
        try:
            app = doc.Application
            fam_doc = app.OpenDocumentFile(rfa_path)
            t2 = Transaction(doc, "Load weld family (fallback)")
            t2.Start()
            try:
                ok = fam_doc.LoadFamily(doc, opts, fam_ref)
                t2.Commit()
            except:
                try:
                    t2.RollBack()
                except:
                    pass
            try:
                fam_doc.Close(False)
            except:
                pass
        except:
            ok = False

    # 3) re-find symbol
    fam = _find_family_by_name(WELD_RFA_FAMILY_NAME)
    sym = _find_symbol_from_family(fam) if fam else None
    return sym


def find_weld_symbol():
    # Backward-compatible wrapper
    return ensure_weld_symbol_loaded()


    col = (FilteredElementCollector(doc)
           .OfClass(FamilySymbol)
           .OfCategory(BuiltInCategory.OST_PipeAccessory))
    symbols = []
    for s in col:
        try:
            fam = s.Family
            if fam and fam.Name == WELD_RFA_FAMILY_NAME:
                if WELD_RFA_TYPE_NAME:
                    if s.Name == WELD_RFA_TYPE_NAME:
                        return s
                else:
                    symbols.append(s)
        except:
            continue
    if not symbols:
        # try to load the family from disk if it isn't in the model yet
        ensure_weld_family_loaded()
        # re-query
        col = (FilteredElementCollector(doc)
               .OfClass(FamilySymbol)
               .OfCategory(BuiltInCategory.OST_PipeAccessory))
        for s in col:
            try:
                fam = s.Family
                if fam and fam.Name == WELD_RFA_FAMILY_NAME:
                    if WELD_RFA_TYPE_NAME:
                        if s.Name == WELD_RFA_TYPE_NAME:
                            return s
                    else:
                        symbols.append(s)
            except:
                continue

    return symbols[0] if symbols else None


def marker_exists_near(pt, existing_pts):
    for (_, ipt) in existing_pts:
        if dist(pt, ipt) <= DUP_TOL_FT:
            return True
    return False


def bbox_thickness_along_axis(el, axis_dir):
    axis_dir = unitize(axis_dir)
    if not axis_dir:
        return 1e18
    bb = el.get_BoundingBox(None)
    if not bb:
        return 1e18

    mn = bb.Min
    mx = bb.Max
    corners = [
        XYZ(mn.X, mn.Y, mn.Z), XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mn.Z), XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mn.Y, mn.Z), XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mx.X, mx.Y, mn.Z), XYZ(mx.X, mx.Y, mx.Z),
    ]
    projs = [p.DotProduct(axis_dir) for p in corners]
    return max(projs) - min(projs)


def rotate_align(inst, pt, axis_name, target_dir):
    tr = inst.GetTransform()
    if axis_name == "X":
        base = tr.BasisX
    elif axis_name == "Y":
        base = tr.BasisY
    else:
        base = tr.BasisZ

    base = unitize(base)
    target_dir = unitize(target_dir)
    if (not base) or (not target_dir):
        return

    dot = max(-1.0, min(1.0, base.DotProduct(target_dir)))
    angle = math.acos(dot)
    if angle <= 1e-6:
        return

    axis_dir = unitize(base.CrossProduct(target_dir))
    if not axis_dir:
        return

    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, axis_dir), angle)


def rotate_roll(inst, pt, axis_dir, roll_deg):
    if abs(roll_deg) < 1e-9:
        return
    axis_dir = unitize(axis_dir)
    if not axis_dir:
        return
    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, axis_dir), math.radians(roll_deg))


def place_instance(pt, sym):
    """Place at exact XYZ in model coords.
    In plan/section views, using the Level overload can force Z to the view's level.
    So: try a no-level overload first, then fall back to level-based placement and set offset.
    """
    inst = None

    # 1) Prefer placing directly at XYZ (keeps Z correct)
    try:
        inst = doc.Create.NewFamilyInstance(pt, sym, StructuralType.NonStructural)
        return inst
    except:
        inst = None

    # 2) Fallback: level-based overload, then push offset to match pt.Z
    lvl = get_level_for_point(pt)
    try:
        if lvl:
            inst = doc.Create.NewFamilyInstance(pt, sym, lvl, StructuralType.NonStructural)
        else:
            inst = doc.Create.NewFamilyInstance(pt, sym, StructuralType.NonStructural)
    except:
        try:
            if lvl:
                inst = doc.Create.NewFamilyInstance(pt, sym, lvl)
            else:
                inst = doc.Create.NewFamilyInstance(pt, sym)
        except:
            inst = None

    # If level-based, attempt to set offset so the instance lands at pt.Z
    if inst and lvl:
        try:
            # Common param names/ids for offset vary by family;
            # try built-in first then name.
            p_off = inst.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
            if p_off and (not p_off.IsReadOnly):
                p_off.Set(pt.Z - lvl.Elevation)
            else:
                p2 = inst.LookupParameter("Offset")
                if p2 and (not p2.IsReadOnly):
                    p2.Set(pt.Z - lvl.Elevation)
        except:
            pass

    return inst


def solve_orientation_keep_best(base_inst, pt, target_axis):
    candidates = [base_inst]
    total = len(AXES) * len(ROLLS_DEG)

    for _ in range(total - 1):
        try:
            new_ids = ElementTransformUtils.CopyElement(doc, base_inst.Id, XYZ(0, 0, 0))
            if new_ids and new_ids.Count > 0:
                candidates.append(doc.GetElement(new_ids[0]))
        except:
            pass

    best_inst = None
    best_score = 1e18

    idx = 0
    for ax in AXES:
        for roll in ROLLS_DEG:
            if idx >= len(candidates):
                continue
            inst = candidates[idx]
            idx += 1
            if not inst:
                continue
            try:
                rotate_align(inst, pt, ax, target_axis)
                rotate_roll(inst, pt, target_axis, roll)
            except:
                continue

            sc = bbox_thickness_along_axis(inst, target_axis)
            if sc < best_score:
                best_score = sc
                best_inst = inst

    for inst in candidates:
        if inst and best_inst and inst.Id != best_inst.Id:
            try:
                doc.Delete(inst.Id)
            except:
                pass

    return best_inst if best_inst else base_inst


def is_tap_like(el):
    if not isinstance(el, FabricationPart):
        return False
    try:
        if el.IsATap():
            return True
    except:
        pass
    pn = (get_param_valstr(el, PARAM_PRODUCT_NAME) or get_param_str(el, PARAM_PRODUCT_NAME) or "").lower()
    for k in TAP_KEYWORDS:
        if k in pn:
            return True
    return False


def make_perp_basis(axis_dir):
    w = unitize(axis_dir)
    if not w:
        return (None, None, None)

    helper = XYZ(0, 0, 1)
    if abs(w.DotProduct(helper)) > 0.95:
        helper = XYZ(0, 1, 0)

    u = unitize(w.CrossProduct(helper))
    if not u:
        helper = XYZ(1, 0, 0)
        u = unitize(w.CrossProduct(helper))
        if not u:
            return (None, None, None)

    v = unitize(w.CrossProduct(u))
    if not v:
        return (None, None, None)

    return (u, v, w)


def cross_section_od_from_bbox(el, axis_origin, axis_dir):
    u, v, _ = make_perp_basis(axis_dir)
    if not u or not v:
        return None

    bb = el.get_BoundingBox(None)
    if not bb:
        return None

    mn = bb.Min
    mx = bb.Max
    corners = [
        XYZ(mn.X, mn.Y, mn.Z), XYZ(mn.X, mn.Y, mx.Z),
        XYZ(mn.X, mx.Y, mn.Z), XYZ(mn.X, mx.Y, mx.Z),
        XYZ(mx.X, mn.Y, mn.Z), XYZ(mx.X, mn.Y, mx.Z),
        XYZ(mx.X, mx.Y, mn.Z), XYZ(mx.X, mx.Y, mx.Z),
    ]

    u_vals, v_vals = [], []
    for p in corners:
        try:
            d = p - axis_origin
            u_vals.append(d.DotProduct(u))
            v_vals.append(d.DotProduct(v))
        except:
            continue

    if not u_vals or not v_vals:
        return None

    du = max(u_vals) - min(u_vals)
    dv = max(v_vals) - min(v_vals)
    od = max(du, dv)
    return od if od > 1e-9 else None


def compute_tap_face_point_and_axis_and_visual_od(tap_part, run_part):
    tap_conns = get_connectors(tap_part)
    if len(tap_conns) < 2:
        return (None, None, None)

    run_id = run_part.Id.IntegerValue

    run_conn = None
    for c in tap_conns:
        try:
            owners = connected_owner_ids(c)
            if run_id in owners:
                run_conn = c
                break
        except:
            continue
    if not run_conn:
        return (None, None, None)

    branch_conn = None
    for c in tap_conns:
        if c != run_conn:
            branch_conn = c
            break
    if not branch_conn:
        return (None, None, None)

    axis_dir = unitize(branch_conn.Origin - run_conn.Origin)
    if not axis_dir:
        return (None, None, None)

    run_od = get_param_len_ft(run_part, "Outside Diameter") or get_param_len_ft(run_part, "OutsideDiameter")
    if not run_od or run_od <= 1e-9:
        return (None, None, None)

    place_pt = run_conn.Origin + (axis_dir * (run_od / 2.0))
    vis_od = cross_section_od_from_bbox(tap_part, run_conn.Origin, axis_dir)

    return (place_pt, axis_dir, vis_od)


# ---------------- KEY / HOST-ID STAMPING ----------------
def joint_key_from(pt, participant_ids):
    rk = round_key(pt)
    pid = "-".join([str(i) for i in sorted(participant_ids)])
    return "J:{}:{}:{}|P:{}".format(rk[0], rk[1], rk[2], pid)


def host_ids_str_from(participant_ids):
    return ",".join([str(i) for i in sorted(participant_ids)])


def collect_existing_marker_keys_and_pts():
    keys = set()
    legacy_pts = []

    col = (FilteredElementCollector(doc)
           .OfCategory(BuiltInCategory.OST_PipeAccessory)
           .WhereElementIsNotElementType())

    for e in col:
        try:
            if not (e.Symbol and e.Symbol.Family and e.Symbol.Family.Name == WELD_RFA_FAMILY_NAME):
                continue
        except:
            continue

        has_key = False
        try:
            p = e.LookupParameter(WELD_KEY_PARAM)
            if p and p.HasValue:
                k = (p.AsString() or "").strip()
                if k:
                    keys.add(k)
                    has_key = True
        except:
            pass

        # Only use proximity-duplicate logic for legacy markers that don't have a key stamped.
        if not has_key:
            try:
                legacy_pts.append((e, e.Location.Point))
            except:
                pass

    return keys, legacy_pts


def stamp_text_param(inst, pname, value):
    try:
        p = inst.LookupParameter(pname)
        if p and (not p.IsReadOnly):
            p.Set(str(value))
            return True
    except:
        pass
    return False

# ---------------- MAIN ----------------

# Preselect OR postselect:
# - If you preselect elements and hit the button, we use that selection.
# - If nothing is selected, we prompt you to pick scope elements.
pre_ids = list(uidoc.Selection.GetElementIds())
if pre_ids and len(pre_ids) > 0:
    picked = [doc.GetElement(eid) for eid in pre_ids]
else:
    try:
        refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select scope (ITMs + RFAs).")
    except:
        forms.alert("Selection cancelled.", exitscript=True)
    picked = [doc.GetElement(r) for r in refs if r]

picked = [e for e in picked if e is not None]
picked_ids = set([e.Id.IntegerValue for e in picked])
picked_ids = set([e.Id.IntegerValue for e in picked])

# Build joints from selection scope
joints = {}  # key -> {"pt": XYZ, "elem_ids": set(int)}
def joint_add(key, pt, eid_int):
    if key not in joints:
        joints[key] = {"pt": pt, "elem_ids": set()}
    joints[key]["elem_ids"].add(eid_int)

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

# Find missing joints (no weld ITM already present)
missing = []
rows = []
for key, data in joints.items():
    pt = data["pt"]
    elem_ids = sorted(list(data["elem_ids"]))
    elems = [doc.GetElement(ElementId(i)) for i in elem_ids]
    elems = [e for e in elems if e is not None]

    welds = [e for e in elems if isinstance(e, FabricationPart) and is_weld_itm(e)]
    has_weld = len(welds) > 0

    participants = [e for e in elems if not (isinstance(e, FabricationPart) and is_weld_itm(e))]
    if (len(participants) >= 2) and (not has_weld):
        missing.append({"pt": pt, "participants": participants, "participant_ids": elem_ids})
        rows.append([xyz_str(pt), ", ".join(["{}({})".format(display_name(e), e.Id.IntegerValue) for e in participants[:8]])])

# Debug report suppressed

if not missing:
    forms.alert("No missing-weld joints found. Nothing to place.", warn_icon=False)
    script.exit()

sym = find_weld_symbol()
if not sym:
    forms.alert("Could not find Pipe Accessory symbol for '{}'.".format(WELD_RFA_FAMILY_NAME), exitscript=True)

existing_marker_keys, legacy_marker_pts = collect_existing_marker_keys_and_pts()

placed = 0
skipped_dupe = 0
skipped_err = 0
added_to_asm = 0

t = Transaction(doc, "Place Pipe Weld markers at missing weld joints (host ids + key)")
t.Start()
try:
    if not sym.IsActive:
        sym.Activate()
        doc.Regenerate()

    for j in missing:
        base_pt = j["pt"]
        participants = j["participants"]
        participant_ids = j["participant_ids"]

        jkey = joint_key_from(base_pt, participant_ids)
        hostids_str = host_ids_str_from(participant_ids)

        # primary idempotence
        if jkey in existing_marker_keys:
            skipped_dupe += 1
            continue

        place_pt = base_pt
        target_axis = None
        size_od = None

        # TAP/OLET handling
        tap_part = None
        for e in participants:
            if is_tap_like(e):
                tap_part = e
                break

        if tap_part:
            # Choose run part as the largest Revit Outside Diameter participant besides tap
            run_part = None
            run_best = 0.0
            for e in participants:
                if e.Id.IntegerValue == tap_part.Id.IntegerValue:
                    continue
                od = get_param_len_ft(e, "Outside Diameter") or get_param_len_ft(e, "OutsideDiameter") or 0.0
                if od > run_best:
                    run_best = od
                    run_part = e

            if run_part:
                p2, ax2, vis_od = compute_tap_face_point_and_axis_and_visual_od(tap_part, run_part)
                if p2 and ax2:
                    place_pt = p2
                    target_axis = ax2
                    size_od = vis_od

        # Non-tap fallback axis: derive from any element with 2+ connectors (near->far)
        if not target_axis:
            best_dir = None
            best_len = 0.0
            for el in participants:
                conns = get_connectors(el)
                if len(conns) < 2:
                    continue
                near_c = None
                near_d = 1e9
                for c in conns:
                    try:
                        d = dist(c.Origin, base_pt)
                        if d < near_d:
                            near_d = d
                            near_c = c
                    except:
                        continue
                if not near_c or near_d > (LEN_TOL_FT * 3.0):
                    continue
                far_c = None
                far_d = 0.0
                for c in conns:
                    if c == near_c:
                        continue
                    try:
                        d = dist(c.Origin, near_c.Origin)
                        if d > far_d:
                            far_d = d
                            far_c = c
                    except:
                        continue
                if far_c and far_d > best_len:
                    v = unitize(far_c.Origin - near_c.Origin)
                    if v:
                        best_dir = v
                        best_len = far_d
            target_axis = best_dir

        if not target_axis:
            skipped_err += 1
            continue

        # Sizing:
        # - tap: keep bbox OD; if missing -> non-tap sizing
        # - non-tap: max Revit Outside Diameter of participants
        if size_od is None or (not size_od) or size_od <= 1e-9:
            max_od = 0.0
            for e in participants:
                od = get_param_len_ft(e, "Outside Diameter") or get_param_len_ft(e, "OutsideDiameter") or 0.0
                if od > max_od:
                    max_od = od
            size_od = max_od if max_od > 1e-9 else None

        if not size_od:
            skipped_err += 1
            continue

        # secondary duplicate safety (older markers without key)
        if marker_exists_near(place_pt, legacy_marker_pts):
            skipped_dupe += 1
            continue

        base = place_instance(place_pt, sym)
        if not base:
            skipped_err += 1
            continue

        # Size
        target_cp = size_od + SIZE_OFFSET_FT
        p = base.LookupParameter(WELD_SIZE_PARAM)
        if p and (not p.IsReadOnly) and target_cp > 0.0:
            try:
                p.Set(target_cp)
            except:
                pass

        best_inst = solve_orientation_keep_best(base, place_pt, target_axis)

        # Stamp key + host ids
        stamp_text_param(best_inst, WELD_KEY_PARAM, jkey)
        stamp_text_param(best_inst, WELD_HOSTIDS_PARAM, hostids_str)
        # If participants are already in an Assembly, add the new marker to that Assembly
        asm_id = None
        for e in participants:
            try:
                aid = e.AssemblyInstanceId
                if aid and aid != ElementId.InvalidElementId:
                    asm_id = aid
                    break
            except:
                continue
        if asm_id:
            try:
                asm = doc.GetElement(asm_id)
                if asm:
                    ids = List[ElementId]()
                    ids.Add(best_inst.Id)
                    asm.AddMemberIds(ids)
                    added_to_asm += 1
            except:
                pass


        existing_marker_keys.add(jkey)

        placed += 1

finally:
    t.Commit()

forms.alert(
    "Placed: {0}\nAdded to assembly: {1}\nSkipped (already had marker): {2}\nSkipped (error): {3}".format(
        placed, added_to_asm, skipped_dupe, skipped_err
    ),
    warn_icon=False
)