# -*- coding: utf-8 -*-
# ===========================
# SCRIPT 2/2 RESNAP EXISTING WELD MARKERS (USES CP WeldHostIds)
# ===========================
# -*- coding: utf-8 -*-
"""
Re-Snap Pipe Weld Markers (Button Tool)

- Scans active view FabricationParts and builds current connector joint map
- For each existing 'Pipe Weld' marker in the active view:
    * Reads CP_WeldHostIds (comma element ids)
    * Finds the current joint whose participant ids include those host ids
    * Recomputes correct placement + size (tap/olet logic supported)
    * Moves + resizes + (light) aligns

This is the "locked to companion pipes" behavior without an IUpdater.
"""

from __future__ import print_function

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FabricationPart, FilteredElementCollector, BuiltInCategory, XYZ, Line,
    Transaction, ElementTransformUtils
)
import math

output = script.get_output()
doc = revit.doc
av = doc.ActiveView

# ---------------- CONFIG ----------------
LEN_TOL_FT = 0.005208333          # same as placer
SIZE_OFFSET_FT = 1.0 / 96.0       # +1/8" brim
MOVE_EPS_FT = 1e-6

WELD_RFA_FAMILY_NAME = "Pipe Weld"
WELD_SIZE_PARAM = "CP_Diameter"
WELD_HOSTIDS_PARAM = "CP_WeldHostIds"

PARAM_PRODUCT_NAME = "Product Name"
PRODUCT_WELDS_VALUE = "Welds"
TAP_KEYWORDS = ["branch connection", "branch connections", "weldolet", "o-let", "olet", "tap", "fishmouth"]
# --------------------------------------


def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None


def get_param_len_ft(el, name):
    p = get_param(el, name)
    if not p or (not p.HasValue):
        return None
    try:
        return float(p.AsDouble())
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


def dist(a, b):
    try:
        return a.DistanceTo(b)
    except:
        return 1e9


def unitize(v):
    try:
        l = v.GetLength()
        if l < 1e-12:
            return None
        return v / l
    except:
        return None


def round_key(xyz):
    rx = int(round(xyz.X / LEN_TOL_FT))
    ry = int(round(xyz.Y / LEN_TOL_FT))
    rz = int(round(xyz.Z / LEN_TOL_FT))
    return (rx, ry, rz)


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


def is_weld_itm(fp):
    pn = (get_param_valstr(fp, PARAM_PRODUCT_NAME) or get_param_str(fp, PARAM_PRODUCT_NAME) or "").strip()
    return pn.lower() == PRODUCT_WELDS_VALUE.lower()


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


def rotate_align_Z(inst, pt, target_dir):
    try:
        tr = inst.GetTransform()
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
    except:
        pass



def project_onto_plane(v, plane_normal):
    """Component of v perpendicular to plane_normal."""
    n = unitize(plane_normal)
    if not n:
        return None
    return v - n.Multiply(v.DotProduct(n))


def rotate_align_axis(inst, pt, from_vec, to_vec):
    """Rotate inst so from_vec aligns to to_vec (shortest arc)."""
    from_vec = unitize(from_vec)
    to_vec = unitize(to_vec)
    if (not from_vec) or (not to_vec):
        return
    dot = max(-1.0, min(1.0, from_vec.DotProduct(to_vec)))
    ang = math.acos(dot)
    if ang <= 1e-6:
        return
    ax = unitize(from_vec.CrossProduct(to_vec))
    if not ax:
        return
    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, ax), ang)


def rotate_match_roll(inst, pt, axis_dir, target_roll_ref):
    """
    After Z has been aligned to axis_dir, rotate about axis_dir so the *closest* of (X or Y)
    matches target_roll_ref (projected into plane perp axis_dir).
    This avoids consistent 90deg offsets when the family's 'roll' axis is Y instead of X.
    """
    axis_dir = unitize(axis_dir)
    if not axis_dir:
        return
    target_roll_ref = project_onto_plane(target_roll_ref, axis_dir) if target_roll_ref else None
    target_roll_ref = unitize(target_roll_ref) if target_roll_ref else None
    if not target_roll_ref:
        return

    tr = inst.GetTransform()

    def signed_angle_about_axis(from_vec, to_vec, axis):
        dot = max(-1.0, min(1.0, from_vec.DotProduct(to_vec)))
        ang = math.acos(dot)
        if ang <= 1e-6:
            return 0.0
        cross = from_vec.CrossProduct(to_vec)
        sign = -1.0 if cross.DotProduct(axis) < 0 else 1.0
        return ang * sign

    # Candidate current directions (projected into plane perp axis_dir)
    cur_x = project_onto_plane(tr.BasisX, axis_dir)
    cur_y = project_onto_plane(tr.BasisY, axis_dir)
    cur_x = unitize(cur_x) if cur_x else None
    cur_y = unitize(cur_y) if cur_y else None

    candidates = []
    if cur_x:
        candidates.append(("X", signed_angle_about_axis(cur_x, target_roll_ref, axis_dir)))
    if cur_y:
        candidates.append(("Y", signed_angle_about_axis(cur_y, target_roll_ref, axis_dir)))

    if not candidates:
        return

    # Choose the smaller rotation magnitude
    axis_name, ang = min(candidates, key=lambda t: abs(t[1]))
    if abs(ang) <= 1e-6:
        return

    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, axis_dir), ang)


def bbox_thickness_along_axis(inst, axis):
    # thickness of element's bounding box projected onto axis direction
    axis = unitize(axis)
    if not axis:
        return 1e18
    bb = inst.get_BoundingBox(None)
    if not bb:
        return 1e18
    pts = [
        XYZ(bb.Min.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Min.X, bb.Max.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Min.Y, bb.Max.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Min.Z),
        XYZ(bb.Max.X, bb.Max.Y, bb.Max.Z),
    ]
    projs = [p.DotProduct(axis) for p in pts]
    return max(projs) - min(projs)


def rotate_align(inst, pt, from_vec, to_vec):
    # rotate element so from_vec aligns to to_vec (both in model coords)
    from_vec = unitize(from_vec)
    to_vec = unitize(to_vec)
    if (not from_vec) or (not to_vec):
        return
    dot = max(-1.0, min(1.0, from_vec.DotProduct(to_vec)))
    ang = math.acos(dot)
    if ang <= 1e-6:
        return
    ax = unitize(from_vec.CrossProduct(to_vec))
    if not ax:
        return
    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, ax), ang)


def rotate_roll_deg(inst, pt, axis_dir, deg):
    axis_dir = unitize(axis_dir)
    if not axis_dir:
        return
    ang = math.radians(float(deg))
    if abs(ang) <= 1e-9:
        return
    ElementTransformUtils.RotateElement(doc, inst.Id, Line.CreateUnbound(pt, axis_dir), ang)


def solve_orientation_like_placer(base_inst, pt, target_axis, roll_ref):
    """
    Mirrors Script 1's behavior:
      - Try mapping family X/Y/Z to target_axis
      - Try roll 0/90/180/270 about target_axis
      - Keep the candidate with best (smallest) bbox thickness along target_axis,
        with a tie-breaker to best match roll_ref when available.
    Uses temporary copies and deletes the losers.
    """
    AXES = [XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ]
    ROLLS = [0, 90, 180, 270]

    roll_ref = unitize(project_onto_plane(roll_ref, target_axis)) if roll_ref else None

    total = len(AXES) * len(ROLLS)
    candidates = [base_inst]

    # make copies so we can test safely without accumulating rotations
    for _ in range(total - 1):
        try:
            new_ids = ElementTransformUtils.CopyElement(doc, base_inst.Id, XYZ(0, 0, 0))
            if new_ids and new_ids.Count > 0:
                candidates.append(doc.GetElement(new_ids[0]))
        except:
            pass

    best_inst = base_inst
    best_score = 1e18
    best_roll_err = 1e18

    idx = 0
    for ax in AXES:
        for roll in ROLLS:
            if idx >= len(candidates):
                continue
            inst = candidates[idx]
            idx += 1
            if not inst:
                continue
            try:
                tr = inst.GetTransform()
                # align chosen family axis to target_axis
                if ax.IsAlmostEqualTo(XYZ.BasisX):
                    from_v = tr.BasisX
                elif ax.IsAlmostEqualTo(XYZ.BasisY):
                    from_v = tr.BasisY
                else:
                    from_v = tr.BasisZ

                rotate_align(inst, pt, from_v, target_axis)

                # apply roll offset about target_axis
                rotate_roll_deg(inst, pt, target_axis, roll)

                # if we have a roll reference, snap roll to it (X/Y whichever is closer)
                if roll_ref:
                    rotate_match_roll(inst, pt, target_axis, roll_ref)

                sc = bbox_thickness_along_axis(inst, target_axis)

                # compute roll error (tie-breaker)
                roll_err = 0.0
                if roll_ref:
                    tr2 = inst.GetTransform()
                    cx = unitize(project_onto_plane(tr2.BasisX, target_axis))
                    cy = unitize(project_onto_plane(tr2.BasisY, target_axis))
                    best_dot = -1.0
                    for c in [cx, cy]:
                        if c:
                            best_dot = max(best_dot, c.DotProduct(roll_ref))
                    best_dot = max(-1.0, min(1.0, best_dot))
                    roll_err = math.acos(best_dot)

                if (sc < best_score - 1e-9) or (abs(sc - best_score) <= 1e-9 and roll_err < best_roll_err):
                    best_score = sc
                    best_roll_err = roll_err
                    best_inst = inst
            except:
                continue

    # delete losers
    for inst in candidates:
        if inst and best_inst and inst.Id != best_inst.Id:
            try:
                doc.Delete(inst.Id)
            except:
                pass

    return best_inst


def choose_roll_ref_for_joint(joint_pt, axis_dir, elems):
    """
    Pick a stable roll reference from a host connector CoordinateSystem near joint_pt.
    This prevents 'flattening' (loss of roll) when re-snapping.
    """
    axis_dir = unitize(axis_dir)
    if not axis_dir:
        return None

    best = None
    best_score = -1.0

    for el in elems:
        conns = get_connectors(el)
        if not conns:
            continue

        # connector nearest the joint
        near = None
        near_d = 1e9
        for c in conns:
            try:
                d = dist(c.Origin, joint_pt)
                if d < near_d:
                    near_d = d
                    near = c
            except:
                continue

        # must be close enough to be the actual joint connector
        if (not near) or (near_d > (LEN_TOL_FT * 6.0)):
            continue

        try:
            cs = near.CoordinateSystem
            bx = cs.BasisX
            by = cs.BasisY
        except:
            continue

        # choose the basis that is most perpendicular to axis_dir
        for cand in (bx, by):
            p = project_onto_plane(cand, axis_dir)
            p = unitize(p) if p else None
            if not p:
                continue
            # score = |p| (should be 1 after unitize, but keep for clarity)
            score = p.GetLength()
            if score > best_score:
                best_score = score
                best = p

    return best

def parse_host_ids(hostids_text):
    if not hostids_text:
        return []
    parts = [p.strip() for p in str(hostids_text).split(",") if p.strip()]
    out = []
    for p in parts:
        try:
            out.append(int(p))
        except:
            continue
    return sorted(list(set(out)))


def collect_fab_parts_in_view():
    col = (FilteredElementCollector(doc, av.Id)
           .OfClass(FabricationPart)
           .WhereElementIsNotElementType())
    return list(col)


def collect_weld_markers_in_view():
    col = (FilteredElementCollector(doc, av.Id)
           .OfCategory(BuiltInCategory.OST_PipeAccessory)
           .WhereElementIsNotElementType())
    out = []
    for e in col:
        try:
            if e.Symbol and e.Symbol.Family and e.Symbol.Family.Name == WELD_RFA_FAMILY_NAME:
                _ = e.Location.Point
                out.append(e)
        except:
            continue
    return out


def build_joint_map(fab_parts):
    """
    joints[key] = {"pt": XYZ, "elem_ids": set(int), "elems": list(Element), "tap": FabricationPart|None}
    """
    fab_ids = set([e.Id.IntegerValue for e in fab_parts])
    joints = {}

    def add(key, pt, el):
        if key not in joints:
            joints[key] = {"pt": pt, "elem_ids": set(), "elems": [], "tap": None}
        eid = el.Id.IntegerValue
        if eid not in joints[key]["elem_ids"]:
            joints[key]["elem_ids"].add(eid)
            joints[key]["elems"].append(el)
            if is_tap_like(el):
                joints[key]["tap"] = el

    for el in fab_parts:
        conns = get_connectors(el)
        if not conns:
            continue
        for c in conns:
            try:
                pt = c.Origin
            except:
                continue

            other_ids = [oid for oid in connected_owner_ids(c) if oid in fab_ids]
            if not other_ids:
                continue

            key = round_key(pt)
            add(key, pt, el)
            for oid in other_ids:
                try:
                    other_el = doc.GetElement(ElementId(oid))
                    if other_el:
                        add(key, pt, other_el)
                except:
                    continue

    return joints


def choose_axis_for_joint(joint_pt, elems):
    best_dir = None
    best_len = 0.0

    for el in elems:
        conns = get_connectors(el)
        if len(conns) < 2:
            continue

        near = None
        near_d = 1e9
        for c in conns:
            try:
                d = dist(c.Origin, joint_pt)
                if d < near_d:
                    near_d = d
                    near = c
            except:
                continue

        if not near or near_d > (LEN_TOL_FT * 6.0):
            continue

        far = None
        far_d = 0.0
        for c in conns:
            if c == near:
                continue
            try:
                d = dist(c.Origin, near.Origin)
                if d > far_d:
                    far_d = d
                    far = c
            except:
                continue

        if far and far_d > best_len:
            v = unitize(far.Origin - near.Origin)
            if v:
                best_dir = v
                best_len = far_d

    return best_dir


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


def solve_target_for_joint(joint):
    joint_pt = joint["pt"]
    elems = [e for e in joint["elems"] if e is not None]

    participants = [e for e in elems if not (isinstance(e, FabricationPart) and is_weld_itm(e))]
    if len(participants) < 2:
        return (None, None, None)

    tap_part = joint.get("tap", None)

    place_pt = joint_pt
    axis_dir = choose_axis_for_joint(joint_pt, participants)
    size_od = None

    if tap_part:
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
            p2, ax2, od2 = compute_tap_face_point_and_axis_and_visual_od(tap_part, run_part)
            if p2 and ax2:
                place_pt = p2
                axis_dir = ax2
                size_od = od2

    if not axis_dir:
        return (None, None, None)

    if (not size_od) or size_od <= 1e-9:
        max_od = 0.0
        for e in participants:
            od = get_param_len_ft(e, "Outside Diameter") or get_param_len_ft(e, "OutsideDiameter") or 0.0
            if od > max_od:
                max_od = od
        size_od = max_od if max_od > 1e-9 else None

    if not size_od:
        return (None, None, None)

    return (place_pt, axis_dir, size_od)


def find_joint_for_hostids(joints, host_ids):
    """
    Returns best matching joint where all host_ids are contained in joint elem_ids.
    """
    if not host_ids:
        return None

    host_set = set(host_ids)
    best = None
    best_extra = 1e18

    for j in joints.values():
        jset = j["elem_ids"]
        if host_set.issubset(jset):
            extra = len(jset) - len(host_set)
            if extra < best_extra:
                best_extra = extra
                best = j

    return best


# ---------------- MAIN ----------------
fab_parts = collect_fab_parts_in_view()
markers = collect_weld_markers_in_view()

if not markers:
    forms.alert("No '{}' markers found in the active view.".format(WELD_RFA_FAMILY_NAME), warn_icon=False)
    script.exit()

if not fab_parts:
    forms.alert("No FabricationParts found in the active view.", warn_icon=False)
    script.exit()

joints = build_joint_map(fab_parts)
if not joints:
    forms.alert("No connected joints found in the active view scope.", warn_icon=False)
    script.exit()

moved = 0
resized = 0
skipped_nohost = 0
skipped_nomatch = 0
skipped_solve = 0

t = Transaction(doc, "Re-Snap Pipe Weld markers to host joints (active view)")
t.Start()
try:
    for mk in markers:
        # read host ids
        host_text = ""
        try:
            p = mk.LookupParameter(WELD_HOSTIDS_PARAM)
            if p and p.HasValue:
                host_text = (p.AsString() or "").strip()
        except:
            host_text = ""

        host_ids = parse_host_ids(host_text)
        if not host_ids:
            skipped_nohost += 1
            continue

        joint = find_joint_for_hostids(joints, host_ids)
        if not joint:
            skipped_nomatch += 1
            continue

        place_pt, axis_dir, size_od = solve_target_for_joint(joint)
        if not place_pt or not axis_dir or not size_od:
            skipped_solve += 1
            continue

        # move
        try:
            mk_pt = mk.Location.Point
        except:
            skipped_solve += 1
            continue

        move_vec = place_pt - mk_pt
        if move_vec.GetLength() > MOVE_EPS_FT:
            try:
                ElementTransformUtils.MoveElement(doc, mk.Id, move_vec)
                moved += 1
            except:
                skipped_solve += 1
                continue

        # resize
        try:
            ps = mk.LookupParameter(WELD_SIZE_PARAM)
            if ps and (not ps.IsReadOnly):
                target_cp = size_od + SIZE_OFFSET_FT
                if target_cp > 0.0:
                    ps.Set(target_cp)
                    resized += 1
        except:
            pass

        # align like Script 1 (axis + roll) and also match host roll
        try:
            new_pt = mk.Location.Point
            participants = [e for e in joint["elems"] if e is not None]
            roll_ref = choose_roll_ref_for_joint(joint.get("pt"), axis_dir, participants)
            mk = solve_orientation_like_placer(mk, new_pt, axis_dir, roll_ref)
        except:
            pass

finally:
    t.Commit()

forms.alert(
    "Re-Snap complete (Active View)\n\n"
    "Markers found: {0}\n"
    "Moved: {1}\n"
    "Resized: {2}\n"
    "Skipped (no {3}): {4}\n"
    "Skipped (no matching joint in view): {5}\n"
    "Skipped (could not solve/update): {6}\n".format(
        len(markers), moved, resized, WELD_HOSTIDS_PARAM, skipped_nohost, skipped_nomatch, skipped_solve
    ),
    warn_icon=(moved > 0)
)
