# -*- coding: utf-8 -*-
"""Convert Fab to Native - Unified conversion tool.

Consolidates Convert Straights, Elbows, Transitions, and Taps into a single
branded WPF window using the XamlReader/ShowDialog pick-loop pattern.

All four engines are live:
  - Straights:   Fab straight -> native Duct.Create (point-to-point)
  - Elbows:      Fab elbow -> NewElbowFitting (stub-bridge or native-endpoint)
  - Transitions: Fab transition -> NewTransitionFitting (oriented-stub bridge)
  - Taps:        Fab tap -> native branch duct (bbox-span, roll-locked)

Recommended conversion order:  Straights > Elbows > Transitions > Taps

Author: Jeremiah Griffith
Version: 0.2.0
"""

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.ObjectModel import ObservableCollection
from collections import defaultdict
import os
import math
import traceback
import System.Windows
import System.Windows.Input

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

uidoc = revit.uidoc
doc   = revit.doc
out   = script.get_output()
log   = script.get_logger()
Mech  = DB.Mechanical

XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")

# ================================================================
#  TUNABLES
# ================================================================
DEBUG = False

# Elbows
ELBOW_SEARCH_TOL_FT = 0.08
ELBOW_STUB_LEN_FT   = 0.25

# Taps
SEARCH_TOL_CURVE_FT  = 8.0
CLAMP_BRANCH_END     = True
CLAMP_BACKOFF_FT     = 1.0 / 96.0       # 1/8"
MIN_LEN_FT           = 1.0 / 120.0      # 1/10"
ROLL_LOCK_TOL        = 1e-6
RESIDUAL_LOCK_TOL    = 1e-5
MAX_ACCEPT_DEG       = 45.0


# ================================================================
#  DATA CLASSES
# ================================================================
class ResultItem(object):
    """Row in the results DataGrid. Properties must match XAML bindings."""
    def __init__(self, elem_id, status, message):
        self.ElemId  = str(elem_id)
        self.Status  = status
        self.Message = message


# ================================================================
#  SELECTION FILTER
# ================================================================
class FabDuctSelFilter(ISelectionFilter):
    """Allows only Fabrication Ductwork elements."""
    def AllowElement(self, e):
        try:
            return (e.Category and
                    e.Category.Id.IntegerValue == int(
                        DB.BuiltInCategory.OST_FabricationDuctwork))
        except:
            return False

    def AllowReference(self, ref, pt):
        return False


# ================================================================
#  SHARED MATH HELPERS
# ================================================================
def _add(a, b):
    return DB.XYZ(a.X + b.X, a.Y + b.Y, a.Z + b.Z)

def _sub(a, b):
    return DB.XYZ(a.X - b.X, a.Y - b.Y, a.Z - b.Z)

def _mul(v, s):
    return DB.XYZ(v.X * s, v.Y * s, v.Z * s)

def _dot(a, b):
    return a.X * b.X + a.Y * b.Y + a.Z * b.Z

def _crs(a, b):
    return DB.XYZ(a.Y*b.Z - a.Z*b.Y,
                  a.Z*b.X - a.X*b.Z,
                  a.X*b.Y - a.Y*b.X)

def _veclen(a):
    return math.sqrt(_dot(a, a))

def _norm(v):
    try:
        L = _veclen(v)
        return DB.XYZ(v.X/L, v.Y/L, v.Z/L) if L > 1e-9 else DB.XYZ(1, 0, 0)
    except:
        return DB.XYZ(1, 0, 0)

def _proj_on_plane(v, n):
    return _sub(v, _mul(n, _dot(v, n)))

def _signed_angle_in_plane(v_from, v_to, n):
    a = max(min(_dot(_norm(v_from), _norm(v_to)), 1.0), -1.0)
    ang = math.acos(a)
    s = _dot(_crs(v_from, v_to), n)
    return ang if s >= 0.0 else -ang

def _rot_about_axis(v, n_axis_unit, ang):
    vpar = _mul(n_axis_unit, _dot(v, n_axis_unit))
    vper = _sub(v, vpar)
    w    = _crs(n_axis_unit, vper)
    return _add(_add(vpar, _mul(vper, math.cos(ang))), _mul(w, math.sin(ang)))

def _wrap_pi(a):
    while a <= -math.pi: a += 2.0 * math.pi
    while a >   math.pi: a -= 2.0 * math.pi
    return a

def _deg(a):
    return a * 180.0 / math.pi

def closest_points_on_rays(p0, d0, p1, d1):
    """Midpoint of closest approach between two rays."""
    d0 = _norm(d0); d1 = _norm(d1)
    w0 = _sub(p0, p1)
    a = _dot(d0, d0); b = _dot(d0, d1); c = _dot(d1, d1)
    d = _dot(d0, w0); e = _dot(d1, w0)
    denom = a*c - b*b
    if abs(denom) < 1e-9:
        return DB.XYZ((p0.X+p1.X)/2.0, (p0.Y+p1.Y)/2.0, (p0.Z+p1.Z)/2.0)
    s = (b*e - c*d) / denom
    t = (a*e - b*d) / denom
    cp0 = _add(p0, _mul(d0, s))
    cp1 = _add(p1, _mul(d1, t))
    return DB.XYZ((cp0.X+cp1.X)/2.0, (cp0.Y+cp1.Y)/2.0, (cp0.Z+cp1.Z)/2.0)


# ================================================================
#  SHARED CONNECTOR UTILITIES
# ================================================================
def iter_connectors(elem):
    """Yield all connectors from an element."""
    cm = getattr(elem, "ConnectorManager", None)
    if cm:
        for c in cm.Connectors:
            yield c
        return
    try:
        mep = elem.MEPModel
        if mep:
            for c in mep.ConnectorManager.Connectors:
                yield c
    except:
        pass


def two_ends(elem):
    """Return the first two connectors, or (None, None)."""
    conns = [c for c in iter_connectors(elem)]
    return (conns[0], conns[1]) if len(conns) >= 2 else (None, None)


def two_farthest_connectors(elem):
    """Return the two connectors with maximum distance between them."""
    cons = list(iter_connectors(elem))
    if len(cons) < 2:
        return (None, None)
    best = None; pair = (cons[0], cons[1])
    for i in range(len(cons)):
        for j in range(i+1, len(cons)):
            try:
                d = cons[i].Origin.DistanceTo(cons[j].Origin)
            except:
                d = -1.0
            if best is None or d > best:
                best = d; pair = (cons[i], cons[j])
    return pair


def connector_area(c):
    """Cross-sectional area of a connector."""
    try:
        if c.Shape == DB.ConnectorProfileType.Round:
            d = float(c.Radius) * 2.0
            return math.pi * (d * d) / 4.0
        if c.Shape == DB.ConnectorProfileType.Oval:
            a = float(c.Radius); b = float(c.RadiusTwo)
            return math.pi * a * b
        return float(c.Width) * float(c.Height)
    except:
        return 0.0


def _as_xyz_direction(maybe_dir, fallback_xyz=None):
    if (hasattr(maybe_dir, 'X') and hasattr(maybe_dir, 'Y')
            and hasattr(maybe_dir, 'Z')):
        return _norm(maybe_dir)
    if fallback_xyz is not None:
        return _norm(fallback_xyz)
    return None


def connector_dir(c):
    """Best-effort direction vector for a connector."""
    try:
        d = _as_xyz_direction(c.Direction)
        if d is not None:
            return d
    except:
        pass
    try:
        return _norm(c.CoordinateSystem.BasisZ)
    except:
        pass
    try:
        allc = list(iter_connectors(c.Owner))
        if len(allc) >= 2:
            a, b = None, None
            best = -1.0
            for i in range(len(allc)):
                for j in range(i+1, len(allc)):
                    d = _veclen(_sub(allc[i].Origin, allc[j].Origin))
                    if d > best:
                        best = d; a, b = allc[i], allc[j]
            if a and b:
                v = _sub(b.Origin, a.Origin)
                return _norm(v) if _veclen(v) > 1e-9 else DB.XYZ(1, 0, 0)
    except:
        pass
    return DB.XYZ(1, 0, 0)


def _connector_closest_to(elem, point):
    """Return the connector on elem nearest to point."""
    pick = None; dmin = None
    for c in iter_connectors(elem):
        d = c.Origin.DistanceTo(point)
        if pick is None or d < dmin:
            pick = c; dmin = d
    return pick


def shape_key(conn):
    """(shape_name, dim1, dim2) tuple for size comparison."""
    shp = getattr(conn, "Shape", None)
    if shp == DB.ConnectorProfileType.Round:
        try:    return ("Round", round(2.0 * conn.Radius, 6), None)
        except: return ("Round", None, None)
    elif shp == DB.ConnectorProfileType.Rectangular:
        try:    return ("Rect", round(conn.Width, 6), round(conn.Height, 6))
        except: return ("Rect", None, None)
    elif shp == DB.ConnectorProfileType.Oval:
        try:    return ("Oval", round(conn.Width, 6), round(conn.Height, 6))
        except: return ("Oval", None, None)
    return ("Other", None, None)


def conn_size_tuple(c):
    """(shape, dim1_ft, dim2_ft) for sizing stubs."""
    shp = getattr(c, "Shape", None)
    if shp == DB.ConnectorProfileType.Round:
        return ("Round", 2.0 * c.Radius, None)
    elif shp == DB.ConnectorProfileType.Rectangular:
        return ("Rect", c.Width, c.Height)
    elif shp == DB.ConnectorProfileType.Oval:
        return ("Oval", c.Width, c.Height)
    return ("Other", None, None)


def sizes_compatible(s0, s1, tol=1e-4):
    if s0[0] != s1[0]:
        return False
    if s0[0] == "Round":
        if s0[1] is None or s1[1] is None: return True
        return abs(s0[1] - s1[1]) <= tol
    if s0[0] in ("Rect", "Oval"):
        if None in (s0[1], s0[2], s1[1], s1[2]): return True
        return ((abs(s0[1]-s1[1]) <= tol and abs(s0[2]-s1[2]) <= tol) or
                (abs(s0[1]-s1[2]) <= tol and abs(s0[2]-s1[1]) <= tol))
    return False


def shape_name_from_connector(c):
    try:
        s = c.Shape
        if s == DB.ConnectorProfileType.Round:       return "Round"
        if s == DB.ConnectorProfileType.Rectangular: return "Rectangular"
        if s == DB.ConnectorProfileType.Oval:        return "Oval"
    except:
        pass
    return "Unknown"


# ================================================================
#  SHARED MODEL QUERIES
# ================================================================
def get_mech_system_type_id():
    """Return any available MechanicalSystemType Id."""
    coll = DB.FilteredElementCollector(doc).OfClass(Mech.MechanicalSystemType)
    for st in coll:
        return st.Id
    return DB.ElementId.InvalidElementId


def find_ducttype_for_shape(shape):
    """Find first native DuctType matching the connector shape."""
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        try:
            if dt.Shape == shape:
                return dt
        except:
            continue
    return None


def ensure_ducttype(shape):
    """Return DuctType Id for shape, or InvalidElementId with alert."""
    dt = find_ducttype_for_shape(shape)
    if dt:
        return dt.Id
    forms.alert(
        "No native Duct Type found for shape: {0}.\n"
        "Draw a tiny native duct of this shape "
        "(or Transfer Project Standards > Duct Types) and rerun."
        .format(shape),
        title="Missing Duct Type", warn_icon=True)
    return DB.ElementId.InvalidElementId


def level_for(elem):
    """Best-guess Level Id for an element."""
    try:
        if elem.LevelId and elem.LevelId != DB.ElementId.InvalidElementId:
            return elem.LevelId
    except:
        pass
    av = doc.ActiveView
    if hasattr(av, "GenLevel") and av.GenLevel:
        return av.GenLevel.Id
    bb = elem.get_BoundingBox(None)
    z  = (bb.Min.Z + bb.Max.Z) * 0.5 if bb else 0.0
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    if levels:
        return min(levels, key=lambda L: abs(L.Elevation - z)).Id
    return DB.ElementId.InvalidElementId


def set_size_from_connector(duct, conn):
    """Copy cross-section dimensions from a connector onto a new duct."""
    try:
        if conn.Shape == DB.ConnectorProfileType.Round:
            diam = 2.0 * conn.Radius
            p = duct.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            if p and not p.IsReadOnly:
                p.Set(diam)
        elif conn.Shape in (DB.ConnectorProfileType.Rectangular,
                            DB.ConnectorProfileType.Oval):
            pw = duct.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            ph = duct.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            if pw and not pw.IsReadOnly and hasattr(conn, "Width"):
                pw.Set(conn.Width)
            if ph and not ph.IsReadOnly and hasattr(conn, "Height"):
                ph.Set(conn.Height)
    except Exception as ex:
        log.warn("Failed to set size on {0}: {1}".format(duct.Id, ex))


def copy_common_metadata(src, dst):
    """Copy workset, phase, comments, mark from src to dst."""
    try:
        if doc.IsWorkshared:
            wsparam = dst.get_Parameter(
                DB.BuiltInParameter.ELEM_PARTITION_PARAM)
            if wsparam and not wsparam.IsReadOnly:
                wsparam.Set(src.WorksetId.IntegerValue)
    except:
        pass
    for bip in (DB.BuiltInParameter.PHASE_CREATED,
                DB.BuiltInParameter.PHASE_DEMOLISHED):
        try:
            sp = src.get_Parameter(bip)
            dp = dst.get_Parameter(bip)
            if (sp and dp and not dp.IsReadOnly
                    and sp.AsElementId() != DB.ElementId.InvalidElementId):
                dp.Set(sp.AsElementId())
        except:
            pass
    for bip in (DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS,
                DB.BuiltInParameter.ALL_MODEL_MARK):
        try:
            sp = src.get_Parameter(bip)
            dp = dst.get_Parameter(bip)
            if sp and dp and not dp.IsReadOnly and sp.AsString():
                dp.Set(sp.AsString())
        except:
            pass


def connect_to_nearby(duct):
    """If new duct ends coincide with a native connector, connect them."""
    tol = 1.0 / 12.0
    outline_pad = 0.05
    viewid = doc.ActiveView.Id
    for c in iter_connectors(duct):
        pt = c.Origin
        minp = DB.XYZ(pt.X - outline_pad,
                       pt.Y - outline_pad,
                       pt.Z - outline_pad)
        maxp = DB.XYZ(pt.X + outline_pad,
                       pt.Y + outline_pad,
                       pt.Z + outline_pad)
        bbfilter = DB.BoundingBoxIntersectsFilter(
            DB.Outline(minp, maxp))
        nearby = (DB.FilteredElementCollector(doc, viewid)
                  .WherePasses(bbfilter)
                  .WhereElementIsNotElementType())
        connected = False
        for e in nearby:
            if connected:
                break
            if e.Id == duct.Id:
                continue
            for oc in iter_connectors(e):
                try:
                    if (oc.Origin.DistanceTo(pt) <= tol
                            and not c.IsConnectedTo(oc)):
                        c.ConnectTo(oc)
                        connected = True
                        break
                except:
                    pass


# ================================================================
#  STRAIGHTS ENGINE
# ================================================================
def convert_straight(part, sys_type_id, delete_original):
    """Convert one fabrication straight to native duct."""
    c0, c1 = two_ends(part)
    if not c0 or not c1:
        return None, "skip: not a straight (needs 2 end connectors)"
    if c0.Shape != c1.Shape:
        return None, "skip: end shapes differ"

    ducttype_id = ensure_ducttype(c0.Shape)
    if ducttype_id == DB.ElementId.InvalidElementId:
        return None, "skip: missing duct type"
    if sys_type_id == DB.ElementId.InvalidElementId:
        return None, "skip: no Mechanical System Type"

    sp, ep = c0.Origin, c1.Origin
    lvl_id = level_for(part)

    try:
        new_duct = Mech.Duct.Create(
            doc, sys_type_id, ducttype_id, lvl_id, sp, ep)
        if not new_duct:
            return None, "create failed"
        set_size_from_connector(new_duct, c0)
        copy_common_metadata(part, new_duct)
        sn = part.LookupParameter("System Name")
        if sn and sn.AsString():
            try:
                new_duct.LookupParameter("System Name").Set(sn.AsString())
            except:
                pass
        connect_to_nearby(new_duct)
        if delete_original:
            doc.Delete(part.Id)
        return new_duct, "ok"
    except Exception as ex:
        return None, "create failed: {0}".format(ex)


# ================================================================
#  ELBOWS ENGINE - HELPERS
# ================================================================
def elbow_angle_deg(c0, c1):
    """Angle between two connector directions in degrees."""
    try:
        v0 = getattr(c0, "Direction", None) or c0.CoordinateSystem.BasisZ
        v1 = getattr(c1, "Direction", None) or c1.CoordinateSystem.BasisZ
        v0 = v0.Normalize(); v1 = v1.Normalize()
        dot = v0.DotProduct(v1)
        dot = max(min(dot, 1.0), -1.0)
        return abs(math.degrees(math.acos(dot)))
    except:
        return 180.0


def is_elbow(elem, minang=5.0, maxang=175.0):
    """True if element has exactly 2 connectors at a non-straight angle."""
    c0, c1 = two_ends(elem)
    if not c0 or not c1:
        return False
    ang = elbow_angle_deg(c0, c1)
    return (ang > minang and ang < maxang)


def _elbow_group_type():
    for nm in ("Elbows", "Elbow"):
        try:
            return getattr(DB.RoutingPreferenceRuleGroupType, nm)
        except:
            pass
    return None


def _ducttype_with_elbow_for_shape(shape):
    """Find a DuctType with an elbow routing preference rule for the shape."""
    grp_type = _elbow_group_type()
    best = None
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        try:
            if dt.Shape != shape: continue
        except:
            continue
        if best is None:
            best = dt
        try:
            rpm = dt.RoutingPreferenceManager
            n = 0
            try:
                grp = rpm.GetRuleGroup(grp_type) if grp_type else None
                if grp:
                    try:    n = grp.GetNumberOfRules()
                    except:
                        try:    n = grp.NumberOfRules
                        except: n = 0
            except:
                try:    n = rpm.GetNumberOfRules(grp_type) if grp_type else 0
                except: n = 0
            if n and n > 0:
                return dt
        except:
            pass
    return best


def find_nearby_native(conn, tol_ft=0.05):
    """Find the nearest native duct connector to a fab connector."""
    viewid = doc.ActiveView.Id
    pt = conn.Origin
    pad = tol_ft
    minp = DB.XYZ(pt.X - pad, pt.Y - pad, pt.Z - pad)
    maxp = DB.XYZ(pt.X + pad, pt.Y + pad, pt.Z + pad)
    bbfilter = DB.BoundingBoxIntersectsFilter(DB.Outline(minp, maxp))

    best = None; bestd = None
    shp_sz = shape_key(conn)

    cand = (DB.FilteredElementCollector(doc, viewid)
            .OfCategory(DB.BuiltInCategory.OST_DuctCurves)
            .WhereElementIsNotElementType()
            .WherePasses(bbfilter))
    for e in cand:
        for oc in iter_connectors(e):
            try:
                d = oc.Origin.DistanceTo(pt)
                if d <= tol_ft and sizes_compatible(shp_sz, shape_key(oc)):
                    if best is None or d < bestd:
                        best = oc; bestd = d
            except:
                pass
    return best, bestd


def _create_elbow_stub(conn, part_for_level, sys_type_id, ducttype_id,
                       stub_len_ft=0.25):
    """Create a short native duct stub at a fab connector.

    Runs inside an active transaction (no internal transaction management).
    Returns (nearest_connector, stub_element) or (None, None).
    """
    if ducttype_id == DB.ElementId.InvalidElementId:
        return None, None

    lvl_id = level_for(part_for_level)

    try:
        dvec = (getattr(conn, "Direction", None)
                or conn.CoordinateSystem.BasisZ)
        dvec = dvec.Normalize()
    except:
        dvec = DB.XYZ.BasisZ

    sp = conn.Origin
    ep = _add(sp, _mul(dvec, stub_len_ft))

    try:
        stub = Mech.Duct.Create(
            doc, sys_type_id, ducttype_id, lvl_id, sp, ep)
        if not stub:
            return None, None

        set_size_from_connector(stub, conn)
        doc.Regenerate()

        near = _connector_closest_to(stub, sp)
        return near, stub
    except:
        return None, None


# ================================================================
#  ELBOWS ENGINE
# ================================================================
def convert_elbow(part, sys_type_id, delete_original):
    """Convert one fabrication elbow to a native elbow fitting.

    Runs inside an active transaction from run_conversion().
    """
    c0, c1 = two_ends(part)
    if not c0 or not c1:
        return None, "skip: missing connectors"

    ang = elbow_angle_deg(c0, c1)
    if ang <= 5.0 or ang >= 175.0:
        return None, "skip: not an elbow (angle {0:.1f} deg)".format(ang)

    sp0 = c0.Origin
    sp1 = c1.Origin

    # Find a duct type with an elbow routing preference
    dt = _ducttype_with_elbow_for_shape(getattr(c0, "Shape", None))
    if not dt:
        return None, "failed: no Duct Type of this shape found in model"

    try:
        grp_type = _elbow_group_type()
        rpm = dt.RoutingPreferenceManager
        n_rules = 0
        try:
            grp = rpm.GetRuleGroup(grp_type) if grp_type else None
            if grp:
                try:    n_rules = grp.GetNumberOfRules()
                except:
                    try:    n_rules = grp.NumberOfRules
                    except: n_rules = 0
        except:
            try:    n_rules = rpm.GetNumberOfRules(grp_type) if grp_type else 0
            except: n_rules = 0
        if n_rules == 0:
            return None, ("failed: Duct Type '{}' has no Elbow in "
                          "Routing Preferences".format(dt.Name))
    except:
        pass

    # --- Look for existing native connectors at each end ---
    nat0, _ = find_nearby_native(c0, tol_ft=ELBOW_SEARCH_TOL_FT)
    stub0 = None
    if nat0 is None:
        nat0, stub0 = _create_elbow_stub(
            c0, part, sys_type_id, dt.Id, ELBOW_STUB_LEN_FT)

    if nat0 is None:
        return None, ("skip: no native endpoint near conn[0] "
                       "(convert straights first)")

    nat1, _ = find_nearby_native(c1, tol_ft=ELBOW_SEARCH_TOL_FT)
    stub1 = None
    if nat1 is None:
        nat1, stub1 = _create_elbow_stub(
            c1, part, sys_type_id, dt.Id, ELBOW_STUB_LEN_FT)

    if nat1 is None:
        return None, ("skip: no native endpoint near conn[1] "
                       "(convert straights first)")

    if nat0 == nat1:
        return None, "skip: both ends matched the same native connector"

    # --- Case A: both ends are temporary stubs (reshape to meet) ---
    if stub0 and stub1:
        d0 = _mul(_norm(getattr(c0, "Direction", None)
                        or c0.CoordinateSystem.BasisZ), -1.0)
        d1 = _mul(_norm(getattr(c1, "Direction", None)
                        or c1.CoordinateSystem.BasisZ), -1.0)
        meet = closest_points_on_rays(sp0, d0, sp1, d1)

        try:
            stub0.Location.Curve = DB.Line.CreateBound(sp0, meet)
            stub1.Location.Curve = DB.Line.CreateBound(meet, sp1)

            # Unify sizes based on c0
            try:
                if c0.Shape == DB.ConnectorProfileType.Round:
                    dp0 = stub0.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                    dp1 = stub1.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
                    if (dp0 and not dp0.IsReadOnly
                            and dp1 and not dp1.IsReadOnly):
                        target = 2.0 * c0.Radius
                        dp0.Set(target); dp1.Set(target)
                elif c0.Shape in (DB.ConnectorProfileType.Rectangular,
                                  DB.ConnectorProfileType.Oval):
                    w0p = stub0.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                    h0p = stub0.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                    w1p = stub1.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                    h1p = stub1.get_Parameter(
                        DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                    if hasattr(c0, "Width") and hasattr(c0, "Height"):
                        if w0p and not w0p.IsReadOnly: w0p.Set(c0.Width)
                        if h0p and not h0p.IsReadOnly: h0p.Set(c0.Height)
                        if w1p and not w1p.IsReadOnly: w1p.Set(c0.Width)
                        if h1p and not h1p.IsReadOnly: h1p.Set(c0.Height)
            except:
                pass

            doc.Regenerate()

            far0 = _connector_closest_to(stub0, meet)
            far1 = _connector_closest_to(stub1, meet)

            # Try NewElbowFitting first
            elbow = None
            try:
                elbow = doc.Create.NewElbowFitting(far0, far1)
            except:
                elbow = None

            # Fallback: connect and find auto-created fitting
            if not elbow:
                try:
                    far0.ConnectTo(far1)
                    doc.Regenerate()
                    for cc in (far0, far1):
                        refs = (list(cc.AllRefs)
                                if hasattr(cc, "AllRefs") else [])
                        for rc in refs:
                            owner = rc.Owner
                            if (owner and owner.Category
                                    and owner.Category.Id.IntegerValue
                                    == int(DB.BuiltInCategory
                                           .OST_DuctFitting)):
                                elbow = owner
                                break
                        if elbow:
                            break
                except:
                    elbow = None

            if not elbow:
                return None, "failed: could not insert elbow fitting"

            doc.Delete(stub0.Id)
            doc.Delete(stub1.Id)
            if delete_original:
                try:
                    doc.Delete(part.Id)
                except:
                    pass

            return elbow, "ok"

        except Exception as ex:
            return None, "failed: elbow from stubs ({0})".format(ex)

    # --- Case B: at least one endpoint is an existing native connector ---
    try:
        elbow = None
        try:
            elbow = doc.Create.NewElbowFitting(nat0, nat1)
        except:
            elbow = None

        if not elbow:
            nat0.ConnectTo(nat1)
            doc.Regenerate()
            for cc in (nat0, nat1):
                refs = (list(cc.AllRefs)
                        if hasattr(cc, "AllRefs") else [])
                for rc in refs:
                    owner = rc.Owner
                    if (owner and owner.Category
                            and owner.Category.Id.IntegerValue
                            == int(DB.BuiltInCategory.OST_DuctFitting)):
                        elbow = owner
                        break
                if elbow:
                    break

        if not elbow:
            return None, "failed: could not insert elbow fitting"

        # Clean up any stubs we created
        if stub0:
            try: doc.Delete(stub0.Id)
            except: pass
        if stub1:
            try: doc.Delete(stub1.Id)
            except: pass

        if delete_original:
            try:
                doc.Delete(part.Id)
            except:
                pass

        return elbow, "ok"

    except Exception as ex:
        return None, "failed: elbow create ({0})".format(ex)


# ================================================================
#  TRANSITIONS ENGINE - HELPERS
# ================================================================
def _create_oriented_stub(conn, sys_type_id, ducttype_id, lvl_id,
                          length_ft):
    """Create a stub duct oriented along a connector's direction.

    Returns (stub_element, start_connector, end_connector).
    Runs inside an active transaction.
    """
    sp = conn.Origin
    try:
        d = (getattr(conn, "Direction", None)
             or conn.CoordinateSystem.BasisZ)
    except:
        d = DB.XYZ.BasisZ
    d = d.Normalize()
    ep = DB.XYZ(sp.X + d.X * length_ft,
                sp.Y + d.Y * length_ft,
                sp.Z + d.Z * length_ft)
    stub = Mech.Duct.Create(doc, sys_type_id, ducttype_id, lvl_id, sp, ep)

    # Size to match
    shp, s1, s2 = conn_size_tuple(conn)
    try:
        if shp == "Round" and s1 is not None:
            p = stub.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            if p and not p.IsReadOnly:
                p.Set(s1)
        elif shp in ("Rect", "Oval") and s1 is not None and s2 is not None:
            pw = stub.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            ph = stub.get_Parameter(
                DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            if pw and not pw.IsReadOnly: pw.Set(s1)
            if ph and not ph.IsReadOnly: ph.Set(s2)
    except:
        pass

    cons = list(iter_connectors(stub))
    startc = None; endc = None
    if cons:
        try:    startc = min(cons, key=lambda c: c.Origin.DistanceTo(sp))
        except: startc = cons[0]
        try:    endc = max(cons, key=lambda c: c.Origin.DistanceTo(sp))
        except: endc = cons[-1]
    return stub, startc, endc


def _place_transition(con0, con1):
    """Place a transition fitting between two connectors."""
    try:
        t = doc.Create.NewTransitionFitting(con0, con1)
        if t:
            return t
    except:
        pass
    try:
        con0.ConnectTo(con1)
        doc.Regenerate()
        for cc in (con0, con1):
            try:
                for rc in cc.AllRefs:
                    owner = rc.Owner
                    if (owner and owner.Category
                            and owner.Category.Id.IntegerValue
                            == int(DB.BuiltInCategory.OST_DuctFitting)):
                        return owner
            except:
                pass
    except:
        pass
    return None


def _set_transition_length(fitting, length_ft):
    """Force a transition fitting to a target centerline length."""
    if fitting is None:
        return False
    ok = False
    try:
        for p in fitting.Parameters:
            try:
                name = p.Definition.Name if p and p.Definition else ""
                if (not p.IsReadOnly) and name and ("length" in name.lower()):
                    p.Set(length_ft)
                    ok = True
            except:
                pass
    except:
        pass
    if not ok:
        try:
            t = doc.GetElement(fitting.GetTypeId())
            if t:
                for p in t.Parameters:
                    try:
                        name = p.Definition.Name if p and p.Definition else ""
                        if ((not p.IsReadOnly) and name
                                and ("length" in name.lower())):
                            p.Set(length_ft)
                            ok = True
                    except:
                        pass
                if ok:
                    doc.Regenerate()
        except:
            pass
    if ok:
        try:
            doc.Regenerate()
        except:
            pass
    return ok


def _first_ducttype_for_shape(shape):
    """First duct type matching the given connector shape."""
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        try:
            if dt.Shape == shape:
                return dt.Id
        except:
            pass
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        return dt.Id
    return DB.ElementId.InvalidElementId


def _first_mech_system_type():
    for st in DB.FilteredElementCollector(doc).OfClass(
            Mech.MechanicalSystemType):
        return st.Id
    return DB.ElementId.InvalidElementId


def _level_id_for(elem):
    """Level Id for an element (returns Id, not element)."""
    try:
        if (hasattr(elem, "LevelId") and elem.LevelId
                and elem.LevelId != DB.ElementId.InvalidElementId):
            return elem.LevelId
    except:
        pass
    try:
        av = doc.ActiveView
        if hasattr(av, "GenLevel") and av.GenLevel:
            return av.GenLevel.Id
    except:
        pass
    lvls = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    return lvls[0].Id if lvls else DB.ElementId.InvalidElementId


# ================================================================
#  TRANSITIONS ENGINE
# ================================================================
def convert_transition(part, sys_type_id, delete_original):
    """Convert one fabrication transition to a native transition fitting.

    Runs inside an active transaction from run_conversion().
    Uses SubTransaction for atomic rollback within the parent transaction.
    """
    c0, c1 = two_farthest_connectors(part)
    if not c0 or not c1:
        return None, "skip: less than 2 connectors"

    p0, p1 = c0.Origin, c1.Origin
    L = p0.DistanceTo(p1)

    sys_el = getattr(part, "MEPSystem", None)
    sys_id = sys_el.GetTypeId() if sys_el else _first_mech_system_type()
    lvl_id = _level_id_for(part)
    dt0 = _first_ducttype_for_shape(c0.Shape)
    dt1 = _first_ducttype_for_shape(c1.Shape)

    if (sys_id == DB.ElementId.InvalidElementId
            or lvl_id == DB.ElementId.InvalidElementId
            or not dt0 or not dt1):
        return None, "failed: missing system/type/level"

    stub_len = max(1.0/192.0, min(0.25, L * 0.1))

    st = DB.SubTransaction(doc)
    st.Start()
    try:
        stub0, s0_start, s0_end = _create_oriented_stub(
            c0, sys_id, dt0, lvl_id, stub_len)
        stub1, s1_start, s1_end = _create_oriented_stub(
            c1, sys_id, dt1, lvl_id, stub_len)
        if not s0_start or not s1_start:
            raise Exception("stub connectors missing")

        trans = _place_transition(s0_start, s1_start)
        if not trans:
            raise Exception("failed to create transition fitting")

        try:
            target_L = c0.Origin.DistanceTo(c1.Origin)
            _set_transition_length(trans, target_L)
        except:
            pass

        # Delete temp stubs
        try:
            if stub0: doc.Delete(stub0.Id)
            if stub1: doc.Delete(stub1.Id)
        except:
            pass

        # Delete original fab part
        if delete_original:
            try:
                doc.Delete(part.Id)
            except:
                pass

        st.Commit()
        return trans, "ok"

    except Exception as ex:
        try:
            st.RollBack()
        except:
            pass
        return None, "error: {0}".format(ex)


# ================================================================
#  TAPS ENGINE - SPATIAL INDEX & CACHES
# ================================================================
_TAP_GRID = SEARCH_TOL_CURVE_FT
_NATIVE_INDEX = defaultdict(list)

def _bucket_key(pt):
    return (int(pt.X / _TAP_GRID), int(pt.Y / _TAP_GRID))

def _build_native_index():
    """Build spatial index of native duct curves in the active view."""
    _NATIVE_INDEX.clear()
    viewid = doc.ActiveView.Id
    for e in (DB.FilteredElementCollector(doc, viewid)
              .OfCategory(DB.BuiltInCategory.OST_DuctCurves)
              .WhereElementIsNotElementType()):
        loc = getattr(e, "Location", None)
        crv = getattr(loc, "Curve", None) if loc else None
        if not crv:
            continue
        try:
            mid = crv.Evaluate(0.5, True)
        except:
            sp2, ep2 = crv.GetEndPoint(0), crv.GetEndPoint(1)
            mid = DB.XYZ((sp2.X+ep2.X)/2.0, (sp2.Y+ep2.Y)/2.0,
                         (sp2.Z+ep2.Z)/2.0)
        _NATIVE_INDEX[_bucket_key(mid)].append((e, crv))


def _nearest_native_duct_cached(pt, tol_ft=SEARCH_TOL_CURVE_FT):
    bx, by = _bucket_key(pt)
    best_e, best_d2 = None, 1e99
    maxd2 = tol_ft * tol_ft
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for e, crv in _NATIVE_INDEX.get((bx+dx, by+dy), ()):
                prj = crv.Project(pt)
                if not prj:
                    continue
                q = crv.Evaluate(prj.Parameter, False)
                dx2 = pt.X-q.X; dy2 = pt.Y-q.Y; dz2 = pt.Z-q.Z
                d2 = dx2*dx2 + dy2*dy2 + dz2*dz2
                if d2 < best_d2 and d2 <= maxd2:
                    best_e, best_d2 = e, d2
    return best_e


# Cached type IDs
_ROUND_TYPE_ID  = None
_RECT_TYPE_ID   = None
_DEFAULT_SYSID  = None
_DEFAULT_LVL    = None


def _init_tap_caches():
    global _ROUND_TYPE_ID, _RECT_TYPE_ID, _DEFAULT_SYSID, _DEFAULT_LVL
    _ROUND_TYPE_ID = _guess_round_ducttype_id()
    _RECT_TYPE_ID  = _guess_rect_ducttype_id()
    _DEFAULT_SYSID = _first_duct_system_type()
    _DEFAULT_LVL   = getattr(doc.ActiveView, "GenLevel", None)
    if not _DEFAULT_LVL:
        _DEFAULT_LVL = (DB.FilteredElementCollector(doc)
                        .OfClass(DB.Level).FirstElement())


def _first_duct_system_type():
    try:
        dt = (DB.FilteredElementCollector(doc)
              .OfClass(Mech.DuctType).FirstElement())
        if dt:
            sid = getattr(dt, "DefaultSystemTypeId", None)
            if sid and sid.IntegerValue != -1:
                return sid
    except:
        pass
    try:
        d = (DB.FilteredElementCollector(doc)
             .OfClass(Mech.Duct).FirstElement())
        if d:
            p = d.get_Parameter(
                DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
            if p:
                sid = p.AsElementId()
                if sid and sid.IntegerValue != -1:
                    return sid
    except:
        pass
    return DB.ElementId.InvalidElementId


def _system_from_host(duct_elem):
    if duct_elem:
        try:
            sys = duct_elem.MEPSystem
            if sys:
                tid = sys.GetTypeId()
                if tid and tid.IntegerValue != -1:
                    return tid
        except:
            pass
        try:
            p = duct_elem.get_Parameter(
                DB.BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM)
            if p:
                eid = p.AsElementId()
                if eid and eid.IntegerValue != -1:
                    return eid
        except:
            pass
    return DB.ElementId.InvalidElementId


def _guess_round_ducttype_id(pref_elem=None):
    if pref_elem and isinstance(pref_elem, Mech.Duct):
        c0t, c1t = two_farthest_connectors(pref_elem)
        if c0t and ("Round" in {shape_name_from_connector(c0t),
                                shape_name_from_connector(c1t)}):
            try:
                return pref_elem.DuctType.Id
            except:
                pass
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        nm = (getattr(dt, "FamilyName", "") + " "
              + getattr(dt, "Name", "")).lower()
        if "round" in nm:
            return dt.Id
    any_dt = (DB.FilteredElementCollector(doc)
              .OfClass(Mech.DuctType).FirstElement())
    return any_dt.Id if any_dt else DB.ElementId.InvalidElementId


def _guess_rect_ducttype_id():
    for dt in DB.FilteredElementCollector(doc).OfClass(Mech.DuctType):
        nm = (getattr(dt, "FamilyName", "") + " "
              + getattr(dt, "Name", "")).lower()
        if "rect" in nm or "rectangular" in nm:
            return dt.Id
    any_dt = (DB.FilteredElementCollector(doc)
              .OfClass(Mech.DuctType).FirstElement())
    return any_dt.Id if any_dt else DB.ElementId.InvalidElementId


def _level_for_tap(elem):
    try:
        return doc.GetElement(elem.ReferenceLevel.Id)
    except:
        try:
            p = elem.get_Parameter(
                DB.BuiltInParameter.RBS_START_LEVEL_PARAM)
            if p:
                eid = p.AsElementId()
                if eid and eid.IntegerValue != -1:
                    return doc.GetElement(eid)
        except:
            pass
    try:
        return doc.ActiveView.GenLevel
    except:
        pass
    return (DB.FilteredElementCollector(doc)
            .OfClass(DB.Level).FirstElement())


def _nearest_level_at_z(z):
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
    if not levels:
        return None
    return min(levels, key=lambda L: abs(float(L.Elevation) - float(z)))


# ================================================================
#  TAPS ENGINE - PARAM HELPERS
# ================================================================
def _set_param_if_exists(elem, bipname, val):
    bip = getattr(DB.BuiltInParameter, bipname, None)
    if not bip:
        return False
    try:
        p = elem.get_Parameter(bip)
        if p:
            p.Set(float(val))
            return True
    except:
        pass
    return False


def _set_round_diameter(elem, diam_ft):
    for name in ("RBS_CURVE_DIAMETER_PARAM",
                 "RBS_DUCT_DIAMETER_PARAM",
                 "RBS_PIPE_DIAMETER_PARAM",
                 "RBS_FLEXDUCT_DIAMETER_PARAM"):
        if _set_param_if_exists(elem, name, diam_ft):
            return
    try:
        elem.Diameter = float(diam_ft)
    except:
        pass


def _set_rect_size(elem, w_ft, h_ft):
    _set_param_if_exists(elem, "RBS_CURVE_WIDTH_PARAM", w_ft)
    _set_param_if_exists(elem, "RBS_CURVE_HEIGHT_PARAM", h_ft)


# ================================================================
#  TAPS ENGINE - BBOX / SPAN
# ================================================================
def _get_bbox(elem):
    bb = None
    try:
        bb = elem.get_BoundingBox(None)
    except:
        pass
    if not bb:
        try:
            bb = elem.get_BoundingBox(doc.ActiveView)
        except:
            bb = None
    return bb


def _bbox_span_about_origin(elem, origin, dirv):
    bb = _get_bbox(elem)
    if not bb:
        return 0.0, 0.0
    mins, maxs = bb.Min, bb.Max
    corners = [DB.XYZ(x, y, z)
               for x in (mins.X, maxs.X)
               for y in (mins.Y, maxs.Y)
               for z in (mins.Z, maxs.Z)]
    tmin, tmax = 1e9, -1e9
    for c in corners:
        t = _dot(_sub(c, origin), dirv)
        if t < tmin: tmin = t
        if t > tmax: tmax = t
    return float(tmin), float(tmax)


# ================================================================
#  TAPS ENGINE - DIRECTION & NEIGHBOR HELPERS
# ================================================================
def _outward_branch_dir_from(conn_small, host_like):
    d = connector_dir(conn_small)
    if host_like:
        away = _norm(_sub(conn_small.Origin, host_like.Origin))
        if _dot(d, away) < 0:
            d = _mul(d, -1.0)
        p0 = conn_small.Origin
        p1 = _add(p0, _mul(d, 0.1))
        if (_veclen(_sub(p1, host_like.Origin))
                < _veclen(_sub(p0, host_like.Origin))):
            d = _mul(d, -1.0)
    return _norm(d)


def _find_round_neighbor(conn):
    try:
        for r in conn.AllRefs:
            try:
                oc = r
                if oc.Owner.Id == conn.Owner.Id:
                    continue
                if oc.Shape == DB.ConnectorProfileType.Round:
                    return oc, oc.Owner
            except:
                pass
    except:
        pass
    return None, None


# ================================================================
#  TAPS ENGINE - ROUND ROLL LOCK
# ================================================================
def _pick_best_created_end(created_duct, fab_dir):
    """Pick created connector whose axis direction best matches fab_dir."""
    cons = list(iter_connectors(created_duct)) or []
    if not cons:
        return None
    best = None; bestdp = -9e9
    for c in cons:
        try:
            d = connector_dir(c)
        except:
            d = None
        if d is None:
            continue
        dp = _dot(_norm(d), _norm(fab_dir))
        if dp > bestdp:
            bestdp = dp; best = c
    return best


def _round_roll_lock(created_duct, axis_start, axis_end, fab_conn_small):
    """Lock roll of round stub so it matches fabrication connector CS."""
    try:
        axis = _norm(_sub(axis_end, axis_start))
        rot_line = DB.Line.CreateBound(axis_start, axis_end)
        doc.Regenerate()

        fab_axis = connector_dir(fab_conn_small)
        c_new = _pick_best_created_end(created_duct, fab_axis)
        if not c_new:
            cands = list(iter_connectors(created_duct)) or []
            if not cands:
                return created_duct
            c_new = min(cands,
                        key=lambda c: _veclen(_sub(c.Origin, axis_start)))

        cs_new = c_new.CoordinateSystem
        cs_fab = fab_conn_small.CoordinateSystem
        if (not cs_new) or (not cs_fab):
            return created_duct

        nx = _norm(_proj_on_plane(cs_new.BasisX, axis))
        fx = _norm(_proj_on_plane(cs_fab.BasisX, axis))
        fy = _norm(_proj_on_plane(cs_fab.BasisY, axis))

        base = _signed_angle_in_plane(nx, fx, axis)

        offsets = (0.0, math.pi/2.0, -math.pi/2.0,
                   math.pi, -math.pi,
                   3.0*math.pi/2.0, -3.0*math.pi/2.0)
        candidates = [_wrap_pi(base + off) for off in offsets]

        def err(theta):
            rx = _norm(_proj_on_plane(
                _rot_about_axis(cs_new.BasisX, axis, theta), axis))
            ry = _norm(_proj_on_plane(
                _rot_about_axis(cs_new.BasisY, axis, theta), axis))
            return (_veclen(_sub(rx, fx))**2
                    + _veclen(_sub(ry, fy))**2)

        ranked = sorted(
            ((err(t), abs(_wrap_pi(t)), t) for t in candidates),
            key=lambda x: (x[0], x[1]))
        theta = ranked[0][2]

        if abs(_deg(theta)) > MAX_ACCEPT_DEG:
            crv = created_duct.Location.Curve
            a, b = crv.GetEndPoint(0), crv.GetEndPoint(1)
            doc.Delete(created_duct.Id)
            created_duct = Mech.Duct.Create(
                doc,
                created_duct.MEPSystem.GetTypeId(),
                created_duct.DuctType.Id,
                created_duct.ReferenceLevel.Id, b, a)
            doc.Regenerate()

            c_new = (_pick_best_created_end(created_duct, fab_axis)
                     or c_new)
            cs_new = c_new.CoordinateSystem
            nx = _norm(_proj_on_plane(cs_new.BasisX, axis))
            base = _signed_angle_in_plane(nx, fx, axis)
            candidates = [_wrap_pi(base + off) for off in offsets]
            ranked = sorted(
                ((err(t), abs(_wrap_pi(t)), t) for t in candidates),
                key=lambda x: (x[0], x[1]))
            theta = ranked[0][2]

        if abs(theta) > ROLL_LOCK_TOL:
            DB.ElementTransformUtils.RotateElement(
                doc, created_duct.Id, rot_line, theta)
            doc.Regenerate()

        # Residual correction passes
        for _ in range(2):
            cons2 = list(iter_connectors(created_duct)) or []
            if not cons2:
                break
            c2 = (_pick_best_created_end(created_duct, fab_axis)
                  or min(cons2,
                         key=lambda c: _veclen(
                             _sub(c.Origin, axis_start))))
            cs2 = c2.CoordinateSystem
            if not cs2:
                break
            nx2 = _norm(_proj_on_plane(cs2.BasisX, axis))
            a_new = _signed_angle_in_plane(nx2, fx, axis)
            res = _wrap_pi(-a_new)
            if abs(res) <= RESIDUAL_LOCK_TOL:
                break
            DB.ElementTransformUtils.RotateElement(
                doc, created_duct.Id, rot_line, res)
            doc.Regenerate()

        return created_duct

    except:
        return created_duct


# ================================================================
#  TAPS ENGINE
# ================================================================
def convert_tap(part, sys_type_id, delete_original):
    """Convert one fabrication tap to a native branch duct.

    Creates a native duct along the branch axis with proper sizing
    and roll-locked orientation. Runs inside an active transaction
    from run_conversion().
    """
    cons = list(iter_connectors(part)) or []
    if len(cons) == 0:
        return None, "skip: no connectors"

    # Identify host vs branch by area (smallest = branch)
    c_sorted = sorted(cons, key=lambda c: connector_area(c))
    c_small  = c_sorted[0]
    c_big    = c_sorted[1] if len(c_sorted) > 1 else None

    shape = shape_name_from_connector(c_small)
    branch_conn_for_dir = c_small
    dt_id      = None
    size_tuple = (None, None)

    # Handle rectangular-to-round takeoffs
    if shape == "Rectangular":
        rconn, rowner = _find_round_neighbor(c_small)
        if rconn:
            shape      = "Round"
            size_tuple = (float(rconn.Radius) * 2.0, None)
            host_near  = _nearest_native_duct_cached(
                rconn.Origin, SEARCH_TOL_CURVE_FT)
            dt_id = (_guess_round_ducttype_id(host_near)
                     if host_near else _guess_round_ducttype_id())

    if dt_id is None:
        if shape == "Round":
            size_tuple = (float(branch_conn_for_dir.Radius) * 2.0, None)
            host_near = _nearest_native_duct_cached(
                branch_conn_for_dir.Origin, SEARCH_TOL_CURVE_FT)
            dt_id = (_guess_round_ducttype_id(host_near)
                     if host_near else _guess_round_ducttype_id())
        elif shape == "Rectangular":
            size_tuple = (float(c_small.Width), float(c_small.Height))
            dt_id = _guess_rect_ducttype_id()
        else:
            return None, "skip: unsupported shape"

    if not dt_id or dt_id.IntegerValue == -1:
        return None, "error: no duct type"

    ref_pt      = branch_conn_for_dir.Origin
    host_native = _nearest_native_duct_cached(ref_pt, SEARCH_TOL_CURVE_FT)
    tap_sys_id  = (_system_from_host(host_native)
                   if host_native else _first_duct_system_type())
    lvl = (_level_for_tap(host_native) if host_native
           else (_nearest_level_at_z(ref_pt.Z) or _DEFAULT_LVL))

    if not (lvl and tap_sys_id and tap_sys_id.IntegerValue != -1):
        return None, "skip: no system/level"

    # Axis direction: away from host
    if c_big is not None:
        axis_dir = _outward_branch_dir_from(c_small, c_big)
        origin   = c_big.Origin
    else:
        axis_dir = connector_dir(c_small)
        origin   = c_small.Origin

    # Span from bounding box
    tmin, tmax = _bbox_span_about_origin(part, origin, axis_dir)
    t_branch = _dot(_sub(c_small.Origin, origin), axis_dir)
    if CLAMP_BRANCH_END and t_branch is not None:
        tmax = min(tmax, t_branch - CLAMP_BACKOFF_FT)
    if tmax - tmin < MIN_LEN_FT:
        return None, "skip: too short"

    p_start = _add(origin, _mul(axis_dir, tmin))
    p_end   = _add(origin, _mul(axis_dir, tmax))

    try:
        created = Mech.Duct.Create(
            doc, tap_sys_id, dt_id, lvl.Id, p_start, p_end)

        if shape == "Round":
            desired = size_tuple[0]
            _set_round_diameter(created, desired)
            created = (_round_roll_lock(
                created, p_start, p_end, c_small) or created)
        else:
            w_ft, h_ft = size_tuple
            _set_rect_size(created, w_ft, h_ft)

        if delete_original:
            try:
                doc.Delete(part.Id)
            except:
                pass

        return created, "ok"

    except Exception as ex:
        return None, "error: {0}".format(ex)


# ================================================================
#  ENGINE DISPATCH
# ================================================================
ENGINES = {
    "Straights":   convert_straight,
    "Elbows":      convert_elbow,
    "Transitions": convert_transition,
    "Taps":        convert_tap,
}


# ================================================================
#  SELECTION HELPERS
# ================================================================
def pick_elements():
    """Interactive pick of fabrication ductwork."""
    flt = FabDuctSelFilter()
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, flt,
        "Select fabrication ductwork (click or box-select, "
        "then Finish or press Escape)")
    return [doc.GetElement(r.ElementId) for r in refs]


def collect_all_in_view():
    """All fabrication ductwork in active view."""
    viewid = doc.ActiveView.Id
    coll = (DB.FilteredElementCollector(doc, viewid)
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
            .WhereElementIsNotElementType())
    return list(coll)


def collect_visible_in_view():
    """Visible fabrication ductwork in active view."""
    viewid = doc.ActiveView.Id
    visfilter = DB.VisibleInViewFilter(doc, viewid)
    coll = (DB.FilteredElementCollector(doc, viewid)
            .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)
            .WhereElementIsNotElementType()
            .WherePasses(visfilter))
    return list(coll)


def collect_current_selection():
    """Fabrication ductwork from current Revit selection."""
    cat_id = int(DB.BuiltInCategory.OST_FabricationDuctwork)
    result = []
    for eid in uidoc.Selection.GetElementIds():
        e = doc.GetElement(eid)
        if (e and e.Category
                and e.Category.Id.IntegerValue == cat_id):
            result.append(e)
    return result


# ================================================================
#  FAILURE SUPPRESSION
# ================================================================
class _SuppressWarnings(DB.IFailuresPreprocessor):
    """Delete all warnings (including Duplicate Mark) during conversion."""
    def PreprocessFailures(self, failuresAccessor):
        has_errors = False
        try:
            for msg in failuresAccessor.GetFailureMessages():
                try:
                    sev = msg.GetSeverity()
                    if sev == DB.FailureSeverity.Warning:
                        failuresAccessor.DeleteWarning(msg)
                    else:
                        has_errors = True
                except:
                    has_errors = True
        except:
            pass
        if has_errors:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# ================================================================
#  CONVERSION RUNNER
# ================================================================
def run_conversion(elements, engine_name, delete_originals):
    """Run the selected conversion engine on elements.

    Returns list of ResultItem for the DataGrid.
    """
    engine_fn = ENGINES.get(engine_name)
    if not engine_fn:
        return [ResultItem("--", "Error",
                           "Unknown engine: {}".format(engine_name))]

    sys_type_id = get_mech_system_type_id()

    # Initialize caches for Taps engine (spatial index + type caches)
    if engine_name == "Taps":
        _init_tap_caches()
        _build_native_index()

    results = []
    ok = skip = err = 0

    tg = DB.TransactionGroup(
        doc, "Fab>Native {}".format(engine_name))
    tg.Start()

    try:
        with forms.ProgressBar(
                title="Converting {} {}...".format(
                    len(elements), engine_name.lower()),
                cancellable=True, step=1) as pb:
            for i, part in enumerate(elements):
                pb.update_progress(i, len(elements))
                if pb.cancelled:
                    results.append(ResultItem(
                        "--", "Skip", "Cancelled by user"))
                    break

                eid = part.Id.IntegerValue
                t = DB.Transaction(
                    doc, "Convert {}".format(engine_name))

                opts = t.GetFailureHandlingOptions()
                opts.SetFailuresPreprocessor(_SuppressWarnings())
                opts.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(opts)

                t.Start()
                try:
                    new_elem, status = engine_fn(
                        part, sys_type_id, delete_originals)
                    if status == "ok":
                        ok += 1
                        t.Commit()
                        new_id = (new_elem.Id.IntegerValue
                                  if new_elem else "")
                        results.append(ResultItem(
                            eid, "OK",
                            "Created {}".format(new_id)))
                    elif status.startswith("skip"):
                        skip += 1
                        t.RollBack()
                        results.append(ResultItem(
                            eid, "Skip", status))
                    else:
                        err += 1
                        t.RollBack()
                        results.append(ResultItem(
                            eid, "Error", status))
                except Exception as ex:
                    err += 1
                    if t.HasStarted() and not t.HasEnded():
                        t.RollBack()
                    results.append(ResultItem(
                        eid, "Error", str(ex)))
    finally:
        if ok > 0:
            tg.Assimilate()
        else:
            tg.RollBack()

    return results


# ================================================================
#  WINDOW CLASS (XamlReader + pick-loop)
# ================================================================
class ConvertWindow(object):
    """Main window for the unified Convert tool."""

    def __init__(self, xaml_path, preloaded_elements=None):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        self.elements     = list(preloaded_elements or [])
        self.confirmed    = False
        self._pending_pick = None

        self._get_controls()
        self._wire_events()
        self._update_element_count()
        self._update_convert_button()

        if self.elements:
            self.txt_status.Text = "{} pre-selected elements loaded.".format(
                len(self.elements))

    def _get_controls(self):
        find = self.window.FindName
        self.btn_pick_elements = find("BtnPickElements")
        self.btn_all           = find("BtnAllInView")
        self.btn_visible       = find("BtnVisible")
        self.txt_elem_count    = find("TxtElementCount")
        self.btn_clear_sel     = find("BtnClearSelection")
        self.rb_straights   = find("RbStraights")
        self.rb_elbows      = find("RbElbows")
        self.rb_transitions = find("RbTransitions")
        self.rb_taps        = find("RbTaps")
        self.chk_delete     = find("ChkDeleteOriginals")
        self.results_card   = find("ResultsCard")
        self.dg_results     = find("DgResults")
        self.txt_results    = find("TxtResultsSummary")
        self.btn_convert    = find("BtnConvert")
        self.btn_cancel     = find("BtnCancel")
        self.txt_status     = find("TxtStatus")

    def _wire_events(self):
        self.btn_pick_elements.Click += self._on_pick_elements
        self.btn_all.Click           += self._on_all_in_view
        self.btn_visible.Click       += self._on_visible
        self.btn_clear_sel.Click     += self._on_clear_selection
        self.btn_convert.Click += self._on_convert
        self.btn_cancel.Click  += self._on_cancel
        self.window.KeyDown    += self._on_key

    def _on_pick_elements(self, sender, args):
        if self.elements:
            if not forms.alert(
                    "Select different elements?",
                    title="Re-pick Elements?",
                    ok=False, yes=True, no=True):
                return
        self._pending_pick = "pick"
        self.window.Hide()

    def _on_all_in_view(self, sender, args):
        try:
            self.elements = collect_all_in_view()
        except:
            self.elements = []
        self._update_element_count()
        self._update_convert_button()
        self.txt_status.Text = "Collected {} elements from view.".format(
            len(self.elements))

    def _on_visible(self, sender, args):
        try:
            self.elements = collect_visible_in_view()
        except:
            self.elements = []
        self._update_element_count()
        self._update_convert_button()
        self.txt_status.Text = "Collected {} visible elements.".format(
            len(self.elements))

    def _on_clear_selection(self, sender, args):
        count = len(self.elements)
        self.elements = []
        self._update_element_count()
        self._update_convert_button()
        self.txt_status.Text = "Cleared {} element(s).".format(count)

    def _on_convert(self, sender, args):
        if not self.elements:
            self.txt_status.Text = "No elements selected."
            return
        self._run_conversion()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.window.Close()

    def _on_key(self, sender, args):
        if args.Key == System.Windows.Input.Key.Escape:
            self.confirmed = False
            self.window.Close()

    def _update_element_count(self):
        count = len(self.elements)
        if count == 0:
            self.txt_elem_count.Text = "(none)"
        elif count == 1:
            self.txt_elem_count.Text = "1 element"
        else:
            self.txt_elem_count.Text = "{} elements".format(count)

    def _update_convert_button(self):
        self.btn_convert.IsEnabled = len(self.elements) > 0

    def _selected_engine(self):
        if self.rb_straights.IsChecked:
            return "Straights"
        if self.rb_elbows.IsChecked:
            return "Elbows"
        if self.rb_transitions.IsChecked:
            return "Transitions"
        if self.rb_taps.IsChecked:
            return "Taps"
        return "Straights"

    def _run_conversion(self):
        engine_name = self._selected_engine()
        delete_originals = self.chk_delete.IsChecked

        self.txt_status.Text = "Converting {} {}...".format(
            len(self.elements), engine_name.lower())

        results = run_conversion(
            self.elements, engine_name, delete_originals)

        items = ObservableCollection[ResultItem]()
        ok_count = skip_count = err_count = 0
        for r in results:
            items.Add(r)
            if r.Status == "OK":
                ok_count += 1
            elif r.Status == "Skip":
                skip_count += 1
            else:
                err_count += 1

        self.dg_results.ItemsSource = items
        self.results_card.Visibility = Visibility.Visible

        parts = []
        if ok_count:
            parts.append("{} converted".format(ok_count))
        if skip_count:
            parts.append("{} skipped".format(skip_count))
        if err_count:
            parts.append("{} errors".format(err_count))
        summary = " | ".join(parts) if parts else "No results"
        self.txt_results.Text = summary

        if err_count > 0:
            self.txt_status.Text = "Done with {} error(s).".format(
                err_count)
        elif ok_count > 0:
            self.txt_status.Text = "Converted {} element(s).".format(
                ok_count)
        else:
            self.txt_status.Text = "No elements converted."

        self.elements = []
        self._update_element_count()
        self._update_convert_button()

    def show(self):
        while True:
            self._pending_pick = None
            self.window.ShowDialog()

            if self._pending_pick == "pick":
                self._do_pick_elements()
                continue
            else:
                break

        return self.confirmed

    def _do_pick_elements(self):
        try:
            new_elems = pick_elements()
            if new_elems:
                self.elements = new_elems
                self._update_element_count()
                self._update_convert_button()
                self.txt_status.Text = "Picked {} elements.".format(
                    len(self.elements))
        except Exception:
            pass


# ================================================================
#  MAIN
# ================================================================
def main():
    preselection = collect_current_selection()
    win = ConvertWindow(XAML_PATH, preselection)
    win.show()


if __name__ == "__main__":
    main()
