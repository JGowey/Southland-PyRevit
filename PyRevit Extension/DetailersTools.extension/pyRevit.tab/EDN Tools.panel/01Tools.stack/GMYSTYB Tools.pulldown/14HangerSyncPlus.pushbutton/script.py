# -*- coding: utf-8 -*-
"""Hanger Sync Plus  (v1.0.0)

Combined Lower Attachment + Upper Attachment sync for RFA (Generic Model) hangers.

LOWER ATTACHMENT — content-first workflow:
  Select duct(s) or pipe(s). Tool auto-discovers nearby Generic Model hangers.
  Writes: BOH elevation (slope-interpolated), hanger clamp/trapeze size,
          insulation thickness, XY recenter.

UPPER ATTACHMENT — slope-line workflow:
  Same auto-discovered hangers. Pick a sloped model line.
  Writes: FAMILY_TOP_LEVEL_OFFSET_PARAM so rod-top elevations follow the slope.

SYNC BOTH: runs Lower then Upper in one click.

No fabrication FilteredElementCollectors during hanger scan.
No transactions until a Sync button is pressed.

Author: Jeremiah Griffith  |  Version: 1.0.0
"""

from __future__ import division
import os, math, clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import SolidColorBrush, Color
from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc   = revit.doc
uidoc = revit.uidoc

# ---------------------------------------------------------------------------
# Content categories (what the user selects as the pipe/duct "host")
# ---------------------------------------------------------------------------

HOST_CATS = [
    DB.BuiltInCategory.OST_FabricationDuctwork,
    DB.BuiltInCategory.OST_DuctCurves,
    DB.BuiltInCategory.OST_FabricationPipework,
    DB.BuiltInCategory.OST_PipeCurves,
]
PROX_FT = 1.0   # default XY proximity radius (feet)

import json as _json

_SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "hanger_sync_settings.json")
_DEFAULT_SETTINGS = {
    "proximity_ft": 1.0, "insulation_on": False, "insulation_param": "",
    "col_family": True, "col_type": True, "col_status": False,
}

def _load_settings():
    try:
        with open(_SETTINGS_FILE, "r") as f:
            data = _json.load(f)
        s = dict(_DEFAULT_SETTINGS)
        s.update(data)
        return s
    except Exception:
        return dict(_DEFAULT_SETTINGS)

def _save_settings(s):
    try:
        with open(_SETTINGS_FILE, "w") as f:
            _json.dump(s, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Geometry helpers  (read-only, no transactions)
# ---------------------------------------------------------------------------

def get_curve(elem):
    try:
        loc = elem.Location
        return loc.Curve if hasattr(loc, "Curve") else None
    except Exception:
        return None


def xy_project_onto_curve(pt, curve):
    """
    Project pt onto curve in XY only (ignoring Z for sloped elements).
    Returns (t, projected_pt) where t is the normalized parameter [0,1]
    along the segment and projected_pt carries the interpolated Z at that t.
    Uses the unbounded line so callers can detect projections past the endpoints.
    """
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    length_xy = math.sqrt(dx * dx + dy * dy)
    if length_xy < 0.001:
        return 0.0, p0
    ux = dx / length_xy
    uy = dy / length_xy
    dot = (pt.X - p0.X) * ux + (pt.Y - p0.Y) * uy
    t   = dot / length_xy
    cx  = p0.X + dot * ux
    cy  = p0.Y + dot * uy
    cz  = p0.Z + t * (p1.Z - p0.Z)
    return t, DB.XYZ(cx, cy, cz)


def slope_label(curve):
    """Return a human-readable slope string for a model line."""
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dz_in    = (p1.Z - p0.Z) * 12.0
    horiz_in = ((p1.X - p0.X) ** 2 + (p1.Y - p0.Y) ** 2) ** 0.5 * 12.0
    length_in = curve.Length * 12.0
    if horiz_in > 0.001:
        slope_str = '{:.3f}" per 12" ({})'.format(
            abs(dz_in / horiz_in) * 12.0,
            "up" if dz_in > 0 else "down")
    else:
        slope_str = "vertical"
    return '{:.1f}" long, {}'.format(length_in, slope_str)


# ---------------------------------------------------------------------------
# Lower attachment helpers
# ---------------------------------------------------------------------------

def get_bottom_z(elem):
    """
    Return bottom-of-body elevation in feet.
    Prefers 'Lower End Bottom Elevation' (fabrication params) over bounding box
    because connection flanges extend below the body and inflate bbox.Min.Z.
    """
    try:
        p = elem.LookupParameter("Lower End Bottom Elevation")
        if p and p.StorageType == DB.StorageType.Double and p.AsDouble() != 0.0:
            return p.AsDouble()
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        return bb.Min.Z if bb else None
    except Exception:
        return None


def get_pipe_bottom_at_t(content, t):
    """
    Return the bottom elevation at normalized position t along the content slope.

    For fabrication parts Revit pre-computes Lower/Upper End Bottom Elevation
    for each end; interpolating between them at t is accurate for round,
    rectangular, and insulated profiles without any profile math on our part.

    For native duct/pipe falls back to centerline Z minus half profile depth.
    t=0 maps to GetEndPoint(0), t=1 maps to GetEndPoint(1).
    """
    try:
        cat         = content.Category.Id.IntegerValue
        is_fab_pipe = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)

        if is_fab_pipe or is_fab_duct:
            p_lo = content.LookupParameter("Lower End Bottom Elevation")
            p_hi = content.LookupParameter("Upper End Bottom Elevation")
            if (p_lo and p_hi and
                    p_lo.StorageType == DB.StorageType.Double and
                    p_hi.StorageType == DB.StorageType.Double):
                lo = p_lo.AsDouble()
                hi = p_hi.AsDouble()
                return lo + t * (hi - lo)
            return get_bottom_z(content)

        curve = get_curve(content)
        if curve is None:
            return get_bottom_z(content)
        p0   = curve.GetEndPoint(0)
        p1   = curve.GetEndPoint(1)
        cl_z = p0.Z + t * (p1.Z - p0.Z)

        is_nat_pipe = cat == int(DB.BuiltInCategory.OST_PipeCurves)
        is_nat_duct = cat == int(DB.BuiltInCategory.OST_DuctCurves)
        half_depth  = 0.0
        if is_nat_pipe:
            for pname in ("Outside Diameter", "Diameter"):
                p = content.LookupParameter(pname)
                if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                    half_depth = p.AsDouble() / 2.0
                    break
        elif is_nat_duct:
            p = content.LookupParameter("Height")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                half_depth = p.AsDouble() / 2.0
        return cl_z - half_depth

    except Exception:
        return get_bottom_z(content)


def xy_dist(pt, elem):
    try:
        curve = get_curve(elem)
        if curve:
            _t, p = xy_project_onto_curve(pt, curve)
        elif hasattr(elem.Location, "Point"):
            p = elem.Location.Point
        else:
            return float("inf")
        return math.sqrt((pt.X - p.X) ** 2 + (pt.Y - p.Y) ** 2)
    except Exception:
        return float("inf")


def get_size(elem):
    """Human-readable nominal size string (no insulation)."""
    try:
        cat = elem.Category.Id.IntegerValue
        is_fab_pipe = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)
        if is_fab_pipe:
            p = elem.LookupParameter("Product Entry")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip().replace('"', "")
                if s:
                    return s + '"'
            p = elem.LookupParameter("Main Primary Diameter")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return '%s"' % int(round(p.AsDouble() * 12))
        if is_fab_duct:
            p = elem.LookupParameter("Overall Size")
            if p and p.StorageType == DB.StorageType.String:
                return p.AsString() or ""
        parts = []
        for n in ("Width", "Height", "Diameter", "Outside Diameter"):
            p = elem.LookupParameter(n)
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                parts.append('%s"' % int(round(p.AsDouble() * 12)))
        return "x".join(parts)
    except Exception:
        return ""


def get_nominal_dia(content):
    """Return nominal diameter in feet, or None."""
    try:
        cat = content.Category.Id.IntegerValue
        is_fab_pipe = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)
        if is_fab_pipe:
            p = content.LookupParameter("Product Entry")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip().replace('"', "").strip()
                if s:
                    try: return float(s) / 12.0
                    except Exception: pass
            p = content.LookupParameter("Main Primary Diameter")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()
        if is_fab_duct:
            p = content.LookupParameter("Overall Size")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip().lower().replace(" ", "")
                import re as _re
                if "x" in s:
                    nums = []
                    for part in s.split("x"):
                        digits = _re.sub(r"[^0-9.]", "", part)
                        if digits: nums.append(float(digits))
                    if nums: return max(nums) / 12.0
                else:
                    digits = _re.sub(r"[^0-9.]", "", s)
                    if digits: return float(digits) / 12.0
        for name in ("Diameter", "Outside Diameter"):
            p = content.LookupParameter(name)
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()
        pw = content.LookupParameter("Width")
        ph = content.LookupParameter("Height")
        if pw and ph and pw.StorageType == DB.StorageType.Double:
            w = pw.AsDouble()
            h = ph.AsDouble()
            if w > 0 or h > 0: return max(w, h)
    except Exception:
        pass
    return None


def get_duct_width(content):
    """Return duct width in feet for CP_Trapeze Width Desired, or None."""
    try:
        cat = content.Category.Id.IntegerValue
        if cat == int(DB.BuiltInCategory.OST_FabricationDuctwork):
            p = content.LookupParameter("Main Primary Width")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()
        elif cat == int(DB.BuiltInCategory.OST_DuctCurves):
            p = content.LookupParameter("Width")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()
    except Exception:
        pass
    return None


def get_insulation_thickness(content):
    """
    Derive insulation thickness in feet from fabrication end-elevation params.
    Lower End Bottom Elevation minus Lower End Bottom of Insulation Elevation.
    Falls back to top-of-insulation minus top-of-pipe.
    """
    try:
        p_bot     = content.LookupParameter("Lower End Bottom Elevation")
        p_bot_ins = content.LookupParameter("Lower End Bottom of Insulation Elevation")
        if (p_bot and p_bot_ins and
                p_bot.StorageType == DB.StorageType.Double and
                p_bot_ins.StorageType == DB.StorageType.Double):
            thickness = p_bot.AsDouble() - p_bot_ins.AsDouble()
            if thickness >= 0: return thickness
        p_top     = content.LookupParameter("Lower End Top Elevation")
        p_top_ins = content.LookupParameter("Lower End Top of Insulation Elevation")
        if (p_top and p_top_ins and
                p_top.StorageType == DB.StorageType.Double and
                p_top_ins.StorageType == DB.StorageType.Double):
            thickness = p_top_ins.AsDouble() - p_top.AsDouble()
            if thickness >= 0: return thickness
    except Exception:
        pass
    return 0.0


def write_elevation(hanger, bottom_z):
    """Write bottom_z to the hanger. Tries param names in priority order."""
    for name in ("BOHElevation", "CP_Bottom Elevation", "CP_Hung Object Bottom Elev"):
        p = hanger.LookupParameter(name)
        if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
            p.Set(bottom_z)
            return name
    return None


def write_insulation(hanger, thickness_ft):
    """Write insulation thickness to CP_Host Insulation Thickness and CP_Hanger Insulation Adjust."""
    written = []
    p = hanger.LookupParameter("CP_Host Insulation Thickness")
    if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
        p.Set(thickness_ft)
        written.append("CP_Host Insulation Thickness")
    p2 = hanger.LookupParameter("CP_Hanger Insulation Adjust")
    if p2 and p2.StorageType == DB.StorageType.String and not p2.IsReadOnly:
        p2.Set("%.4g" % (thickness_ft * 12) if thickness_ft > 0 else "0")
        written.append("CP_Hanger Insulation Adjust")
    return written[0] if written else None


def write_trapeze_width(hanger, width_ft):
    p = hanger.LookupParameter("CP_Trapeze Width Desired")
    if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
        p.Set(width_ft)
        return "CP_Trapeze Width Desired"
    return None


def write_size(hanger, nominal_dia_ft):
    """Write to CP_Hung Object Diameter and CP_Host Nominal Diameter. Can trigger nested regen."""
    written = []
    for name in ("CP_Hung Object Diameter", "CP_Host Nominal Diameter"):
        p = hanger.LookupParameter(name)
        if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
            p.Set(nominal_dia_ft)
            written.append(name)
    return written


def recenter_xy(hanger, content):
    """
    Move hanger perpendicular-only to the pipe centerline (XY).
    Uses the unbounded line projection so hangers near segment ends
    don't jump to the coupler joint.
    """
    try:
        hanger_loc = hanger.Location
        if not hasattr(hanger_loc, "Point"):
            return False
        pt = hanger_loc.Point
        content_loc = content.Location
        if hasattr(content_loc, "Curve"):
            curve  = content_loc.Curve
            p0     = curve.GetEndPoint(0)
            p1     = curve.GetEndPoint(1)
            dx_p   = p1.X - p0.X
            dy_p   = p1.Y - p0.Y
            pipe_l = (dx_p ** 2 + dy_p ** 2) ** 0.5
            if pipe_l < 0.001: return False
            ux = dx_p / pipe_l
            uy = dy_p / pipe_l
            dot    = (pt.X - p0.X) * ux + (pt.Y - p0.Y) * uy
            cx     = p0.X + dot * ux
            cy     = p0.Y + dot * uy
            move_x = cx - pt.X
            move_y = cy - pt.Y
        elif hasattr(content_loc, "Point"):
            target = content_loc.Point
            move_x = target.X - pt.X
            move_y = target.Y - pt.Y
        else:
            return False
        if (move_x ** 2 + move_y ** 2) < 0.0001:
            return False
        DB.ElementTransformUtils.MoveElement(doc, hanger.Id, DB.XYZ(move_x, move_y, 0.0))
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Hanger discovery  (Generic Model only — never fabrication collectors)
# ---------------------------------------------------------------------------

def find_hangers_near(content_elem, prox_ft=None):
    """
    Return list of (hanger_elem, xy_dist, t) for Generic Model FamilyInstances
    within prox_ft XY of content's centerline, on-segment only.
    """
    if prox_ft is None:
        prox_ft = PROX_FT
    results = []
    try:
        bb = content_elem.get_BoundingBox(None)
        if bb is None:
            return results
        mn = DB.XYZ(bb.Min.X - prox_ft, bb.Min.Y - prox_ft, bb.Min.Z - 2.0)
        mx = DB.XYZ(bb.Max.X + prox_ft, bb.Max.Y + prox_ft, bb.Max.Z + 15.0)
        bb_filter  = DB.BoundingBoxIntersectsFilter(DB.Outline(mn, mx))
        candidates = (DB.FilteredElementCollector(doc)
                      .OfCategory(DB.BuiltInCategory.OST_GenericModel)
                      .WhereElementIsNotElementType()
                      .WherePasses(bb_filter))
        curve = get_curve(content_elem)
        for elem in candidates:
            if not isinstance(elem, DB.FamilyInstance):
                continue
            try:
                pt = elem.Location.Point
            except Exception:
                continue
            if curve is not None:
                try:
                    t, proj_xy = xy_project_onto_curve(pt, curve)
                    d = math.sqrt((pt.X - proj_xy.X) ** 2 + (pt.Y - proj_xy.Y) ** 2)
                    if d > prox_ft: continue
                    if t < -0.01 or t > 1.01: continue
                except Exception:
                    continue
            else:
                d = xy_dist(pt, content_elem)
                t = 0.5
                if d > PROX_FT: continue
            results.append((elem, d, t))
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

class LowerRow(object):
    """One row for the Lower Attachment tab (content → hanger pair)."""

    def __init__(self, content, hanger, dist_ft, t=0.5):
        self.ContentId    = str(content.Id.IntegerValue)
        self.ContentType  = content.Category.Name if content.Category else ""
        self.ContentSize  = get_size(content)
        self.HangerId     = str(hanger.Id.IntegerValue)
        self.HangerFamily = ""
        self.HangerType   = ""
        try:
            type_id = hanger.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type:
                    if hasattr(elem_type, 'FamilyName'):
                        try: self.HangerFamily = str(elem_type.FamilyName) or ""
                        except Exception: pass
                    if not self.HangerFamily and hanger.Category:
                        self.HangerFamily = str(hanger.Category.Name) or ""
                    try:
                        if hasattr(elem_type, 'Name'):
                            self.HangerType = str(elem_type.Name) or ""
                    except Exception: pass
                    if not self.HangerType:
                        try:
                            p = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                            if p: self.HangerType = p.AsString() or ""
                        except Exception: pass
        except Exception:
            pass
        self.DistFt         = "%.2f'" % dist_ft
        self.DistWarning    = dist_ft > 0.5
        self._content       = content
        self._hanger        = hanger
        self._t             = t
        self._bottom_z      = get_pipe_bottom_at_t(content, t)
        self._nominal_dia   = get_nominal_dia(content)
        self._duct_width    = get_duct_width(content)
        self._insulation_ft = get_insulation_thickness(content)
        self.Insulation     = ""
        self.Status         = "Ready" if self._bottom_z is not None else "No elevation"


# Upper attachment status constants
_UP_READY    = "Ready"
_UP_OUTSIDE  = "Outside Range"
_UP_NO_LEVEL = "No Top Level"
_UP_READONLY = "Read Only"
_UP_SYNCED   = "Synced"
_UP_NO_LINE  = "No slope line"

_UP_COLORS = {
    _UP_READY:    Color.FromArgb(255,  76, 175,  80),   # green
    _UP_OUTSIDE:  Color.FromArgb(255, 255, 152,   0),   # orange
    _UP_NO_LEVEL: Color.FromArgb(255, 244,  67,  54),   # red
    _UP_READONLY: Color.FromArgb(255, 244,  67,  54),   # red
    _UP_SYNCED:   Color.FromArgb(255, 187, 134, 252),   # purple
    _UP_NO_LINE:  Color.FromArgb(255, 158, 158, 158),   # gray
}

def _brush(color):
    return SolidColorBrush(color)


class UpperRow(object):
    """One row for the Upper Attachment tab (hanger + slope line)."""

    def __init__(self, hanger_id, slope_curve):
        self.element_id = hanger_id
        hanger = doc.GetElement(hanger_id)

        try:
            type_id   = hanger.GetTypeId()
            elem_type = doc.GetElement(type_id)
            try:    self.Family = elem_type.FamilyName or "?"
            except Exception: self.Family = "?"
            try:    self.Type = elem_type.Name or "?"
            except Exception:
                try:
                    p = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    self.Type = p.AsString() if p else "?"
                except Exception: self.Type = "?"
        except Exception:
            self.Family = "?"
            self.Type   = "?"

        self.HangerID      = str(hanger_id.IntegerValue)
        self.TopLevel      = "?"
        self.CurrentOffset = "?"
        self.NewOffset     = "?"
        self._status       = _UP_NO_LEVEL
        self._new_offset_ft = None

        if hanger is None:
            self._finalize(); return

        loc = hanger.Location
        if not isinstance(loc, DB.LocationPoint):
            self._finalize(); return

        p_top_lv = hanger.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
        if p_top_lv is None:
            self._finalize(); return
        top_lv = doc.GetElement(p_top_lv.AsElementId())
        if not isinstance(top_lv, DB.Level):
            self._finalize(); return
        self.TopLevel   = top_lv.Name
        top_level_z     = top_lv.Elevation

        p_offset = hanger.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
        if p_offset is None or p_offset.IsReadOnly:
            self._status = _UP_READONLY
            self._finalize(); return
        cur_off_ft         = p_offset.AsDouble()
        self.CurrentOffset = '{:.3f}"'.format(cur_off_ft * 12.0)

        if slope_curve is None:
            self._status = _UP_NO_LINE
            self._finalize(); return

        pt    = loc.Point
        t_val, proj_pt = xy_project_onto_curve(pt, slope_curve)
        if t_val < -0.01 or t_val > 1.01:
            self._status = _UP_OUTSIDE
            self._finalize(); return

        new_off_ft          = proj_pt.Z - top_level_z
        self._new_offset_ft = new_off_ft
        self.NewOffset      = '{:.3f}"'.format(new_off_ft * 12.0)
        self._status        = _UP_READY
        self._finalize()

    def _finalize(self):
        c = _UP_COLORS.get(self._status, Color.FromArgb(255, 200, 200, 200))
        self.StatusColor = _brush(c)
        self.Status      = self._status

    def mark_synced(self):
        self._status = _UP_SYNCED
        self._finalize()


def _build_upper_rows(hanger_ids, slope_curve):
    col = ObservableCollection[object]()
    for hid in hanger_ids:
        col.Add(UpperRow(hid, slope_curve))
    return col


# ---------------------------------------------------------------------------
# SI_Tools warning suppressor  (Upper Attachment)
# ---------------------------------------------------------------------------

def _auto_click_yes():
    """
    Spin a background thread that finds the SI_Tools warning window and sends
    Alt+Y via keybd_event. Must be called immediately before t.Commit().
    The GIL releases when Commit() blocks on the native modal.
    """
    import threading, ctypes

    def _watcher():
        import time
        user32 = ctypes.windll.user32
        VK_MENU = 0x12
        VK_Y    = 0x59
        KEYEVENTF_KEYUP = 0x0002
        for _ in range(150):
            hwnd = user32.FindWindowW(None, u"SI_Tools - Warning")
            if hwnd:
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                user32.keybd_event(VK_MENU, 0, 0, 0)
                user32.keybd_event(VK_Y,    0, 0, 0)
                user32.keybd_event(VK_Y,    0, KEYEVENTF_KEYUP, 0)
                user32.keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0)
                return
            time.sleep(0.1)

    t = threading.Thread(target=_watcher)
    t.daemon = True
    t.start()


# ---------------------------------------------------------------------------
# Selection filters
# ---------------------------------------------------------------------------

class HostFilter(ISelectionFilter):
    def AllowElement(self, e):
        if e is None or e.Category is None: return False
        return any(e.Category.Id.IntegerValue == int(c) for c in HOST_CATS)
    def AllowReference(self, r, p): return False


class HangerFilter(ISelectionFilter):
    def AllowElement(self, e):
        if e is None or e.Category is None: return False
        if e.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel): return False
        return isinstance(e, DB.FamilyInstance)
    def AllowReference(self, r, p): return False


class ModelLineFilter(ISelectionFilter):
    def AllowElement(self, elem): return isinstance(elem, DB.CurveElement)
    def AllowReference(self, ref, pt): return False


# ---------------------------------------------------------------------------
# Sync functions  (all model writes go here)
# ---------------------------------------------------------------------------

def sync_lower_rows(rows, get_insul_ft_fn):
    """
    Write BOH elevation, hanger size, insulation, and XY for all Ready rows.
    One transaction per hanger to avoid a single failure cascading.
    Returns (done, errors).
    """
    done = errors = 0
    for row in rows:
        if row.Status not in ("Ready", "Ready (picked)"):
            continue
        if row._bottom_z is None:
            continue
        t = DB.Transaction(doc, "Hanger Sync Lower - %s" % row.HangerId)
        try:
            t.Start()
            ins_ft        = get_insul_ft_fn(row._content) or 0.0
            hung_bottom_z = row._bottom_z - ins_ft
            elev_param    = write_elevation(row._hanger, hung_bottom_z)
            size_params   = []
            if row._nominal_dia is not None:
                insulated_dia = row._nominal_dia + 2.0 * ins_ft
                try:
                    size_params = write_size(row._hanger, insulated_dia)
                except Exception:
                    pass
            if getattr(row, "_duct_width", None) is not None:
                try:
                    tw = write_trapeze_width(row._hanger, row._duct_width)
                    if tw and tw not in size_params: size_params.append(tw)
                except Exception:
                    pass
            xy_moved = False
            try:
                xy_moved = recenter_xy(row._hanger, row._content)
            except Exception:
                pass
            if ins_ft > 0:
                try:
                    write_insulation(row._hanger, ins_ft)
                except Exception:
                    pass
            if elev_param:
                t.Commit()
                parts = [elev_param] + size_params if size_params else [elev_param]
                if ins_ft > 0: parts.append("ins=%.4g\"" % (ins_ft * 12))
                if xy_moved:   parts.append("XY")
                row.Status = "Synced: " + ", ".join(parts)
                done += 1
            else:
                t.RollBack()
                row.Status = "No writable elevation param"
                errors += 1
        except Exception as ex:
            try:
                if t.GetStatus() == DB.TransactionStatus.Started:
                    t.RollBack()
            except Exception:
                pass
            row.Status = "Error: %s" % str(ex)[:50]
            errors += 1
    return done, errors


def sync_upper_rows(upper_rows):
    """
    Write Top Offset for all Ready rows in a single transaction.
    SI_Tools warning suppressor starts before Commit().
    Returns (synced, skipped).
    """
    synced = skipped = 0
    try:
        t = DB.Transaction(doc, "Hanger Sync Upper")
        t.Start()
        for row in upper_rows:
            if row._status != _UP_READY or row._new_offset_ft is None:
                skipped += 1
                continue
            hanger   = doc.GetElement(row.element_id)
            p_offset = hanger.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
            if p_offset is None or p_offset.IsReadOnly:
                skipped += 1
                continue
            p_offset.Set(row._new_offset_ft)
            row.mark_synced()
            synced += 1
        _auto_click_yes()
        t.Commit()
    except Exception as ex:
        try: t.RollBack()
        except Exception: pass
        forms.alert("Upper sync error:\n{}".format(str(ex)), title="Hanger Sync")
    return synced, skipped


# ---------------------------------------------------------------------------
# SelectContentWindow  (inline XAML — temp file pattern)
# ---------------------------------------------------------------------------

class SelectContentWindow(forms.WPFWindow):

    XAML = """<?xml version="1.0" encoding="utf-8"?>
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Hanger Sync Plus" Width="520" Height="310"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
    FontFamily="Segoe UI" Background="#F5F7FB">
    <Grid>
        <Border VerticalAlignment="Top" Height="4" Background="#8F6C05"/>
        <StackPanel Margin="28,20,28,16">

            <!-- Title row -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <Border Width="36" Height="36" CornerRadius="18"
                        Background="#FDF5DC" Margin="0,0,12,0">
                    <TextBlock Text="+" FontSize="22" FontWeight="Bold"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               Foreground="#8F6C05"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Hanger Sync Plus" FontSize="15" FontWeight="SemiBold"
                               Foreground="#0F1729"/>
                    <TextBlock x:Name="txtSubtitle" Text="Choose a workflow"
                               FontSize="11" Foreground="#6B7280"/>
                </StackPanel>
            </StackPanel>

            <!-- Three workflow cards -->
            <Grid Margin="0,0,0,14">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>

                <!-- Pick Hangers -->
                <Border Grid.Column="0" Background="White" CornerRadius="6"
                        BorderBrush="#E5E7EB" BorderThickness="1" Padding="10,10">
                    <StackPanel>
                        <TextBlock Text="[ H ]" FontSize="13" FontWeight="Bold"
                                   Foreground="#8F6C05"
                                   HorizontalAlignment="Center" Margin="0,0,0,4"/>
                        <TextBlock Text="Pick Hangers" FontSize="11" FontWeight="SemiBold"
                                   Foreground="#1F2937" HorizontalAlignment="Center"
                                   Margin="0,0,0,4"/>
                        <TextBlock Text="Select hangers then pick a slope line -- Upper Attachment only"
                                   FontSize="10" Foreground="#6B7280" TextWrapping="Wrap"
                                   HorizontalAlignment="Center" TextAlignment="Center"
                                   Margin="0,0,0,10"/>
                        <Button x:Name="btnPickHangers" Content="Pick Hangers"
                                Height="26" Background="#8F6C05" Foreground="White"
                                BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                                Click="btnPickHangers_Click"/>
                    </StackPanel>
                </Border>

                <!-- Pick in Viewport -->
                <Border Grid.Column="2" Background="White" CornerRadius="6"
                        BorderBrush="#E5E7EB" BorderThickness="1" Padding="10,10">
                    <StackPanel>
                        <TextBlock Text="[ V ]" FontSize="13" FontWeight="Bold"
                                   Foreground="#8F6C05"
                                   HorizontalAlignment="Center" Margin="0,0,0,4"/>
                        <TextBlock Text="Pick in Viewport" FontSize="11" FontWeight="SemiBold"
                                   Foreground="#1F2937" HorizontalAlignment="Center"
                                   Margin="0,0,0,4"/>
                        <TextBlock Text="Select duct or pipe -- auto-discovers nearby hangers"
                                   FontSize="10" Foreground="#6B7280" TextWrapping="Wrap"
                                   HorizontalAlignment="Center" TextAlignment="Center"
                                   Margin="0,0,0,10"/>
                        <Button x:Name="btnPick" Content="Pick in Viewport"
                                Height="26" Background="#8F6C05" Foreground="White"
                                BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                                Click="btnPick_Click"/>
                    </StackPanel>
                </Border>

                <!-- Pick Part + Hanger -->
                <Border Grid.Column="4" Background="White" CornerRadius="6"
                        BorderBrush="#E5E7EB" BorderThickness="1" Padding="10,10">
                    <StackPanel>
                        <TextBlock Text="[ P ]" FontSize="13" FontWeight="Bold"
                                   Foreground="#5C6BC0"
                                   HorizontalAlignment="Center" Margin="0,0,0,4"/>
                        <TextBlock Text="Pick Part + Hanger" FontSize="11" FontWeight="SemiBold"
                                   Foreground="#1F2937" HorizontalAlignment="Center"
                                   Margin="0,0,0,4"/>
                        <TextBlock Text="Pick one duct/pipe then one hanger -- single pair sync"
                                   FontSize="10" Foreground="#6B7280" TextWrapping="Wrap"
                                   HorizontalAlignment="Center" TextAlignment="Center"
                                   Margin="0,0,0,10"/>
                        <Button x:Name="btnPickPair" Content="Pick Part + Hanger"
                                Height="26" Background="#5C6BC0" Foreground="White"
                                BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                                Click="btnPickPair_Click"/>
                    </StackPanel>
                </Border>
            </Grid>

            <!-- Footer -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnUseSel" Content="Use Current Selection"
                        Height="26" Margin="0,0,10,0" Padding="12,0"
                        Visibility="Collapsed"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                        Click="btnUseSel_Click"/>
                <Button x:Name="btnCancel" Content="Cancel"
                        Width="76" Height="26"
                        Background="White" BorderBrush="#D1D5DB"
                        BorderThickness="1" Cursor="Hand"
                        Click="btnCancel_Click"/>
            </StackPanel>

        </StackPanel>
    </Grid>
</Window>"""

    def __init__(self, sel_count=0):
        import tempfile, io
        tmp = tempfile.NamedTemporaryFile(suffix=".xaml", delete=False)
        tmp.close()
        with io.open(tmp.name, "w", encoding="utf-8") as f:
            f.write(self.XAML)
        self._xaml_path   = tmp.name
        self._sel_count   = sel_count
        self.picked       = False
        self.pick_pair    = False
        self.pick_hangers = False
        self.use_sel      = False
        forms.WPFWindow.__init__(self, self._xaml_path)
        self._apply_mode()

    def _apply_mode(self):
        from System.Windows import Visibility
        if self._sel_count > 0:
            self.txtSubtitle.Text     = "%d element(s) already selected" % self._sel_count
            self.btnUseSel.Visibility = Visibility.Visible
        else:
            self.txtSubtitle.Text     = "Choose a workflow"
            self.btnUseSel.Visibility = Visibility.Collapsed

    def btnUseSel_Click(self, sender, e):      self.use_sel      = True; self.Close()
    def btnPick_Click(self, sender, e):        self.picked       = True; self.Close()
    def btnPickHangers_Click(self, sender, e): self.pick_hangers = True; self.Close()
    def btnPickPair_Click(self, sender, e):    self.pick_pair    = True; self.Close()
    def btnCancel_Click(self, sender, e):      self.Close()

    def show(self):
        self.ShowDialog()
        try:
            os.unlink(self._xaml_path)
        except Exception:
            pass
        return self.picked, self.pick_pair, self.use_sel, self.pick_hangers


# ---------------------------------------------------------------------------
# SelectLineWindow  (notice shown before the Pick Line pick context)
# ---------------------------------------------------------------------------

class SelectLineWindow(forms.WPFWindow):

    XAML = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Hanger Sync" Width="420" Height="230"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
    FontFamily="Segoe UI" Background="#F5F7FB">
    <Grid>
        <Border VerticalAlignment="Top" Height="4" Background="#8F6C05"/>
        <StackPanel Margin="32,22,32,20" VerticalAlignment="Top">
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <Border Width="40" Height="40" CornerRadius="20"
                        Background="#FDF5DC" Margin="0,0,14,0">
                    <TextBlock Text="&#x2571;" FontSize="20"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               Foreground="#8F6C05"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Upper Attachment" FontSize="16" FontWeight="SemiBold"
                               Foreground="#0F1729"/>
                    <TextBlock Text="Slope line selection required"
                               FontSize="11" Foreground="#6B7280"/>
                </StackPanel>
            </StackPanel>
            <Border Background="White" CornerRadius="6"
                    Padding="16,12" Margin="0,0,0,20"
                    BorderBrush="#E5E7EB" BorderThickness="1">
                <StackPanel Orientation="Horizontal">
                    <Border Width="22" Height="22" CornerRadius="11"
                            Background="#8F6C05" Margin="0,0,10,0">
                        <TextBlock Text="1" FontSize="11" FontWeight="Bold"
                                   Foreground="White"
                                   HorizontalAlignment="Center"
                                   VerticalAlignment="Center"/>
                    </Border>
                    <TextBlock VerticalAlignment="Center"
                               Foreground="#1F2937" FontSize="12" TextWrapping="Wrap">
                        Click a <Bold>sloped model line</Bold> in the viewport
                    </TextBlock>
                </StackPanel>
            </Border>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnCancel" Content="Cancel"
                        Width="80" Height="30" Margin="0,0,10,0"
                        Background="White" BorderBrush="#D1D5DB"
                        BorderThickness="1" Cursor="Hand"
                        Click="btnCancel_Click"/>
                <Button x:Name="btnPick" Content="Pick in Viewport"
                        Width="130" Height="30"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                        Click="btnPick_Click"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>"""

    def __init__(self):
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".xaml", delete=False, mode="w")
        tmp.write(self.XAML)
        tmp.close()
        self._xaml_path = tmp.name
        self.picked     = False
        forms.WPFWindow.__init__(self, self._xaml_path)

    def btnPick_Click(self, sender, e):   self.picked = True; self.Close()
    def btnCancel_Click(self, sender, e): self.Close()

    def show(self):
        self.ShowDialog()
        try: os.unlink(self._xaml_path)
        except Exception: pass
        return self.picked


# ---------------------------------------------------------------------------
# InsulationChoiceWindow  (quick-pair workflow only)
# ---------------------------------------------------------------------------

class InsulationChoiceWindow(forms.WPFWindow):

    XAML = """<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Insulation Setting" Width="440" Height="340"
    WindowStartupLocation="CenterScreen" ResizeMode="NoResize"
    FontFamily="Segoe UI" Background="#F5F7FB">
    <Grid>
        <Border VerticalAlignment="Top" Height="4" Background="#8F6C05"/>
        <StackPanel Margin="28,20,28,16" VerticalAlignment="Top">
            <StackPanel Orientation="Horizontal" Margin="0,0,0,14">
                <Border Width="36" Height="36" CornerRadius="18"
                        Background="#FDF5DC" Margin="0,0,12,0">
                    <TextBlock Text="&#x29C9;" FontSize="18"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               Foreground="#8F6C05"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Insulation detected" FontSize="14" FontWeight="SemiBold"
                               Foreground="#0F1729"/>
                    <TextBlock x:Name="txtThickness" Text="Thickness: unknown"
                               FontSize="11" Foreground="#6B7280"/>
                </StackPanel>
            </StackPanel>
            <Border x:Name="pnlParamPicker" Background="White" CornerRadius="6"
                    Padding="12,10" Margin="0,0,0,12"
                    BorderBrush="#E5E7EB" BorderThickness="1"
                    Visibility="Collapsed">
                <StackPanel>
                    <TextBlock Text="Select the parameter that holds insulation thickness:"
                               FontSize="11" Foreground="#374151" Margin="0,0,0,6"
                               TextWrapping="Wrap"/>
                    <ComboBox x:Name="cmbParam" Height="26"/>
                </StackPanel>
            </Border>
            <Border Background="White" CornerRadius="6"
                    Padding="12,10" Margin="0,0,0,16"
                    BorderBrush="#E5E7EB" BorderThickness="1">
                <StackPanel>
                    <TextBlock Text="Where should the hanger snap to?"
                               FontSize="11" FontWeight="SemiBold"
                               Foreground="#374151" Margin="0,0,0,8"/>
                    <RadioButton x:Name="rdoOutsideInsulation"
                                 Content="Outside of insulation  (hanger wraps insulated OD)"
                                 IsChecked="True" Margin="0,0,0,6" Foreground="#1F2937"/>
                    <RadioButton x:Name="rdoPipeSurface"
                                 Content="Pipe / duct surface  (bare metal)"
                                 Foreground="#1F2937"/>
                </StackPanel>
            </Border>
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnCancel" Content="Cancel"
                        Width="80" Height="28" Margin="0,0,10,0"
                        Background="White" BorderBrush="#D1D5DB"
                        BorderThickness="1" Cursor="Hand" Click="btnCancel_Click"/>
                <Button x:Name="btnApply" Content="Apply"
                        Width="90" Height="28"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand" FontWeight="SemiBold"
                        Click="btnApply_Click"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>"""

    def __init__(self, content_elem, detected_thickness_ft):
        import tempfile, io
        tmp = tempfile.NamedTemporaryFile(suffix=".xaml", delete=False)
        tmp.close()
        with io.open(tmp.name, "w", encoding="utf-8") as f:
            f.write(self.XAML)
        self._xaml_path          = tmp.name
        self._content            = content_elem
        self._detected_thickness = detected_thickness_ft
        self.snap_to_insulation  = True
        self.insulation_ft       = detected_thickness_ft or 0.0
        self.confirmed           = False
        forms.WPFWindow.__init__(self, self._xaml_path)
        self._setup()

    def _setup(self):
        if self._detected_thickness and self._detected_thickness > 0:
            self.txtThickness.Text = "Thickness: %.4g in" % (self._detected_thickness * 12)
        else:
            from System.Windows import Visibility
            self.pnlParamPicker.Visibility = Visibility.Visible
            self.txtThickness.Text = "Thickness: could not be determined"
            self._populate_params()

    def _populate_params(self):
        self.cmbParam.Items.Clear()
        param_names = set()
        try:
            for p in self._content.Parameters:
                if p.StorageType == DB.StorageType.Double:
                    param_names.add(p.Definition.Name)
        except Exception:
            pass
        def sort_key(n):
            nl = n.lower()
            priority = 0 if any(kw in nl for kw in ("insul","lining","thickness")) else 1
            return (priority, nl)
        for name in sorted(param_names, key=sort_key):
            self.cmbParam.Items.Add(name)
        if self.cmbParam.Items.Count > 0:
            self.cmbParam.SelectedIndex = 0

    def btnApply_Click(self, sender, e):
        self.snap_to_insulation = bool(self.rdoOutsideInsulation.IsChecked)
        if self._detected_thickness and self._detected_thickness > 0:
            self.insulation_ft = self._detected_thickness
        else:
            try:
                param_name = self.cmbParam.SelectedItem
                if param_name:
                    p = self._content.LookupParameter(param_name)
                    if p and p.StorageType == DB.StorageType.Double:
                        self.insulation_ft = p.AsDouble()
            except Exception:
                self.insulation_ft = 0.0
        self.confirmed = True
        self.Close()

    def btnCancel_Click(self, sender, e): self.Close()

    def show(self):
        self.ShowDialog()
        try: os.unlink(self._xaml_path)
        except Exception: pass
        return self.confirmed


# ---------------------------------------------------------------------------
# Pick-pair workflow  (single part + single hanger, no auto-discovery)
# ---------------------------------------------------------------------------

def run_pick_pair_workflow():
    try:
        ref_content = uidoc.Selection.PickObject(
            ObjectType.Element, HostFilter(), "Step 1 of 2 — Pick the duct or pipe")
        content = doc.GetElement(ref_content.ElementId)
    except Exception:
        return
    if content is None:
        forms.alert("Could not read selected element.", title="Hanger Sync"); return

    try:
        ref_hanger = uidoc.Selection.PickObject(
            ObjectType.Element, HangerFilter(), "Step 2 of 2 — Pick the hanger to update")
        hanger = doc.GetElement(ref_hanger.ElementId)
    except Exception:
        return
    if hanger is None:
        forms.alert("Could not read selected hanger.", title="Hanger Sync"); return

    try:
        hpt   = hanger.Location.Point
        curve = get_curve(content)
        if curve:
            t, _ = xy_project_onto_curve(hpt, curve)
            t    = max(0.0, min(1.0, t))
        else:
            t = 0.5
    except Exception:
        t = 0.5

    detected_ins = get_insulation_thickness(content)
    snap_ins_ft  = 0.0
    dlg = InsulationChoiceWindow(content, detected_ins)
    if not dlg.show():
        return
    snap_ins_ft = dlg.insulation_ft if dlg.snap_to_insulation else 0.0

    bottom_z    = get_pipe_bottom_at_t(content, t)
    nominal_dia = get_nominal_dia(content)
    duct_width  = get_duct_width(content)
    if bottom_z is None:
        forms.alert("Could not determine elevation from selected content.", title="Hanger Sync")
        return

    hung_bottom_z = bottom_z - snap_ins_ft
    insulated_dia = (nominal_dia + 2.0 * snap_ins_ft) if nominal_dia else None

    t_write = DB.Transaction(doc, "Hanger Sync - Pick Pair")
    try:
        t_write.Start()
        written = []
        elev_p = write_elevation(hanger, hung_bottom_z)
        if elev_p: written.append(elev_p)
        if insulated_dia is not None:
            try:
                written.extend(write_size(hanger, insulated_dia))
            except Exception: pass
        if duct_width is not None:
            try:
                tw = write_trapeze_width(hanger, duct_width)
                if tw: written.append(tw)
            except Exception: pass
        try:
            recenter_xy(hanger, content)
            written.append("XY")
        except Exception: pass
        if snap_ins_ft > 0:
            try:
                write_insulation(hanger, snap_ins_ft)
                written.append("Insulation")
            except Exception: pass
        t_write.Commit()
        forms.alert(
            "Hanger %s updated.\nWritten: %s" % (
                hanger.Id.IntegerValue, ", ".join(written) if written else "none"),
            title="Hanger Sync — Done")
    except Exception as ex:
        try:
            if t_write.GetStatus() == DB.TransactionStatus.Started:
                t_write.RollBack()
        except Exception: pass
        forms.alert("Sync failed: %s" % str(ex), title="Hanger Sync")


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------

_XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")


class HangerManagerWindow(forms.WPFWindow):

    def __init__(self, xaml_path, lower_rows, slope_line_id=None, direct_hanger_ids=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self._lower_rows        = list(lower_rows)
        self._slope_line_id     = slope_line_id
        self._direct_hanger_ids = list(direct_hanger_ids) if direct_hanger_ids else None
        self._upper_rows        = ObservableCollection[object]()
        self._settings          = _load_settings()
        self._result            = "close"

        self._apply_settings()
        self._populate_lower()
        self._populate_insul_params()
        self._init_col_visibility()
        self._rebuild_upper_rows()
        self._populate_upper()
        self._update_header()

        if self._direct_hanger_ids is not None:
            self._apply_upper_only_mode()

    def _apply_upper_only_mode(self):
        """Hide Lower tab, Re-Scan, Sync Lower, Sync Both when opened via Pick Hangers."""
        from System.Windows import Visibility
        try: self.tabMain.Items[0].Visibility = Visibility.Collapsed
        except Exception: pass
        try: self.tabMain.SelectedIndex = 1
        except Exception: pass
        try: self.btnRescan.Visibility    = Visibility.Collapsed
        except Exception: pass
        try: self.btnSyncLower.Visibility = Visibility.Collapsed
        except Exception: pass
        try: self.btnSyncBoth.Visibility  = Visibility.Collapsed
        except Exception: pass
        try:
            n = len(self._direct_hanger_ids)
            self.txtContentInfo.Text = "{} hanger(s) selected via Pick Hangers".format(n)
        except Exception: pass

    # ── settings ────────────────────────────────────────────────────────────

    def _apply_settings(self):
        s = self._settings
        try: self.txtProximity.Text = str(s.get("proximity_ft", 1.0))
        except Exception: pass
        try:
            self.chkInsulation.IsChecked = bool(s.get("insulation_on", False))
            self.cmbInsulParam.IsEnabled = bool(s.get("insulation_on", False))
        except Exception: pass
        try:
            self.chkColFamily.IsChecked = bool(s.get("col_family", True))
            self.chkColType.IsChecked   = bool(s.get("col_type",   True))
            self.chkColStatus.IsChecked = bool(s.get("col_status", False))
        except Exception: pass

    def _collect_settings(self):
        s = dict(_DEFAULT_SETTINGS)
        try: s["proximity_ft"]    = float(self.txtProximity.Text.strip())
        except Exception: pass
        try: s["insulation_on"]   = bool(self.chkInsulation.IsChecked)
        except Exception: pass
        try: s["insulation_param"] = self.cmbInsulParam.SelectedItem or ""
        except Exception: pass
        try:
            s["col_family"] = bool(self.chkColFamily.IsChecked)
            s["col_type"]   = bool(self.chkColType.IsChecked)
            s["col_status"] = bool(self.chkColStatus.IsChecked)
        except Exception: pass
        return s

    # ── Lower tab ────────────────────────────────────────────────────────────

    def _populate_lower(self):
        col = ObservableCollection[object]()
        for r in self._lower_rows:
            col.Add(r)
        self.dgLower.ItemsSource = col
        self._update_status_bar()

    def _populate_insul_params(self):
        self.cmbInsulParam.Items.Clear()
        from collections import defaultdict as _dd
        usage = _dd(int)
        param_names = set()
        for row in self._lower_rows:
            try:
                for p in row._content.Parameters:
                    if p.StorageType == DB.StorageType.Double:
                        name = p.Definition.Name
                        param_names.add(name)
                        try:
                            if p.AsDouble() != 0.0: usage[name] += 1
                        except Exception: pass
            except Exception: pass

        def sort_key(n):
            nl = n.lower()
            is_insul = 1 if any(kw in nl for kw in ("insul","lining","thickness","adjust")) else 0
            return (1 - is_insul, -usage.get(n, 0), nl)

        for name in sorted(param_names, key=sort_key):
            self.cmbInsulParam.Items.Add(name)

        self.cmbInsulParam.SelectionChanged += self._insul_selection_changed
        saved_param = self._settings.get("insulation_param", "")
        selected    = False
        if saved_param:
            for i in range(self.cmbInsulParam.Items.Count):
                if self.cmbInsulParam.Items[i] == saved_param:
                    self.cmbInsulParam.SelectedIndex = i; selected = True; break
        if not selected:
            for i in range(self.cmbInsulParam.Items.Count):
                if any(kw in self.cmbInsulParam.Items[i].lower()
                       for kw in ("insul","thickness","lining")):
                    self.cmbInsulParam.SelectedIndex = i; selected = True; break
        if not selected and self.cmbInsulParam.Items.Count > 0:
            self.cmbInsulParam.SelectedIndex = 0
        self._update_insul_preview()
        self._refresh_insulation_col()

    def _insul_selection_changed(self, sender, e):
        self._update_insul_preview()
        self._refresh_insulation_col()

    def chkInsulation_Changed(self, sender, e):
        self.cmbInsulParam.IsEnabled = bool(self.chkInsulation.IsChecked)
        self._update_insul_preview()
        self._refresh_insulation_col()

    def _refresh_insulation_col(self):
        for row in self._lower_rows:
            ins_ft = self._get_insulation_ft(row._content)
            row.Insulation = "none" if not ins_ft else '%.4g"' % (ins_ft * 12)
        self._populate_lower()

    def _update_insul_preview(self):
        try:
            if not bool(self.chkInsulation.IsChecked):
                self.txtInsulPreview.Text = ""; return
            param_name = self.cmbInsulParam.SelectedItem
            if not param_name or not self._lower_rows:
                self.txtInsulPreview.Text = ""; return
            p = self._lower_rows[0]._content.LookupParameter(param_name)
            if p and p.StorageType == DB.StorageType.Double:
                self.txtInsulPreview.Text = "= %.4g\" on first element" % (p.AsDouble() * 12)
            else:
                self.txtInsulPreview.Text = "(not found)"
        except Exception:
            self.txtInsulPreview.Text = ""

    def _get_insulation_ft(self, content_elem):
        if not bool(self.chkInsulation.IsChecked): return None
        param_name = self.cmbInsulParam.SelectedItem
        if not param_name: return None
        try:
            p = content_elem.LookupParameter(param_name)
            if p and p.StorageType == DB.StorageType.Double:
                return p.AsDouble()
        except Exception: pass
        return None

    def _get_proximity_ft(self):
        try:
            val = float(self.txtProximity.Text.strip())
            if val > 0: return val
        except Exception: pass
        return PROX_FT

    def _init_col_visibility(self):
        from System.Windows import Visibility
        self.colFamily.Visibility = (Visibility.Visible
            if bool(self.chkColFamily.IsChecked) else Visibility.Collapsed)
        self.colType.Visibility   = (Visibility.Visible
            if bool(self.chkColType.IsChecked) else Visibility.Collapsed)
        self.colStatus.Visibility = Visibility.Collapsed  # hidden until explicitly enabled

    def ColVisibility_Changed(self, sender, e):
        from System.Windows import Visibility
        def vis(chk):
            return Visibility.Visible if bool(chk.IsChecked) else Visibility.Collapsed
        self.colFamily.Visibility = vis(self.chkColFamily)
        self.colType.Visibility   = vis(self.chkColType)
        self.colStatus.Visibility = vis(self.chkColStatus)

    # ── Upper tab ────────────────────────────────────────────────────────────

    def _get_slope_curve(self):
        if self._slope_line_id is None:
            return None
        line_elem = doc.GetElement(self._slope_line_id)
        if line_elem and isinstance(line_elem, DB.CurveElement):
            return line_elem.GeometryCurve
        return None

    def _rebuild_upper_rows(self):
        """Rebuild UpperRow list from direct_hanger_ids (if set) or current lower rows + slope curve."""
        curve = self._get_slope_curve()
        if self._direct_hanger_ids is not None:
            hids = list(self._direct_hanger_ids)
        else:
            seen = set()
            hids = []
            for row in self._lower_rows:
                hid = row._hanger.Id.IntegerValue
                if hid not in seen:
                    seen.add(hid)
                    hids.append(row._hanger.Id)
        self._upper_rows = _build_upper_rows(hids, curve)

    def _populate_upper(self):
        self.dgUpper.ItemsSource = self._upper_rows
        self._update_status_bar()

    def _update_line_label(self):
        if self._slope_line_id is not None:
            line_elem = doc.GetElement(self._slope_line_id)
            if line_elem and isinstance(line_elem, DB.CurveElement):
                self.LineLabel.Text = slope_label(line_elem.GeometryCurve)
            else:
                self.LineLabel.Text = "Line not found"
                self._slope_line_id = None
        else:
            self.LineLabel.Text = "No slope line selected"

    # ── Header / status ──────────────────────────────────────────────────────

    def _update_header(self):
        n               = len(self._lower_rows)
        unique_content  = len(set(r._content.Id.IntegerValue for r in self._lower_rows))
        self.txtContentInfo.Text = "{} hanger(s)  matched to  {} content element(s)".format(
            n, unique_content)
        self._update_line_label()

    def _update_status_bar(self):
        lower_ready  = sum(1 for r in self._lower_rows if r.Status == "Ready")
        lower_synced = sum(1 for r in self._lower_rows if "Synced" in r.Status)
        upper_ready  = sum(1 for r in self._upper_rows if r._status == _UP_READY)
        upper_synced = sum(1 for r in self._upper_rows if r._status == _UP_SYNCED)
        self.txtStatus.Text = (
            "Lower: {} ready, {} synced   |   Upper: {} ready, {} synced"
            .format(lower_ready, lower_synced, upper_ready, upper_synced))
        self.btnSyncLower.IsEnabled = lower_ready > 0
        self.btnSyncUpper.IsEnabled = upper_ready > 0
        self.btnSyncBoth.IsEnabled  = (lower_ready > 0 or upper_ready > 0)

    # ── Button handlers ───────────────────────────────────────────────────────

    def pick_line_click(self, sender, e):
        self._result = "pick_line"
        self.Close()

    def btnRescan_Click(self, sender, e):
        prox_ft     = self._get_proximity_ft()
        content_map = {}
        for row in self._lower_rows:
            eid = row._content.Id.IntegerValue
            if eid not in content_map:
                content_map[eid] = row._content
        new_rows = []
        seen = set()
        for content in content_map.values():
            for hanger, dist, t in find_hangers_near(content, prox_ft):
                if hanger.Id.IntegerValue not in seen:
                    seen.add(hanger.Id.IntegerValue)
                    new_rows.append(LowerRow(content, hanger, dist, t))
        self._lower_rows = new_rows
        self._populate_lower()
        self._populate_insul_params()
        self._rebuild_upper_rows()
        self._populate_upper()
        self._update_header()
        self.txtStatus.Text = "Rescanned at %.2f ft — %d hanger(s) found." % (prox_ft, len(new_rows))

    def btnRemoveRow_Click(self, sender, e):
        """Remove selected rows from both tabs (matched by Hanger ID)."""
        selected_lower = list(self.dgLower.SelectedItems)
        selected_upper = list(self.dgUpper.SelectedItems)
        ids_to_remove  = set()
        for r in selected_lower:
            ids_to_remove.add(r.HangerId)
        for r in selected_upper:
            ids_to_remove.add(r.HangerID)
        if not ids_to_remove:
            return
        self._lower_rows = [r for r in self._lower_rows if r.HangerId not in ids_to_remove]
        self._rebuild_upper_rows()
        self._populate_lower()
        self._populate_upper()
        self._update_header()
        self.btnRemoveRow.IsEnabled = False

    def dgLower_SelectionChanged(self, sender, e):
        has_sel = len(list(self.dgLower.SelectedItems)) > 0
        self.btnRemoveRow.IsEnabled = has_sel

    def dgUpper_SelectionChanged(self, sender, e):
        has_sel = len(list(self.dgUpper.SelectedItems)) > 0
        self.btnRemoveRow.IsEnabled = has_sel

    def btnSyncLower_Click(self, sender, e):
        done, errors = sync_lower_rows(self._lower_rows, self._get_insulation_ft)
        self._populate_lower()
        self._update_status_bar()
        msg = "Lower: {} synced.".format(done)
        if errors: msg += "  {} error(s).".format(errors)
        self.txtStatus.Text = msg

    def btnSyncUpper_Click(self, sender, e):
        if self._slope_line_id is None:
            forms.alert("No slope line selected. Click Pick Line first.", title="Hanger Sync")
            return
        synced, skipped = sync_upper_rows(list(self._upper_rows))
        self._populate_upper()
        self._update_status_bar()
        msg = "Upper: {} synced.".format(synced)
        if skipped: msg += "  {} skipped.".format(skipped)
        self.txtStatus.Text = msg

    def btnSyncBoth_Click(self, sender, e):
        lower_done, lower_errors = sync_lower_rows(self._lower_rows, self._get_insulation_ft)
        upper_synced = upper_skipped = 0
        if self._slope_line_id is not None:
            upper_synced, upper_skipped = sync_upper_rows(list(self._upper_rows))
        self._populate_lower()
        self._populate_upper()
        self._update_status_bar()
        msg = "Lower: {} synced".format(lower_done)
        if lower_errors: msg += " ({} errors)".format(lower_errors)
        if self._slope_line_id is not None:
            msg += ".  Upper: {} synced".format(upper_synced)
            if upper_skipped: msg += " ({} skipped)".format(upper_skipped)
        else:
            msg += ".  Upper skipped (no slope line)."
        self.txtStatus.Text = msg

    def btnClose_Click(self, sender, e):
        _save_settings(self._collect_settings())
        self._result = "close"
        self.Close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

sel_ids       = uidoc.Selection.GetElementIds()
valid_cat_ids = [int(c) for c in HOST_CATS]
content_elems = [
    doc.GetElement(i) for i in sel_ids
    if doc.GetElement(i) is not None
    and doc.GetElement(i).Category is not None
    and doc.GetElement(i).Category.Id.IntegerValue in valid_cat_ids
]

dlg                                  = SelectContentWindow(len(content_elems))
result, do_pair, use_sel, do_hangers = dlg.show()

# ── Pick Hangers -> Upper Attachment only ─────────────────────────────────────
if do_hangers:
    try:
        hanger_refs = uidoc.Selection.PickObjects(
            ObjectType.Element, HangerFilter(),
            "Pick hangers (Generic Model families) -- multi-select OK")
        hanger_ids = [r.ElementId for r in hanger_refs]
    except Exception:
        hanger_ids = []

    slope_line_id = None
    if hanger_ids:
        try:
            line_ref = uidoc.Selection.PickObject(
                ObjectType.Element, ModelLineFilter(),
                "Pick the sloped model line for Upper Attachment")
            slope_line_id = line_ref.ElementId
        except Exception:
            pass

        while True:
            win = HangerManagerWindow(
                _XAML_PATH, [], slope_line_id,
                direct_hanger_ids=hanger_ids)
            win.ShowDialog()

            slope_line_id = win._slope_line_id

            if win._result == "pick_line":
                try:
                    ref = uidoc.Selection.PickObject(
                        ObjectType.Element, ModelLineFilter(),
                        "Pick the sloped model line for Upper Attachment")
                    slope_line_id = ref.ElementId
                except Exception:
                    pass
            else:
                break

# ── Pick Part + Hanger ────────────────────────────────────────────────────────
elif do_pair:
    run_pick_pair_workflow()

# ── Pick in Viewport / Use Selection ─────────────────────────────────────────
else:
    if use_sel:
        pass  # content_elems already populated
    elif result:
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element, HostFilter(),
                "Pick duct / pipe to sync hangers for (multi-select OK)")
            content_elems = [
                doc.GetElement(r.ElementId)
                for r in refs
                if doc.GetElement(r.ElementId) is not None
            ]
        except Exception:
            content_elems = []
    else:
        content_elems = []

    if content_elems:
        _init_settings = _load_settings()
        _init_prox_ft  = float(_init_settings.get("proximity_ft", PROX_FT))

        lower_rows = []
        seen = set()
        for content in content_elems:
            for hanger, dist, t in find_hangers_near(content, _init_prox_ft):
                if hanger.Id.IntegerValue not in seen:
                    seen.add(hanger.Id.IntegerValue)
                    lower_rows.append(LowerRow(content, hanger, dist, t))

        if not lower_rows:
            forms.alert(
                "No Generic Model hangers found within %.2f' XY of the selected content.\n\n"
                "Adjust Max hanger distance and re-run, or select different content."
                % _init_prox_ft,
                title="Hanger Sync Plus")
        else:
            slope_line_id = None
            while True:
                win = HangerManagerWindow(_XAML_PATH, lower_rows, slope_line_id)
                win.ShowDialog()

                lower_rows    = win._lower_rows
                slope_line_id = win._slope_line_id

                if win._result == "pick_line":
                    _notice = SelectLineWindow()
                    if _notice.show():
                        try:
                            ref = uidoc.Selection.PickObject(
                                ObjectType.Element, ModelLineFilter(),
                                "Pick the sloped model line for Upper Attachment")
                            slope_line_id = ref.ElementId
                        except Exception:
                            pass
                    # loop — reopen window with slope_line_id pre-loaded
                else:
                    break
