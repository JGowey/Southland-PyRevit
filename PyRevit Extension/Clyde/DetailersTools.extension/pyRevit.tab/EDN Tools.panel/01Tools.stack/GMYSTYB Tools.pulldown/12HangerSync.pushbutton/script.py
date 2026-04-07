# -*- coding: utf-8 -*-
"""
RFA Hanger Sync
Author: Jeremiah Griffith  |  Version: 4.0.0

Content-first workflow:
  1. User selects duct(s) or pipe(s) in Revit before running the tool.
  2. Tool reads geometry from those elements (pure read, no transactions).
  3. Finds Generic Model hangers within 3 ft XY of each selected element.
  4. User can pick a host manually for any unmatched hanger.
  5. Sync All writes BOHElevation to each matched hanger.

No fabrication FilteredElementCollector during scan.
No GTP writes. No transactions until Sync All.
"""

from __future__ import division
import os, math, clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")

from System.Collections.ObjectModel import ObservableCollection
from pyrevit import revit, DB, forms
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc   = revit.doc
uidoc = revit.uidoc

# Categories the user can pick as content (duct/pipe)
HOST_CATS = [
    DB.BuiltInCategory.OST_FabricationDuctwork,
    DB.BuiltInCategory.OST_DuctCurves,
    DB.BuiltInCategory.OST_FabricationPipework,
    DB.BuiltInCategory.OST_PipeCurves,
]
PROX_FT = 1.0   # XY radius -- hangers should be within ~1 ft of centerline

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
# GEOMETRY HELPERS  (read-only, no transactions)
# ---------------------------------------------------------------------------

def get_bottom_z(elem):
    """
    Return the bottom elevation of the element body in feet.

    For fabrication parts, prefer 'Lower End Bottom Elevation' -- this is
    the bottom of the duct/pipe body and excludes connection flanges that
    extend below the body and inflate the bounding box minimum.
    Fall back to bounding box only if the param is absent or zero.
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


def get_curve(elem):
    try:
        loc = elem.Location
        return loc.Curve if hasattr(loc, "Curve") else None
    except Exception:
        return None


def xy_project_onto_curve(pt, curve):
    """
    Project pt onto curve in 2D (XY only), ignoring Z.
    Returns (t, projected_pt) where t is the normalized parameter [0,1]
    along the segment, and projected_pt is the XY centerline point at t
    (Z matches the interpolated pipe centerline at that position).

    Using the unbounded line through the curve endpoints so we can
    detect projections beyond the segment endpoints separately.
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

    # Scalar projection of (pt - p0) onto pipe XY direction
    dot = (pt.X - p0.X) * ux + (pt.Y - p0.Y) * uy
    t   = dot / length_xy            # normalized 0-1 along segment

    # XY position on centerline
    cx = p0.X + dot * ux
    cy = p0.Y + dot * uy

    # Z interpolated along the 3D slope
    cz = p0.Z + t * (p1.Z - p0.Z)

    return t, DB.XYZ(cx, cy, cz)


def get_pipe_bottom_at_t(content, t):
    """
    Return the bottom elevation in feet at normalized position t along
    the content's slope.

    For fabrication parts (pipe and duct -- round or rectangular):
      Revit already computes 'Lower End Bottom Elevation' and
      'Upper End Bottom Elevation' for each end of the part. These
      account for the actual profile shape (round, rectangular, insulated)
      without any manual diameter or depth math on our part.
      Interpolating between them at parameter t gives the exact bottom
      elevation at the hanger's position along the run.

    For native duct/pipe (no fabrication end-elevation params):
      Fall back to centerline interpolation minus half the profile depth.

    t=0 = GetEndPoint(0) end, t=1 = GetEndPoint(1) end.
    """
    try:
        cat          = content.Category.Id.IntegerValue
        is_fab_pipe  = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct  = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)

        if is_fab_pipe or is_fab_duct:
            # Primary method: interpolate between the two end bottom elevations.
            # Works for round duct, rectangular duct, and pipe regardless of
            # profile -- Revit calculates these correctly already.
            p_lo = content.LookupParameter("Lower End Bottom Elevation")
            p_hi = content.LookupParameter("Upper End Bottom Elevation")
            if (p_lo and p_hi and
                    p_lo.StorageType == DB.StorageType.Double and
                    p_hi.StorageType == DB.StorageType.Double):
                lo = p_lo.AsDouble()
                hi = p_hi.AsDouble()
                # t=0 maps to the lower end (GetEndPoint(0) of the curve)
                return lo + t * (hi - lo)

            # Fallback: use static Lower End Bottom Elevation
            return get_bottom_z(content)

        # Native duct / pipe: no fabrication end params.
        # Interpolate centerline Z then subtract half the profile depth.
        curve = get_curve(content)
        if curve is None:
            return get_bottom_z(content)

        p0   = curve.GetEndPoint(0)
        p1   = curve.GetEndPoint(1)
        cl_z = p0.Z + t * (p1.Z - p0.Z)

        is_nat_pipe = cat == int(DB.BuiltInCategory.OST_PipeCurves)
        is_nat_duct = cat == int(DB.BuiltInCategory.OST_DuctCurves)

        half_depth = 0.0
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
    """XY-only distance from pt to elem's location curve or point."""
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
    """Return a human-readable size string using the NOMINAL size (no insulation)."""
    try:
        cat = elem.Category.Id.IntegerValue
        is_fab_pipe = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)

        if is_fab_pipe:
            # Product Entry is the nominal size string (e.g. "8"), never includes insulation
            p = elem.LookupParameter("Product Entry")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip().replace('"', "")
                if s:
                    return s + '"'
            # Fallback: Main Primary Diameter (nominal, feet)
            p = elem.LookupParameter("Main Primary Diameter")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return '%s"' % int(round(p.AsDouble() * 12))

        if is_fab_duct:
            # Overall Size for duct is nominal (width x height, no insulation layer)
            p = elem.LookupParameter("Overall Size")
            if p and p.StorageType == DB.StorageType.String:
                return p.AsString() or ""

        # Native duct / pipe
        parts = []
        for n in ("Width", "Height", "Diameter", "Outside Diameter"):
            p = elem.LookupParameter(n)
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                parts.append('%s"' % int(round(p.AsDouble() * 12)))
        return "x".join(parts)
    except Exception:
        return ""


def write_elevation(hanger, bottom_z):
    """Write bottom_z to the hanger. Returns param name written or None."""
    for name in ("BOHElevation", "CP_Bottom Elevation", "CP_Hung Object Bottom Elev"):
        p = hanger.LookupParameter(name)
        if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
            p.Set(bottom_z)
            return name
    return None


def get_duct_width(content):
    """
    Return duct width in feet for CP_Trapeze Width Desired.
    Main Primary Width on fabrication ductwork is stored in feet
    (Revit stores all length params internally in feet).
    """
    try:
        cat = content.Category.Id.IntegerValue
        if cat == int(DB.BuiltInCategory.OST_FabricationDuctwork):
            p = content.LookupParameter("Main Primary Width")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()   # already feet
        elif cat == int(DB.BuiltInCategory.OST_DuctCurves):
            p = content.LookupParameter("Width")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()
    except Exception:
        pass
    return None


def get_nominal_dia(content):
    """
    Return nominal diameter in feet from content element, or None.

    For fabrication pipe: use Main Primary Diameter (double, feet).
    For fabrication duct: parse Overall Size string (e.g. "24x12").
    For native duct/pipe: use Diameter / Width / Height double params.

    Debug report confirmed:
      - Fab pipe has 'Main Primary Diameter' = nominal dia in feet.
      - 'Overall Size' = "8"o" -- the degree/circle char kills float parse.
      - 'Outside Diameter' = OD (8.125"), not nominal (8.0").
    """
    try:
        cat = content.Category.Id.IntegerValue
        is_fab_pipe = cat == int(DB.BuiltInCategory.OST_FabricationPipework)
        is_fab_duct = cat == int(DB.BuiltInCategory.OST_FabricationDuctwork)

        if is_fab_pipe:
            # Product Entry = nominal size string like "8" (inches, no insulation).
            # Always use this first -- Main Primary Diameter can reflect the
            # insulated OD on insulated pipes, giving the wrong value.
            p = content.LookupParameter("Product Entry")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip().replace('"', "").strip()
                if s:
                    try:
                        return float(s) / 12.0
                    except Exception:
                        pass
            # Fallback: Main Primary Diameter
            p = content.LookupParameter("Main Primary Diameter")
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()

        if is_fab_duct:
            # Overall Size string e.g. "24x12" or "14"o" -- parse carefully
            p = content.LookupParameter("Overall Size")
            if p and p.StorageType == DB.StorageType.String:
                s = (p.AsString() or "").strip()
                # Strip non-numeric chars except x and dot and minus
                import re as _re
                s = s.lower().replace(" ", "")
                if "x" in s:
                    parts = s.split("x")
                    nums = []
                    for part in parts:
                        digits = _re.sub(r"[^0-9.]", "", part)
                        if digits:
                            nums.append(float(digits))
                    if nums:
                        return max(nums) / 12.0
                else:
                    digits = _re.sub(r"[^0-9.]", "", s)
                    if digits:
                        return float(digits) / 12.0

        # Native duct / pipe -- double params already in feet
        for name in ("Diameter", "Outside Diameter"):
            p = content.LookupParameter(name)
            if p and p.StorageType == DB.StorageType.Double and p.AsDouble() > 0:
                return p.AsDouble()

        # Native rectangular duct
        pw = content.LookupParameter("Width")
        ph = content.LookupParameter("Height")
        if pw and ph and pw.StorageType == DB.StorageType.Double:
            w = pw.AsDouble()
            h = ph.AsDouble()
            if w > 0 or h > 0:
                return max(w, h)

    except Exception:
        pass
    return None


def get_insulation_thickness(content):
    """
    Derive insulation thickness in feet from fabrication pipe/duct geometry.

    The fabrication part stores per-end elevation params that already encode
    the insulation layer. Insulation thickness = pipe bottom elevation minus
    the bottom-of-insulation elevation (they're equal when uninsulated).

    FabricationConfiguration.GetInsulationThickness() would be cleaner but
    the IronPython namespace path is broken in Revit 2024 pyRevit; geometry
    derivation is reliable and works for both pipe and duct.
    """
    try:
        # Use lower-end params (consistent for straight runs)
        p_bot      = content.LookupParameter("Lower End Bottom Elevation")
        p_bot_ins  = content.LookupParameter("Lower End Bottom of Insulation Elevation")
        if (p_bot and p_bot_ins and
                p_bot.StorageType == DB.StorageType.Double and
                p_bot_ins.StorageType == DB.StorageType.Double):
            thickness = p_bot.AsDouble() - p_bot_ins.AsDouble()
            if thickness >= 0:
                return thickness

        # Fallback: top-of-insulation minus top-of-pipe
        p_top     = content.LookupParameter("Lower End Top Elevation")
        p_top_ins = content.LookupParameter("Lower End Top of Insulation Elevation")
        if (p_top and p_top_ins and
                p_top.StorageType == DB.StorageType.Double and
                p_top_ins.StorageType == DB.StorageType.Double):
            thickness = p_top_ins.AsDouble() - p_top.AsDouble()
            if thickness >= 0:
                return thickness
    except Exception:
        pass
    return 0.0


def write_insulation(hanger, thickness_ft):
    """
    Write insulation thickness to the hanger.

    Two params must be set for the family formula to actually extend the rod:
      CP_Host Insulation Thickness  -- double (feet): the thickness value
      CP_Hanger Insulation Adjust   -- string: must be set to the numeric
                                       inch value (e.g. "1") to override
                                       the "By Spec" default and use our value.

    Returns list of param names written.
    """
    written = []
    p = hanger.LookupParameter("CP_Host Insulation Thickness")
    if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
        p.Set(thickness_ft)
        written.append("CP_Host Insulation Thickness")

    # Set CP_Hanger Insulation Adjust to the numeric inch string.
    # "By Spec" defers to the fabrication spec and ignores our value.
    # A numeric string (e.g. "1") tells the family to use that thickness.
    p2 = hanger.LookupParameter("CP_Hanger Insulation Adjust")
    if p2 and p2.StorageType == DB.StorageType.String and not p2.IsReadOnly:
        if thickness_ft > 0:
            inch_str = "%.4g" % (thickness_ft * 12)
            p2.Set(inch_str)
        else:
            p2.Set("0")
        written.append("CP_Hanger Insulation Adjust")

    return written[0] if written else None


def write_trapeze_width(hanger, width_ft):
    """Write duct width to CP_Trapeze Width Desired. Returns param name or None."""
    p = hanger.LookupParameter("CP_Trapeze Width Desired")
    if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
        p.Set(width_ft)
        return "CP_Trapeze Width Desired"
    return None


def recenter_xy(hanger, content):
    """
    Move the hanger laterally so it sits on the content centerline,
    but do NOT move it along the pipe/duct axis.

    Problem with plain curve.Project(): it snaps to the nearest point
    on the bounded curve segment, so hangers near segment ends move
    to the coupler/joint rather than staying at their along-pipe position.

    Fix: compute the pipe direction unit vector, decompose the offset
    into along-pipe and perpendicular components, apply only the
    perpendicular component. The hanger stays at its current position
    along the run and only closes the lateral gap to the centerline.
    """
    try:
        hanger_loc = hanger.Location
        if not hasattr(hanger_loc, "Point"):
            return False
        pt = hanger_loc.Point

        content_loc = content.Location
        if hasattr(content_loc, "Curve"):
            curve = content_loc.Curve
            # Get pipe direction unit vector (XY only)
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            dx_pipe = p1.X - p0.X
            dy_pipe = p1.Y - p0.Y
            pipe_len = (dx_pipe**2 + dy_pipe**2) ** 0.5
            if pipe_len < 0.001:
                return False
            # Unit vector along pipe
            ux = dx_pipe / pipe_len
            uy = dy_pipe / pipe_len

            # Project onto the UNBOUNDED line through p0 in pipe direction
            # to find the centerline point at the hanger's along-pipe position
            dot = (pt.X - p0.X) * ux + (pt.Y - p0.Y) * uy
            # Centerline point directly across from hanger
            cx = p0.X + dot * ux
            cy = p0.Y + dot * uy

            # Offset = perpendicular only
            move_x = cx - pt.X
            move_y = cy - pt.Y

        elif hasattr(content_loc, "Point"):
            # Point-located element: full XY snap is appropriate
            target  = content_loc.Point
            move_x  = target.X - pt.X
            move_y  = target.Y - pt.Y
        else:
            return False

        if (move_x**2 + move_y**2) < 0.0001:   # already within ~1/8"
            return False

        DB.ElementTransformUtils.MoveElement(
            doc, hanger.Id, DB.XYZ(move_x, move_y, 0.0))
        return True
    except Exception:
        return False


def write_size(hanger, nominal_dia_ft):
    """
    Write nominal diameter to CP_Hung Object Diameter and CP_Host Nominal Diameter.
    These params drive nested sub-family type selection via family formulas and
    can trigger MODIFICATION IS FORBIDDEN during regen. We attempt the write and
    return a list of param names that were actually set. Caller wraps in try/except.
    """
    written = []
    for name in ("CP_Hung Object Diameter", "CP_Host Nominal Diameter"):
        p = hanger.LookupParameter(name)
        if p and p.StorageType == DB.StorageType.Double and not p.IsReadOnly:
            p.Set(nominal_dia_ft)
            written.append(name)
    return written


# ---------------------------------------------------------------------------
# HANGER DISCOVERY  (Generic Model only, no fabrication collectors)
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

        # Expand bbox only enough to catch hangers sitting just above/beside
        # the pipe -- 0.5 ft lateral expansion, generous Z above for rods
        mn = DB.XYZ(bb.Min.X - prox_ft, bb.Min.Y - prox_ft, bb.Min.Z - 2.0)
        mx = DB.XYZ(bb.Max.X + prox_ft, bb.Max.Y + prox_ft, bb.Max.Z + 15.0)
        bb_filter = DB.BoundingBoxIntersectsFilter(DB.Outline(mn, mx))

        candidates = (DB.FilteredElementCollector(doc)
                      .OfCategory(DB.BuiltInCategory.OST_GenericModel)
                      .WhereElementIsNotElementType()
                      .WherePasses(bb_filter))

        # Get bounded curve for on-segment test
        curve = get_curve(content_elem)

        for elem in candidates:
            if not isinstance(elem, DB.FamilyInstance):
                continue
            try:
                pt = elem.Location.Point
            except Exception:
                continue

            # All projections are XY-only -- sloped pipes must use 2D
            # projection so Z difference between hanger (at slab) and
            # pipe (mid-run) doesn't pull the projection to an endpoint.
            if curve is not None:
                try:
                    t, proj_xy = xy_project_onto_curve(pt, curve)

                    # XY perpendicular distance
                    d = math.sqrt((pt.X - proj_xy.X)**2 + (pt.Y - proj_xy.Y)**2)
                    if d > prox_ft:
                        continue

                    # On-segment test: t must be in [0, 1]
                    if t < -0.01 or t > 1.01:
                        continue

                except Exception:
                    continue
            else:
                d = xy_dist(pt, content_elem)
                t = 0.5
                if d > PROX_FT:
                    continue

            results.append((elem, d, t))
    except Exception:
        pass
    return results


# ---------------------------------------------------------------------------
# DATA MODEL
# ---------------------------------------------------------------------------

class Row(object):
    def __init__(self, content, hanger, dist_ft, t=0.5):
        self.ContentId   = str(content.Id.IntegerValue)
        self.ContentType = content.Category.Name if content.Category else ""
        self.ContentSize = get_size(content)
        self.HangerId    = str(hanger.Id.IntegerValue)
        self.HangerFamily = ""
        self.HangerType   = ""
        try:
            type_id = hanger.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type:
                    # FamilyName on the type element is the reliable path
                    if hasattr(elem_type, 'FamilyName'):
                        try:
                            self.HangerFamily = str(elem_type.FamilyName) or ""
                        except Exception:
                            pass
                    if not self.HangerFamily and hanger.Category:
                        self.HangerFamily = str(hanger.Category.Name) or ""
                    # Type name
                    try:
                        if hasattr(elem_type, 'Name'):
                            self.HangerType = str(elem_type.Name) or ""
                    except Exception:
                        pass
                    if not self.HangerType:
                        try:
                            p = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                            if p:
                                self.HangerType = p.AsString() or ""
                        except Exception:
                            pass
        except Exception:
            pass
        self.DistFt      = "%.2f'" % dist_ft
        self._content       = content
        self._hanger        = hanger
        self._t             = t      # normalized position [0,1] along pipe
        # Elevation interpolated at hanger's position along the slope
        self._bottom_z      = get_pipe_bottom_at_t(content, t)
        self._nominal_dia   = get_nominal_dia(content)
        self._duct_width    = get_duct_width(content)
        self._insulation_ft = get_insulation_thickness(content)
        self.Insulation     = ""   # populated after window opens with param selection
        # Flag rows where the match distance is suspicious
        self.DistWarning    = dist_ft > 0.5
        self.Status         = "Ready" if self._bottom_z is not None else "No elevation"


class HostFilter(ISelectionFilter):
    def AllowElement(self, e):
        if e is None or e.Category is None:
            return False
        return any(e.Category.Id.IntegerValue == int(c) for c in HOST_CATS)
    def AllowReference(self, r, p):
        return False


# ---------------------------------------------------------------------------
# WINDOW
# ---------------------------------------------------------------------------

class SelectContentWindow(forms.WPFWindow):
    """
    Polished instruction dialog shown when no duct/pipe is pre-selected.
    Replaces the default forms.alert() warning box.
    """
    XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="RFA Hanger Sync"
    Width="460" Height="300"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    FontFamily="Segoe UI"
    Background="#F5F7FB">

    <Grid>
        <!-- Accent bar top -->
        <Border VerticalAlignment="Top" Height="4" Background="#8F6C05"/>

        <StackPanel Margin="32,24,32,20" VerticalAlignment="Top">

            <!-- Icon + Title row -->
            <StackPanel Orientation="Horizontal" Margin="0,0,0,16">
                <Border Width="40" Height="40" CornerRadius="20"
                        Background="#FDF5DC" Margin="0,0,14,0">
                    <TextBlock Text="&#x2316;" FontSize="20"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"
                               Foreground="#8F6C05"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="RFA Hanger Sync"
                               FontSize="16" FontWeight="SemiBold"
                               Foreground="#0F1729"/>
                    <TextBlock x:Name="txtSubtitle"
                               Text="Content selection required"
                               FontSize="11" Foreground="#6B7280"/>
                </StackPanel>
            </StackPanel>

            <!-- Instruction steps -->
            <Border x:Name="pnlSteps" Background="White" CornerRadius="6"
                    Padding="16,12" Margin="0,0,0,20"
                    BorderBrush="#E5E7EB" BorderThickness="1">
                <StackPanel>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <Border Width="22" Height="22" CornerRadius="11"
                                Background="#8F6C05" Margin="0,0,10,0">
                            <TextBlock Text="1" FontSize="11" FontWeight="Bold"
                                       Foreground="White"
                                       HorizontalAlignment="Center"
                                       VerticalAlignment="Center"/>
                        </Border>
                        <TextBlock VerticalAlignment="Center"
                                   Foreground="#1F2937" FontSize="12">
                            Select one or more <Bold>ducts or pipes</Bold> in the model
                        </TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <Border Width="22" Height="22" CornerRadius="11"
                                Background="#8F6C05" Margin="0,0,10,0">
                            <TextBlock Text="2" FontSize="11" FontWeight="Bold"
                                       Foreground="White"
                                       HorizontalAlignment="Center"
                                       VerticalAlignment="Center"/>
                        </Border>
                        <TextBlock VerticalAlignment="Center"
                                   Foreground="#1F2937" FontSize="12">
                            Click <Bold>Pick in Viewport</Bold> to select now
                        </TextBlock>
                    </StackPanel>
                </StackPanel>
            </Border>

            <!-- Buttons -->
            <StackPanel Orientation="Horizontal"
                        HorizontalAlignment="Right">
                <Button x:Name="btnCancel" Content="Cancel"
                        Width="80" Height="30" Margin="0,0,10,0"
                        Background="White" BorderBrush="#D1D5DB"
                        BorderThickness="1" Cursor="Hand"
                        Click="btnCancel_Click"/>
                <Button x:Name="btnUseSel" Content="Use Selection"
                        Width="120" Height="30" Margin="0,0,8,0"
                        Visibility="Collapsed"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand"
                        FontWeight="SemiBold"
                        Click="btnUseSel_Click"
                        ToolTip="Proceed with the currently selected duct/pipe elements."/>
                <Button x:Name="btnPick" Content="Pick in Viewport"
                        Width="130" Height="30" Margin="0,0,8,0"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand"
                        FontWeight="SemiBold"
                        Click="btnPick_Click"/>
                <Button x:Name="btnPickPair" Content="Pick Part + Hanger"
                        Width="140" Height="30"
                        Background="#5C6BC0" Foreground="White"
                        BorderThickness="0" Cursor="Hand"
                        FontWeight="SemiBold"
                        Click="btnPickPair_Click"
                        ToolTip="Pick one duct/pipe, then one hanger. Syncs that single pair directly."/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"""

    def __init__(self, sel_count=0):
        import tempfile, os
        tmp = tempfile.NamedTemporaryFile(suffix=".xaml", delete=False, mode="w")
        tmp.write(self.XAML)
        tmp.close()
        self._xaml_path  = tmp.name
        self._sel_count  = sel_count
        self.picked      = False
        self.pick_pair   = False
        self.use_sel     = False
        forms.WPFWindow.__init__(self, self._xaml_path)
        self._apply_mode()

    def _apply_mode(self):
        from System.Windows import Visibility
        if self._sel_count > 0:
            self.txtSubtitle.Text = "%d element(s) already selected" % self._sel_count
            self.pnlSteps.Visibility   = Visibility.Collapsed
            self.btnUseSel.Visibility  = Visibility.Visible
            self.btnPick.Content       = "Re-pick in Viewport"
        else:
            self.txtSubtitle.Text = "Content selection required"
            self.pnlSteps.Visibility   = Visibility.Visible
            self.btnUseSel.Visibility  = Visibility.Collapsed

    def btnUseSel_Click(self, sender, e):
        self.use_sel = True
        self.Close()

    def btnPick_Click(self, sender, e):
        self.picked = True
        self.Close()

    def btnPickPair_Click(self, sender, e):
        self.picked     = False
        self.pick_pair  = True
        self.Close()

    def btnCancel_Click(self, sender, e):
        self.picked = False
        self.Close()

    def show(self):
        """Show dialog and return (wants_pick, wants_pair, wants_use_sel)."""
        self.ShowDialog()
        try:
            import os; os.unlink(self._xaml_path)
        except Exception:
            pass
        return self.picked, self.pick_pair, self.use_sel


class HangerSyncWindow(forms.WPFWindow):

    def __init__(self, xaml_path, rows):
        forms.WPFWindow.__init__(self, xaml_path)
        self._rows     = rows
        self._settings = _load_settings()
        self._apply_settings()
        self._populate()
        self._populate_insul_params()
        self._init_col_visibility()

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
            self.chkColType.IsChecked   = bool(s.get("col_type", True))
            self.chkColStatus.IsChecked = bool(s.get("col_status", False))
        except Exception: pass

    def _collect_settings(self):
        s = dict(_DEFAULT_SETTINGS)
        try: s["proximity_ft"] = float(self.txtProximity.Text.strip())
        except Exception: pass
        try: s["insulation_on"] = bool(self.chkInsulation.IsChecked)
        except Exception: pass
        try: s["insulation_param"] = self.cmbInsulParam.SelectedItem or ""
        except Exception: pass
        try:
            s["col_family"] = bool(self.chkColFamily.IsChecked)
            s["col_type"]   = bool(self.chkColType.IsChecked)
            s["col_status"] = bool(self.chkColStatus.IsChecked)
        except Exception: pass
        return s

    def ColVisibility_Changed(self, sender, e):
        from System.Windows import Visibility
        def vis(chk):
            return Visibility.Visible if bool(chk.IsChecked) else Visibility.Collapsed
        self.colFamily.Visibility = vis(self.chkColFamily)
        self.colType.Visibility   = vis(self.chkColType)
        self.colStatus.Visibility = vis(self.chkColStatus)

    def _init_col_visibility(self):
        """Collapse Status column on startup (checkbox unchecked by default)."""
        from System.Windows import Visibility
        self.colStatus.Visibility = Visibility.Collapsed

    def chkInsulation_Changed(self, sender, e):
        self.cmbInsulParam.IsEnabled = bool(self.chkInsulation.IsChecked)
        self._update_insul_preview()
        self._refresh_insulation_col()

    def _populate_insul_params(self):
        """
        Fill cmbInsulParam with all Double parameters from content elements.

        Sort order:
          1. Most-used first -- params with a non-zero value on the most
             content elements appear at the top (most relevant for insulation).
          2. Alphabetically within the same usage count.

        Insulation-related keywords still get a priority boost so they
        cluster at the top even when usage counts are equal.
        """
        self.cmbInsulParam.Items.Clear()

        # Count non-zero occurrences of each double param across all content
        from collections import defaultdict as _dd
        usage = _dd(int)
        param_names = set()
        for row in self._rows:
            try:
                for p in row._content.Parameters:
                    if p.StorageType == DB.StorageType.Double:
                        name = p.Definition.Name
                        param_names.add(name)
                        try:
                            if p.AsDouble() != 0.0:
                                usage[name] += 1
                        except Exception:
                            pass
            except Exception:
                pass

        def sort_key(n):
            nl      = n.lower()
            is_insul = 1 if any(kw in nl for kw in
                ("insul", "lining", "thickness", "adjust")) else 0
            count   = usage.get(n, 0)
            # Sort: insulation-related first, then by usage desc, then alpha
            return (1 - is_insul, -count, nl)

        for name in sorted(param_names, key=sort_key):
            self.cmbInsulParam.Items.Add(name)

        # Default selection: saved param > first insulation keyword match > index 0
        self.cmbInsulParam.SelectionChanged += self._insul_selection_changed
        saved_param = self._settings.get("insulation_param", "")
        selected = False
        if saved_param:
            for i in range(self.cmbInsulParam.Items.Count):
                if self.cmbInsulParam.Items[i] == saved_param:
                    self.cmbInsulParam.SelectedIndex = i
                    selected = True
                    break
        if not selected:
            for i in range(self.cmbInsulParam.Items.Count):
                if any(kw in self.cmbInsulParam.Items[i].lower()
                       for kw in ("insul", "thickness", "lining")):
                    self.cmbInsulParam.SelectedIndex = i
                    selected = True
                    break
        if not selected and self.cmbInsulParam.Items.Count > 0:
            self.cmbInsulParam.SelectedIndex = 0

        self._update_insul_preview()
        self._refresh_insulation_col()

    def _insul_selection_changed(self, sender, e):
        self._update_insul_preview()
        self._refresh_insulation_col()

    def _refresh_insulation_col(self):
        """Update the Insulation display value on each row from the selected param."""
        for row in self._rows:
            ins_ft = self._get_insulation_ft(row._content)
            if ins_ft is None or ins_ft == 0.0:
                row.Insulation = "none"
            else:
                row.Insulation = '%.4g"' % (ins_ft * 12)
        self._populate()

    def _update_insul_preview(self):
        """Show resolved value for selected param on first row content."""
        try:
            if not bool(self.chkInsulation.IsChecked):
                self.txtInsulPreview.Text = ""
                return
            param_name = self.cmbInsulParam.SelectedItem
            if not param_name or not self._rows:
                self.txtInsulPreview.Text = ""
                return
            p = self._rows[0]._content.LookupParameter(param_name)
            if p and p.StorageType == DB.StorageType.Double:
                self.txtInsulPreview.Text = "= %.4g\" on first element" % (p.AsDouble() * 12)
            else:
                self.txtInsulPreview.Text = "(not found on first element)"
        except Exception:
            self.txtInsulPreview.Text = ""

    def _get_insulation_ft(self, content_elem):
        """Return insulation thickness in feet from selected param, or None."""
        if not bool(self.chkInsulation.IsChecked):
            return None
        param_name = self.cmbInsulParam.SelectedItem
        if not param_name:
            return None
        try:
            p = content_elem.LookupParameter(param_name)
            if p and p.StorageType == DB.StorageType.Double:
                return p.AsDouble()
        except Exception:
            pass
        return None

    def _populate(self):
        col = ObservableCollection[object]()
        for r in self._rows:
            col.Add(r)
        self.dgHangers.ItemsSource = col
        ready  = sum(1 for r in self._rows if r.Status == "Ready")
        no_elv = sum(1 for r in self._rows if r.Status == "No elevation")
        synced = sum(1 for r in self._rows if r.Status == "Synced")
        self.txtStatus.Text = (
            "%d hanger(s)  |  %d ready  |  %d no elevation  |  %d synced"
            % (len(self._rows), ready, no_elv, synced))
        self.btnSyncAll.IsEnabled  = ready > 0
        

    def _get_proximity_ft(self):
        """Read proximity radius from txtProximity, fall back to PROX_FT."""
        try:
            val = float(self.txtProximity.Text.strip())
            if val > 0:
                return val
        except Exception:
            pass
        return PROX_FT

    def btnRescan_Click(self, sender, e):
        """Re-run proximity search using the current txtProximity value."""
        prox_ft = self._get_proximity_ft()
        content_map = {}
        for row in self._rows:
            eid = row._content.Id.IntegerValue
            if eid not in content_map:
                content_map[eid] = row._content
        rows = []
        seen = set()
        for content in content_map.values():
            for hanger, dist, t in find_hangers_near(content, prox_ft):
                if hanger.Id.IntegerValue not in seen:
                    seen.add(hanger.Id.IntegerValue)
                    rows.append(Row(content, hanger, dist, t))
        self._rows = rows
        self._populate()
        self._populate_insul_params()
        self.txtStatus.Text = ("Rescanned at %.2f ft -- %d hanger(s) found."
                               % (prox_ft, len(rows)))

    def btnRemoveRow_Click(self, sender, e):
        """Remove selected rows from the list without any Revit transaction."""
        selected = list(self.dgHangers.SelectedItems)
        if not selected:
            return
        for item in selected:
            if item in self._rows:
                self._rows.remove(item)
        self._populate()

    def btnSyncAll_Click(self, sender, e):
        done = errors = 0
        for row in self._rows:
            if row.Status not in ("Ready", "Ready (picked)"):
                continue
            if row._bottom_z is None:
                continue
            t = DB.Transaction(doc, "Hanger Sync - %s" % row.HangerId)
            try:
                t.Start()
                # Get insulation thickness from selected param (if enabled)
                ins_ft = self._get_insulation_ft(row._content) or 0.0

                # BOHElevation = pipe bottom minus insulation thickness.
                hung_bottom_z = row._bottom_z - ins_ft
                elev_param = write_elevation(row._hanger, hung_bottom_z)

                # Hanger clamp/trapeze size
                size_params = []
                # Clevis: nominal dia + 2x insulation
                if row._nominal_dia is not None:
                    insulated_dia = row._nominal_dia + 2.0 * ins_ft
                    try:
                        size_params = write_size(row._hanger, insulated_dia)
                    except Exception:
                        pass
                # Trapeze: duct width -> CP_Trapeze Width Desired
                if getattr(row, "_duct_width", None) is not None:
                    try:
                        tw = write_trapeze_width(row._hanger, row._duct_width)
                        if tw and tw not in size_params:
                            size_params.append(tw)
                    except Exception:
                        pass

                # XY re-center
                xy_moved = False
                try:
                    xy_moved = recenter_xy(row._hanger, row._content)
                except Exception:
                    pass

                # Write insulation thickness params for reference
                ins_param = None
                if ins_ft > 0:
                    try:
                        ins_param = write_insulation(row._hanger, ins_ft)
                    except Exception:
                        pass

                if elev_param:
                    t.Commit()
                    parts = [elev_param] + size_params if size_params else [elev_param]
                    if ins_ft > 0:
                        parts.append("ins=%.4g\"" % (ins_ft * 12))
                    if xy_moved:
                        parts.append("XY")
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
        self._populate()

    def btnClose_Click(self, sender, e):
        _save_settings(self._collect_settings())
        self.Close()

    def dgHangers_SelectionChanged(self, sender, e):
        has_selection = len(list(self.dgHangers.SelectedItems)) > 0
        self.btnRemoveRow.IsEnabled = has_selection


# ---------------------------------------------------------------------------
# INSULATION CHOICE WINDOW
# ---------------------------------------------------------------------------

class InsulationChoiceWindow(forms.WPFWindow):
    """
    Ask the user whether the hanger should snap to the pipe/duct surface
    or to the outside of the insulation.
    Also handles the case where insulation thickness cannot be auto-detected
    and shows a param dropdown for the user to select.
    """
    XAML = """
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Insulation Setting"
    Width="440" Height="380"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    FontFamily="Segoe UI"
    Background="#F5F7FB">

    <Grid>
        <Border VerticalAlignment="Top" Height="4" Background="#8F6C05"/>

        <StackPanel Margin="28,20,28,16" VerticalAlignment="Top">

            <StackPanel Orientation="Horizontal" Margin="0,0,0,14">
                <Border Width="36" Height="36" CornerRadius="18"
                        Background="#FDF5DC" Margin="0,0,12,0">
                    <TextBlock Text="&#x29C9;" FontSize="18"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"
                               Foreground="#8F6C05"/>
                </Border>
                <StackPanel VerticalAlignment="Center">
                    <TextBlock Text="Insulation detected"
                               FontSize="14" FontWeight="SemiBold"
                               Foreground="#0F1729"/>
                    <TextBlock x:Name="txtThickness"
                               Text="Thickness: unknown"
                               FontSize="11" Foreground="#6B7280"/>
                </StackPanel>
            </StackPanel>

            <!-- Param dropdown (shown only when thickness cant be detected) -->
            <Border x:Name="pnlParamPicker"
                    Background="White" CornerRadius="6"
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

            <!-- Snap choice -->
            <Border Background="White" CornerRadius="6"
                    Padding="12,10" Margin="0,0,0,16"
                    BorderBrush="#E5E7EB" BorderThickness="1">
                <StackPanel>
                    <TextBlock Text="Where should the hanger snap to?"
                               FontSize="11" FontWeight="SemiBold"
                               Foreground="#374151" Margin="0,0,0,8"/>
                    <RadioButton x:Name="rdoOutsideInsulation"
                                 Content="Outside of insulation  (hanger wraps the insulated OD)"
                                 IsChecked="True" Margin="0,0,0,6"
                                 Foreground="#1F2937"/>
                    <RadioButton x:Name="rdoPipeSurface"
                                 Content="Pipe / duct surface  (bare metal, ignore insulation)"
                                 Foreground="#1F2937"/>
                </StackPanel>
            </Border>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnCancel" Content="Cancel"
                        Width="80" Height="28" Margin="0,0,10,0"
                        Background="White" BorderBrush="#D1D5DB"
                        BorderThickness="1" Cursor="Hand"
                        Click="btnCancel_Click"/>
                <Button x:Name="btnApply" Content="Apply"
                        Width="90" Height="28"
                        Background="#8F6C05" Foreground="White"
                        BorderThickness="0" Cursor="Hand"
                        FontWeight="SemiBold"
                        Click="btnApply_Click"/>
            </StackPanel>
        </StackPanel>
    </Grid>
</Window>
"""

    def __init__(self, content_elem, detected_thickness_ft):
        import tempfile
        tmp = tempfile.NamedTemporaryFile(suffix=".xaml", delete=False, mode="w")
        tmp.write(self.XAML)
        tmp.close()
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
            self.txtThickness.Text = "Thickness: %.4g in" % (
                self._detected_thickness * 12)
        else:
            # Can't detect -- show param picker
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
            priority = 0 if any(kw in nl for kw in
                ("insul", "lining", "thickness")) else 1
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
            # Read from chosen param
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

    def btnCancel_Click(self, sender, e):
        self.confirmed = False
        self.Close()

    def show(self):
        self.ShowDialog()
        try:
            import os as _os
            _os.unlink(self._xaml_path)
        except Exception:
            pass
        return self.confirmed


# ---------------------------------------------------------------------------
# PICK-PAIR WORKFLOW
# ---------------------------------------------------------------------------

class HangerFilter(ISelectionFilter):
    """Allow only Generic Model FamilyInstances."""
    def AllowElement(self, e):
        if e is None or e.Category is None:
            return False
        if e.Category.Id.IntegerValue != int(DB.BuiltInCategory.OST_GenericModel):
            return False
        return isinstance(e, DB.FamilyInstance)
    def AllowReference(self, r, p):
        return False


def run_pick_pair_workflow():
    """
    Pick-one-part + pick-one-hanger workflow.
    Independent of content pre-selection.
    """
    # ── Step 1: pick content ──────────────────────────────────────────────────
    try:
        ref_content = uidoc.Selection.PickObject(
            ObjectType.Element,
            HostFilter(),
            "Step 1 of 2 -- Pick the duct or pipe")
        content = doc.GetElement(ref_content.ElementId)
    except Exception:
        return  # cancelled

    if content is None:
        forms.alert("Could not read selected element.", title="Hanger Sync")
        return

    # ── Step 2: pick hanger ───────────────────────────────────────────────────
    try:
        ref_hanger = uidoc.Selection.PickObject(
            ObjectType.Element,
            HangerFilter(),
            "Step 2 of 2 -- Pick the hanger to update")
        hanger = doc.GetElement(ref_hanger.ElementId)
    except Exception:
        return  # cancelled

    if hanger is None:
        forms.alert("Could not read selected hanger.", title="Hanger Sync")
        return

    # ── Compute hanger position along content curve ───────────────────────────
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

    # ── Detect insulation ─────────────────────────────────────────────────────
    detected_ins = get_insulation_thickness(content)
    snap_ins_ft  = 0.0

    if detected_ins and detected_ins > 0:
        # Ask user: snap to insulation surface or bare surface?
        dlg = InsulationChoiceWindow(content, detected_ins)
        confirmed = dlg.show()
        if not confirmed:
            return
        snap_ins_ft = dlg.insulation_ft if dlg.snap_to_insulation else 0.0
    else:
        # Try to detect from params -- if still nothing, offer dropdown
        dlg = InsulationChoiceWindow(content, 0.0)
        confirmed = dlg.show()
        if not confirmed:
            return
        snap_ins_ft = dlg.insulation_ft if dlg.snap_to_insulation else 0.0

    # ── Compute values ────────────────────────────────────────────────────────
    bottom_z    = get_pipe_bottom_at_t(content, t)
    nominal_dia = get_nominal_dia(content)
    duct_width  = get_duct_width(content)

    if bottom_z is None:
        forms.alert("Could not determine elevation from selected content.",
                    title="Hanger Sync")
        return

    hung_bottom_z  = bottom_z - snap_ins_ft
    insulated_dia  = (nominal_dia + 2.0 * snap_ins_ft) if nominal_dia else None

    # ── Write ─────────────────────────────────────────────────────────────────
    t_write = DB.Transaction(doc, "Hanger Sync - Pick Pair")
    try:
        t_write.Start()
        written = []

        elev_p = write_elevation(hanger, hung_bottom_z)
        if elev_p:
            written.append(elev_p)

        if insulated_dia is not None:
            try:
                size_written = write_size(hanger, insulated_dia)
                written.extend(size_written)
            except Exception:
                pass

        if duct_width is not None:
            try:
                tw = write_trapeze_width(hanger, duct_width)
                if tw:
                    written.append(tw)
            except Exception:
                pass

        try:
            recenter_xy(hanger, content)
            written.append("XY")
        except Exception:
            pass

        if snap_ins_ft > 0:
            try:
                write_insulation(hanger, snap_ins_ft)
                written.append("Insulation")
            except Exception:
                pass

        t_write.Commit()
        forms.alert(
            "Hanger %s updated. Written: %s" % (
                hanger.Id.IntegerValue, ", ".join(written) if written else "none"),
            title="Hanger Sync -- Done")

    except Exception as ex:
        try:
            if t_write.GetStatus() == DB.TransactionStatus.Started:
                t_write.RollBack()
        except Exception:
            pass
        forms.alert("Sync failed: %s" % str(ex), title="Hanger Sync")


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

# Check current selection for valid content first
sel_ids = uidoc.Selection.GetElementIds()
valid_cat_ids = [int(c) for c in HOST_CATS]
content_elems = [
    doc.GetElement(i) for i in sel_ids
    if doc.GetElement(i) is not None
    and doc.GetElement(i).Category is not None
    and doc.GetElement(i).Category.Id.IntegerValue in valid_cat_ids
]

# Always show the launch dialog -- even with pre-selection,
# user may want to switch to Pick Part + Hanger workflow.
dlg                       = SelectContentWindow(len(content_elems))
result, do_pair, use_sel  = dlg.show()

if do_pair:
    run_pick_pair_workflow()
else:
    if use_sel:
        pass  # content_elems already populated from pre-selection
    elif result:
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                HostFilter(),
                "Pick duct / pipe to sync hangers for (multi-select OK)")
            content_elems = [
                doc.GetElement(r.ElementId)
                for r in refs
                if doc.GetElement(r.ElementId) is not None
            ]
        except Exception:
            content_elems = []
    else:
        content_elems = []  # cancelled

    if not content_elems:
        pass  # user cancelled -- exit silently
    else:
        # Use saved proximity setting so the user's last value is active immediately
        _init_settings = _load_settings()
        _init_prox_ft  = float(_init_settings.get("proximity_ft", PROX_FT))

        # Find Generic Model hangers near the selected content
        rows = []
        seen = set()
        for content in content_elems:
            for hanger, dist, t in find_hangers_near(content, _init_prox_ft):
                if hanger.Id.IntegerValue not in seen:
                    seen.add(hanger.Id.IntegerValue)
                    rows.append(Row(content, hanger, dist, t))

        if not rows:
            forms.alert(
                "No Generic Model hangers found within %.2f' XY of the selected content.\n\n"
                "Adjust the Max hanger distance setting and re-run, or select "
                "different content." % _init_prox_ft,
                title="Hanger Sync")
        else:
            xaml_path = os.path.join(os.path.dirname(__file__), "window.xaml")
            w = HangerSyncWindow(xaml_path, rows)
            w.dgHangers.SelectionChanged += w.dgHangers_SelectionChanged
            w.ShowDialog()
