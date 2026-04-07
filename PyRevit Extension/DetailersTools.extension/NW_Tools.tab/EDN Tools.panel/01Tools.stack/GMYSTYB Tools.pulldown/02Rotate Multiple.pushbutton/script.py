# -*- coding: utf-8 -*-
# Rotate Multiple v0.2.0 -- pyRevit tool
# Single WPF dialog with ReNumber-style selection management.
# Pre-selects current Revit selection, supports Pick Elements re-pick,
# and deselect by Individual / Category / Family / Type.

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from System.Collections.ObjectModel import ObservableCollection
import math
import os
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
log   = script.get_logger()

XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")


# ================================================================
#  DATA CLASSES
# ================================================================
class DeselectItem(object):
    def __init__(self, name, count, element_ids=None):
        self.Name = name
        self.Count = "({} elements)".format(count) if count > 1 else "(1 element)"
        self.ElementIds = element_ids or []


# ================================================================
#  SELECTION FILTERS
# ================================================================
class RotatableSelFilter(ISelectionFilter):
    def AllowElement(self, e):
        try:
            if e is None or isinstance(e, DB.ElementType):
                return False
            cat = e.Category
            if cat and cat.CategoryType in (DB.CategoryType.Model,
                                            DB.CategoryType.Annotation):
                if hasattr(DB, "RevitLinkInstance") and isinstance(e, DB.RevitLinkInstance):
                    return False
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, ref, pt):
        return False


class ReferenceSelFilter(ISelectionFilter):
    def AllowElement(self, e):
        if e is None:
            return False
        try:
            if isinstance(e, DB.Grid):
                return True
            if isinstance(e, DB.CurveElement):
                return True
            cat = e.Category
            if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_VolumeOfInterest):
                return True
            loc = getattr(e, "Location", None)
            if isinstance(loc, DB.LocationCurve):
                return True
            if isinstance(e, DB.FamilyInstance):
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, ref, pt):
        return False


# ================================================================
#  DIRECTION EXTRACTION
# ================================================================
def extract_reference_direction(element):
    if isinstance(element, DB.Grid):
        curve = element.Curve
        if curve:
            return _direction_from_curve(curve)

    if isinstance(element, DB.CurveElement):
        curve = element.GeometryCurve
        if curve:
            return _direction_from_curve(curve)

    cat = element.Category
    if cat and cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_VolumeOfInterest):
        d = _scope_box_direction(element)
        if d:
            return d

    loc = getattr(element, "Location", None)
    if isinstance(loc, DB.LocationCurve):
        return _direction_from_curve(loc.Curve)

    if isinstance(element, DB.FamilyInstance):
        try:
            facing = element.FacingOrientation
            if facing and facing.GetLength() > 1e-6:
                return facing.Normalize()
        except Exception:
            pass
    return None


def _direction_from_curve(curve):
    try:
        d = curve.GetEndPoint(1) - curve.GetEndPoint(0)
        if d.GetLength() > 1e-6:
            return d.Normalize()
    except Exception:
        pass
    return None


def _scope_box_direction(element):
    opt = DB.Options()
    geom = element.get_Geometry(opt)
    if not geom:
        return None
    best_dir, best_len = None, 0.0
    for obj in geom:
        candidates = []
        if isinstance(obj, DB.Line):
            candidates.append(obj)
        elif isinstance(obj, DB.GeometryInstance):
            try:
                for sub in obj.GetInstanceGeometry():
                    if isinstance(sub, DB.Line):
                        candidates.append(sub)
            except Exception:
                pass
        for ln in candidates:
            d = ln.GetEndPoint(1) - ln.GetEndPoint(0)
            h_len = math.sqrt(d.X ** 2 + d.Y ** 2)
            if abs(d.Z) < 0.01 * max(h_len, 1e-6) and h_len > best_len:
                best_len = h_len
                best_dir = DB.XYZ(d.X, d.Y, 0).Normalize()
    return best_dir


# ================================================================
#  ELEMENT ORIENTATION
# ================================================================
def get_element_current_angle(element, axis_dir):
    direction = None
    if isinstance(element, DB.FamilyInstance):
        try:
            facing = element.FacingOrientation
            if facing and facing.GetLength() > 1e-6:
                direction = facing
        except Exception:
            pass

    if direction is None:
        loc = getattr(element, "Location", None)
        if isinstance(loc, DB.LocationCurve):
            try:
                c = loc.Curve
                direction = c.GetEndPoint(1) - c.GetEndPoint(0)
            except Exception:
                pass

    if direction is None:
        loc = getattr(element, "Location", None)
        if isinstance(loc, DB.LocationPoint):
            try:
                rot = loc.Rotation
                direction = DB.XYZ(math.cos(rot), math.sin(rot), 0)
            except Exception:
                pass

    if direction is None:
        return None
    proj = _project_to_plane(direction, axis_dir)
    if proj is None:
        return None
    return _angle_in_plane(proj, axis_dir)


def _project_to_plane(vec, normal):
    dot = vec.X * normal.X + vec.Y * normal.Y + vec.Z * normal.Z
    proj = DB.XYZ(vec.X - normal.X * dot,
                  vec.Y - normal.Y * dot,
                  vec.Z - normal.Z * dot)
    if proj.GetLength() < 1e-6:
        return None
    return proj.Normalize()


def _angle_in_plane(vec, normal):
    if normal.IsAlmostEqualTo(DB.XYZ.BasisZ):
        return math.atan2(vec.Y, vec.X)
    ref = DB.XYZ.BasisX
    dot = ref.X * normal.X + ref.Y * normal.Y + ref.Z * normal.Z
    ref_proj = DB.XYZ(ref.X - normal.X * dot,
                      ref.Y - normal.Y * dot,
                      ref.Z - normal.Z * dot)
    if ref_proj.GetLength() < 1e-6:
        ref = DB.XYZ.BasisY
        dot = ref.X * normal.X + ref.Y * normal.Y + ref.Z * normal.Z
        ref_proj = DB.XYZ(ref.X - normal.X * dot,
                          ref.Y - normal.Y * dot,
                          ref.Z - normal.Z * dot)
    ref_proj = ref_proj.Normalize()
    cross = DB.XYZ(
        normal.Y * ref_proj.Z - normal.Z * ref_proj.Y,
        normal.Z * ref_proj.X - normal.X * ref_proj.Z,
        normal.X * ref_proj.Y - normal.Y * ref_proj.X
    ).Normalize()
    return math.atan2(
        vec.X * cross.X + vec.Y * cross.Y + vec.Z * cross.Z,
        vec.X * ref_proj.X + vec.Y * ref_proj.Y + vec.Z * ref_proj.Z
    )


# ================================================================
#  ROTATION HELPERS
# ================================================================
def get_center_point(el, active_view):
    loc = getattr(el, "Location", None)
    if isinstance(loc, DB.LocationPoint):
        return loc.Point
    try:
        bbox = el.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass
    try:
        bboxv = el.get_BoundingBox(active_view)
        if bboxv:
            return (bboxv.Min + bboxv.Max) * 0.5
    except Exception:
        pass
    return None


def get_bbox_center(el, active_view):
    try:
        bbox = el.get_BoundingBox(None) or el.get_BoundingBox(active_view)
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass
    return None


def make_axis_through(point, direction_xyz):
    d = direction_xyz.Normalize()
    return DB.Line.CreateBound(point - d, point + d)


def has_pinned_elements(elements):
    for el in elements:
        try:
            if getattr(el, "Pinned", False):
                return True
        except Exception:
            pass
    return False


def friendly_ref_label(element):
    cat = element.Category
    cat_name = cat.Name if cat else "Element"
    name = getattr(element, "Name", "")
    eid = element.Id.IntegerValue
    if name:
        return "{}: {} (id {})".format(cat_name, name, eid)
    return "{} (id {})".format(cat_name, eid)


# ================================================================
#  ELEMENT HELPERS (for deselect grouping)
# ================================================================
def get_element_family(elem):
    try:
        et = doc.GetElement(elem.GetTypeId())
        if et:
            fp = et.get_Parameter(
                DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            if fp and fp.HasValue:
                return fp.AsString()
    except Exception:
        pass
    return "Unknown Family"


def get_element_type_name(elem):
    try:
        et = doc.GetElement(elem.GetTypeId())
        if et:
            tn = et.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if tn and tn.HasValue:
                return tn.AsString()
            return et.Name if hasattr(et, "Name") else "Unknown Type"
    except Exception:
        pass
    return "Unknown Type"


def get_element_display_name(elem):
    eid = elem.Id.IntegerValue
    cat = elem.Category.Name if elem.Category else "Unknown"
    name = ""
    try:
        np = elem.get_Parameter(DB.BuiltInParameter.ELEM_NAME_PARAM)
        if np and np.HasValue:
            name = np.AsString()
    except Exception:
        pass
    if name:
        return "ID {} - {} - {}".format(eid, cat, name)
    return "ID {} - {}".format(eid, cat)


# ================================================================
#  ELEMENT PICKER
# ================================================================
def pick_multiple():
    flt = RotatableSelFilter()
    picked = []
    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, flt,
            "Pick elements to rotate (Finish to confirm)")
        for r in refs:
            elem = doc.GetElement(r.ElementId)
            if elem:
                picked.append(elem)
    except Exception:
        pass
    return picked


def filter_preselection(elements):
    """Filter a raw preselection list to only rotatable elements."""
    flt = RotatableSelFilter()
    return [e for e in elements if flt.AllowElement(e)]


# ================================================================
#  WPF WINDOW CONTROLLER
# ================================================================
class RotateMultipleWindow(object):

    def __init__(self, xaml_path, preloaded_elements=None):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        # Element lists
        self.elements = list(preloaded_elements or [])

        # State
        self.ref_direction  = None
        self.ref_label      = ""
        self.confirmed      = False
        self._pending_pick  = None   # "ref", "pts", or "elements"
        self._next_action   = None

        self._get_controls()
        self._wire_events()
        self._update_element_count()
        self._refresh_deselect_list()
        self._update_go_button()

    def _get_controls(self):
        find = self.window.FindName
        self.txt_elem_count     = find("TxtElementCount")
        self.btn_pick_elements  = find("BtnPickElements")
        # Deselect
        self.rb_desel_individual = find("RbDeselectIndividual")
        self.rb_desel_category   = find("RbDeselectCategory")
        self.rb_desel_family     = find("RbDeselectFamily")
        self.rb_desel_type       = find("RbDeselectType")
        self.lst_deselect        = find("LstDeselect")
        self.btn_deselect        = find("BtnDeselect")
        self.btn_deselect_all    = find("BtnDeselectAll")
        # Angle
        self.rb_manual      = find("RbManual")
        self.rb_align        = find("RbAlign")
        self.pnl_manual      = find("PnlManual")
        self.pnl_align       = find("PnlAlign")
        self.txt_angle       = find("TxtAngle")
        self.btn_pick_ref    = find("BtnPickRef")
        self.btn_pick_pts    = find("BtnPickPts")
        self.txt_ref_status  = find("TxtRefStatus")
        self.rb_parallel     = find("RbParallel")
        self.rb_perp         = find("RbPerpendicular")
        # Axis / Center / Options
        self.rb_world_z      = find("RbWorldZ")
        self.rb_view_normal  = find("RbViewNormal")
        self.rb_center_auto  = find("RbCenterAuto")
        self.rb_center_bbox  = find("RbCenterBBox")
        self.chk_unpin       = find("ChkUnpin")
        # Footer
        self.btn_go     = find("BtnGo")
        self.btn_cancel = find("BtnCancel")
        self.txt_status = find("TxtStatus")

    def _wire_events(self):
        self.rb_manual.Checked  += self._on_mode_changed
        self.rb_align.Checked   += self._on_mode_changed
        self.btn_pick_ref.Click += self._on_pick_ref
        self.btn_pick_pts.Click += self._on_pick_pts
        self.btn_pick_elements.Click += self._on_pick_elements
        self.btn_deselect.Click     += self._on_deselect
        self.btn_deselect_all.Click += self._on_deselect_all
        self.rb_desel_individual.Checked += self._on_deselect_mode_changed
        self.rb_desel_category.Checked   += self._on_deselect_mode_changed
        self.rb_desel_family.Checked     += self._on_deselect_mode_changed
        self.rb_desel_type.Checked       += self._on_deselect_mode_changed
        self.btn_go.Click     += self._on_go
        self.btn_cancel.Click += self._on_cancel
        self.window.KeyDown   += self._on_key

    # ---- angle mode panel toggle ----
    def _on_mode_changed(self, sender, args):
        if self.rb_manual.IsChecked:
            self.pnl_manual.Visibility = Visibility.Visible
            self.pnl_align.Visibility  = Visibility.Collapsed
        else:
            self.pnl_manual.Visibility = Visibility.Collapsed
            self.pnl_align.Visibility  = Visibility.Visible

    # ---- pick-loop triggers (flag + hide) ----
    def _on_pick_ref(self, sender, args):
        self._pending_pick = "ref"
        self.window.Hide()

    def _on_pick_pts(self, sender, args):
        self._pending_pick = "pts"
        self.window.Hide()

    def _on_pick_elements(self, sender, args):
        self._pending_pick = "elements"
        self.window.Hide()

    # ---- go / cancel ----
    def _on_go(self, sender, args):
        if not self.elements:
            self.txt_status.Text = "No elements selected."
            return
        if self.rb_manual.IsChecked:
            try:
                float(self.txt_angle.Text)
            except Exception:
                self.txt_status.Text = "Angle must be a number."
                return
        else:
            if self.ref_direction is None:
                self.txt_status.Text = "Pick a reference first."
                return
        self.confirmed = True
        self.window.Close()

    def _on_cancel(self, sender, args):
        self.confirmed = False
        self.window.Close()

    def _on_key(self, sender, args):
        if args.Key == System.Windows.Input.Key.Escape:
            self.confirmed = False
            self.window.Close()

    # ---- element count ----
    def _update_element_count(self):
        count = len(self.elements)
        if count == 0:
            self.txt_elem_count.Text = "(none)"
        else:
            self.txt_elem_count.Text = "{} elements".format(count)

    def _update_go_button(self):
        self.btn_go.IsEnabled = len(self.elements) > 0

    # ---- deselect logic ----
    def _on_deselect_mode_changed(self, sender, args):
        self._refresh_deselect_list()

    def _refresh_deselect_list(self):
        if self.rb_desel_individual.IsChecked:
            self._populate_individual()
        elif self.rb_desel_category.IsChecked:
            self._populate_category()
        elif self.rb_desel_family.IsChecked:
            self._populate_family()
        elif self.rb_desel_type.IsChecked:
            self._populate_type()

    def _populate_individual(self):
        items = ObservableCollection[DeselectItem]()
        for elem in self.elements:
            display = get_element_display_name(elem)
            items.Add(DeselectItem(display, 1,
                                   [str(elem.Id.IntegerValue)]))
        self.lst_deselect.ItemsSource = items

    def _populate_category(self):
        groups = {}
        for elem in self.elements:
            cn = elem.Category.Name if elem.Category else "No Category"
            groups.setdefault(cn, []).append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for cn in sorted(groups.keys()):
            items.Add(DeselectItem(cn, len(groups[cn]), groups[cn]))
        self.lst_deselect.ItemsSource = items

    def _populate_family(self):
        groups = {}
        for elem in self.elements:
            fn = get_element_family(elem)
            groups.setdefault(fn, []).append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for fn in sorted(groups.keys()):
            items.Add(DeselectItem(fn, len(groups[fn]), groups[fn]))
        self.lst_deselect.ItemsSource = items

    def _populate_type(self):
        groups = {}
        for elem in self.elements:
            tn = get_element_type_name(elem)
            groups.setdefault(tn, []).append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for tn in sorted(groups.keys()):
            items.Add(DeselectItem(tn, len(groups[tn]), groups[tn]))
        self.lst_deselect.ItemsSource = items

    def _on_deselect(self, sender, args):
        selected = list(self.lst_deselect.SelectedItems)
        if not selected:
            return
        ids_to_remove = set()
        if self.rb_desel_individual.IsChecked:
            for item in selected:
                for eid in item.ElementIds:
                    ids_to_remove.add(str(eid))
        elif self.rb_desel_category.IsChecked:
            cats = set(item.Name for item in selected)
            for elem in self.elements:
                cn = elem.Category.Name if elem.Category else "No Category"
                if cn in cats:
                    ids_to_remove.add(str(elem.Id.IntegerValue))
        elif self.rb_desel_family.IsChecked:
            fams = set(item.Name for item in selected)
            for elem in self.elements:
                if get_element_family(elem) in fams:
                    ids_to_remove.add(str(elem.Id.IntegerValue))
        elif self.rb_desel_type.IsChecked:
            types = set(item.Name for item in selected)
            for elem in self.elements:
                if get_element_type_name(elem) in types:
                    ids_to_remove.add(str(elem.Id.IntegerValue))

        self.elements = [e for e in self.elements
                         if str(e.Id.IntegerValue) not in ids_to_remove]
        self._update_element_count()
        self._refresh_deselect_list()
        self._update_go_button()

    def _on_deselect_all(self, sender, args):
        count = len(self.elements)
        self.elements = []
        self._update_element_count()
        self._refresh_deselect_list()
        self._update_go_button()
        self.txt_status.Text = "Deselected all {} element(s).".format(count)

    # ---- show loop (pick-loop pattern) ----
    def show(self):
        while True:
            self._pending_pick = None
            self.window.ShowDialog()

            if self._pending_pick == "ref":
                self._do_pick_ref()
                continue
            elif self._pending_pick == "pts":
                self._do_pick_pts()
                continue
            elif self._pending_pick == "elements":
                self._do_pick_elements()
                continue
            else:
                break
        return self.confirmed

    def _do_pick_ref(self):
        try:
            ref_obj = uidoc.Selection.PickObject(
                ObjectType.Element,
                ReferenceSelFilter(),
                "Pick a reference element "
                "(duct, pipe, line, grid, scope box, family)"
            )
            el = doc.GetElement(ref_obj.ElementId)
            if el:
                raw = extract_reference_direction(el)
                if raw:
                    axis_dir = self._current_axis()
                    proj = _project_to_plane(raw, axis_dir)
                    if proj:
                        self.ref_direction = proj
                        self.ref_label = friendly_ref_label(el)
                        self.txt_ref_status.Text = self.ref_label
                        self.txt_status.Text = "Reference set."
                    else:
                        self.txt_ref_status.Text = \
                            "(direction parallel to axis, cannot align)"
                        self.ref_direction = None
                else:
                    self.txt_ref_status.Text = \
                        "(could not extract direction)"
                    self.ref_direction = None
        except Exception:
            pass

    def _do_pick_pts(self):
        try:
            pt1 = uidoc.Selection.PickPoint(
                "Pick the FIRST point (origin)")
            pt2 = uidoc.Selection.PickPoint(
                "Pick the SECOND point (direction)")
            raw = pt2 - pt1
            if raw.GetLength() < 1e-6:
                self.txt_ref_status.Text = "(points too close together)"
                self.ref_direction = None
            else:
                raw = raw.Normalize()
                axis_dir = self._current_axis()
                proj = _project_to_plane(raw, axis_dir)
                if proj:
                    self.ref_direction = proj
                    deg = math.degrees(math.atan2(proj.Y, proj.X))
                    self.ref_label = \
                        "Two-point direction ({:.1f} deg)".format(deg)
                    self.txt_ref_status.Text = self.ref_label
                    self.txt_status.Text = "Reference set."
                else:
                    self.txt_ref_status.Text = \
                        "(direction parallel to axis)"
                    self.ref_direction = None
        except Exception:
            pass

    def _do_pick_elements(self):
        new_elems = pick_multiple()
        if new_elems:
            self.elements = new_elems
            self._update_element_count()
            self._refresh_deselect_list()
            self._update_go_button()
            self.txt_status.Text = "Picked {} elements.".format(
                len(self.elements))

    # ---- helpers ----
    def _current_axis(self):
        if self.rb_world_z.IsChecked:
            return DB.XYZ.BasisZ
        return doc.ActiveView.ViewDirection

    # ---- collected settings ----
    @property
    def is_align_mode(self):
        return self.rb_align.IsChecked

    @property
    def manual_angle_rad(self):
        return math.radians(float(self.txt_angle.Text))

    @property
    def axis_dir(self):
        return self._current_axis()

    @property
    def is_perpendicular(self):
        return self.rb_perp.IsChecked

    @property
    def force_bbox(self):
        return self.rb_center_bbox.IsChecked

    @property
    def auto_unpin(self):
        return self.chk_unpin.IsChecked

    def get_target_direction(self):
        d = self.ref_direction
        if d is None:
            return None
        if self.is_perpendicular:
            ax = self.axis_dir
            perp = DB.XYZ(
                ax.Y * d.Z - ax.Z * d.Y,
                ax.Z * d.X - ax.X * d.Z,
                ax.X * d.Y - ax.Y * d.X
            )
            if perp.GetLength() < 1e-6:
                return None
            return perp.Normalize()
        return d


# ================================================================
#  MAIN
# ================================================================
def main():
    # Grab pre-selection
    preselection = filter_preselection(
        list(revit.get_selection() or []))

    # Show dialog
    win = RotateMultipleWindow(XAML_PATH, preselection)
    if not win.show():
        return

    # Settings
    align_mode  = win.is_align_mode
    axis_dir    = win.axis_dir
    force_bbox  = win.force_bbox
    auto_unpin  = win.auto_unpin
    elements    = win.elements

    if not elements:
        forms.alert("No elements to rotate.", warn_icon=True)
        return

    fixed_angle   = None
    ref_angle_val = None

    if align_mode:
        target_dir = win.get_target_direction()
        if target_dir is None:
            forms.alert("Could not compute target direction.",
                        exitscript=True)
        ref_angle_val = _angle_in_plane(target_dir, axis_dir)
    else:
        fixed_angle = win.manual_angle_rad

    # Pinned handling
    unpin_choice = auto_unpin
    if not auto_unpin and has_pinned_elements(elements):
        unpin_choice = forms.alert(
            "Some elements are pinned. Temporarily unpin them?",
            yes=True, no=True, title="Rotate Multiple"
        )

    # Execute
    ok, skipped, no_orient, failed = 0, 0, 0, 0

    with forms.ProgressBar(title="Rotating...",
                           cancellable=True, step=1) as pb:
        tgroup = DB.TransactionGroup(doc, "Rotate Multiple")
        tgroup.Start()
        try:
            for i, el in enumerate(elements):
                pb.update_progress(i, len(elements))
                if pb.cancelled:
                    break

                pt = (get_bbox_center(el, doc.ActiveView) if force_bbox
                      else get_center_point(el, doc.ActiveView))
                if not pt:
                    skipped += 1
                    continue

                if align_mode:
                    cur_angle = get_element_current_angle(el, axis_dir)
                    if cur_angle is None:
                        no_orient += 1
                        continue
                    angle_rad = ref_angle_val - cur_angle
                    angle_rad = ((angle_rad + math.pi)
                                 % (2 * math.pi) - math.pi)
                    if abs(angle_rad) < math.radians(0.1):
                        ok += 1
                        continue
                else:
                    angle_rad = fixed_angle

                axis = make_axis_through(pt, axis_dir)

                t = DB.Transaction(doc, "Rotate Element")
                t.Start()
                try:
                    was_pinned = False
                    try:
                        if getattr(el, "Pinned", False):
                            if not unpin_choice:
                                t.RollBack()
                                skipped += 1
                                continue
                            was_pinned = True
                            el.Pinned = False
                    except Exception:
                        pass

                    DB.ElementTransformUtils.RotateElement(
                        doc, el.Id, axis, angle_rad)

                    if was_pinned:
                        try:
                            el.Pinned = True
                        except Exception:
                            pass

                    t.Commit()
                    ok += 1
                except Exception as ex:
                    log.error("Rotate failed on {}: {}".format(
                        el.Id.IntegerValue, ex))
                    t.RollBack()
                    failed += 1
        finally:
            if ok > 0:
                tgroup.Assimilate()
            else:
                tgroup.RollBack()

    # Summary
    parts = ["Rotated: {}".format(ok)]
    if skipped:
        parts.append("Skipped (no center): {}".format(skipped))
    if no_orient:
        parts.append("Skipped (no orientation): {}".format(no_orient))
    if failed:
        parts.append("Errors: {}".format(failed))
    msg = "\n".join(parts)
    forms.alert(msg, title="Rotate Multiple",
                warn_icon=(failed > 0 or no_orient > 0))


if __name__ == "__main__":
    main()
