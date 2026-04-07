# -*- coding: utf-8 -*-
# Rotate Multiple — pyRevit tool
# Rotates each chosen element about its own center.
# Selection modes: Pick | Box | All In View | Visible In View | Use Current Selection
# Axis: World Z (Up) or Active View normal
# Center: Auto (Location then BBox) or Force BBox center
# Pinned handling: optional auto-unpin/re-pin (via Yes/No prompt, only if pinned elements exist)

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
import math

uidoc = revit.uidoc
doc   = revit.doc
log   = script.get_logger()

# ---------------- selection filter ----------------
class RotatableSelFilter(ISelectionFilter):
    def AllowElement(self, e):
        try:
            if e is None:
                return False
            # Exclude element types and links (common non-rotatables)
            if isinstance(e, DB.ElementType):
                return False
            cat = e.Category
            # Model + Annotation are usually fine; leave system/links out
            if cat and (cat.CategoryType in (DB.CategoryType.Model, DB.CategoryType.Annotation)):
                # Skip Revit Links explicitly
                if hasattr(DB, "RevitLinkInstance") and isinstance(e, DB.RevitLinkInstance):
                    return False
                return True
        except Exception:
            pass
        return False

    def AllowReference(self, ref, pt):
        return False

# ---------------- selection helpers ----------------
def choose_mode():
    modes = [
        "Pick",
        "Box Select",
        "All In View (all elements)",
        "Visible In View (respects visibility/filters)",
        "Use Current Selection",
    ]
    return forms.CommandSwitchWindow.show(modes, message="Select elements to rotate")

def collect_targets(mode):
    ids = set()
    flt = RotatableSelFilter()

    if mode == "Pick":
        refs = uidoc.Selection.PickObjects(ObjectType.Element, flt, "Pick elements to rotate")
        ids.update(r.ElementId.IntegerValue for r in refs)

    elif mode == "Box Select":
        elems = uidoc.Selection.PickElementsByRectangle(flt, "Drag a box to select elements to rotate")
        ids.update(e.Id.IntegerValue for e in elems)

    elif mode.startswith("All In View"):
        viewid = doc.ActiveView.Id
        coll = DB.FilteredElementCollector(doc, viewid).WhereElementIsNotElementType()
        for e in coll:
            if flt.AllowElement(e):
                ids.add(e.Id.IntegerValue)

    elif mode.startswith("Visible In View"):
        viewid = doc.ActiveView.Id
        coll = (DB.FilteredElementCollector(doc, viewid)
                .WhereElementIsNotElementType()
                .WherePasses(DB.VisibleInViewFilter(doc, viewid)))
        for e in coll:
            if flt.AllowElement(e):
                ids.add(e.Id.IntegerValue)

    elif mode == "Use Current Selection":
        for eid in uidoc.Selection.GetElementIds():
            e = doc.GetElement(eid)
            if e and flt.AllowElement(e):
                ids.add(eid.IntegerValue)

    return [doc.GetElement(DB.ElementId(i)) for i in sorted(ids)]

# ---------------- rotation helpers ----------------
def get_center_point(el, active_view):
    """Prefer a true insertion point; else bbox center (model, then view)."""
    loc = getattr(el, "Location", None)
    if isinstance(loc, DB.LocationPoint):
        return loc.Point

    # model bbox first
    try:
        bbox = el.get_BoundingBox(None)
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass

    # fallback: view bbox
    try:
        bboxv = el.get_BoundingBox(active_view)
        if bboxv:
            return (bboxv.Min + bboxv.Max) * 0.5
    except Exception:
        pass

    return None

def get_bbox_center(el, active_view):
    """Force BBox center (model, then view)."""
    try:
        bbox = el.get_BoundingBox(None) or el.get_BoundingBox(active_view)
        if bbox:
            return (bbox.Min + bbox.Max) * 0.5
    except Exception:
        pass
    return None

def make_axis_through(point, direction_xyz):
    d = direction_xyz.Normalize()
    # Short line around the point defines the rotation axis
    return DB.Line.CreateBound(point - d, point + d)

def has_pinned_elements(elements):
    """Check if any elements in the list are pinned."""
    for el in elements:
        try:
            if getattr(el, "Pinned", False):
                return True
        except Exception:
            pass
    return False

# ---------------- main ----------------
def main():
    # Angle (deg) - DEFAULT CHANGED TO 90
    angle_str = forms.ask_for_string(
        default="90",
        prompt="Angle in degrees (positive = CCW)",
        title="Rotate Multiple"
    )
    if not angle_str:
        return
    
    try:
        angle_rad = math.radians(float(angle_str))
    except Exception:
        forms.alert("Angle must be a number.", exitscript=True)

    # Axis
    axis_choice = forms.CommandSwitchWindow.show(
        ["World Z (Up)", "View Normal (Active View)"],
        message="Choose rotation axis for each element:",
        default="World Z (Up)"
    )
    if not axis_choice:
        return
    axis_dir = DB.XYZ.BasisZ if axis_choice.startswith("World") else doc.ActiveView.ViewDirection

    # Center strategy
    center_pref = forms.CommandSwitchWindow.show(
        ["Auto (Location then BBox)", "Force BBox Center"],
        message="Pick rotation center strategy:",
        default="Auto (Location then BBox)"
    )
    if not center_pref:
        return
    force_bbox = center_pref.startswith("Force")

    # Selection mode
    mode = choose_mode()
    if not mode:
        return
    elements = collect_targets(mode)
    if not elements:
        forms.alert("No valid elements found to rotate.", warn_icon=True)
        return

    # Check if there are any pinned elements - only ask if there are
    unpin_choice = True  # Default: allow unpinning
    if has_pinned_elements(elements):
        unpin_choice = forms.alert(
            "Some elements are pinned. Temporarily unpin them?",
            yes=True,
            no=True,
            title="Rotate Multiple"
        )  # Returns True if Yes, False if No

    ok, skipped, failed = 0, 0, 0

    with forms.ProgressBar(title="Rotating…", cancellable=True, step=1) as pb:
        tgroup = DB.TransactionGroup(doc, "Rotate Multiple")
        tgroup.Start()
        try:
            for i, el in enumerate(elements):
                pb.update_progress(i, len(elements))
                if pb.cancelled:
                    break

                # center point
                if force_bbox:
                    pt = get_bbox_center(el, doc.ActiveView)
                else:
                    pt = get_center_point(el, doc.ActiveView)

                if not pt:
                    skipped += 1
                    continue

                axis = make_axis_through(pt, axis_dir)

                t = DB.Transaction(doc, "Rotate Element")
                t.Start()
                try:
                    # pinned handling
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

                    DB.ElementTransformUtils.RotateElement(doc, el.Id, axis, angle_rad)

                    if was_pinned:
                        try:
                            el.Pinned = True
                        except Exception:
                            pass

                    t.Commit()
                    ok += 1
                except Exception as ex:
                    log.error("Rotate failed on {}: {}".format(el.Id.IntegerValue, ex))
                    t.RollBack()
                    failed += 1
        finally:
            if ok > 0:
                tgroup.Assimilate()
            else:
                tgroup.RollBack()

    # Result dialog
    msg = "Rotated: {}\nSkipped: {}\nErrors: {}".format(ok, skipped, failed)
    forms.alert(msg, title="Rotate Multiple", warn_icon=(failed > 0))

if __name__ == "__main__":
    main()