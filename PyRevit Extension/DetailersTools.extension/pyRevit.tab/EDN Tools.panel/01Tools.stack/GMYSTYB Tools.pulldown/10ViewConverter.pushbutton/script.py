# -*- coding: utf-8 -*-
"""
Legend <-> Detail Converter

Rebuilds supported view-specific annotation elements directly in the
destination view.

Supported: detail curves, text notes, filled regions,
           point-based annotation symbols, dimensions
Skipped:   legend components, unrecognized internal element types
"""
import os

from pyrevit import revit, forms
from Autodesk.Revit import DB
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    ViewType,
    Transaction,
    BuiltInCategory,
    FilledRegion,
    TextNote,
    TextNoteOptions,
    FamilySymbol,
    CurveElement,
    ReferenceArray,
)

doc    = revit.doc
uidoc  = revit.uidoc

SCRIPT_DIR = os.path.dirname(__file__)
XAML_PATH  = os.path.join(SCRIPT_DIR, "window.xaml")
SKIP_BICS  = {int(BuiltInCategory.OST_LegendComponents)}


# ── view helpers ──────────────────────────────────────────────────────

def get_legends():
    return sorted(
        [v for v in FilteredElementCollector(doc).OfClass(DB.View).ToElements()
         if v.ViewType == ViewType.Legend and not v.IsTemplate],
        key=lambda v: v.Name,
    )


def get_detail_views():
    return sorted(
        [v for v in FilteredElementCollector(doc).OfClass(DB.View).ToElements()
         if v.ViewType == ViewType.DraftingView and not v.IsTemplate],
        key=lambda v: v.Name,
    )


def get_view_elements(view):
    elems, skipped = [], []
    for elem in FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements():
        if elem.Id == view.Id or isinstance(elem, DB.View):
            continue
        cat    = elem.Category
        cat_id = int(cat.Id.IntegerValue) if cat is not None else None
        if cat_id in SKIP_BICS:
            skipped.append(elem)
        else:
            elems.append(elem)
    return elems, skipped


def _unique_name(base_name):
    existing = set(v.Name for v in FilteredElementCollector(doc).OfClass(DB.View).ToElements())
    if base_name not in existing:
        return base_name
    i = 1
    while "{} ({})".format(base_name, i) in existing:
        i += 1
    return "{} ({})".format(base_name, i)


def create_detail_view(name, scale):
    for vft in FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if vft.ViewFamily == DB.ViewFamily.Drafting:
            view = DB.ViewDrafting.Create(doc, vft.Id)
            view.Name  = _unique_name(name)
            view.Scale = scale
            return view
    raise Exception("No Drafting ViewFamilyType found.")


def create_legend_view(name, scale):
    """
    View.CreateLegend is inaccessible in IronPython.
    Workaround: duplicate an existing legend and clear its content.
    """
    existing = [
        v for v in FilteredElementCollector(doc).OfClass(DB.View).ToElements()
        if v.ViewType == ViewType.Legend and not v.IsTemplate
    ]
    if not existing:
        raise Exception(
            "No existing Legend views found to use as template. "
            "Create at least one Legend view in the project first."
        )

    new_view_id = existing[0].Duplicate(DB.ViewDuplicateOption.Duplicate)
    doc.Regenerate()

    new_view = doc.GetElement(new_view_id)
    if new_view is None:
        raise Exception("Duplicate returned id={} but GetElement returned None.".format(
            new_view_id.IntegerValue))

    # Only delete recognized annotation classes.
    # class=Element with no category is Revit's internal view anchor --
    # deleting it cascades and destroys the view.
    SAFE_CLASSES = {
        "DetailLine", "DetailArc", "DetailEllipse", "DetailNurbSpline",
        "DetailCurve", "TextNote", "FilledRegion", "AnnotationSymbol", "Dimension",
    }
    to_delete = [
        e.Id for e in
        FilteredElementCollector(doc, new_view_id).WhereElementIsNotElementType().ToElements()
        if not isinstance(e, DB.View)
        and e.Id.IntegerValue != new_view_id.IntegerValue
        and e.GetType().Name in SAFE_CLASSES
    ]
    for eid in to_delete:
        try:
            doc.Delete(eid)
        except Exception:
            pass

    doc.Regenerate()

    # Re-fetch -- Python reference goes stale after doc.Delete calls
    new_view = doc.GetElement(new_view_id)
    if new_view is None:
        raise Exception("Legend view missing after element clear.")

    new_view.Name  = _unique_name(name)
    new_view.Scale = scale
    return new_view


def activate_view(view):
    try:
        uidoc.ActiveView = view
    except Exception:
        pass


def frame_result_view(view):
    try:
        uidoc.RefreshActiveView()
        for ui_view in uidoc.GetOpenUIViews():
            if ui_view.ViewId == view.Id:
                ui_view.ZoomToFit()
                uidoc.RefreshActiveView()
                break
    except Exception:
        pass


# ── rebuild helpers ───────────────────────────────────────────────────

def _xyz(pt):
    return DB.XYZ(pt.X, pt.Y, pt.Z)


def classify_rebuild_support(elem):
    if isinstance(elem, CurveElement):
        return "curve"
    if isinstance(elem, TextNote):
        return "text"
    if isinstance(elem, FilledRegion):
        return "filled_region"
    t = elem.GetType().Name
    if t == "AnnotationSymbol":
        return "annotation_symbol"
    if t == "Dimension":
        return "dimension"
    return "unsupported"


def summarize_transferability(elems):
    supported, unsupported, breakdown = 0, 0, {}
    for elem in elems:
        key = classify_rebuild_support(elem)
        breakdown[key] = breakdown.get(key, 0) + 1
        if key in ("curve", "text", "filled_region", "annotation_symbol", "dimension"):
            supported += 1
        else:
            unsupported += 1
    return supported, unsupported, breakdown


def rebuild_curve_element(elem, dest_view):
    curve = getattr(elem, "GeometryCurve", None)
    if curve is None:
        raise Exception("CurveElement missing GeometryCurve")
    return doc.Create.NewDetailCurve(dest_view, curve.Clone()).Id


def rebuild_text_note(elem, dest_view):
    text    = elem.Text if hasattr(elem, "Text") else ""
    type_id = elem.GetTypeId()
    pt = None
    try:
        pt = _xyz(elem.Coord)
    except Exception:
        pass
    if pt is None:
        loc = getattr(elem, "Location", None)
        if isinstance(loc, DB.LocationPoint):
            pt = _xyz(loc.Point)
    if pt is None:
        raise Exception("TextNote missing placement point")
    opts = TextNoteOptions(type_id)
    try:
        opts.HorizontalAlignment = elem.HorizontalAlignment
    except Exception:
        pass
    try:
        opts.Rotation = elem.Rotation
    except Exception:
        pass
    width = getattr(elem, "Width", None)
    if width and width > 0:
        new_note = TextNote.Create(doc, dest_view.Id, pt, width, text, opts)
    else:
        new_note = TextNote.Create(doc, dest_view.Id, pt, text, opts)

    # Copy leaders
    try:
        src_leaders = elem.GetLeaders()
        if src_leaders and src_leaders.Count > 0:
            for src_leader in src_leaders:
                new_leader = new_note.AddLeader(DB.TextNoteLeaderTypes.TNLT_STRAIGHT_L)
                try:
                    new_leader.End = _xyz(src_leader.End)
                except Exception:
                    pass
                try:
                    if src_leader.HasElbow:
                        new_leader.Elbow = _xyz(src_leader.Elbow)
                except Exception:
                    pass
    except Exception:
        pass

    return new_note.Id


def rebuild_filled_region(elem, dest_view):
    boundaries = elem.GetBoundaries()
    if boundaries is None or boundaries.Count == 0:
        raise Exception("FilledRegion has no boundaries")
    from System.Collections.Generic import List
    loop_list = List[DB.CurveLoop]()
    for loop in boundaries:
        loop_list.Add(loop)
    return FilledRegion.Create(doc, elem.GetTypeId(), dest_view.Id, loop_list).Id


def rebuild_annotation_symbol(elem, dest_view):
    loc = getattr(elem, "Location", None)
    if not isinstance(loc, DB.LocationPoint):
        raise Exception("AnnotationSymbol does not have point location")
    symbol = doc.GetElement(elem.GetTypeId())
    if not isinstance(symbol, FamilySymbol):
        raise Exception("AnnotationSymbol type is not a FamilySymbol")
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    new_inst = doc.Create.NewFamilyInstance(_xyz(loc.Point), symbol, dest_view)

    # Copy instance parameter values
    try:
        for src_param in elem.Parameters:
            if src_param.IsReadOnly:
                continue
            if not src_param.HasValue:
                continue
            dest_param = new_inst.get_Parameter(src_param.Definition)
            if dest_param is None or dest_param.IsReadOnly:
                continue
            try:
                storage = src_param.StorageType
                if storage == DB.StorageType.String:
                    dest_param.Set(src_param.AsString() or "")
                elif storage == DB.StorageType.Integer:
                    dest_param.Set(src_param.AsInteger())
                elif storage == DB.StorageType.Double:
                    dest_param.Set(src_param.AsDouble())
                elif storage == DB.StorageType.ElementId:
                    dest_param.Set(src_param.AsElementId())
            except Exception:
                pass
    except Exception:
        pass

    return new_inst.Id


def rebuild_dimension(elem, dest_view, id_map):
    orig_refs     = elem.References
    new_ref_array = ReferenceArray()
    mapped        = 0

    for ref in orig_refs:
        old_id = ref.ElementId
        if old_id is None or old_id == ElementId.InvalidElementId:
            continue
        new_elem_id = id_map.get(old_id.IntegerValue)
        if new_elem_id is None:
            continue
        new_elem = doc.GetElement(new_elem_id)
        if new_elem is None:
            continue
        new_ref = None
        try:
            new_ref = DB.Reference(new_elem)
        except Exception:
            pass
        if new_ref is None and isinstance(new_elem, CurveElement):
            for ep_idx in (0, 1):
                try:
                    new_ref = new_elem.GeometryCurve.GetEndPointReference(ep_idx)
                    if new_ref is not None:
                        break
                except Exception:
                    continue
        if new_ref is not None:
            try:
                new_ref_array.Append(new_ref)
                mapped += 1
            except Exception:
                pass

    if new_ref_array.Size < 2:
        raise Exception("Need >= 2 mapped references, got {}.".format(new_ref_array.Size))

    try:
        curve = elem.Curve
        dim_line = curve if isinstance(curve, DB.Line) else DB.Line.CreateUnbound(
            curve.GetEndPoint(0),
            (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize()
        )
    except Exception as ex:
        raise Exception("Could not get dimension line: {}".format(ex))

    new_dim = doc.Create.NewDimension(dest_view, dim_line, new_ref_array)
    try:
        type_id = elem.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            new_dim.ChangeTypeId(type_id)
    except Exception:
        pass
    return new_dim.Id


def rebuild_elements(source_view, dest_view, elems):
    created_ids, skipped, id_map = [], [], {}

    def sort_key(e):
        return 1 if e.GetType().Name == "Dimension" else 0

    for elem in sorted(elems, key=sort_key):
        cls_name = elem.GetType().Name
        mode     = classify_rebuild_support(elem)
        try:
            if mode == "curve":
                new_id = rebuild_curve_element(elem, dest_view)
            elif mode == "text":
                new_id = rebuild_text_note(elem, dest_view)
            elif mode == "filled_region":
                new_id = rebuild_filled_region(elem, dest_view)
            elif mode == "annotation_symbol":
                new_id = rebuild_annotation_symbol(elem, dest_view)
            elif mode == "dimension":
                new_id = rebuild_dimension(elem, dest_view, id_map)
            else:
                skipped.append((elem, "unsupported class '{}'".format(cls_name)))
                continue

            created_ids.append(new_id)
            id_map[elem.Id.IntegerValue] = new_id

        except Exception as ex:
            skipped.append((elem, str(ex)))

    return created_ids, skipped, id_map


# ── window ────────────────────────────────────────────────────────────

class ConverterWindow(forms.WPFWindow):

    def __init__(self, xaml_path):
        forms.WPFWindow.__init__(self, xaml_path)
        self._source_view = None
        self._populate_sources()

    def _is_legend_to_detail(self):
        return bool(self.DirLegendToDetail.IsChecked)

    def _populate_sources(self):
        self.SourceViewCombo.Items.Clear()
        self._source_view = None
        views = get_legends() if self._is_legend_to_detail() else get_detail_views()
        for v in views:
            self.SourceViewCombo.Items.Add(v.Name)
        self.SourceViewCombo.Tag = {v.Name: v for v in views}
        self._update_ui()

    def _update_ui(self):
        from System.Windows import Visibility
        self.SourceLabel.Text = (
            "Source Legend" if self._is_legend_to_detail() else "Source Detail View"
        )
        name     = self.SourceViewCombo.SelectedItem
        view_map = self.SourceViewCombo.Tag or {}
        self._source_view = view_map.get(name) if name else None
        self.ConvertBtn.IsEnabled = self._source_view is not None
        if self._source_view is not None:
            self._update_summary()
        else:
            self.SummaryCard.Visibility    = Visibility.Collapsed
            self.WarnCard.Visibility       = Visibility.Collapsed
            self.ResultHintCard.Visibility = Visibility.Collapsed

    def _update_summary(self):
        from System.Windows import Visibility
        view      = self._source_view
        elems, skipped_bics = get_view_elements(view)
        dest_type = "Drafting Detail" if self._is_legend_to_detail() else "Legend"
        supported, unsupported, breakdown = summarize_transferability(elems)

        lines = [
            "Source:              {}".format(view.Name),
            "Scale:               1:{}".format(view.Scale),
            "Elements found:      {}".format(len(elems)),
            "Supported rebuilds:  {}".format(supported),
            "New view type:       {}".format(dest_type),
        ]
        if skipped_bics:
            lines.append("Legend Components:   {} (excluded)".format(len(skipped_bics)))

        detail_parts = ["{}={}".format(k, breakdown[k]) for k in sorted(breakdown)]
        if detail_parts:
            lines.append("Breakdown:           {}".format(", ".join(detail_parts)))

        self.SummaryText.Text = "\n".join(lines)
        self.SummaryCard.Visibility = Visibility.Visible

        self.ResultHintText.Text = (
            "Elements are rebuilt directly in the destination view. "
            "Supported: detail curves, text notes, filled regions, "
            "annotations, and dimensions."
        )
        self.ResultHintCard.Visibility = Visibility.Visible

        warns = []
        if skipped_bics and self._is_legend_to_detail():
            warns.append("{} Legend Component(s) will be excluded.".format(len(skipped_bics)))
        if unsupported:
            warns.append("{} element(s) of unsupported type will be skipped.".format(unsupported))
        if warns:
            self.WarnText.Text = "\n\n".join(warns)
            self.WarnCard.Visibility = Visibility.Visible
        else:
            self.WarnCard.Visibility = Visibility.Collapsed

    def OnDirectionChanged(self, sender, e):
        self._populate_sources()

    def OnSourceViewChanged(self, sender, e):
        self._update_ui()

    def OnConvert(self, sender, e):
        source = self._source_view
        if source is None:
            return

        elems, skipped_bics = get_view_elements(source)
        if not elems:
            forms.alert("No transferable elements found.", title="Converter")
            return

        delete_source = bool(self.DeleteSourceChk.IsChecked)
        open_result   = bool(self.OpenResultChk.IsChecked)
        zoom_result   = bool(self.ZoomResultChk.IsChecked)

        source_view_id = source.Id
        source_name    = source.Name
        source_scale   = source.Scale
        created_dest   = None

        try:
            with Transaction(doc, "Convert Legend / Detail View") as t:
                t.Start()

                temp_name = _unique_name(source_name + "_converting")
                if self._is_legend_to_detail():
                    created_dest = create_detail_view(temp_name, source_scale)
                else:
                    created_dest = create_legend_view(temp_name, source_scale)

                rebuild_elements(source, created_dest, elems)
                doc.Regenerate()

                if delete_source:
                    doc.Delete(source_view_id)

                # Rename to exact source name
                try:
                    created_dest.Name = source_name
                except Exception:
                    try:
                        created_dest.Name = _unique_name(source_name)
                    except Exception:
                        pass

                t.Commit()

        except Exception as ex:
            forms.alert("Conversion failed.\n\n{}".format(ex), title="Converter")
            return

        # Close window first, then open result -- same flow for both directions
        self.Close()

        if open_result and created_dest is not None:
            activate_view(created_dest)
            if zoom_result:
                frame_result_view(created_dest)

        forms.toast(
            "Converted to '{}'.".format(created_dest.Name if created_dest else ""),
            title="Converter"
        )

    def OnCancel(self, sender, e):
        self.Close()


# ── entry point ───────────────────────────────────────────────────────

if not os.path.exists(XAML_PATH):
    forms.alert("window.xaml not found:\n" + XAML_PATH, title="Converter")
else:
    win = ConverterWindow(XAML_PATH)
    win.ShowDialog()
