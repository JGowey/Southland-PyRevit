# -*- coding: utf-8 -*-
"""Upper Attachment Sync - Update RFA hanger Z elevations from a sloped model line.

Selects RFA hangers (Generic Model families), picks a sloped model line,
and writes (target_Z - top_level_Z) to FAMILY_TOP_LEVEL_OFFSET_PARAM for each hanger.
"""
from __future__ import division
import os
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import SolidColorBrush, Color
from pyrevit import forms, script

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

XAML_PATH = os.path.join(os.path.dirname(__file__), 'window.xaml')

# Status constants
STATUS_READY    = "Ready"
STATUS_OUTSIDE  = "Outside Range"
STATUS_NO_LEVEL = "No Top Level"
STATUS_READONLY = "Read Only"
STATUS_SYNCED   = "Synced"

_STATUS_COLORS = {
    STATUS_READY:    Color.FromArgb(255, 76,  175, 80),   # green
    STATUS_OUTSIDE:  Color.FromArgb(255, 255, 152,  0),   # orange
    STATUS_NO_LEVEL: Color.FromArgb(255, 244,  67, 54),   # red
    STATUS_READONLY: Color.FromArgb(255, 244,  67, 54),   # red
    STATUS_SYNCED:   Color.FromArgb(255, 187, 134, 252),  # purple
}


def _brush(color):
    return SolidColorBrush(color)


# ---------------------------------------------------------------------------
# Geometry
# ---------------------------------------------------------------------------

def xy_project_onto_curve(pt, curve):
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    length_xy = (dx * dx + dy * dy) ** 0.5
    if length_xy < 1e-9:
        return 0.5, p0
    ux = dx / length_xy
    uy = dy / length_xy
    dot = (pt.X - p0.X) * ux + (pt.Y - p0.Y) * uy
    t   = dot / length_xy
    cx  = p0.X + dot * ux
    cy  = p0.Y + dot * uy
    cz  = p0.Z + t * (p1.Z - p0.Z)
    return t, DB.XYZ(cx, cy, cz)


def slope_label(curve):
    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)
    dz_in    = (p1.Z - p0.Z) * 12.0
    horiz_in = ((p1.X - p0.X) ** 2 + (p1.Y - p0.Y) ** 2) ** 0.5 * 12.0
    length_in = curve.Length * 12.0
    if horiz_in > 0.001:
        slope_str = "{:.3f}\" per 12\" ({})".format(
            abs(dz_in / horiz_in) * 12.0,
            "up" if dz_in > 0 else "down")
    else:
        slope_str = "vertical"
    return '{:.1f}" long, {}'.format(length_in, slope_str)


# ---------------------------------------------------------------------------
# Selection filters
# ---------------------------------------------------------------------------

class GenericModelFilter(Selection.ISelectionFilter):
    def AllowElement(self, elem):
        if isinstance(elem, DB.FamilyInstance):
            bic = DB.BuiltInCategory.OST_GenericModel
            return (elem.Category is not None and
                    elem.Category.Id.IntegerValue == int(bic))
        return False
    def AllowReference(self, ref, pt):
        return False


class ModelLineFilter(Selection.ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, DB.CurveElement)
    def AllowReference(self, ref, pt):
        return False


# ---------------------------------------------------------------------------
# Row data model
# ---------------------------------------------------------------------------

class HangerRow(object):
    """One row in the DataGrid representing a single hanger."""

    def __init__(self, hanger_id, curve):
        self.element_id = hanger_id
        hanger = doc.GetElement(hanger_id)

        # Family / type name
        try:
            type_id   = hanger.GetTypeId()
            elem_type = doc.GetElement(type_id)
            try:
                self.Family = elem_type.FamilyName or "?"
            except Exception:
                self.Family = "?"
            try:
                self.Type = elem_type.Name or "?"
            except Exception:
                # Fall back to SYMBOL_NAME_PARAM
                try:
                    p = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                    self.Type = p.AsString() if p else "?"
                except Exception:
                    self.Type = "?"
        except Exception:
            self.Family = "?"
            self.Type   = "?"

        self.HangerID = str(hanger_id.IntegerValue)

        # Defaults
        self.TopLevel      = "?"
        self.CurrentOffset = "?"
        self.NewOffset     = "?"
        self._status       = STATUS_NO_LEVEL
        self._new_offset_ft = None

        if hanger is None:
            self._status = STATUS_NO_LEVEL
            self._finalize()
            return

        loc = hanger.Location
        if not isinstance(loc, DB.LocationPoint):
            self._status = STATUS_NO_LEVEL
            self._finalize()
            return

        # Top Level
        p_top_lv = hanger.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)
        if p_top_lv is None:
            self._status = STATUS_NO_LEVEL
            self._finalize()
            return
        top_lv = doc.GetElement(p_top_lv.AsElementId())
        if not isinstance(top_lv, DB.Level):
            self._status = STATUS_NO_LEVEL
            self._finalize()
            return
        self.TopLevel   = top_lv.Name
        top_level_z     = top_lv.Elevation

        # Current offset
        p_offset = hanger.get_Parameter(DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
        if p_offset is None or p_offset.IsReadOnly:
            self._status = STATUS_READONLY
            self._finalize()
            return
        cur_off_ft = p_offset.AsDouble()
        self.CurrentOffset = '{:.3f}"'.format(cur_off_ft * 12.0)

        if curve is None:
            self._finalize()
            return

        # Slope projection
        pt    = loc.Point
        t_val, proj_pt = xy_project_onto_curve(pt, curve)
        if t_val < -0.01 or t_val > 1.01:
            self._status = STATUS_OUTSIDE
            self._finalize()
            return

        new_off_ft          = proj_pt.Z - top_level_z
        self._new_offset_ft = new_off_ft
        self.NewOffset      = '{:.3f}"'.format(new_off_ft * 12.0)
        self._status        = STATUS_READY
        self._finalize()

    def _finalize(self):
        c = _STATUS_COLORS.get(self._status, Color.FromArgb(255, 200, 200, 200))
        self.StatusColor = _brush(c)
        self.Status      = self._status

    def mark_synced(self):
        self._status  = STATUS_SYNCED
        self._finalize()


def build_rows(hanger_ids, line_id):
    curve = None
    if line_id is not None:
        line_elem = doc.GetElement(line_id)
        if line_elem and isinstance(line_elem, DB.CurveElement):
            curve = line_elem.GeometryCurve
    rows = ObservableCollection[object]()
    for hid in hanger_ids:
        rows.Add(HangerRow(hid, curve))
    return rows


# ---------------------------------------------------------------------------
# Failure suppressor
# ---------------------------------------------------------------------------

def _auto_click_yes():
    """
    Spin a background thread that finds the SI_Tools warning window and
    sends Alt+Y (Yes button accelerator) via keybd_event.
    """
    import threading
    import ctypes

    def _watcher():
        import time
        user32 = ctypes.windll.user32
        VK_MENU = 0x12   # Alt
        VK_Y    = 0x59
        KEYEVENTF_KEYUP = 0x0002

        for _ in range(150):   # poll up to 15 seconds
            hwnd = user32.FindWindowW(None, u"SI_Tools - Warning")
            if hwnd:
                user32.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                # Alt down, Y down, Y up, Alt up
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
# Core update
# ---------------------------------------------------------------------------

def update_rows(rows):
    """Write new Top Offset to all Ready rows. Returns (synced, skipped) counts."""
    synced  = 0
    skipped = 0
    try:
        t = DB.Transaction(doc, "Upper Attachment Sync")
        t.Start()
        for row in rows:
            if row._status != STATUS_READY or row._new_offset_ft is None:
                skipped += 1
                continue
            hanger   = doc.GetElement(row.element_id)
            p_offset = hanger.get_Parameter(
                DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
            if p_offset is None or p_offset.IsReadOnly:
                skipped += 1
                continue
            p_offset.Set(row._new_offset_ft)
            row.mark_synced()
            synced += 1
        _auto_click_yes()
        t.Commit()
    except Exception as ex:
        try:
            t.RollBack()
        except Exception:
            pass
        forms.alert("Error during update:\n{}".format(str(ex)), exitscript=True)
    return synced, skipped


# ---------------------------------------------------------------------------
# Debug report
# ---------------------------------------------------------------------------

def debug_report(hanger_ids, line_id=None):
    out   = script.get_output()
    curve = None
    out.print_md("# Upper Attachment -- Debug Report")
    if line_id is not None:
        line_elem = doc.GetElement(line_id)
        if line_elem and isinstance(line_elem, DB.CurveElement):
            curve = line_elem.GeometryCurve
            out.print_md("**Slope line:** {}".format(slope_label(curve)))

    constraint_bips = [
        ("Base Level",  DB.BuiltInParameter.FAMILY_LEVEL_PARAM),
        ("Base Offset", DB.BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM),
        ("Top Level",   DB.BuiltInParameter.FAMILY_TOP_LEVEL_PARAM),
        ("Top Offset",  DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM),
    ]

    for hid in hanger_ids:
        hanger = doc.GetElement(hid)
        if hanger is None:
            out.print_md("**Element {} not found.**".format(hid))
            continue
        type_id   = hanger.GetTypeId()
        elem_type = doc.GetElement(type_id)
        fam_name  = elem_type.FamilyName if hasattr(elem_type, 'FamilyName') else "?"
        type_name = elem_type.Name       if hasattr(elem_type, 'Name')       else "?"
        loc       = hanger.Location
        loc_z     = loc.Point.Z if isinstance(loc, DB.LocationPoint) else None

        out.print_md("---")
        out.print_md("## Element {}  |  {} : {}".format(hid, fam_name, type_name))
        if loc_z is not None:
            out.print_md("**Insertion Z:** {:.4f} ft  ({:.3f}\")".format(
                loc_z, loc_z * 12.0))

        out.print_md("### Constraint Parameters")
        crow = []
        top_level_z = None
        for label, bip in constraint_bips:
            p = hanger.get_Parameter(bip)
            if p is None:
                crow.append([label, "-- not found --", "--", "--"])
                continue
            ro = "Yes" if p.IsReadOnly else "No"
            if p.StorageType == DB.StorageType.ElementId:
                lv = doc.GetElement(p.AsElementId())
                if isinstance(lv, DB.Level):
                    val = "{} (elev {:.4f} ft)".format(lv.Name, lv.Elevation)
                    if label == "Top Level":
                        top_level_z = lv.Elevation
                else:
                    val = "None"
            elif p.StorageType == DB.StorageType.Double:
                val = "{:.4f} ft  ({:.3f}\")".format(p.AsDouble(), p.AsDouble() * 12.0)
            else:
                val = str(p.AsInteger())
            crow.append([label, val, str(p.StorageType).replace("StorageType.", ""), ro])
        out.print_table(table_data=crow, title="",
                        columns=["Parameter", "Value", "Storage Type", "Read Only"])

        if curve is not None and loc_z is not None and top_level_z is not None:
            pt    = loc.Point
            t_val, proj_pt = xy_project_onto_curve(pt, curve)
            if -0.01 <= t_val <= 1.01:
                target_z      = proj_pt.Z
                needed_offset = target_z - top_level_z
                cur_p         = hanger.get_Parameter(
                    DB.BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)
                cur_off = cur_p.AsDouble() if cur_p else 0.0
                out.print_md("### Slope Projection Preview")
                out.print_md("- t along line: {:.4f}".format(t_val))
                out.print_md("- Target Z: {:.4f} ft  ({:.3f}\")".format(
                    target_z, target_z * 12.0))
                out.print_md("- Top Level Z: {:.4f} ft".format(top_level_z))
                out.print_md("- Current Top Offset: {:.4f} ft  ({:.3f}\")".format(
                    cur_off, cur_off * 12.0))
                out.print_md("- **Needed Top Offset: {:.4f} ft  ({:.3f}\")**".format(
                    needed_offset, needed_offset * 12.0))
            else:
                out.print_md("*Outside slope line extent (t={:.3f}).*".format(t_val))

        out.print_md("### All Instance Parameters")
        prows = []
        for p in hanger.Parameters:
            try:
                name = p.Definition.Name
                ro   = "Yes" if p.IsReadOnly else "No"
                st   = str(p.StorageType).replace("StorageType.", "")
                if p.StorageType == DB.StorageType.Double:
                    val = "{:.6f} ft".format(p.AsDouble())
                elif p.StorageType == DB.StorageType.Integer:
                    val = str(p.AsInteger())
                elif p.StorageType == DB.StorageType.String:
                    val = p.AsString() or ""
                elif p.StorageType == DB.StorageType.ElementId:
                    val = str(p.AsElementId().IntegerValue)
                else:
                    val = "?"
                prows.append((name, val, st, ro))
            except Exception:
                pass
        prows.sort(key=lambda r: r[0].lower())
        out.print_table(table_data=[[r[0], r[1], r[2], r[3]] for r in prows],
                        title="",
                        columns=["Parameter Name", "Value (raw/ft)", "Storage Type", "Read Only"])

    out.print_md("---")
    out.print_md("*Double values are in Revit internal units (feet). Multiply by 12 for inches.*")


# ---------------------------------------------------------------------------
# WPF Window
# ---------------------------------------------------------------------------

class UpperAttachmentWindow(forms.WPFWindow):

    def __init__(self, hanger_ids, line_id):
        forms.WPFWindow.__init__(self, XAML_PATH)
        self._result    = None
        self.hanger_ids = list(hanger_ids)
        self.line_id    = line_id
        self._rows      = ObservableCollection[object]()
        self._refresh_ui()

    # -----------------------------------------------------------------------

    def _refresh_ui(self):
        # Slope line label
        if self.line_id is not None:
            line_elem = doc.GetElement(self.line_id)
            if line_elem and isinstance(line_elem, DB.CurveElement):
                self.LineLabel.Text = slope_label(line_elem.GeometryCurve)
            else:
                self.LineLabel.Text = "Line not found"
                self.line_id = None
        else:
            self.LineLabel.Text = "No slope line selected"

        # Build rows
        self._rows = build_rows(self.hanger_ids, self.line_id)
        self.HangerGrid.ItemsSource = self._rows

        # Status bar counts
        total   = len(self._rows)
        ready   = sum(1 for r in self._rows if r._status == STATUS_READY)
        outside = sum(1 for r in self._rows if r._status == STATUS_OUTSIDE)
        no_lv   = sum(1 for r in self._rows if r._status in (STATUS_NO_LEVEL, STATUS_READONLY))
        synced  = sum(1 for r in self._rows if r._status == STATUS_SYNCED)

        self.StatusBar.Text = (
            "{}  hanger(s)   |   {} ready   |   {} outside range   |"
            "   {} no top level   |   {} synced"
        ).format(total, ready, outside, no_lv, synced)

        self.UpdateBtn.IsEnabled = ready > 0
        self.DebugBtn.IsEnabled  = total > 0
        self.RemoveBtn.IsEnabled = False  # enabled by grid selection

    # -----------------------------------------------------------------------

    def grid_selection_changed(self, sender, e):
        self.RemoveBtn.IsEnabled = self.HangerGrid.SelectedItem is not None

    def pick_hangers_click(self, sender, e):
        self._result = 'pick_hangers'
        self.Close()

    def pick_line_click(self, sender, e):
        self._result = 'pick_line'
        self.Close()

    def remove_selected_click(self, sender, e):
        selected = self.HangerGrid.SelectedItem
        if selected is not None and selected in self._rows:
            self.hanger_ids = [hid for hid in self.hanger_ids
                               if hid.IntegerValue != int(selected.HangerID)]
            self._refresh_ui()

    def update_click(self, sender, e):
        self._result = 'update'
        self.Close()

    def debug_click(self, sender, e):
        self._result = 'debug'
        self.Close()

    def cancel_click(self, sender, e):
        self._result = 'cancel'
        self.Close()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    hanger_ids = []
    try:
        pre_sel = list(uidoc.Selection.GetElementIds())
        if pre_sel:
            bic = DB.BuiltInCategory.OST_GenericModel
            for eid in pre_sel:
                elem = doc.GetElement(eid)
                if (isinstance(elem, DB.FamilyInstance) and
                        elem.Category is not None and
                        elem.Category.Id.IntegerValue == int(bic)):
                    hanger_ids.append(eid)
    except Exception:
        pass

    line_id = None

    while True:
        win = UpperAttachmentWindow(hanger_ids, line_id)
        win.ShowDialog()

        if win._result == 'pick_hangers':
            try:
                refs = uidoc.Selection.PickObjects(
                    Selection.ObjectType.Element,
                    GenericModelFilter(),
                    "Select hangers -- Generic Model families only"
                )
                hanger_ids = [r.ElementId for r in refs]
            except Exception:
                pass

        elif win._result == 'pick_line':
            try:
                ref = uidoc.Selection.PickObject(
                    Selection.ObjectType.Element,
                    ModelLineFilter(),
                    "Pick the sloped model line"
                )
                line_id = ref.ElementId
            except Exception:
                pass

        elif win._result == 'debug':
            debug_report(win.hanger_ids, win.line_id)

        elif win._result == 'update':
            hanger_ids = win.hanger_ids
            line_id    = win.line_id
            rows       = build_rows(hanger_ids, line_id)
            synced, skipped = update_rows(rows)
            msg = "{} hanger(s) updated.".format(synced)
            if skipped:
                msg += "\n{} skipped.".format(skipped)
            forms.alert(msg, title="Upper Attachment Sync")
            break

        else:
            break


main()
