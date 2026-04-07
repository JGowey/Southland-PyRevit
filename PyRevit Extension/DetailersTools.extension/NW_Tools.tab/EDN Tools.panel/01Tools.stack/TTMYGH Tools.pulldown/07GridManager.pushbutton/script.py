# -*- coding: utf-8 -*-
"""Grid Manager - Advanced grid manipulation tool.

Three operations in one window:
  Move Grid    -- shift grids 2D (view-specific visual offset) or 3D (model relocation)
  Bubble Vis.  -- show/hide grid bubbles at Top/Bottom/Left/Right end across selected views
  Grid Length  -- extend or shrink grid extents 2D (view-specific) or 3D (model)

Bubble visibility uses positional ends (Top/Bottom/Left/Right) rather than raw
End0/End1.  Classification is based on grid geometry with a 44-degree threshold:
grids within 44 degrees of vertical are N/S and have Top/Bottom ends;
all others are E/W and have Left/Right ends.

Author: Jeremiah Griffith
Version: 1.1.0
"""

from __future__ import print_function
from pyrevit import revit, DB, script
import os, math, traceback

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows        import Visibility, Thickness
from System.Windows.Markup import XamlReader
from System.IO             import StreamReader
import System.Windows.Controls as Ctrl
import System.Windows.Media    as Media

uidoc     = revit.uidoc
doc       = revit.doc
OUT       = script.get_output()
XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")

FEET_PER_INCH = 1.0 / 12.0

_BLUE_BG  = Media.SolidColorBrush(Media.Color.FromRgb(0x25, 0x63, 0xEB))
_GRAY_BG  = Media.SolidColorBrush(Media.Color.FromRgb(0xF0, 0xF2, 0xF7))
_WHITE_FG = Media.Brushes.White
_DARK_FG  = Media.SolidColorBrush(Media.Color.FromRgb(0x37, 0x41, 0x51))
_TEXT_FG  = Media.SolidColorBrush(Media.Color.FromRgb(0x1F, 0x29, 0x37))


# ══════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════

def _parse_inches(text):
    try:
        return float(str(text).strip()) * FEET_PER_INCH
    except (ValueError, TypeError):
        return 0.0


def _get_applicable_views(doc):
    result = []
    valid = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
             DB.ViewType.Section,   DB.ViewType.Elevation,
             DB.ViewType.Detail)
    for v in (DB.FilteredElementCollector(doc)
                .OfClass(DB.View)
                .WhereElementIsNotElementType()):
        if not v.IsTemplate and v.ViewType in valid:
            result.append(v)
    return sorted(result, key=lambda v: v.Name)


def _get_all_grids(doc):
    grids = list(DB.FilteredElementCollector(doc)
                   .OfClass(DB.Grid)
                   .WhereElementIsNotElementType())
    return sorted(grids, key=lambda g: g.Name)


def _orient_tag(grid):
    """Short orientation label: [N/S], [E/W], or [angle]."""
    try:
        s = grid.Curve.GetEndPoint(0)
        e = grid.Curve.GetEndPoint(1)
        dx, dy = abs(e.X - s.X), abs(e.Y - s.Y)
        if   dy > dx * 5: return "[N/S]"
        elif dx > dy * 5: return "[E/W]"
        else:             return "[angle]"
    except Exception:
        return ""


def _grid_end_map(grid):
    """
    Return a dict mapping positional keys to DatumEnds enum values for this grid.

    Classification uses a 44-degree threshold from each axis:
      - Within 44 deg of vertical  -> N/S grid -> keys: 'top', 'bottom'
      - Within 44 deg of horizontal -> E/W grid -> keys: 'left', 'right'

    'top'    = end with higher Y  (North in standard orientation)
    'bottom' = end with lower  Y  (South)
    'left'   = end with lower  X  (West)
    'right'  = end with higher X  (East)

    These map to End0 or End1 depending on how the grid was originally drawn --
    that's the whole point of this abstraction.
    """
    try:
        p0 = grid.Curve.GetEndPoint(0)
        p1 = grid.Curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        # tan(44°) ≈ 0.9657  -- if |dy|/|dx| > tan44, angle from horizontal > 44°
        # meaning the grid is within 44° of vertical -> N/S
        tan44 = math.tan(math.radians(44.0))
        is_ns = (abs(dx) < 1e-9) or (abs(dx) > 0 and abs(dy) / abs(dx) > tan44)

        if is_ns:
            # Top = higher Y, Bottom = lower Y
            if p0.Y >= p1.Y:
                return {'top':    DB.DatumEnds.End0,
                        'bottom': DB.DatumEnds.End1}
            else:
                return {'top':    DB.DatumEnds.End1,
                        'bottom': DB.DatumEnds.End0}
        else:
            # Left = lower X, Right = higher X
            if p0.X <= p1.X:
                return {'left':  DB.DatumEnds.End0,
                        'right': DB.DatumEnds.End1}
            else:
                return {'left':  DB.DatumEnds.End1,
                        'right': DB.DatumEnds.End0}
    except Exception:
        return {}


def _ensure_view_specific(grid, end, view):
    try:
        if grid.GetDatumExtentTypeInView(end, view) != DB.DatumExtentType.ViewSpecific:
            grid.SetDatumExtentType(end, view, DB.DatumExtentType.ViewSpecific)
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════════
#  WINDOW
# ══════════════════════════════════════════════════════════════════

class GridManagerWindow(object):

    def __init__(self, xaml_path):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        self._all_grids = _get_all_grids(doc)
        self._all_views = _get_applicable_views(doc)
        self._active_tab = "Move"

        self._get_controls()
        self._wire_events()
        self._populate_grids()
        self._populate_view_lists()
        self._sync_tab_ui()
        self._sync_move_scope()
        self._sync_length_scope()
        self._update_count()

    # ── control references ────────────────────────────────────────────────────

    def _get_controls(self):
        f = self.window.FindName

        self.btn_tab_move   = f("BtnTabMove")
        self.btn_tab_bubble = f("BtnTabBubble")
        self.btn_tab_length = f("BtnTabLength")

        self.panel_move   = f("PanelMove")
        self.panel_bubble = f("PanelBubble")
        self.panel_length = f("PanelLength")

        # grids list
        self.grid_list_panel  = f("GridListPanel")
        self.txt_grid_search  = f("TxtGridSearch")
        self.btn_grid_all     = f("BtnGridAll")
        self.btn_grid_none    = f("BtnGridNone")
        self.btn_grid_from    = f("BtnGridFromSel")
        self.lbl_grid_count   = f("LblGridCount")

        # Move panel
        self.rb_move_2d          = f("RbMove2D")
        self.rb_move_3d          = f("RbMove3D")
        self.txt_offset_x        = f("TxtOffsetX")
        self.txt_offset_y        = f("TxtOffsetY")
        self.move_scope_card     = f("MoveScopeCard")
        self.rb_move_all         = f("RbMoveScopeAll")
        self.rb_move_active      = f("RbMoveScopeActive")
        self.rb_move_pick        = f("RbMoveScopePick")
        self.move_view_list_bdr  = f("MoveViewListBorder")
        self.view_list_move      = f("ViewListMove")
        self.btn_move_view_all   = f("BtnMoveViewAll")
        self.btn_move_view_none  = f("BtnMoveViewNone")

        # Bubble panel -- 4 positional ends
        self.rb_top_show    = f("RbTopShow")
        self.rb_top_hide    = f("RbTopHide")
        self.rb_top_nc      = f("RbTopNoChange")
        self.rb_bot_show    = f("RbBotShow")
        self.rb_bot_hide    = f("RbBotHide")
        self.rb_bot_nc      = f("RbBotNoChange")
        self.rb_left_show   = f("RbLeftShow")
        self.rb_left_hide   = f("RbLeftHide")
        self.rb_left_nc     = f("RbLeftNoChange")
        self.rb_right_show  = f("RbRightShow")
        self.rb_right_hide  = f("RbRightHide")
        self.rb_right_nc    = f("RbRightNoChange")

        self.rb_bubble_all        = f("RbBubbleScopeAll")
        self.rb_bubble_active     = f("RbBubbleScopeActive")
        self.rb_bubble_pick       = f("RbBubbleScopePick")
        self.bubble_view_list_bdr = f("BubbleViewListBorder")
        self.view_list_bubble     = f("ViewListBubble")
        self.btn_bubble_view_all  = f("BtnBubbleViewAll")
        self.btn_bubble_view_none = f("BtnBubbleViewNone")

        # Length panel
        self.rb_length_2d         = f("RbLength2D")
        self.rb_length_3d         = f("RbLength3D")
        self.txt_end0_delta       = f("TxtEnd0Delta")
        self.txt_end1_delta       = f("TxtEnd1Delta")
        self.length_scope_card    = f("LengthScopeCard")
        self.rb_length_all        = f("RbLengthScopeAll")
        self.rb_length_active     = f("RbLengthScopeActive")
        self.rb_length_pick       = f("RbLengthScopePick")
        self.length_view_list_bdr = f("LengthViewListBorder")
        self.view_list_length     = f("ViewListLength")
        self.btn_length_view_all  = f("BtnLengthViewAll")
        self.btn_length_view_none = f("BtnLengthViewNone")

        # Footer
        self.txt_status = f("TxtStatus")
        self.btn_apply  = f("BtnApply")
        self.btn_close  = f("BtnClose")

    # ── event wiring ─────────────────────────────────────────────────────────

    def _wire_events(self):
        self.btn_tab_move.Click   += lambda s, e: self._set_tab("Move")
        self.btn_tab_bubble.Click += lambda s, e: self._set_tab("Bubble")
        self.btn_tab_length.Click += lambda s, e: self._set_tab("Length")

        self.txt_grid_search.TextChanged += lambda s, e: self._filter_grids(s.Text)
        self.btn_grid_all.Click   += lambda s, e: self._select_all_grids(True)
        self.btn_grid_none.Click  += lambda s, e: self._select_all_grids(False)
        self.btn_grid_from.Click  += lambda s, e: self._grids_from_selection()

        self.rb_move_2d.Checked  += lambda s, e: self._sync_move_scope()
        self.rb_move_3d.Checked  += lambda s, e: self._sync_move_scope()
        self.rb_move_all.Checked    += lambda s, e: self._show_view_list(self.move_view_list_bdr, False)
        self.rb_move_active.Checked += lambda s, e: self._show_view_list(self.move_view_list_bdr, False)
        self.rb_move_pick.Checked   += lambda s, e: self._show_view_list(self.move_view_list_bdr, True)
        self.btn_move_view_all.Click  += lambda s, e: self._check_all_views(self.view_list_move, True)
        self.btn_move_view_none.Click += lambda s, e: self._check_all_views(self.view_list_move, False)

        self.rb_bubble_all.Checked    += lambda s, e: self._show_view_list(self.bubble_view_list_bdr, False)
        self.rb_bubble_active.Checked += lambda s, e: self._show_view_list(self.bubble_view_list_bdr, False)
        self.rb_bubble_pick.Checked   += lambda s, e: self._show_view_list(self.bubble_view_list_bdr, True)
        self.btn_bubble_view_all.Click  += lambda s, e: self._check_all_views(self.view_list_bubble, True)
        self.btn_bubble_view_none.Click += lambda s, e: self._check_all_views(self.view_list_bubble, False)

        self.rb_length_2d.Checked += lambda s, e: self._sync_length_scope()
        self.rb_length_3d.Checked += lambda s, e: self._sync_length_scope()
        self.rb_length_all.Checked    += lambda s, e: self._show_view_list(self.length_view_list_bdr, False)
        self.rb_length_active.Checked += lambda s, e: self._show_view_list(self.length_view_list_bdr, False)
        self.rb_length_pick.Checked   += lambda s, e: self._show_view_list(self.length_view_list_bdr, True)
        self.btn_length_view_all.Click  += lambda s, e: self._check_all_views(self.view_list_length, True)
        self.btn_length_view_none.Click += lambda s, e: self._check_all_views(self.view_list_length, False)

        self.btn_apply.Click += lambda s, e: self._apply()
        self.btn_close.Click += lambda s, e: self.window.Close()

    # ── population ────────────────────────────────────────────────────────────

    def _make_cb(self, content, tag, checked=False):
        cb = Ctrl.CheckBox()
        cb.Content    = content
        cb.Tag        = tag
        cb.IsChecked  = checked
        cb.FontSize   = 11
        cb.Margin     = Thickness(0, 2, 0, 2)
        cb.Foreground = _TEXT_FG
        return cb

    def _populate_grids(self):
        self.grid_list_panel.Children.Clear()
        for g in self._all_grids:
            label = "{0}  {1}".format(g.Name, _orient_tag(g))
            cb = self._make_cb(label, g.Id.IntegerValue, checked=True)
            cb.Checked   += lambda s, e: self._update_count()
            cb.Unchecked += lambda s, e: self._update_count()
            self.grid_list_panel.Children.Add(cb)

    def _populate_view_lists(self):
        for panel in (self.view_list_move, self.view_list_bubble, self.view_list_length):
            panel.Children.Clear()
            for v in self._all_views:
                panel.Children.Add(self._make_cb(v.Name, v.Id.IntegerValue, checked=False))

    # ── grid list helpers ─────────────────────────────────────────────────────

    def _filter_grids(self, text):
        text = text.lower().strip() if text else ""
        for cb in self.grid_list_panel.Children:
            hit = (not text) or (text in str(cb.Content).lower())
            cb.Visibility = Visibility.Visible if hit else Visibility.Collapsed

    def _select_all_grids(self, state):
        for cb in self.grid_list_panel.Children:
            if cb.Visibility == Visibility.Visible:
                cb.IsChecked = state
        self._update_count()

    def _grids_from_selection(self):
        try:
            sel_ids = set(int(eid.IntegerValue)
                          for eid in uidoc.Selection.GetElementIds())
            for cb in self.grid_list_panel.Children:
                cb.IsChecked = int(cb.Tag) in sel_ids
            self._update_count()
            n = sum(1 for cb in self.grid_list_panel.Children if cb.IsChecked)
            self.txt_status.Text = "Matched {0} grid(s) from current selection.".format(n)
        except Exception as ex:
            self.txt_status.Text = "Could not read selection: " + str(ex)

    def _update_count(self):
        checked = sum(1 for cb in self.grid_list_panel.Children if cb.IsChecked)
        total   = self.grid_list_panel.Children.Count
        self.lbl_grid_count.Text = "{0} / {1} selected".format(checked, total)

    def _get_selected_grids(self):
        id_map = {g.Id.IntegerValue: g for g in self._all_grids}
        return [id_map[int(cb.Tag)]
                for cb in self.grid_list_panel.Children
                if cb.IsChecked and int(cb.Tag) in id_map]

    # ── tab / scope UI sync ───────────────────────────────────────────────────

    def _set_tab(self, name):
        self._active_tab = name
        self._sync_tab_ui()

    def _sync_tab_ui(self):
        pairs = [
            ("Move",   self.panel_move,   self.btn_tab_move),
            ("Bubble", self.panel_bubble, self.btn_tab_bubble),
            ("Length", self.panel_length, self.btn_tab_length),
        ]
        for name, panel, btn in pairs:
            active = (name == self._active_tab)
            panel.Visibility = Visibility.Visible if active else Visibility.Collapsed
            btn.Background   = _BLUE_BG if active else _GRAY_BG
            btn.Foreground   = _WHITE_FG if active else _DARK_FG

    def _show_view_list(self, border, show):
        border.Visibility = Visibility.Visible if show else Visibility.Collapsed

    def _check_all_views(self, panel, state):
        for cb in panel.Children:
            cb.IsChecked = state

    def _sync_move_scope(self):
        is3d = bool(self.rb_move_3d.IsChecked)
        self.move_scope_card.Visibility = Visibility.Collapsed if is3d else Visibility.Visible

    def _sync_length_scope(self):
        is3d = bool(self.rb_length_3d.IsChecked)
        self.length_scope_card.Visibility = Visibility.Collapsed if is3d else Visibility.Visible

    # ── view scope resolver ───────────────────────────────────────────────────

    def _resolve_views(self, rb_all, rb_active, rb_pick, view_list):
        if rb_all.IsChecked:
            return list(self._all_views)
        if rb_active.IsChecked:
            av    = doc.ActiveView
            valid = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
                     DB.ViewType.Section,   DB.ViewType.Elevation,
                     DB.ViewType.Detail)
            if av.ViewType in valid and not av.IsTemplate:
                return [av]
            self.txt_status.Text = "Active view is not a supported view type."
            return []
        id_map = {v.Id.IntegerValue: v for v in self._all_views}
        views  = [id_map[int(cb.Tag)]
                  for cb in view_list.Children
                  if cb.IsChecked and int(cb.Tag) in id_map]
        if not views:
            self.txt_status.Text = "No views checked in the pick list."
        return views

    # ══════════════════════════════════════════════════════
    #  OPERATIONS
    # ══════════════════════════════════════════════════════

    def _op_move_3d(self, grids, dx_in, dy_in):
        dx = _parse_inches(dx_in)
        dy = _parse_inches(dy_in)
        if dx == 0.0 and dy == 0.0:
            self.txt_status.Text = "Offset is zero — nothing to do."
            return 0
        vec = DB.XYZ(dx, dy, 0.0)
        t   = DB.Transaction(doc, "Grid Manager - Move 3D")
        t.Start()
        count = 0
        try:
            for g in grids:
                try:
                    DB.ElementTransformUtils.MoveElement(doc, g.Id, vec)
                    count += 1
                except Exception:
                    pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            self.txt_status.Text = "Transaction error: " + str(ex)
            return 0
        return count

    def _op_move_2d(self, grids, views, dx_in, dy_in):
        dx = _parse_inches(dx_in)
        dy = _parse_inches(dy_in)
        if dx == 0.0 and dy == 0.0:
            self.txt_status.Text = "Offset is zero — nothing to do."
            return 0
        t     = DB.Transaction(doc, "Grid Manager - Move 2D")
        t.Start()
        count = 0
        try:
            for view in views:
                for g in grids:
                    try:
                        _ensure_view_specific(g, DB.DatumEnds.End0, view)
                        _ensure_view_specific(g, DB.DatumEnds.End1, view)
                        curves = g.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)
                        if curves.Count == 0:
                            continue
                        crv = curves[0]
                        p0  = crv.GetEndPoint(0)
                        p1  = crv.GetEndPoint(1)
                        np0 = DB.XYZ(p0.X + dx, p0.Y + dy, p0.Z)
                        np1 = DB.XYZ(p1.X + dx, p1.Y + dy, p1.Z)
                        if (np1 - np0).GetLength() < 0.01:
                            continue
                        g.SetCurveInView(DB.DatumExtentType.ViewSpecific, view,
                                         DB.Line.CreateBound(np0, np1))
                        count += 1
                    except Exception:
                        pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            self.txt_status.Text = "Transaction error: " + str(ex)
            return 0
        return count

    def _op_bubble_visibility(self, grids, views, actions):
        """
        Set bubble visibility using positional actions.
        actions: dict with keys 'top','bottom','left','right'
                 values: 'show' | 'hide' | 'nochange'

        For each grid, _grid_end_map() resolves which DatumEnds corresponds
        to each position.  Positions that don't apply to a grid's orientation
        are silently skipped (e.g. 'top'/'bottom' for an E/W grid).
        """
        all_nc = all(v == "nochange" for v in actions.values())
        if all_nc:
            self.txt_status.Text = "All positions set to No Change — nothing to do."
            return 0

        t     = DB.Transaction(doc, "Grid Manager - Bubble Visibility")
        t.Start()
        count = 0
        try:
            for view in views:
                for g in grids:
                    try:
                        end_map = _grid_end_map(g)
                        for pos, datum_end in end_map.items():
                            action = actions.get(pos, "nochange")
                            if action == "show":
                                g.ShowBubbleInView(datum_end, view)
                            elif action == "hide":
                                g.HideBubbleInView(datum_end, view)
                        count += 1
                    except Exception:
                        pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            self.txt_status.Text = "Transaction error: " + str(ex)
            return 0
        return count

    def _op_adjust_length(self, grids, views, d0_in, d1_in, is_3d):
        d0 = _parse_inches(d0_in)
        d1 = _parse_inches(d1_in)
        if d0 == 0.0 and d1 == 0.0:
            self.txt_status.Text = "Both deltas are zero — nothing to do."
            return 0
        t     = DB.Transaction(doc, "Grid Manager - Adjust Length")
        t.Start()
        count = 0
        try:
            if is_3d:
                for g in grids:
                    try:
                        if g.Pinned:
                            g.Pinned = False
                        loc = g.Location
                        old = loc.Curve
                        p0  = old.GetEndPoint(0)
                        p1  = old.GetEndPoint(1)
                        d   = (p1 - p0).Normalize()
                        np0 = p0 - d * d0
                        np1 = p1 + d * d1
                        if (np1 - np0).GetLength() < 0.01:
                            continue
                        loc.Curve = DB.Line.CreateBound(np0, np1)
                        count += 1
                    except Exception:
                        pass
            else:
                for view in views:
                    for g in grids:
                        try:
                            _ensure_view_specific(g, DB.DatumEnds.End0, view)
                            _ensure_view_specific(g, DB.DatumEnds.End1, view)
                            curves = g.GetCurvesInView(DB.DatumExtentType.ViewSpecific, view)
                            if curves.Count == 0:
                                continue
                            crv = curves[0]
                            p0  = crv.GetEndPoint(0)
                            p1  = crv.GetEndPoint(1)
                            d   = (p1 - p0).Normalize()
                            np0 = p0 - d * d0
                            np1 = p1 + d * d1
                            if (np1 - np0).GetLength() < 0.01:
                                continue
                            g.SetCurveInView(DB.DatumExtentType.ViewSpecific, view,
                                             DB.Line.CreateBound(np0, np1))
                            count += 1
                        except Exception:
                            pass
            t.Commit()
        except Exception as ex:
            t.RollBack()
            self.txt_status.Text = "Transaction error: " + str(ex)
            return 0
        return count

    # ══════════════════════════════════════════════════════
    #  APPLY DISPATCH
    # ══════════════════════════════════════════════════════

    def _apply(self):
        grids = self._get_selected_grids()
        if not grids:
            self.txt_status.Text = "No grids selected — check at least one grid on the left."
            return

        try:
            tab = self._active_tab

            if tab == "Move":
                dx = self.txt_offset_x.Text or "0"
                dy = self.txt_offset_y.Text or "0"
                if self.rb_move_3d.IsChecked:
                    n = self._op_move_3d(grids, dx, dy)
                    if n:
                        self.txt_status.Text = "3D moved {0} grid(s).".format(n)
                else:
                    views = self._resolve_views(
                        self.rb_move_all, self.rb_move_active,
                        self.rb_move_pick, self.view_list_move)
                    if views:
                        n = self._op_move_2d(grids, views, dx, dy)
                        if n:
                            self.txt_status.Text = (
                                "2D moved {0} grid(s) across {1} view(s)."
                                .format(len(grids), len(views)))

            elif tab == "Bubble":
                def _action(rb_show, rb_hide):
                    if rb_show.IsChecked: return "show"
                    if rb_hide.IsChecked: return "hide"
                    return "nochange"

                actions = {
                    'top':    _action(self.rb_top_show,   self.rb_top_hide),
                    'bottom': _action(self.rb_bot_show,   self.rb_bot_hide),
                    'left':   _action(self.rb_left_show,  self.rb_left_hide),
                    'right':  _action(self.rb_right_show, self.rb_right_hide),
                }
                views = self._resolve_views(
                    self.rb_bubble_all, self.rb_bubble_active,
                    self.rb_bubble_pick, self.view_list_bubble)
                if views:
                    n = self._op_bubble_visibility(grids, views, actions)
                    if n:
                        self.txt_status.Text = (
                            "Bubble visibility set for {0} grid(s) in {1} view(s)."
                            .format(len(grids), len(views)))

            elif tab == "Length":
                d0 = self.txt_end0_delta.Text or "0"
                d1 = self.txt_end1_delta.Text or "0"
                if self.rb_length_3d.IsChecked:
                    n = self._op_adjust_length(grids, [], d0, d1, True)
                    if n:
                        self.txt_status.Text = "3D extents adjusted for {0} grid(s).".format(n)
                else:
                    views = self._resolve_views(
                        self.rb_length_all, self.rb_length_active,
                        self.rb_length_pick, self.view_list_length)
                    if views:
                        n = self._op_adjust_length(grids, views, d0, d1, False)
                        if n:
                            self.txt_status.Text = (
                                "2D extents adjusted for {0} grid(s) in {1} view(s)."
                                .format(len(grids), len(views)))

        except Exception:
            self.txt_status.Text = "Unexpected error — see pyRevit output."
            OUT.print_md("**Grid Manager error:**\n```\n{0}\n```".format(
                traceback.format_exc()))

    def show(self):
        self.window.ShowDialog()


# ══════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════

win = GridManagerWindow(XAML_PATH)
win.show()
