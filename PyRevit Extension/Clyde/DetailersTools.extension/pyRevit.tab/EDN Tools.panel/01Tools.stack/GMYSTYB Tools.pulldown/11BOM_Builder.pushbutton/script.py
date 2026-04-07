# -*- coding: utf-8 -*-
"""BOM Builder - Multi-category Bill of Materials with advanced filtering.

MODE A - Grab Elements:
    Collects live Revit elements from the model (or active view) by category,
    applies unlimited filter rules, then displays results in a DataGrid.

MODE B - Existing Schedule:
    Reads rows from an existing Revit schedule view (any schedule) and
    displays them in the DataGrid for further filtering and export.
    Useful when you already have a schedule set up but need cross-column
    filtering, OR / NOT logic, or a CSV that looks different from what
    Revit lets you export natively.

Supported categories (Mode A):
    Fabrication Parts (ITMs), Mechanical Equipment, Pipe/Duct Accessories,
    Pipe/Duct Fittings, Structural Framing, Structural Connection plates /
    bolts / welds, Conduit Fittings.

Author:  Jeremiah Griffith
Version: 2.0.0
"""

from __future__ import print_function
from pyrevit import revit, DB, script
from System.Collections.Generic import List
import os, sys, csv, codecs, traceback

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Data")
clr.AddReference("System.Windows.Forms")

from System.Windows          import Window, Visibility, Thickness, VerticalAlignment
from System.Windows.Markup   import XamlReader
from System.IO               import StreamReader
from System.Data             import DataTable
from System.Windows.Forms    import SaveFileDialog, DialogResult
import System.Windows.Controls as Ctrl
import System.Windows.Input    as WpfInput
import System.Windows.Media    as Media

uidoc     = revit.uidoc
doc       = revit.doc
app       = doc.Application
OUT       = script.get_output()
log       = script.get_logger()
XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")

# ================================================================
#  CATEGORY MAP  (OST_StructuralConnections does not exist;
#                 use the OST_StructConnection* family instead)
# ================================================================

CAT_MAP = {
    "Fabrication Parts (ITMs)": [
        DB.BuiltInCategory.OST_FabricationPipework,
        DB.BuiltInCategory.OST_FabricationDuctwork,
        DB.BuiltInCategory.OST_FabricationHangers,
        DB.BuiltInCategory.OST_FabricationContainment,
    ],
    "Mechanical Equipment":   [DB.BuiltInCategory.OST_MechanicalEquipment],
    "Pipe Accessories":       [DB.BuiltInCategory.OST_PipeAccessory],
    "Pipe Fittings":          [DB.BuiltInCategory.OST_PipeFitting],
    "Duct Accessories":       [DB.BuiltInCategory.OST_DuctAccessory],
    "Duct Fittings":          [DB.BuiltInCategory.OST_DuctFitting],
    "Structural Framing":     [DB.BuiltInCategory.OST_StructuralFraming],
    "Structural Connections": [
        DB.BuiltInCategory.OST_StructConnectionPlates,
        DB.BuiltInCategory.OST_StructConnectionBolts,
        DB.BuiltInCategory.OST_StructConnectionWelds,
        DB.BuiltInCategory.OST_StructConnectionShearStuds,
        DB.BuiltInCategory.OST_StructConnectionAnchors,
    ],
    "Conduit Fittings":       [DB.BuiltInCategory.OST_ConduitFitting],
}

OPERATORS = [
    "equals", "not equals", "contains", "not contains",
    "starts with", "ends with",
    ">", "<", ">=", "<=",
    "is blank", "is not blank",
]

DEFAULT_COLUMNS = ["Category", "Family Name", "Type Name", "Level"]

_SKIP_PARAMS = {
    "Edited by", "Child Elements", "Worksharing Display Settings",
    "BIM Update Date", "Model Updates Status",
    "Phase Created", "Phase Demolished",
}


# ================================================================
#  ELEMENT HELPERS  (Mode A)
# ================================================================

def _safe_cat(elem):
    try:
        return elem.Category.Name if elem.Category else "Unknown"
    except Exception:
        return "Unknown"


def _family_name(elem):
    try:
        t = doc.GetElement(elem.GetTypeId())
        if t and hasattr(t, 'FamilyName') and t.FamilyName:
            return t.FamilyName
    except Exception:
        pass
    try:
        if hasattr(elem, 'ProductName') and elem.ProductName:
            return elem.ProductName
    except Exception:
        pass
    return ""


def _type_name(elem):
    try:
        t = doc.GetElement(elem.GetTypeId())
        if t:
            return t.Name or ""
    except Exception:
        pass
    return ""


def _level_name(elem):
    try:
        if elem.LevelId and elem.LevelId != DB.ElementId.InvalidElementId:
            lvl = doc.GetElement(elem.LevelId)
            return lvl.Name if lvl else ""
    except Exception:
        pass
    try:
        p = elem.LookupParameter("Level")
        if p and p.HasValue:
            return p.AsString() or ""
    except Exception:
        pass
    return ""


def _param_value(elem, param_name):
    p = elem.LookupParameter(param_name)
    if p is None:
        try:
            te = doc.GetElement(elem.GetTypeId())
            if te:
                p = te.LookupParameter(param_name)
        except Exception:
            pass
    if p is None or not p.HasValue:
        return ""
    try:
        st = p.StorageType
        if st == DB.StorageType.String:
            return p.AsString() or ""
        if st == DB.StorageType.Double:
            vs = p.AsValueString()
            return vs if vs else str(round(p.AsDouble(), 6))
        if st == DB.StorageType.Integer:
            vs = p.AsValueString()
            return vs if vs else str(p.AsInteger())
        if st == DB.StorageType.ElementId:
            eid = p.AsElementId()
            if eid.IntegerValue < 0:
                return ""
            ref = doc.GetElement(eid)
            return ref.Name if (ref and hasattr(ref, 'Name')) else str(eid.IntegerValue)
    except Exception:
        pass
    return ""


def _display_value(elem, col_name):
    if col_name == "Category":    return _safe_cat(elem)
    if col_name == "Family Name": return _family_name(elem)
    if col_name == "Type Name":   return _type_name(elem)
    if col_name == "Level":       return _level_name(elem)
    return _param_value(elem, col_name)


# ================================================================
#  SCHEDULE READER  (Mode B)
# ================================================================

def _list_schedules():
    col = (DB.FilteredElementCollector(doc)
             .OfClass(DB.ViewSchedule)
             .ToElements())
    results = []
    for v in col:
        try:
            if v.IsTemplate:
                continue
            results.append((v.Name, v))
        except Exception:
            pass
    return sorted(results, key=lambda x: x[0].lower())


def _read_schedule(view_schedule):
    td   = view_schedule.GetTableData()
    body = td.GetSectionData(DB.SectionType.Body)
    if body is None:
        return [], []
    num_rows = body.NumberOfRows
    num_cols = body.NumberOfColumns
    if num_rows == 0 or num_cols == 0:
        return [], []

    headers = []
    for c in range(num_cols):
        try:
            headers.append(view_schedule.GetCellText(DB.SectionType.Body, 0, c))
        except Exception:
            headers.append("Col_{}".format(c))

    rows = []
    for r in range(1, num_rows):
        row = {}
        for c in range(num_cols):
            try:
                val = view_schedule.GetCellText(DB.SectionType.Body, r, c)
                row[headers[c]] = val or ""
            except Exception:
                row[headers[c]] = ""
        rows.append(row)

    return headers, rows


# ================================================================
#  FILTER ENGINE  (shared)
# ================================================================

def _eval_op(val, rule):
    op = rule.get("op", "equals").lower()
    fv = rule.get("value", "").strip()
    if op == "is blank":     return val == ""
    if op == "is not blank": return val != ""
    v = val.lower()
    f = fv.lower()
    if op == "equals":       return v == f
    if op == "not equals":   return v != f
    if op == "contains":     return f in v
    if op == "not contains": return f not in v
    if op == "starts with":  return v.startswith(f)
    if op == "ends with":    return v.endswith(f)
    try:
        vn = float(val.replace(",", ""))
        fn = float(fv.replace(",", ""))
        if op == ">":  return vn > fn
        if op == "<":  return vn < fn
        if op == ">=": return vn >= fn
        if op == "<=": return vn <= fn
    except Exception:
        pass
    return True


def _passes_filters(target, rules, use_dict=False):
    if not rules:
        return True
    def _val(t, r):
        return t.get(r.get("param", ""), "").strip() if use_dict \
               else _display_value(t, r.get("param", "")).strip()
    result = _eval_op(_val(target, rules[0]), rules[0])
    for rule in rules[1:]:
        r = _eval_op(_val(target, rule), rule)
        if rule.get("logic", "AND").upper() == "OR":
            result = result or r
        else:
            result = result and r
    return result


# ================================================================
#  ELEMENT COLLECTION  (Mode A)
# ================================================================

def _collect_elements(cat_names, active_view_only):
    bic_list = []
    for name in cat_names:
        bic_list.extend(CAT_MAP.get(name, []))
    if not bic_list:
        return []
    cat_ids = List[DB.ElementId]()
    for bic in bic_list:
        try:
            cat_ids.Add(DB.ElementId(bic))
        except Exception:
            pass
    if cat_ids.Count == 0:
        return []
    try:
        filt = DB.ElementMulticategoryFilter(cat_ids)
        col  = (DB.FilteredElementCollector(doc, uidoc.ActiveView.Id)
                if active_view_only
                else DB.FilteredElementCollector(doc))
        return list(col.WherePasses(filt).WhereElementIsNotElementType().ToElements())
    except Exception as ex:
        log.error("Collection error: {}".format(ex))
        return []


def _discover_params(elements, max_sample=200):
    names = set(DEFAULT_COLUMNS)
    for elem in list(elements)[:max_sample]:
        try:
            for p in elem.Parameters:
                n = p.Definition.Name if p.Definition else None
                if n and len(n) < 80:
                    names.add(n)
        except Exception:
            pass
        try:
            te = doc.GetElement(elem.GetTypeId())
            if te:
                for p in te.Parameters:
                    n = p.Definition.Name if p.Definition else None
                    if n and len(n) < 80:
                        names.add(n)
        except Exception:
            pass
    names -= _SKIP_PARAMS
    extras = sorted(n for n in names if n not in DEFAULT_COLUMNS)
    return DEFAULT_COLUMNS + extras


# ================================================================
#  DATA TABLE BUILDER  (shared)
# ================================================================

def _build_dt_elems(elements, columns, group_by_type):
    dt     = DataTable()
    prefix = "Count" if group_by_type else "Element ID"
    dt.Columns.Add(prefix)
    for c in columns:
        dt.Columns.Add(c)
    if group_by_type:
        groups, order = {}, []
        for elem in elements:
            key = tuple(_display_value(elem, c) for c in columns)
            if key not in groups:
                groups[key] = 0
                order.append(key)
            groups[key] += 1
        order.sort(key=lambda k: (-groups[k], k[0] if k else ""))
        for key in order:
            row = dt.NewRow()
            row["Count"] = str(groups[key])
            for c, v in zip(columns, key):
                row[c] = v
            dt.Rows.Add(row)
    else:
        for elem in elements:
            row = dt.NewRow()
            row["Element ID"] = str(elem.Id.IntegerValue)
            for c in columns:
                row[c] = _display_value(elem, c)
            dt.Rows.Add(row)
    return dt


def _build_dt_rows(rows, columns, group_by_type):
    dt = DataTable()
    if group_by_type and columns:
        dt.Columns.Add("Count")
        for c in columns:
            dt.Columns.Add(c)
        groups, order = {}, []
        for r in rows:
            key = tuple(r.get(c, "") for c in columns)
            if key not in groups:
                groups[key] = 0
                order.append(key)
            groups[key] += 1
        order.sort(key=lambda k: (-groups[k], k[0] if k else ""))
        for key in order:
            row = dt.NewRow()
            row["Count"] = str(groups[key])
            for c, v in zip(columns, key):
                row[c] = v
            dt.Rows.Add(row)
    else:
        for c in columns:
            dt.Columns.Add(c)
        for r in rows:
            row = dt.NewRow()
            for c in columns:
                row[c] = r.get(c, "")
            dt.Rows.Add(row)
    return dt


# ================================================================
#  BRUSH HELPERS
# ================================================================

def _brush(hex_color):
    c = Media.ColorConverter.ConvertFromString(hex_color)
    return Media.SolidColorBrush(c)


# ================================================================
#  WINDOW
# ================================================================

class BomWindow(object):
    def __init__(self, xaml_path):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        self._data_table        = None
        self._rule_rows         = []
        self._column_checkboxes = {}
        self._known_params      = []
        self._schedule_headers  = []
        self._schedule_rows     = []
        self._schedule_lookup   = {}
        self._mode              = "grab"

        self._get_controls()
        self._wire_events()
        self._populate_schedule_list()
        self._sync_mode_ui("grab")

    # ── controls ─────────────────────────────────────────────────

    def _get_controls(self):
        f = self.window.FindName
        self.btn_mode_grab  = f("BtnModeGrab")
        self.btn_mode_sched = f("BtnModeSched")
        self.panel_grab     = f("PanelGrab")
        self.panel_sched    = f("PanelSched")

        self._cat_cbs = {
            "Fabrication Parts (ITMs)": f("ChkFab"),
            "Mechanical Equipment":     f("ChkMech"),
            "Pipe Accessories":         f("ChkPipeAcc"),
            "Pipe Fittings":            f("ChkPipeFit"),
            "Duct Accessories":         f("ChkDuctAcc"),
            "Duct Fittings":            f("ChkDuctFit"),
            "Structural Framing":       f("ChkStructFrm"),
            "Structural Connections":   f("ChkStructConn"),
            "Conduit Fittings":         f("ChkConduitFit"),
        }

        self.chk_group     = f("ChkGroup")
        self.chk_view      = f("ChkView")
        self.cb_schedule   = f("CbSchedule")
        self.btn_load_sched = f("BtnLoadSched")
        self.chk_group_b   = f("ChkGroupB")

        self.rules_panel   = f("RulesPanel")
        self.columns_panel = f("ColumnsPanel")
        self.results_grid  = f("ResultsGrid")
        self.txt_status    = f("TxtStatus")
        self.txt_count     = f("TxtCount")
        self.btn_build     = f("BtnBuild")
        self.btn_export          = f("BtnExport")
        self.btn_export_schedule = f("BtnExportSchedule")
        self.btn_close           = f("BtnClose")
        self.btn_add_rule  = f("BtnAddRule")
        self.btn_all_cols  = f("BtnAllCols")
        self.btn_none_cols = f("BtnNoneCols")
        self.btn_def_cols  = f("BtnDefCols")
        self.lbl_no_rules  = f("LblNoRules")
        self.lbl_no_cols   = f("LblNoCols")

    def _wire_events(self):
        self.btn_mode_grab.Click  += lambda s, e: self._sync_mode_ui("grab")
        self.btn_mode_sched.Click += lambda s, e: self._sync_mode_ui("schedule")
        self.btn_build.Click      += self._on_build
        self.btn_export.Click          += self._on_export
        self.btn_export_schedule.Click += self._on_export_schedule
        self.btn_close.Click      += lambda s, e: self.window.Close()
        self.btn_add_rule.Click   += lambda s, e: self._add_rule_row()
        self.btn_all_cols.Click   += lambda s, e: self._set_all_cols(True)
        self.btn_none_cols.Click  += lambda s, e: self._set_all_cols(False)
        self.btn_def_cols.Click   += lambda s, e: self._reset_default_cols()
        self.btn_load_sched.Click += self._on_load_schedule
        self.window.KeyDown       += self._on_key

    # ── mode switching ────────────────────────────────────────────

    def _sync_mode_ui(self, mode):
        self._mode  = mode
        is_grab = (mode == "grab")

        if is_grab:
            self.btn_mode_grab.Background  = _brush("#2563EB")
            self.btn_mode_grab.Foreground  = _brush("#FFFFFF")
            self.btn_mode_sched.Background = _brush("#F0F2F7")
            self.btn_mode_sched.Foreground = _brush("#374151")
        else:
            self.btn_mode_grab.Background  = _brush("#F0F2F7")
            self.btn_mode_grab.Foreground  = _brush("#374151")
            self.btn_mode_sched.Background = _brush("#2563EB")
            self.btn_mode_sched.Foreground = _brush("#FFFFFF")

        self.panel_grab.Visibility  = Visibility.Visible   if is_grab  else Visibility.Collapsed
        self.panel_sched.Visibility = Visibility.Collapsed if is_grab  else Visibility.Visible
        self.btn_build.Content = "Build BOM" if is_grab else "Apply Filters"
        self._set_status(
            "Select categories and click Build BOM" if is_grab
            else "Load a schedule, then click Apply Filters",
            ""
        )

    # ── schedule picker ───────────────────────────────────────────

    def _populate_schedule_list(self):
        if self.cb_schedule is None:
            return
        self.cb_schedule.Items.Clear()
        self.cb_schedule.Items.Add("-- select a schedule --")
        for name, vs in _list_schedules():
            self.cb_schedule.Items.Add(name)
            self._schedule_lookup[name] = vs
        self.cb_schedule.SelectedIndex = 0

    def _on_load_schedule(self, sender, args):
        sel = str(self.cb_schedule.SelectedItem or "")
        vs  = self._schedule_lookup.get(sel)
        if vs is None:
            self._set_status("Select a schedule from the list first.", "")
            return
        self._set_status("Reading schedule...", "")
        try:
            headers, rows = _read_schedule(vs)
            if not headers:
                self._set_status("Schedule appears to be empty.", "")
                return
            self._schedule_headers = headers
            self._schedule_rows    = rows
            self._populate_filter_params(headers)
            self._populate_columns(headers, preserve_existing=False, all_checked=True)
            self._set_status(
                "Loaded: {}".format(sel),
                "{} rows  |  {} columns".format(len(rows), len(headers))
            )
        except Exception as ex:
            self._set_status("Error reading schedule: {}".format(ex), "")
            log.error(traceback.format_exc())

    # ── rule rows ─────────────────────────────────────────────────

    def _add_rule_row(self):
        first = (len(self._rule_rows) == 0)
        panel = Ctrl.StackPanel()
        panel.Orientation = Ctrl.Orientation.Horizontal
        panel.Margin = Thickness(0, 0, 0, 5)

        logic_cb = Ctrl.ComboBox()
        logic_cb.Width      = 52
        logic_cb.Margin     = Thickness(0, 0, 4, 0)
        logic_cb.FontSize   = 11
        logic_cb.Items.Add("AND")
        logic_cb.Items.Add("OR")
        logic_cb.SelectedIndex = 0
        logic_cb.Visibility = Visibility.Hidden if first else Visibility.Visible
        panel.Children.Add(logic_cb)

        param_cb = Ctrl.ComboBox()
        param_cb.Width      = 150
        param_cb.Margin     = Thickness(0, 0, 4, 0)
        param_cb.FontSize   = 11
        param_cb.IsEditable = True
        for p in self._known_params:
            param_cb.Items.Add(p)
        panel.Children.Add(param_cb)

        op_cb = Ctrl.ComboBox()
        op_cb.Width  = 100
        op_cb.Margin = Thickness(0, 0, 4, 0)
        op_cb.FontSize = 11
        for op in OPERATORS:
            op_cb.Items.Add(op)
        op_cb.SelectedIndex = 0
        panel.Children.Add(op_cb)

        val_tb = Ctrl.TextBox()
        val_tb.Width    = 96
        val_tb.Margin   = Thickness(0, 0, 4, 0)
        val_tb.FontSize = 11
        val_tb.Padding  = Thickness(4, 2, 4, 2)
        val_tb.VerticalContentAlignment = VerticalAlignment.Center
        panel.Children.Add(val_tb)

        rm_btn = Ctrl.Button()
        rm_btn.Content  = "x"
        rm_btn.Width    = 22
        rm_btn.Height   = 22
        rm_btn.FontSize = 10
        rm_btn.ToolTip  = "Remove rule"
        rm_btn.Margin   = Thickness(0)

        row_tuple = (logic_cb, param_cb, op_cb, val_tb, panel)
        rm_btn.Click += self._make_remover(row_tuple, panel)
        panel.Children.Add(rm_btn)

        self.rules_panel.Children.Add(panel)
        self._rule_rows.append(row_tuple)
        if self.lbl_no_rules:
            self.lbl_no_rules.Visibility = Visibility.Collapsed

    def _make_remover(self, row_tuple, panel):
        def _remove(s, e):
            self.rules_panel.Children.Remove(panel)
            if row_tuple in self._rule_rows:
                self._rule_rows.remove(row_tuple)
            if self._rule_rows:
                self._rule_rows[0][0].Visibility = Visibility.Hidden
            if not self._rule_rows and self.lbl_no_rules:
                self.lbl_no_rules.Visibility = Visibility.Visible
        return _remove

    # ── columns ──────────────────────────────────────────────────

    def _populate_columns(self, param_names, preserve_existing=False, all_checked=False):
        existing = set()
        if preserve_existing:
            existing = {n for n, c in self._column_checkboxes.items() if c.IsChecked}
        self.columns_panel.Children.Clear()
        self._column_checkboxes.clear()
        for name in param_names:
            chk = Ctrl.CheckBox()
            chk.Content   = name
            chk.FontSize  = 11
            chk.Margin    = Thickness(0, 0, 0, 3)
            if preserve_existing:
                chk.IsChecked = name in existing
            elif all_checked:
                chk.IsChecked = True
            else:
                chk.IsChecked = name in DEFAULT_COLUMNS
            self.columns_panel.Children.Add(chk)
            self._column_checkboxes[name] = chk
        if self.lbl_no_cols:
            self.lbl_no_cols.Visibility = Visibility.Collapsed

    def _set_all_cols(self, checked):
        for chk in self._column_checkboxes.values():
            chk.IsChecked = checked

    def _reset_default_cols(self):
        for name, chk in self._column_checkboxes.items():
            if self._mode == "schedule":
                chk.IsChecked = True
            else:
                chk.IsChecked = name in DEFAULT_COLUMNS

    def _populate_filter_params(self, param_names):
        self._known_params = list(param_names)
        for (_, param_cb, __, ___, ____) in self._rule_rows:
            cur = str(param_cb.Text) if param_cb.IsEditable else ""
            param_cb.Items.Clear()
            for p in param_names:
                param_cb.Items.Add(p)
            if cur:
                param_cb.Text = cur

    # ── helpers ──────────────────────────────────────────────────

    def _selected_cats(self):
        return [n for n, c in self._cat_cbs.items() if c and c.IsChecked]

    def _get_rules(self):
        rules = []
        for i, (logic_cb, param_cb, op_cb, val_tb, _) in enumerate(self._rule_rows):
            param = str(param_cb.Text).strip() if param_cb.IsEditable else \
                    (str(param_cb.SelectedItem) if param_cb.SelectedItem else "")
            if not param:
                continue
            op    = str(op_cb.SelectedItem) if op_cb.SelectedIndex >= 0 else "equals"
            val   = str(val_tb.Text).strip()
            logic = "AND" if i == 0 else \
                    (str(logic_cb.SelectedItem) if logic_cb.SelectedIndex >= 0 else "AND")
            rules.append({"logic": logic, "param": param, "op": op, "value": val})
        return rules

    def _selected_columns(self):
        return [n for n, c in self._column_checkboxes.items() if c.IsChecked]

    def _set_status(self, msg, count_msg):
        if self.txt_status: self.txt_status.Text = msg
        if self.txt_count:  self.txt_count.Text  = count_msg

    # ── actions ──────────────────────────────────────────────────

    def _on_build(self, sender, args):
        self.btn_build.IsEnabled           = False
        self.btn_export.IsEnabled          = False
        self.btn_export_schedule.IsEnabled = False
        try:
            if self._mode == "grab":
                self._run_grab()
            else:
                self._run_schedule()
        except Exception as ex:
            self._set_status("Error: {}".format(ex), "")
            log.error(traceback.format_exc())
        finally:
            self.btn_build.IsEnabled = True

    def _run_grab(self):
        cats = self._selected_cats()
        if not cats:
            self._set_status("Select at least one category.", "")
            return
        self._set_status("Collecting elements...", "")
        active_view = bool(self.chk_view and self.chk_view.IsChecked)
        elements    = _collect_elements(cats, active_view)
        if not elements:
            self._set_status("No elements found.", "")
            return
        self._set_status("Discovering parameters ({} elements)...".format(len(elements)), "")
        params = _discover_params(elements)
        self._populate_filter_params(params)
        first_run = not self._column_checkboxes
        self._populate_columns(params, preserve_existing=not first_run)
        rules    = self._get_rules()
        filtered = [e for e in elements if _passes_filters(e, rules, use_dict=False)]
        columns  = self._selected_columns() or list(DEFAULT_COLUMNS)
        group    = bool(self.chk_group and self.chk_group.IsChecked)
        dt = _build_dt_elems(filtered, columns, group)
        self._data_table = dt
        self.results_grid.ItemsSource = dt.DefaultView
        note = "  ({} rule{})".format(len(rules), "s" if len(rules) != 1 else "") if rules else ""
        self._set_status("BOM ready{}".format(note),
                         "{} rows  |  {} total  |  {} filtered".format(
                             dt.Rows.Count, len(elements), len(filtered)))
        self.btn_export.IsEnabled          = True
        self.btn_export_schedule.IsEnabled = True

    def _run_schedule(self):
        if not self._schedule_rows:
            self._set_status("Load a schedule first.", "")
            return
        rules    = self._get_rules()
        filtered = [r for r in self._schedule_rows if _passes_filters(r, rules, use_dict=True)]
        columns  = self._selected_columns() or list(self._schedule_headers)
        group    = bool(self.chk_group_b and self.chk_group_b.IsChecked)
        dt = _build_dt_rows(filtered, columns, group)
        self._data_table = dt
        self.results_grid.ItemsSource = dt.DefaultView
        note = "  ({} rule{})".format(len(rules), "s" if len(rules) != 1 else "") if rules else ""
        self._set_status("Filtered{}".format(note),
                         "{} rows  |  {} total rows".format(len(filtered), len(self._schedule_rows)))
        self.btn_export.IsEnabled          = True
        self.btn_export_schedule.IsEnabled = True

    def _on_export(self, sender, args):
        if self._data_table is None:
            return
        sfd = SaveFileDialog()
        sfd.Filter     = "CSV files (*.csv)|*.csv"
        sfd.DefaultExt = "csv"
        sfd.FileName   = "BOM_Export"
        sfd.Title      = "Export BOM to CSV"
        if sfd.ShowDialog() != DialogResult.OK:
            return
        try:
            headers = [col.ColumnName for col in self._data_table.Columns]
            with codecs.open(sfd.FileName, 'w', encoding='utf-8-sig') as f:
                writer = csv.writer(f, lineterminator='\n')
                writer.writerow(headers)
                for row in self._data_table.Rows:
                    writer.writerow([str(row[h]) for h in headers])
            self._set_status("Exported: {}".format(os.path.basename(sfd.FileName)), "")
        except Exception as ex:
            self._set_status("Export error: {}".format(ex), "")

    def _on_export_schedule(self, sender, args):
        """Export results to a new Revit ViewSchedule."""
        import datetime
        timestamp = datetime.datetime.now().strftime("%m-%d %H.%M")

        try:
            if self._mode == "grab":
                cats = self._selected_cats()
                if not cats:
                    self._set_status("Select at least one category first.", "")
                    return
                columns = self._selected_columns() or list(DEFAULT_COLUMNS)
                rules   = self._get_rules()
                name    = "BOM - {} - {}".format(", ".join(cats[:2]), timestamp)
                if len(cats) > 2:
                    name = "BOM - {} cats - {}".format(len(cats), timestamp)

                self.window.Hide()
                try:
                    sched = _build_revit_schedule_from_elems(cats, columns, rules, name)
                    # Open the new schedule view
                    uidoc.ActiveView = sched
                    self._set_status("Schedule created: {}".format(name), "")
                except Exception as ex:
                    self._set_status("Schedule export error: {}".format(ex), "")
                    log.error(traceback.format_exc())
                finally:
                    self.window.ShowDialog()

            else:
                # Schedule mode: duplicate source + add extra filters
                sel = str(self.cb_schedule.SelectedItem or "")
                source_vs = self._schedule_lookup.get(sel)
                if source_vs is None:
                    self._set_status("Load a schedule first.", "")
                    return
                rules = self._get_rules()
                name  = "{} (BOM Filter) {}".format(sel, timestamp)

                self.window.Hide()
                try:
                    new_vs = _duplicate_schedule_with_filters(source_vs, rules, name)
                    uidoc.ActiveView = new_vs
                    self._set_status("Schedule created: {}".format(name), "")
                except Exception as ex:
                    self._set_status("Schedule export error: {}".format(ex), "")
                    log.error(traceback.format_exc())
                finally:
                    self.window.ShowDialog()

        except Exception as ex:
            self._set_status("Error: {}".format(ex), "")
            log.error(traceback.format_exc())


    def _on_key(self, sender, args):
        if args.Key == WpfInput.Key.Escape:
            self.window.Close()
        elif args.Key == WpfInput.Key.Return:
            self._on_build(None, None)

    def show(self):
        self.window.ShowDialog()


# ================================================================
#  SCHEDULE EXPORT ENGINE
# ================================================================

# Revit ScheduleFilterType mapping from our operator strings
_SCHED_FILTER_MAP = {
    "equals":       "Equal",
    "not equals":   "NotEqual",
    "contains":     "Contains",
    "not contains": "NotContains",
    "starts with":  "BeginsWith",
    "ends with":    "EndsWith",
    "is blank":     "HasNoValue",
    "is not blank": "HasValue",
    ">":            "GreaterThan",
    "<":            "LessThan",
    ">=":           "GreaterThanOrEqual",
    "<=":           "LessThanOrEqual",
}


def _sft(op_str):
    """Return a ScheduleFilterType enum value by name, or None."""
    name = _SCHED_FILTER_MAP.get(op_str.lower())
    if name is None:
        return None
    try:
        return getattr(DB.ScheduleFilterType, name)
    except AttributeError:
        return None


def _find_schedulable_field(defn, param_name):
    """Return the first SchedulableField whose display name matches param_name."""
    try:
        for sf in defn.GetSchedulableFields():
            try:
                if sf.GetName(doc) == param_name:
                    return sf
            except Exception:
                pass
    except Exception:
        pass
    return None


def _add_schedule_filter(defn, field_id, rule):
    """Attempt to add a ScheduleFilter for a rule dict. Silently skips on failure."""
    try:
        op    = rule.get("op", "equals").lower()
        val   = rule.get("value", "")
        sft   = _sft(op)
        if sft is None:
            return
        if op in ("is blank", "is not blank"):
            sf = DB.ScheduleFilter(field_id, sft)
        else:
            # Try string filter first; fall back to numeric
            try:
                sf = DB.ScheduleFilter(field_id, sft, val)
            except Exception:
                try:
                    sf = DB.ScheduleFilter(field_id, sft, float(val))
                except Exception:
                    return
        defn.AddFilter(sf)
    except Exception:
        pass


def _build_revit_schedule_from_elems(cat_names, columns, rules, schedule_name):
    """
    Create a multi-category ViewSchedule in Revit with the selected columns
    as fields and the user filter rules applied as schedule filters.
    Returns the new ViewSchedule or raises on failure.
    """
    # Determine category IDs to restrict the schedule
    bic_list = []
    for name in cat_names:
        bic_list.extend(CAT_MAP.get(name, []))

    with revit.Transaction("BOM Builder - Create Schedule"):
        # Multi-category schedule (supports mixed categories)
        sched = DB.ViewSchedule.CreateMultiCategorySchedule(doc)
        sched.Name = schedule_name

        defn = sched.Definition

        # Restrict to selected categories
        if bic_list:
            cat_id_list = List[DB.ElementId]()
            for bic in bic_list:
                try:
                    cat_id_list.Add(DB.ElementId(bic))
                except Exception:
                    pass
            try:
                defn.SetCategoryIds(cat_id_list)
            except Exception:
                pass  # Not all Revit versions support this on multi-cat schedules

        # Add fields for each selected column (skip virtual cols that have no param)
        added_fields = {}  # param_name -> ScheduleField
        for col in columns:
            if col in ("Category", "Family Name", "Type Name", "Level", "Count", "Element ID"):
                # Try to add as built-in equivalents if available
                sf = _find_schedulable_field(defn, col)
                if sf:
                    try:
                        field = defn.AddField(sf)
                        added_fields[col] = field
                    except Exception:
                        pass
                continue
            sf = _find_schedulable_field(defn, col)
            if sf:
                try:
                    field = defn.AddField(sf)
                    added_fields[col] = field
                except Exception:
                    pass

        # Apply filter rules
        for i, rule in enumerate(rules):
            pname = rule.get("param", "")
            field = added_fields.get(pname)
            if field is None:
                # Try to find + add the field just for filtering (hidden)
                sf = _find_schedulable_field(defn, pname)
                if sf:
                    try:
                        field = defn.AddField(sf)
                        field.IsHidden = True
                        added_fields[pname] = field
                    except Exception:
                        pass
            if field:
                _add_schedule_filter(defn, field.FieldId, rule)

        return sched


def _duplicate_schedule_with_filters(source_vs, extra_rules, schedule_name):
    """
    Duplicate source_vs into a new schedule view and add extra_rules on top
    of any existing filters. Returns the new ViewSchedule.
    """
    with revit.Transaction("BOM Builder - Duplicate Schedule"):
        new_id = source_vs.Duplicate(DB.ViewDuplicateOption.Duplicate)
        new_vs = doc.GetElement(new_id)
        new_vs.Name = schedule_name

        defn = new_vs.Definition

        # Build a lookup: param name -> field in the new schedule
        field_lookup = {}
        for i in range(defn.GetFieldCount()):
            try:
                field = defn.GetField(i)
                fname = field.GetName()
                field_lookup[fname] = field
            except Exception:
                pass

        for rule in extra_rules:
            pname = rule.get("param", "")
            field = field_lookup.get(pname)
            if field is None:
                # Try to add it
                sf = _find_schedulable_field(defn, pname)
                if sf:
                    try:
                        field = defn.AddField(sf)
                        field.IsHidden = True
                        field_lookup[pname] = field
                    except Exception:
                        pass
            if field:
                _add_schedule_filter(defn, field.FieldId, rule)

        return new_vs


# ================================================================
#  MAIN
# ================================================================

def main():
    win = BomWindow(XAML_PATH)
    win.show()

main()
