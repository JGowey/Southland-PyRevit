# -*- coding: utf-8 -*-
"""Smart Sheet Splitter
Duplicates a template sheet for each unique value of a chosen parameter,
applying a visibility filter so each sheet shows only its run.

Requires Revit 2023+ (ViewSheet.Duplicate).
"""
from pyrevit import revit, DB, forms
from System.Collections.Generic import List
import System
import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import (
    Window, WindowStartupLocation, SizeToContent,
    Thickness, HorizontalAlignment, MessageBox,
    MessageBoxButton, MessageBoxImage, VerticalAlignment,
    ResizeMode,
)
from System.Windows.Controls import (
    StackPanel, DockPanel, Border, ComboBox,
    ComboBoxItem, Button, TextBlock, Orientation, Dock,
    ListBox, ListBoxItem, CheckBox, TextBox, ScrollViewer,
    ScrollBarVisibility, SelectionMode,
)
from System.Windows.Media import (
    Brushes, SolidColorBrush, Color,
)

doc = revit.doc

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
TOOL_PREFIX = "NW-SS"
TOOL_TITLE = "Smart Sheet Splitter"
REMOVE_VIEW_TEMPLATES = True

CATEGORY_NAMES = [
    "Air Terminals", "Cable Tray Fittings", "Cable Trays",
    "Columns", "Communication Devices", "Conduit Fittings",
    "Conduits", "Data Devices", "Duct Accessories",
    "Duct Fittings", "Ducts", "Electrical Equipment",
    "Electrical Fixtures", "Flex Ducts", "Flex Pipes",
    "Generic Models", "Lighting Devices", "Lighting Fixtures",
    "Mechanical Control Devices", "Mechanical Equipment",
    "Medical Equipment", "MEP Fabrication Ductwork",
    "MEP Fabrication Hangers", "MEP Fabrication Pipework",
    "Pipe Accessories", "Pipe Fittings", "Pipes",
    "Plumbing Equipment", "Plumbing Fixtures",
    "Specialty Equipment", "Sprinklers",
    "Structural Columns", "Structural Connections",
    "Structural Foundations", "Structural Framing",
    "Structural Rebar",
]

MODE_ASSEMBLY = "Assembly Members"
MODE_VIEW     = "View Elements"

SHEET_CURRENT  = "Current Sheet"
SHEET_ASSEMBLY = "Assembly Sheets"
SHEET_ALL      = "All Sheets"

# ── Special built-in parameters that need explicit handling ──────────
# These are pseudo-parameters whose Definition.BuiltInParameter often
# reports INVALID, but they DO have valid ElementIds for filter rules.
# NOTE: Only parameters that Revit actually supports in view filter
# rules are included here.  "Category" is NOT supported by Revit's
# filter system, and "Family" / "Type" are remapped to their string
# equivalents (Family Name / Type Name) via FILTER_REMAP.
SPECIAL_BIPS = {
    "Family":           DB.BuiltInParameter.ELEM_FAMILY_PARAM,
    "Type":             DB.BuiltInParameter.ELEM_TYPE_PARAM,
    "Type Name":        DB.BuiltInParameter.ALL_MODEL_TYPE_NAME,
    "Family Name":      DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME,
}

# ── Parameters to HIDE from the picker ───────────────────────────────
# These appear on elements but cannot be used in Revit view filters.
HIDDEN_PARAMS = {
    "Category",          # ELEM_CATEGORY_PARAM - not filterable
    "Family and Type",   # No single BIP equivalent in filter rules
}

# ── Filter-compatible remaps ─────────────────────────────────────────
# Some BIPs appear on elements but Revit won't accept them in view
# filter rules.  We remap to equivalent BIPs that Revit DOES support
# in ParameterFilterElement rules.
#   "Family"  → ALL_MODEL_FAMILY_NAME  (string, widely supported)
#   "Type"    → ALL_MODEL_TYPE_NAME    (string, widely supported)
FILTER_REMAP = {
    "Family": {
        'bip':  DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME,
        'eid':  DB.ElementId(int(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)),
        'storage_override': DB.StorageType.String,
    },
    "Type": {
        'bip':  DB.BuiltInParameter.ALL_MODEL_TYPE_NAME,
        'eid':  DB.ElementId(int(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)),
        'storage_override': DB.StorageType.String,
    },
}

# ─── Help text ───────────────────────────────────────────────────────
HELP_WORKFLOW = (
    "WORKFLOW:\n"
    "1. Choose how to pull elements\n"
    "   (Assembly Members or View Elements)\n"
    "2. Choose which sheet(s) to use as templates\n"
    "3. Pick the FILTER parameter\n"
    "   (drives filter creation — one filter per value)\n"
    "4. For each template sheet, a preview dialog\n"
    "   shows every filter value and the sheet name\n"
    "   it will receive. Check the values to create.\n"
    "   If 'separate name param' is enabled you can\n"
    "   pick a different param to drive sheet naming.\n"
    "5. Tool duplicates each template sheet, assigns\n"
    "   the sheet number, and applies a view filter\n"
    "   so each sheet only shows matching elements.\n"
    "\n"
    "PREP (before running):\n"
    "- Create your template sheet with the layout\n"
    "- Place views (and schedule if using\n"
    "  Filter by Sheet) on the template\n"
    "- Make sure elements have the driving\n"
    "  parameter filled in\n"
    "\n"
    "RE-RUN:\n"
    "- The tool does not track previously created\n"
    "  sheets. Running again on the same template\n"
    "  will create additional copies. Revit will\n"
    "  auto-suffix the sheet number if it conflicts\n"
    "  (e.g. Rack-001-1-2).\n"
    "- View filters are always reused if one with\n"
    "  the same name already exists — reuse is safe\n"
    "  and keeps the filter count down.\n"
    "\n"
    "VIEW FILTERS:\n"
    "- Filters are named: {}-<Param>-<Value>\n".format(TOOL_PREFIX) +
    "- Uses 'does not equal' + hidden visibility\n"
    "- View templates are removed from duplicated\n"
    "  views so filters can be applied\n"
    "- Schedules with 'Filter by Sheet' will\n"
    "  auto-filter to visible elements\n"
    "\n"
    "KNOWN LIMITATIONS:\n"
    "- Requires Revit 2023+\n"
    "- View templates on new views are removed\n"
    "- Blank parameter values are ignored\n"
    "- 'Assembly Sheets' only with Assembly source"
)

HELP_SOURCE = (
    "Assembly Members:\n"
    "  Pulls from the AssemblyInstance.\n"
    "  Gets ALL members even if outside\n"
    "  the view crop or hidden.\n"
    "  Only works on assembly sheets.\n"
    "\n"
    "View Elements:\n"
    "  Pulls visible elements from placed views.\n"
    "  Works on ANY sheet.\n"
    "  Only finds currently visible elements."
)

HELP_SHEETS = (
    "Current Sheet:\n"
    "  Uses the active sheet you have open.\n"
    "\n"
    "Assembly Sheets:\n"
    "  Shows all sheets with assembly views.\n"
    "  Fast scan - no graphic regeneration.\n"
    "  (Only with Assembly Members mode.)\n"
    "\n"
    "All Sheets:\n"
    "  Shows every sheet in the project.\n"
    "  You pick which ones to process."
)

HELP_FILTER_PARAM = (
    "Select the parameter that DRIVES the split.\n"
    "One sheet is created per unique value, and\n"
    "one view filter is created per value.\n"
    "\n"
    "This parameter must be filterable in Revit\n"
    "view filters (only filterable params shown).\n"
    "\n"
    "Common choices:\n"
    "  - Mark  (short, reusable filters)\n"
    "  - CP_Material Code\n"
    "  - Any shared/project parameter\n"
    "\n"
    "TIP: Use a short, stable parameter here to\n"
    "keep the total filter count manageable.\n"
    "Use a separate Name Parameter (next step)\n"
    "if you want richer sheet numbers/names."
)

HELP_NAME_PARAM = (
    "Select the parameter used for SHEET NAMING.\n"
    "This is separate from the filter parameter.\n"
    "\n"
    "Choose '— Same as filter' to use the filter\n"
    "value as the sheet name (default behavior).\n"
    "\n"
    "Any parameter on the elements works here,\n"
    "including long concatenated values that would\n"
    "create too many view filters if used directly.\n"
    "\n"
    "MISALIGNMENT WARNING:\n"
    "For clean results, each filter value should\n"
    "map to exactly ONE name value. If a filter\n"
    "value maps to multiple names, the first\n"
    "(alphabetically) is used and a ⚠ badge\n"
    "appears on that row — hover it to see all.\n"
    "\n"
    "FILTER REUSE:\n"
    "Filters are always keyed on the FILTER value,\n"
    "never the name. A filter for Mark=1 will be\n"
    "reused across runs regardless of what sheet\n"
    "name it produces."
)

HELP_VALUES = (
    "This dialog shows every unique filter value\n"
    "found on the elements of this template sheet.\n"
    "\n"
    "Check the values you want to create sheets for.\n"
    "Uncheck any you want to skip.\n"
    "\n"
    "NOTE: The tool does not track previously\n"
    "created sheets. If you run again with the\n"
    "same values, additional sheets will be created.\n"
    "Revit will auto-suffix the number if needed.\n"
    "\n"
    "FILTER REUSE:\n"
    "View filters are always reused if a matching\n"
    "filter already exists in the model — this is\n"
    "safe and keeps the filter count manageable.\n"
    "\n"
    "NAME COLUMN (when enabled):\n"
    "Shows what sheet name/number each filter value\n"
    "will produce. Change the name param at the top\n"
    "to update all rows live.\n"
    "  ⚠ badge = multiple name values found for\n"
    "  this filter value — hover for the full list.\n"
    "  First alphabetically will be used."
)

HELP_SHEET_PICK = (
    "Pick template sheets to duplicate.\n"
    "Each selected sheet is duplicated once\n"
    "per value you choose in the next step.\n"
    "\n"
    "The sheet's layout, views, schedules,\n"
    "and detailing are all copied."
)

HELP_NAMING = (
    "Controls how the new sheet number is set.\n"
    "\n"
    "Unchecked (default - append):\n"
    "  Sheet number = <template>-<value>\n"
    "  e.g. A101 + 'Zone1' → A101-Zone1\n"
    "\n"
    "Checked (replace):\n"
    "  Sheet number = <value> only\n"
    "  e.g. A101 + 'Zone1' → Zone1\n"
    "  Use this when the value already\n"
    "  contains the full desired number.\n"
    "\n"
    "In both modes, duplicate numbers are\n"
    "made unique automatically.\n"
    "The sheet NAME is never changed."
)


# ─────────────────────────────────────────────
# ACCENT / COLORS
# ─────────────────────────────────────────────
ACCENT = SolidColorBrush(Color.FromRgb(0, 122, 204))
BG_COLOR = SolidColorBrush(Color.FromRgb(245, 245, 245))


# ─────────────────────────────────────────────
# WPF HELPERS
# ─────────────────────────────────────────────
def _show_help(help_text, title="Help"):
    MessageBox.Show(str(help_text), str(title),
                    MessageBoxButton.OK, MessageBoxImage.Information)


def _make_section(title, help_text, content):
    border = Border()
    border.BorderBrush = SolidColorBrush(Color.FromRgb(200, 200, 200))
    border.BorderThickness = Thickness(1)
    border.CornerRadius = System.Windows.CornerRadius(4)
    border.Padding = Thickness(10)
    border.Margin = Thickness(0, 0, 0, 10)
    border.Background = Brushes.White

    panel = StackPanel()
    header = DockPanel()
    header.Margin = Thickness(0, 0, 0, 8)

    title_tb = TextBlock()
    title_tb.Text = title
    title_tb.FontSize = 12
    title_tb.FontWeight = System.Windows.FontWeights.Bold
    title_tb.Foreground = ACCENT
    title_tb.VerticalAlignment = VerticalAlignment.Center
    DockPanel.SetDock(title_tb, Dock.Left)
    header.Children.Add(title_tb)

    help_btn = Button()
    help_btn.Content = "?"
    help_btn.Width = 22
    help_btn.Height = 22
    help_btn.HorizontalAlignment = HorizontalAlignment.Right
    help_btn.FontWeight = System.Windows.FontWeights.Bold
    help_btn.ToolTip = "Click for help"
    help_btn.Tag = help_text
    help_btn.Click += lambda s, e: _show_help(s.Tag, title)
    header.Children.Add(help_btn)

    panel.Children.Add(header)
    panel.Children.Add(content)
    border.Child = panel
    return border


def _make_ok_cancel(ok_text, on_ok, on_cancel):
    panel = StackPanel()
    panel.Orientation = Orientation.Horizontal
    panel.HorizontalAlignment = HorizontalAlignment.Right
    panel.Margin = Thickness(0, 12, 0, 0)

    ok_btn = Button()
    ok_btn.Content = ok_text
    ok_btn.Width = 100
    ok_btn.Height = 30
    ok_btn.Margin = Thickness(0, 0, 8, 0)
    ok_btn.Background = ACCENT
    ok_btn.Foreground = Brushes.White
    ok_btn.FontWeight = System.Windows.FontWeights.Bold
    ok_btn.BorderBrush = ACCENT
    ok_btn.Click += on_ok
    ok_btn.IsDefault = True
    panel.Children.Add(ok_btn)

    cancel_btn = Button()
    cancel_btn.Content = "Cancel"
    cancel_btn.Width = 90
    cancel_btn.Height = 30
    cancel_btn.Click += on_cancel
    cancel_btn.IsCancel = True
    panel.Children.Add(cancel_btn)
    return panel


# ─────────────────────────────────────────────
# SETTINGS DIALOG
# ─────────────────────────────────────────────
class SettingsDialog(Window):
    def __init__(self):
        self.Title = TOOL_TITLE
        self.Width = 420
        self.SizeToContent = SizeToContent.Height
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize
        self.Background = BG_COLOR
        self.result = None

        root = StackPanel()
        root.Margin = Thickness(12)

        # Title + main help
        title_row = DockPanel()
        title_row.Margin = Thickness(0, 0, 0, 12)
        title_lbl = TextBlock()
        title_lbl.Text = TOOL_TITLE
        title_lbl.FontSize = 15
        title_lbl.FontWeight = System.Windows.FontWeights.Bold
        title_lbl.Foreground = ACCENT
        title_lbl.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(title_lbl, Dock.Left)
        title_row.Children.Add(title_lbl)

        main_help = Button()
        main_help.Content = "?"
        main_help.Width = 26
        main_help.Height = 26
        main_help.HorizontalAlignment = HorizontalAlignment.Right
        main_help.FontWeight = System.Windows.FontWeights.Bold
        main_help.FontSize = 14
        main_help.ToolTip = "How this tool works"
        main_help.Click += lambda s, e: _show_help(HELP_WORKFLOW, TOOL_TITLE)
        title_row.Children.Add(main_help)
        root.Children.Add(title_row)

        # Source combo
        self._source_combo = ComboBox()
        self._source_combo.Height = 28
        for item in [MODE_ASSEMBLY, MODE_VIEW]:
            ci = ComboBoxItem()
            ci.Content = item
            self._source_combo.Items.Add(ci)
        self._source_combo.SelectedIndex = 0
        self._source_combo.SelectionChanged += self._on_source_changed
        root.Children.Add(
            _make_section("Element Source", HELP_SOURCE, self._source_combo))

        # Sheet combo
        self._sheet_combo = ComboBox()
        self._sheet_combo.Height = 28
        self._update_sheet_options()
        root.Children.Add(
            _make_section("Sheet Selection", HELP_SHEETS, self._sheet_combo))

        # Naming mode checkbox
        self._replace_number_cb = CheckBox()
        self._replace_number_cb.Content = (
            "Replace sheet number with value\n"
            "(default: append value with hyphen)")
        self._replace_number_cb.Margin = Thickness(2, 2, 2, 2)
        self._replace_number_cb.IsChecked = False
        root.Children.Add(
            _make_section("Sheet Number Mode", HELP_NAMING, self._replace_number_cb))

        # Separate name parameter checkbox
        self._sep_name_cb = CheckBox()
        self._sep_name_cb.Content = (
            "Use a separate parameter for sheet naming\n"
            "(filter on one param, name sheets from another)")
        self._sep_name_cb.Margin = Thickness(2, 2, 2, 2)
        self._sep_name_cb.IsChecked = False
        root.Children.Add(
            _make_section("Sheet Naming Parameter", HELP_NAME_PARAM, self._sep_name_cb))

        root.Children.Add(
            _make_ok_cancel("Next", self._on_ok, self._on_cancel))
        self.Content = root

    def _update_sheet_options(self):
        self._sheet_combo.Items.Clear()
        source = MODE_ASSEMBLY
        sel = self._source_combo.SelectedItem
        if sel:
            source = sel.Content
        options = [SHEET_CURRENT]
        if source == MODE_ASSEMBLY:
            options.append(SHEET_ASSEMBLY)
        options.append(SHEET_ALL)
        for item in options:
            ci = ComboBoxItem()
            ci.Content = item
            self._sheet_combo.Items.Add(ci)
        self._sheet_combo.SelectedIndex = 0

    def _on_source_changed(self, sender, args):
        self._update_sheet_options()

    def _on_ok(self, sender, args):
        src = self._source_combo.SelectedItem
        sht = self._sheet_combo.SelectedItem
        self.result = {
            'source': src.Content if src else MODE_ASSEMBLY,
            'sheet_mode': sht.Content if sht else SHEET_CURRENT,
            'replace_number': bool(self._replace_number_cb.IsChecked),
            'allow_name_param': bool(self._sep_name_cb.IsChecked),
        }
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ─────────────────────────────────────────────
# SINGLE-SELECT PICKER (parameters)
# ─────────────────────────────────────────────
class SinglePickerDialog(Window):
    def __init__(self, items, title, help_text, button_text="Select"):
        self.Title = title
        self.Width = 460
        self.Height = 580
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.Background = BG_COLOR
        self.result = None
        self._all_items = list(items)

        root = StackPanel()
        root.Margin = Thickness(12)

        # Title + help
        title_row = DockPanel()
        title_row.Margin = Thickness(0, 0, 0, 8)
        title_tb = TextBlock()
        title_tb.Text = title
        title_tb.FontSize = 12
        title_tb.FontWeight = System.Windows.FontWeights.Bold
        title_tb.Foreground = ACCENT
        title_tb.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(title_tb, Dock.Left)
        title_row.Children.Add(title_tb)
        help_btn = Button()
        help_btn.Content = "?"
        help_btn.Width = 22
        help_btn.Height = 22
        help_btn.HorizontalAlignment = HorizontalAlignment.Right
        help_btn.FontWeight = System.Windows.FontWeights.Bold
        help_btn.Click += lambda s, e: _show_help(help_text, title)
        title_row.Children.Add(help_btn)
        root.Children.Add(title_row)

        # Search
        self._search = TextBox()
        self._search.Height = 26
        self._search.Margin = Thickness(0, 0, 0, 6)
        self._search.TextChanged += self._on_search
        root.Children.Add(self._search)

        # List
        self._listbox = ListBox()
        self._listbox.Height = 400
        self._listbox.SelectionMode = SelectionMode.Single
        self._listbox.MouseDoubleClick += self._on_double_click
        self._populate()
        root.Children.Add(self._listbox)

        root.Children.Add(
            _make_ok_cancel(button_text, self._on_ok, self._on_cancel))
        self.Content = root

    def _populate(self, filter_text=""):
        self._listbox.Items.Clear()
        ft = filter_text.lower()
        for item in self._all_items:
            if ft and ft not in item.lower():
                continue
            lbi = ListBoxItem()
            lbi.Content = item
            self._listbox.Items.Add(lbi)

    def _on_search(self, sender, args):
        self._populate(self._search.Text)

    def _on_double_click(self, sender, args):
        self._on_ok(None, None)

    def _on_ok(self, sender, args):
        sel = self._listbox.SelectedItem
        if sel:
            self.result = sel.Content
            self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ─────────────────────────────────────────────
# MULTI-SELECT PICKER (sheets / values)
# ─────────────────────────────────────────────
class MultiPickerDialog(Window):
    def __init__(self, items, title, help_text, button_text="Select"):
        self.Title = title
        self.Width = 500
        self.Height = 600
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.Background = BG_COLOR
        self.result = None
        self._all_items = list(items)
        self._checkboxes = []
        self._checked_set = set()

        root = StackPanel()
        root.Margin = Thickness(12)

        # Title + help
        title_row = DockPanel()
        title_row.Margin = Thickness(0, 0, 0, 8)
        title_tb = TextBlock()
        title_tb.Text = title
        title_tb.FontSize = 12
        title_tb.FontWeight = System.Windows.FontWeights.Bold
        title_tb.Foreground = ACCENT
        title_tb.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(title_tb, Dock.Left)
        title_row.Children.Add(title_tb)
        help_btn = Button()
        help_btn.Content = "?"
        help_btn.Width = 22
        help_btn.Height = 22
        help_btn.HorizontalAlignment = HorizontalAlignment.Right
        help_btn.FontWeight = System.Windows.FontWeights.Bold
        help_btn.Click += lambda s, e: _show_help(help_text, title)
        title_row.Children.Add(help_btn)
        root.Children.Add(title_row)

        # Search
        self._search = TextBox()
        self._search.Height = 26
        self._search.Margin = Thickness(0, 0, 0, 6)
        self._search.TextChanged += self._on_search
        root.Children.Add(self._search)

        # Scrollable checkbox list
        self._list_panel = StackPanel()
        scroll = ScrollViewer()
        scroll.Height = 380
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Content = self._list_panel
        self._populate()
        root.Children.Add(scroll)

        # Check All / Uncheck All
        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.Margin = Thickness(0, 8, 0, 0)
        for label, handler in [
            ("Check All", self._check_all),
            ("Uncheck All", self._uncheck_all),
        ]:
            b = Button()
            b.Content = label
            b.Width = 90
            b.Height = 26
            b.Margin = Thickness(0, 0, 8, 0)
            b.Click += handler
            btn_row.Children.Add(b)
        root.Children.Add(btn_row)

        root.Children.Add(
            _make_ok_cancel(button_text, self._on_ok, self._on_cancel))
        self.Content = root

    def _populate(self, filter_text=""):
        # Save checked state
        for cb in self._checkboxes:
            if cb.IsChecked:
                self._checked_set.add(cb.Tag)
            elif cb.Tag in self._checked_set:
                self._checked_set.discard(cb.Tag)

        self._list_panel.Children.Clear()
        self._checkboxes = []
        ft = filter_text.lower()
        for item in self._all_items:
            if ft and ft not in item.lower():
                continue
            cb = CheckBox()
            cb.Content = item
            cb.Margin = Thickness(4, 2, 0, 2)
            cb.Tag = item
            if item in self._checked_set:
                cb.IsChecked = True
            self._list_panel.Children.Add(cb)
            self._checkboxes.append(cb)

    def _on_search(self, sender, args):
        self._populate(self._search.Text)

    def _check_all(self, sender, args):
        for cb in self._checkboxes:
            cb.IsChecked = True

    def _uncheck_all(self, sender, args):
        for cb in self._checkboxes:
            cb.IsChecked = False

    def _on_ok(self, sender, args):
        # Collect from current visible + previously checked hidden items
        for cb in self._checkboxes:
            if cb.IsChecked:
                self._checked_set.add(cb.Tag)
            elif cb.Tag in self._checked_set:
                self._checked_set.discard(cb.Tag)
        self.result = [item for item in self._all_items
                       if item in self._checked_set]
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ─────────────────────────────────────────────
# COMBINED FILTER + NAME PICKER DIALOG
# ─────────────────────────────────────────────
class FilterNamePickerDialog(Window):
    """Single dialog that lets the user:
      1. Choose a name parameter (or reuse filter param) via a ComboBox
      2. See a live two-column preview: filter value → resolved sheet name
      3. Check/uncheck which values to create
    Misalignment is shown inline per row — no separate warning dialog.
    """
    _SAME_AS_FILTER_PREFIX = u"\u2014 Same as filter"

    def __init__(self, filter_values, elems, filter_param_info,
                 all_name_params, filter_param_name,
                 sheet_label, allow_name_param=True):
        self.Title = u"{} — Pick Values".format(sheet_label)
        self.Width = 660 if allow_name_param else 480
        self.Height = 700 if allow_name_param else 560
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip
        self.Background = BG_COLOR
        self.result = None

        # ── Instance state ────────────────────────────────────────────────
        self._filter_values = filter_values
        self._elems = elems
        self._filter_param_info = filter_param_info
        self._all_name_params = all_name_params
        self._filter_param_name = filter_param_name
        self._allow_name_param = allow_name_param
        self._checkboxes = []       # list of (CheckBox, filter_val)
        self._f2n = {}              # filter_val → set(name_vals)
        self._first_build = True

        # ── Root layout: DockPanel guarantees buttons stay visible ────────
        # DockPanel rule: Bottom/Top items added first; LastChildFill=True
        # means the scroll area gets all remaining space in the middle.
        root = DockPanel()
        root.Margin = Thickness(12)
        root.LastChildFill = True

        # ── TOP: title row ────────────────────────────────────────────────
        title_row = DockPanel()
        title_row.Margin = Thickness(0, 0, 0, 8)
        title_tb = TextBlock()
        title_tb.Text = sheet_label
        title_tb.FontSize = 13
        title_tb.FontWeight = System.Windows.FontWeights.Bold
        title_tb.Foreground = ACCENT
        title_tb.VerticalAlignment = VerticalAlignment.Center
        DockPanel.SetDock(title_tb, Dock.Left)
        title_row.Children.Add(title_tb)
        # Help button shows name-param help when feature is on,
        # general values help when it's off
        help_topic = HELP_NAME_PARAM if allow_name_param else HELP_VALUES
        help_title = "Name Parameter" if allow_name_param else "Pick Values"
        help_btn = Button()
        help_btn.Content = "?"
        help_btn.Width = 22
        help_btn.Height = 22
        help_btn.HorizontalAlignment = HorizontalAlignment.Right
        help_btn.FontWeight = System.Windows.FontWeights.Bold
        help_btn.ToolTip = "Click for help"
        help_btn.Click += lambda s, e: _show_help(help_topic, help_title)
        title_row.Children.Add(help_btn)
        DockPanel.SetDock(title_row, Dock.Top)
        root.Children.Add(title_row)

        # ── TOP: name param selector (only when feature is enabled) ───────
        if allow_name_param:
            name_lbl = TextBlock()
            name_lbl.Text = u"Sheet name param:"
            name_lbl.FontWeight = System.Windows.FontWeights.Bold
            name_lbl.Margin = Thickness(0, 0, 0, 3)
            DockPanel.SetDock(name_lbl, Dock.Top)
            root.Children.Add(name_lbl)

            self._name_search = TextBox()
            self._name_search.Height = 24
            self._name_search.Margin = Thickness(0, 0, 0, 2)
            self._name_search.TextChanged += self._on_name_search
            DockPanel.SetDock(self._name_search, Dock.Top)
            root.Children.Add(self._name_search)

            self._name_list = ListBox()
            self._name_list.Height = 90
            self._name_list.Margin = Thickness(0, 0, 0, 8)
            self._name_list.SelectionMode = SelectionMode.Single
            # SelectionChanged wired AFTER full UI build to avoid firing
            # _rebuild_rows before _hdr_name_tb / _rows_panel are assigned
            self._name_all_items = []
            same_label = u"{} ({})".format(
                self._SAME_AS_FILTER_PREFIX, filter_param_name)
            self._name_same_label = same_label
            self._name_all_items.append(same_label)
            for n in sorted(all_name_params.keys()):
                self._name_all_items.append(n)
            self._name_combo = None     # unused; kept for guard consistency
            self._populate_name_list("")
            self._name_list.SelectedIndex = 0
            DockPanel.SetDock(self._name_list, Dock.Top)
            root.Children.Add(self._name_list)
        else:
            self._name_combo = None
            self._name_list = None
            self._name_search = None
            self._name_all_items = []
            self._name_same_label = u""

        # ── TOP: column header ────────────────────────────────────────────
        self._hdr_border = Border()
        self._hdr_border.Background = SolidColorBrush(Color.FromRgb(225, 225, 225))
        self._hdr_border.Padding = Thickness(4, 3, 4, 3)
        self._hdr_border.Margin = Thickness(0, 0, 0, 2)
        hdr = DockPanel()
        hdr.LastChildFill = True
        hdr_chk = TextBlock()
        hdr_chk.Width = 26
        hdr_chk.Text = ""
        DockPanel.SetDock(hdr_chk, Dock.Left)
        hdr.Children.Add(hdr_chk)
        self._hdr_filter_tb = TextBlock()
        self._hdr_filter_tb.Width = 160
        self._hdr_filter_tb.Text = u"Filter  ({})".format(filter_param_name)
        self._hdr_filter_tb.FontWeight = System.Windows.FontWeights.Bold
        self._hdr_filter_tb.FontSize = 10
        DockPanel.SetDock(self._hdr_filter_tb, Dock.Left)
        hdr.Children.Add(self._hdr_filter_tb)
        if allow_name_param:
            hdr_arrow = TextBlock()
            hdr_arrow.Width = 28
            hdr_arrow.Text = u"\u2192"
            hdr_arrow.Foreground = Brushes.Gray
            hdr_arrow.HorizontalAlignment = HorizontalAlignment.Center
            DockPanel.SetDock(hdr_arrow, Dock.Left)
            hdr.Children.Add(hdr_arrow)
            self._hdr_name_tb = TextBlock()
            self._hdr_name_tb.Text = u"Sheet Name  (same as filter)"
            self._hdr_name_tb.FontWeight = System.Windows.FontWeights.Bold
            self._hdr_name_tb.FontSize = 10
            hdr.Children.Add(self._hdr_name_tb)
        else:
            self._hdr_name_tb = None
        self._hdr_border.Child = hdr
        DockPanel.SetDock(self._hdr_border, Dock.Top)
        root.Children.Add(self._hdr_border)

        # ── BOTTOM: OK / Cancel ───────────────────────────────────────────
        # Must be added before the fill child (scroll) so DockPanel pins it
        ok_cancel = _make_ok_cancel(
            "Create Selected", self._on_ok, self._on_cancel)
        ok_cancel.Margin = Thickness(0, 8, 0, 0)
        DockPanel.SetDock(ok_cancel, Dock.Bottom)
        root.Children.Add(ok_cancel)

        # ── BOTTOM: summary / warning line ───────────────────────────────
        self._summary_tb = TextBlock()
        self._summary_tb.Margin = Thickness(0, 4, 0, 0)
        self._summary_tb.TextWrapping = System.Windows.TextWrapping.Wrap
        DockPanel.SetDock(self._summary_tb, Dock.Bottom)
        root.Children.Add(self._summary_tb)

        # ── BOTTOM: Check all / Uncheck all ──────────────────────────────
        btn_row = StackPanel()
        btn_row.Orientation = Orientation.Horizontal
        btn_row.Margin = Thickness(0, 6, 0, 0)
        for lbl, hdlr in [("Check All", self._check_all),
                           ("Uncheck All", self._uncheck_all)]:
            b = Button()
            b.Content = lbl
            b.Width = 90
            b.Height = 26
            b.Margin = Thickness(0, 0, 8, 0)
            b.Click += hdlr
            btn_row.Children.Add(b)
        DockPanel.SetDock(btn_row, Dock.Bottom)
        root.Children.Add(btn_row)

        # ── FILL: scrollable rows (added last — fills remaining space) ────
        self._rows_panel = StackPanel()
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Content = self._rows_panel
        # No fixed height — DockPanel gives this all remaining space
        root.Children.Add(scroll)

        self.Content = root

        # Wire name list event and do first _rebuild_rows AFTER all
        # UI attributes exist so there are no AttributeError surprises
        if allow_name_param and self._name_list is not None:
            self._name_list.SelectionChanged += self._on_name_changed
        self._rebuild_rows()

    # ── Helpers ───────────────────────────────────────────────────────────
    def _populate_name_list(self, filter_text=""):
        self._name_list.Items.Clear()
        ft = filter_text.lower()
        for item in self._name_all_items:
            if ft and ft not in item.lower():
                continue
            lbi = ListBoxItem()
            lbi.Content = item
            self._name_list.Items.Add(lbi)

    def _is_same_as_filter(self):
        if self._name_list is None:
            return True
        sel = self._name_list.SelectedItem
        if not sel:
            return True
        return str(sel.Content).startswith(self._SAME_AS_FILTER_PREFIX)

    def _get_name_param_info(self):
        if self._name_list is None or self._is_same_as_filter():
            return None, None
        sel = self._name_list.SelectedItem
        if not sel:
            return None, None
        name = str(sel.Content)
        return name, self._all_name_params.get(name)

    def _on_name_search(self, sender, args):
        current = self._name_list.SelectedItem
        current_val = str(current.Content) if current else None
        self._populate_name_list(self._name_search.Text)
        # Re-select previously selected item if still visible
        if current_val:
            for i in range(self._name_list.Items.Count):
                lbi = self._name_list.Items[i]
                if str(lbi.Content) == current_val:
                    self._name_list.SelectedIndex = i
                    break

    def _rebuild_rows(self):
        name_param_name, name_param_info = self._get_name_param_info()
        use_separate = name_param_info is not None

        # Update column header label to reflect selected name param
        if self._hdr_name_tb is not None:
            if use_separate:
                self._hdr_name_tb.Text = u"Sheet Name  ({})".format(name_param_name)
            else:
                self._hdr_name_tb.Text = u"Sheet Name  (same as filter)"

        # Build filter→name map if a separate name param is chosen
        if use_separate:
            self._f2n = build_filter_to_name_map(
                self._elems, self._filter_param_info, name_param_info)
        else:
            self._f2n = {}

        # Preserve checked state across live rebuilds (name param changes)
        checked = set()
        if not self._first_build:
            for cb, fval in self._checkboxes:
                if cb.IsChecked:
                    checked.add(fval)

        self._rows_panel.Children.Clear()
        self._checkboxes = []
        warn_count = 0
        no_name_count = 0

        for fval in self._filter_values:
            row = DockPanel()
            row.Margin = Thickness(2, 3, 2, 3)
            row.LastChildFill = True

            # Checkbox — all rows enabled, checked by default on first build
            cb = CheckBox()
            cb.IsChecked = True if self._first_build else (fval in checked)
            cb.Width = 26
            cb.VerticalAlignment = VerticalAlignment.Center
            DockPanel.SetDock(cb, Dock.Left)
            row.Children.Add(cb)

            # Filter value label
            fval_tb = TextBlock()
            fval_tb.Width = 160
            fval_tb.Text = str(fval)
            fval_tb.VerticalAlignment = VerticalAlignment.Center
            fval_tb.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            fval_tb.ToolTip = str(fval)
            DockPanel.SetDock(fval_tb, Dock.Left)
            row.Children.Add(fval_tb)

            # Arrow + name cell — only shown when separate name param is enabled
            if self._allow_name_param:
                arrow_tb = TextBlock()
                arrow_tb.Text = u"  \u2192  "
                arrow_tb.Width = 28
                arrow_tb.VerticalAlignment = VerticalAlignment.Center
                arrow_tb.Foreground = Brushes.Gray
                DockPanel.SetDock(arrow_tb, Dock.Left)
                row.Children.Add(arrow_tb)

                name_sub = StackPanel()
                name_sub.Orientation = Orientation.Horizontal
                name_sub.VerticalAlignment = VerticalAlignment.Center

                name_tb = TextBlock()
                name_tb.VerticalAlignment = VerticalAlignment.Center
                name_tb.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis

                if use_separate:
                    name_vals = self._f2n.get(fval, set())
                    if len(name_vals) == 0:
                        # Name param blank on all elements for this filter value
                        name_tb.Text = u"(no value — filter value used as fallback)"
                        name_tb.Foreground = Brushes.DarkGray
                        name_tb.FontStyle = System.Windows.FontStyles.Italic
                        no_name_count += 1
                    elif len(name_vals) == 1:
                        # Clean 1:1 mapping
                        nv = list(name_vals)[0]
                        name_tb.Text = nv
                        name_tb.Foreground = Brushes.Black
                        name_tb.ToolTip = nv
                    else:
                        # Multiple name values — amber warning, hover for full list
                        best = sorted(name_vals)[0]
                        name_tb.Text = best
                        name_tb.Foreground = SolidColorBrush(Color.FromRgb(160, 90, 0))
                        name_tb.ToolTip = best
                        warn_tb = TextBlock()
                        warn_tb.Text = u"  \u26a0 {} values".format(len(name_vals))
                        warn_tb.Foreground = SolidColorBrush(Color.FromRgb(200, 70, 0))
                        warn_tb.FontSize = 10
                        warn_tb.VerticalAlignment = VerticalAlignment.Center
                        all_str = u"\n".join(sorted(name_vals)[:12])
                        if len(name_vals) > 12:
                            all_str += u"\n… (+{} more)".format(len(name_vals) - 12)
                        warn_tb.ToolTip = (
                            u"Multiple name values found for filter '{}'\n"
                            u"First (alphabetical) will be used:\n\n{}".format(
                                fval, all_str))
                        name_sub.Children.Add(warn_tb)
                        warn_count += 1
                else:
                    # Same-as-filter: show grayed mirror so user can see the result
                    name_tb.Text = str(fval)
                    name_tb.Foreground = Brushes.Gray

                name_sub.Children.Insert(0, name_tb)
                row.Children.Add(name_sub)

            self._rows_panel.Children.Add(row)
            self._checkboxes.append((cb, fval))

        self._first_build = False

        # Summary warning line below the row list
        parts = []
        if warn_count:
            parts.append(
                u"\u26a0 {} row(s) have mismatched name data — "
                u"hover the \u26a0 badge to see all values.".format(warn_count))
        if no_name_count:
            parts.append(
                u"{} row(s) have no name value — "
                u"filter value will be used.".format(no_name_count))
        if parts:
            self._summary_tb.Text = u"  ".join(parts)
            self._summary_tb.Foreground = SolidColorBrush(Color.FromRgb(200, 70, 0))
        else:
            self._summary_tb.Text = u""

    def _on_name_changed(self, sender, args):
        self._rebuild_rows()

    def _check_all(self, sender, args):
        for cb, _ in self._checkboxes:
            cb.IsChecked = True

    def _uncheck_all(self, sender, args):
        for cb, _ in self._checkboxes:
            cb.IsChecked = False

    def _on_ok(self, sender, args):
        name_param_name, name_param_info = self._get_name_param_info()
        use_separate = name_param_info is not None
        to_create = []
        for cb, fval in self._checkboxes:
            if not cb.IsChecked:
                continue
            if use_separate:
                name_vals = self._f2n.get(fval, set())
                nval = sorted(name_vals)[0] if name_vals else fval
            else:
                nval = fval
            to_create.append((fval, nval))
        self.result = {
            'use_separate': use_separate,
            'name_param': name_param_name,
            'to_create': to_create,   # [(filter_val, name_val), ...]
        }
        self.Close()

    def _on_cancel(self, sender, args):
        self.result = None
        self.Close()


# ─────────────────────────────────────────────
# CORE HELPERS
# ─────────────────────────────────────────────
def alert_exit(msg):
    forms.alert(msg, title=TOOL_TITLE, exitscript=True)


def sanitize_token(s):
    if s is None:
        return ""
    t = str(s).strip().replace(" ", "")
    for ch in r'/\:;|[]{}<>?"':
        t = t.replace(ch, "-")
    return t


def _get_category_id_set():
    ids = set()
    for cname in CATEGORY_NAMES:
        try:
            c = doc.Settings.Categories.get_Item(cname)
            if c:
                ids.add(c.Id.IntegerValue)
        except Exception:
            pass
    return ids


def get_all_filter_cat_ids():
    cat_ids = List[DB.ElementId]()
    for cname in CATEGORY_NAMES:
        try:
            c = doc.Settings.Categories.get_Item(cname)
            if c:
                cat_ids.Add(c.Id)
        except Exception:
            pass
    return cat_ids


def find_assembly_from_sheet(sheet):
    for vp in DB.FilteredElementCollector(doc, sheet.Id) \
                .OfClass(DB.Viewport).ToElements():
        view = doc.GetElement(vp.ViewId)
        try:
            aid = view.AssociatedAssemblyInstanceId
            if aid and aid != DB.ElementId.InvalidElementId:
                asm = doc.GetElement(aid)
                if isinstance(asm, DB.AssemblyInstance):
                    return asm
        except Exception:
            continue
    return None


def collect_assembly_members(asm):
    cat_ids = _get_category_id_set()
    elems = []
    for eid in asm.GetMemberIds():
        e = doc.GetElement(eid)
        if e and e.Category and e.Category.Id.IntegerValue in cat_ids:
            elems.append(e)
    return elems


def _get_assembly_sheet_list():
    view_to_asm = {}
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements():
        try:
            if view.IsTemplate:
                continue
            aid = view.AssociatedAssemblyInstanceId
            if aid and aid != DB.ElementId.InvalidElementId:
                asm = doc.GetElement(aid)
                if isinstance(asm, DB.AssemblyInstance):
                    view_to_asm[view.Id.IntegerValue] = asm
        except Exception:
            continue
    if not view_to_asm:
        return []
    sheet_to_asm = {}
    for vp in DB.FilteredElementCollector(doc).OfClass(DB.Viewport).ToElements():
        try:
            vid = vp.ViewId.IntegerValue
            if vid in view_to_asm and vp.SheetId.IntegerValue not in sheet_to_asm:
                sheet_to_asm[vp.SheetId.IntegerValue] = view_to_asm[vid]
        except Exception:
            continue
    results = []
    for sid_int, asm in sheet_to_asm.items():
        try:
            sheet = doc.GetElement(DB.ElementId(sid_int))
            if sheet and isinstance(sheet, DB.ViewSheet):
                results.append((sheet, asm))
        except Exception:
            continue
    results.sort(key=lambda x: x[0].SheetNumber)
    return results


def collect_elements_from_sheet_views(sheet):
    cat_ids = _get_category_id_set()
    seen = set()
    elems = []
    for vp in DB.FilteredElementCollector(doc, sheet.Id) \
                .OfClass(DB.Viewport).ToElements():
        view = doc.GetElement(vp.ViewId)
        try:
            vtype = view.ViewType
            if vtype in (DB.ViewType.Legend, DB.ViewType.Schedule,
                         DB.ViewType.DrawingSheet):
                continue
        except Exception:
            continue
        try:
            for e in DB.FilteredElementCollector(doc, view.Id).ToElements():
                try:
                    if e.Id.IntegerValue in seen:
                        continue
                    if not e.Category:
                        continue
                    if e.Category.Id.IntegerValue not in cat_ids:
                        continue
                    seen.add(e.Id.IntegerValue)
                    elems.append(e)
                except Exception:
                    continue
        except Exception:
            continue
    return elems


# ─── Parameter discovery ─────────────────────────────────────────────
def _discover_params_from_elements(elems):
    """Walk elements, collect parameters. Handles special BIPs like
    Category, Family and Type, Type that don't self-report their
    BuiltInParameter correctly."""
    name_to_info = {}

    for e in elems:
        try:
            params = e.Parameters
            if not params:
                continue
            for p in params:
                try:
                    defn = p.Definition
                    if not defn:
                        continue
                    pname = defn.Name
                    if not pname or pname in name_to_info:
                        continue

                    info = {
                        'name': pname, 'bip': None,
                        'guid': None, 'eid': None,
                    }

                    # Check for special BIPs first
                    if pname in SPECIAL_BIPS:
                        bip = SPECIAL_BIPS[pname]
                        info['bip'] = bip
                        info['eid'] = DB.ElementId(int(bip))
                        name_to_info[pname] = info
                        continue

                    # Normal BIP check
                    try:
                        if hasattr(defn, 'BuiltInParameter'):
                            bip = defn.BuiltInParameter
                            if bip != DB.BuiltInParameter.INVALID:
                                info['bip'] = bip
                                info['eid'] = DB.ElementId(int(bip))
                    except Exception:
                        pass

                    # Shared parameter
                    if p.IsShared:
                        try:
                            info['guid'] = p.GUID
                            if not info['eid']:
                                spe = DB.SharedParameterElement.Lookup(
                                    doc, p.GUID)
                                if spe:
                                    info['eid'] = spe.Id
                        except Exception:
                            pass

                    # Fallback: ParameterElement by name
                    if not info['eid']:
                        try:
                            for pe in (DB.FilteredElementCollector(doc)
                                       .OfClass(DB.ParameterElement)
                                       .ToElements()):
                                if pe.Name == pname:
                                    info['eid'] = pe.Id
                                    break
                        except Exception:
                            pass

                    name_to_info[pname] = info
                except Exception:
                    continue
        except Exception:
            continue

    return name_to_info


def _get_param_from_element(elem, param_info):
    """Get a parameter from an element, checking both instance and type."""
    # Try instance first
    if param_info['bip'] is not None:
        try:
            p = elem.get_Parameter(param_info['bip'])
            if p:
                return p
        except Exception:
            pass
    if param_info['guid'] is not None:
        try:
            p = elem.get_Parameter(param_info['guid'])
            if p:
                return p
        except Exception:
            pass
    if param_info['name']:
        try:
            p = elem.LookupParameter(param_info['name'])
            if p:
                return p
        except Exception:
            pass

    # Try element type (for type-level params like ALL_MODEL_FAMILY_NAME)
    try:
        etype = doc.GetElement(elem.GetTypeId())
        if etype:
            if param_info['bip'] is not None:
                try:
                    p = etype.get_Parameter(param_info['bip'])
                    if p:
                        return p
                except Exception:
                    pass
            if param_info['guid'] is not None:
                try:
                    p = etype.get_Parameter(param_info['guid'])
                    if p:
                        return p
                except Exception:
                    pass
            if param_info['name']:
                try:
                    p = etype.LookupParameter(param_info['name'])
                    if p:
                        return p
                except Exception:
                    pass
    except Exception:
        pass

    return None



def _build_value_to_eid_map(elems, param_info):
    """Build a mapping from displayed value -> ElementId for ElementId-storage
    parameters (Category, Family, Type, Family and Type, etc.). This lets us
    reliably create filter rules using the correct ElementId even when the
    display value is not the same as Element.Name."""
    value_to_eid = {}
    for e in elems:
        try:
            p = _get_param_from_element(e, param_info)
            if not p or p.StorageType != DB.StorageType.ElementId:
                continue
            eid = p.AsElementId()
            if not eid or eid == DB.ElementId.InvalidElementId:
                continue

            # Prefer Revit's display string (matches what users see / pick)
            disp = None
            try:
                disp = p.AsValueString()
            except Exception:
                disp = None
            if not disp or not str(disp).strip():
                # Fall back to resolved element / category name
                try:
                    el = doc.GetElement(eid)
                    if el:
                        disp = el.Name
                except Exception:
                    disp = None
                if not disp:
                    try:
                        cat = DB.Category.GetCategory(doc, eid)
                        if cat:
                            disp = cat.Name
                    except Exception:
                        pass

            if disp and str(disp).strip():
                value_to_eid[str(disp).strip()] = eid
        except Exception:
            continue
    return value_to_eid


def get_unique_param_values(elems, param_info):
    values = set()

    # If this is an ElementId-storage param, (re)build a stable mapping
    # for THIS element set. (Different sheets/assemblies can contain
    # different types/categories, so we don't want a stale global map.)
    try:
        st = _infer_storage(param_info, elems)
        if st == DB.StorageType.ElementId:
            param_info['value_to_eid'] = _build_value_to_eid_map(elems, param_info)
        else:
            param_info['value_to_eid'] = {}
    except Exception:
        param_info['value_to_eid'] = {}

    pname = param_info.get('name', '')

    for e in elems:
        try:
            p = _get_param_from_element(e, param_info)

            # Direct extraction fallback for Family Name / Type Name
            # These BIPs can be hard to read via get_Parameter on some
            # element types (especially fabrication parts).
            if not p or (p.StorageType == DB.StorageType.String
                         and not (p.AsString() or "").strip()):
                val = None
                if pname in ("Family Name", "Family"):
                    try:
                        tid = e.GetTypeId()
                        if tid and tid != DB.ElementId.InvalidElementId:
                            etype = doc.GetElement(tid)
                            if etype:
                                # FamilySymbol has .Family.Name
                                try:
                                    val = etype.Family.Name
                                except Exception:
                                    # For non-family types, try FamilyName param
                                    try:
                                        fp = etype.get_Parameter(
                                            DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                                        if fp:
                                            val = fp.AsString()
                                    except Exception:
                                        try:
                                            val = etype.FamilyName
                                        except Exception:
                                            pass
                    except Exception:
                        pass
                elif pname in ("Type Name", "Type"):
                    try:
                        tid = e.GetTypeId()
                        if tid and tid != DB.ElementId.InvalidElementId:
                            etype = doc.GetElement(tid)
                            if etype:
                                val = etype.Name
                    except Exception:
                        pass

                if val and str(val).strip():
                    values.add(str(val).strip())
                continue

            if not p:
                continue

            # ElementId-type params (Family and Type, etc.)
            if p.StorageType == DB.StorageType.ElementId:
                try:
                    disp = p.AsValueString()
                except Exception:
                    disp = None
                if disp and str(disp).strip():
                    values.add(str(disp).strip())
                else:
                    eid = p.AsElementId()
                    if eid and eid != DB.ElementId.InvalidElementId:
                        el = doc.GetElement(eid)
                        if el:
                            values.add(el.Name)
                        else:
                            try:
                                cat = DB.Category.GetCategory(doc, eid)
                                if cat:
                                    values.add(cat.Name)
                            except Exception:
                                pass
                continue

            s = p.AsString()
            if s and s.strip():
                values.add(s.strip())
                continue
            vs = p.AsValueString()
            if vs and vs.strip():
                values.add(vs.strip())
                continue
            if p.StorageType == DB.StorageType.Integer:
                values.add(str(p.AsInteger()))
        except Exception:
            pass
    return sorted(v for v in values if v and str(v).strip())


def _infer_storage(param_info, elems):
    for e in elems:
        try:
            p = _get_param_from_element(e, param_info)
            if p:
                return p.StorageType
        except Exception:
            pass
    return DB.StorageType.String


def get_param_string_value(elem, param_info):
    """Return the display string for a parameter on an element.
    Works for any param (not just filterable ones). Returns None if missing."""
    pname = param_info.get('name', '')
    p = _get_param_from_element(elem, param_info)

    if not p or (p.StorageType == DB.StorageType.String
                 and not (p.AsString() or "").strip()):
        # Direct fallback for Family/Type Name on fabrication parts etc.
        if pname in ("Family Name", "Family"):
            try:
                tid = elem.GetTypeId()
                if tid and tid != DB.ElementId.InvalidElementId:
                    etype = doc.GetElement(tid)
                    if etype:
                        try:
                            return etype.Family.Name
                        except Exception:
                            pass
                        try:
                            fp = etype.get_Parameter(
                                DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
                            if fp:
                                return fp.AsString()
                        except Exception:
                            pass
            except Exception:
                pass
        elif pname in ("Type Name", "Type"):
            try:
                tid = elem.GetTypeId()
                if tid and tid != DB.ElementId.InvalidElementId:
                    etype = doc.GetElement(tid)
                    if etype:
                        return etype.Name
            except Exception:
                pass
        if not p:
            return None

    if p.StorageType == DB.StorageType.ElementId:
        try:
            disp = p.AsValueString()
            if disp and disp.strip():
                return disp.strip()
        except Exception:
            pass
        try:
            eid = p.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                el = doc.GetElement(eid)
                if el:
                    return el.Name
        except Exception:
            pass
        return None

    s = p.AsString()
    if s and s.strip():
        return s.strip()
    vs = p.AsValueString()
    if vs and vs.strip():
        return vs.strip()
    if p.StorageType == DB.StorageType.Integer:
        return str(p.AsInteger())
    return None


def build_filter_to_name_map(elems, filter_param_info, name_param_info):
    """Map each filter-param value → set of name-param values found on elements.

    Returns:
        dict: { filter_value(str) → set of name_values(str) }
    Used both for validation (checking 1:1) and for resolving the name to
    write to the new sheet number."""
    f2n = {}
    for e in elems:
        try:
            fval = get_param_string_value(e, filter_param_info)
            if not fval:
                continue
            nval = get_param_string_value(e, name_param_info)
            if not nval:
                continue
            if fval not in f2n:
                f2n[fval] = set()
            f2n[fval].add(nval)
        except Exception:
            continue
    return f2n


# ─── Filter creation / application ───────────────────────────────────
def _build_filter_name(param_name, value):
    return "{}-{}-{}".format(
        TOOL_PREFIX, sanitize_token(param_name), sanitize_token(value))


def _find_existing_filter(name):
    for f in (DB.FilteredElementCollector(doc)
              .OfClass(DB.ParameterFilterElement).ToElements()):
        if f.Name == name:
            return f
    return None


def _get_filter_cat_ids_for_param(param_info, all_cat_ids):
    valid = List[DB.ElementId]()
    # Build set of parameter EIDs to check: the original + any remap
    check_eids = set()
    eid = param_info.get('eid')
    if eid:
        check_eids.add(eid.IntegerValue)
    # Also check the remapped BIP (e.g. Family → ALL_MODEL_FAMILY_NAME)
    remap = FILTER_REMAP.get(param_info.get('name', ''))
    if remap and remap.get('eid'):
        check_eids.add(remap['eid'].IntegerValue)

    for cat_id in all_cat_ids:
        try:
            single_cat = List[DB.ElementId]()
            single_cat.Add(cat_id)
            pids = DB.ParameterFilterUtilities.GetFilterableParametersInCommon(
                doc, single_cat)
            for pid in pids:
                if pid.IntegerValue in check_eids:
                    valid.Add(cat_id)
                    break
        except Exception:
            continue
    return valid



def _resolve_element_id_for_value(param_info, value, elems):
    """For ElementId-type params, resolve the chosen display string back to
    an ElementId for filter rule creation.

    We prefer a cached mapping built from the actual elements because:
    - Type/Family display strings often don't match Element.Name
    - Category can resolve via Category.GetCategory
    """
    try:
        m = param_info.get('value_to_eid') or {}
        if value in m:
            return m[value]
    except Exception:
        pass

    # Fallback: brute-force match against AsValueString() on elements
    for e in elems:
        try:
            p = _get_param_from_element(e, param_info)
            if not p or p.StorageType != DB.StorageType.ElementId:
                continue
            try:
                disp = p.AsValueString()
            except Exception:
                disp = None
            if disp and str(disp).strip() == str(value).strip():
                eid = p.AsElementId()
                if eid and eid != DB.ElementId.InvalidElementId:
                    return eid
        except Exception:
            continue
    return None


def _validate_cat_ids_for_rule(rule, cat_ids):
    """Test which categories from cat_ids actually accept the given rule.
    Revit's SetElementFilter will reject the whole filter if ANY category
    in the list doesn't support the rule's parameter.  We test each
    category individually by creating a temporary ElementParameterFilter
    and checking ParameterFilterUtilities.IsFilterValidForCategories."""
    valid = List[DB.ElementId]()
    test_filter = DB.ElementParameterFilter(rule)
    for cid in cat_ids:
        try:
            single = List[DB.ElementId]()
            single.Add(cid)
            if DB.ParameterFilterUtilities.IsFilterValidForCategories(
                    doc, single, test_filter):
                valid.Add(cid)
        except Exception:
            # If the API call itself fails for this category, skip it
            continue
    return valid


def _validate_cat_ids_for_rule_brute(rule, cat_ids):
    """Fallback: try Create + SetElementFilter one category at a time.
    More expensive but works even if IsFilterValidForCategories is
    unavailable (older Revit builds)."""
    valid = List[DB.ElementId]()
    test_filter = DB.ElementParameterFilter(rule)
    for cid in cat_ids:
        try:
            single = List[DB.ElementId]()
            single.Add(cid)
            tmp_name = "__nwss_test_{}_{}".format(
                cid.IntegerValue, System.Guid.NewGuid().ToString("N")[:8])
            tmp = DB.ParameterFilterElement.Create(doc, tmp_name, single)
            try:
                tmp.SetElementFilter(test_filter)
                valid.Add(cid)
            except Exception:
                pass
            doc.Delete(tmp.Id)
        except Exception:
            continue
    return valid


def get_or_create_filter_inverted(param_info, param_name, value,
                                   filter_cat_ids, elems):
    fname = _build_filter_name(param_name, value)
    existing = _find_existing_filter(fname)
    if existing:
        return existing, False

    # ── Apply remap for BIPs that Revit won't accept in filters ────────
    # e.g. "Family" (ELEM_FAMILY_PARAM) → ALL_MODEL_FAMILY_NAME (string)
    remap = FILTER_REMAP.get(param_name)
    if remap:
        filter_param_id = remap['eid']
        filter_storage = remap['storage_override']
    else:
        filter_param_id = param_info['eid']
        filter_storage = _infer_storage(param_info, elems)

    if not filter_param_id:
        raise Exception("No ElementId for '{}'.".format(param_name))

    rule = None
    rule_errors = []

    # ── Strategy 1: Factory rules (preferred) ──────────────────────────
    try:
        if filter_storage == DB.StorageType.ElementId:
            val_eid = _resolve_element_id_for_value(param_info, value, elems)
            if val_eid:
                rule = DB.ParameterFilterRuleFactory.CreateNotEqualsRule(
                    filter_param_id, val_eid)
    except Exception as ex:
        rule_errors.append("Factory EID: {}".format(str(ex)))

    if not rule:
        try:
            if filter_storage == DB.StorageType.Integer:
                rule = DB.ParameterFilterRuleFactory.CreateNotEqualsRule(
                    filter_param_id, int(str(value).strip()))
        except (ValueError, TypeError, Exception) as ex:
            rule_errors.append("Factory INT: {}".format(str(ex)))

    if not rule:
        try:
            if filter_storage == DB.StorageType.Double:
                rule = DB.ParameterFilterRuleFactory.CreateNotEqualsRule(
                    filter_param_id, float(str(value).strip()), 0.0001)
        except (ValueError, TypeError, Exception) as ex:
            rule_errors.append("Factory DBL: {}".format(str(ex)))

    if not rule:
        try:
            rule = DB.ParameterFilterRuleFactory.CreateNotEqualsRule(
                filter_param_id, str(value), False)
        except Exception as ex:
            rule_errors.append("Factory STR: {}".format(str(ex)))

    # ── Strategy 2: Manual FilterInverseRule with ParameterValueProvider ─
    if not rule and filter_storage == DB.StorageType.ElementId:
        val_eid = _resolve_element_id_for_value(param_info, value, elems)
        if val_eid:
            try:
                provider = DB.ParameterValueProvider(filter_param_id)
                base_rule = DB.FilterElementIdRule(
                    provider, DB.FilterNumericEquals(), val_eid)
                rule = DB.FilterInverseRule(base_rule)
            except Exception as ex:
                rule_errors.append("Manual EID inverse: {}".format(str(ex)))

    # ── Strategy 3: String-based fallback ──────────────────────────────
    if not rule:
        try:
            provider = DB.ParameterValueProvider(filter_param_id)
            try:
                base_rule = DB.FilterStringRule(
                    provider, DB.FilterStringEquals(), str(value))
            except TypeError:
                base_rule = DB.FilterStringRule(
                    provider, DB.FilterStringEquals(), str(value), False)
            rule = DB.FilterInverseRule(base_rule)
        except Exception as ex:
            rule_errors.append("String inverse: {}".format(str(ex)))

    if not rule:
        raise Exception(
            "All filter strategies failed for '{}' = '{}'\n{}".format(
                param_name, value, "\n".join(rule_errors)))

    try:
        elem_filter = DB.ElementParameterFilter(rule)
    except Exception as ex:
        raise Exception(
            "ElementParameterFilter failed for '{}' = '{}': {}".format(
                param_name, value, str(ex)))

    # ── Validate categories against the rule ───────────────────────────
    # For special BIPs (Category, Family, Type, etc.) we may have been
    # given all_cat_ids but not every category supports the parameter.
    # Narrow down to only those that accept the rule.
    validated_cat_ids = filter_cat_ids

    if param_name in SPECIAL_BIPS or param_name in FILTER_REMAP:
        # For remapped params, re-check which categories support the
        # FILTER parameter (not the discovery parameter).
        if remap:
            remap_info = {'eid': remap['eid'], 'name': param_name}
            remap_cats = List[DB.ElementId]()
            for cat_id in filter_cat_ids:
                try:
                    single_cat = List[DB.ElementId]()
                    single_cat.Add(cat_id)
                    pids = DB.ParameterFilterUtilities \
                        .GetFilterableParametersInCommon(doc, single_cat)
                    for pid in pids:
                        if pid.IntegerValue == remap['eid'].IntegerValue:
                            remap_cats.Add(cat_id)
                            break
                except Exception:
                    continue
            if remap_cats.Count > 0:
                validated_cat_ids = remap_cats

        # If still using all_cat_ids (no remap or remap found nothing),
        # try runtime validation
        if validated_cat_ids.Count > 10:
            try:
                tested = _validate_cat_ids_for_rule(rule, validated_cat_ids)
                if tested.Count > 0:
                    validated_cat_ids = tested
            except Exception:
                try:
                    tested = _validate_cat_ids_for_rule_brute(
                        rule, validated_cat_ids)
                    if tested.Count > 0:
                        validated_cat_ids = tested
                except Exception:
                    pass  # keep validated_cat_ids as-is, let Create/Set fail with a clear message

    if validated_cat_ids.Count == 0:
        debug = "param_id={}, storage={}, remap={}".format(
            filter_param_id.IntegerValue, filter_storage,
            "yes" if remap else "no")
        raise Exception(
            "No categories support filter rule for '{}' = '{}'.\n{}".format(
                param_name, value, debug))

    try:
        pf = DB.ParameterFilterElement.Create(doc, fname, validated_cat_ids)
    except Exception as ex:
        raise Exception(
            "ParameterFilterElement.Create failed for '{}': {}".format(
                fname, str(ex)))
    try:
        pf.SetElementFilter(elem_filter)
    except Exception as ex:
        raise Exception(
            "SetElementFilter failed for '{}' = '{}': {}".format(
                param_name, value, str(ex)))
    return pf, True


def remove_view_template_from_view(view):
    view_name = "?"
    try:
        view_name = view.Name
    except Exception:
        pass
    try:
        current_tid = view.ViewTemplateId
        if current_tid == DB.ElementId.InvalidElementId:
            return True, "{}: no template".format(view_name)
        view.ViewTemplateId = DB.ElementId.InvalidElementId
        after_tid = view.ViewTemplateId
        if after_tid == DB.ElementId.InvalidElementId:
            return True, "{}: template removed".format(view_name)
        else:
            return False, "{}: template set failed".format(view_name)
    except Exception as ex:
        return False, "{}: {}".format(view_name, str(ex))


def apply_filter_to_view(view, pf):
    view_name = "?"
    try:
        view_name = view.Name
    except Exception:
        pass
    try:
        vtype = view.ViewType
        if vtype in (DB.ViewType.Legend, DB.ViewType.Schedule,
                     DB.ViewType.DrawingSheet, DB.ViewType.Internal,
                     DB.ViewType.Undefined):
            return True, "{}: skipped ({})".format(view_name, vtype)
    except Exception:
        pass
    try:
        tid = view.ViewTemplateId
        if tid != DB.ElementId.InvalidElementId:
            tpl = doc.GetElement(tid)
            tpl_name = tpl.Name if tpl else str(tid.IntegerValue)
            return False, "{}: has template '{}'".format(view_name, tpl_name)
    except Exception:
        pass
    filter_id = pf.Id
    try:
        current_filters = DB.View.GetFilters(view)
        already_has = any(
            fid.IntegerValue == filter_id.IntegerValue
            for fid in current_filters)
    except Exception as ex:
        return False, "{}: GetFilters: {}".format(view_name, str(ex))
    if not already_has:
        try:
            DB.View.AddFilter(view, filter_id)
        except Exception as ex:
            return False, "{}: AddFilter: {}".format(view_name, str(ex))
    try:
        DB.View.SetFilterVisibility(view, filter_id, False)
    except Exception as ex:
        return False, "{}: SetFilterVisibility: {}".format(view_name, str(ex))
    return True, "{}: OK".format(view_name)


def all_sheets():
    return list(DB.FilteredElementCollector(doc)
                .OfClass(DB.ViewSheet).ToElements())


def sheet_number_exists(num):
    return any(s.SheetNumber == num for s in all_sheets())


def next_unique_sheet_number(base_num, value, sep="-"):
    desired = "{}{}{}".format(base_num, sep, sanitize_token(value))
    if not sheet_number_exists(desired):
        return desired
    i = 2
    while True:
        alt = "{}{}{}".format(desired, sep, i)
        if not sheet_number_exists(alt):
            return alt
        i += 1


def next_unique_sheet_number_replace(value, sep="-"):
    """Replace mode: sheet number is just the sanitized value, no base prefix."""
    desired = sanitize_token(value)
    if not sheet_number_exists(desired):
        return desired
    i = 2
    while True:
        alt = "{}{}{}".format(desired, sep, i)
        if not sheet_number_exists(alt):
            return alt
        i += 1


def view_name_exists(name):
    return any(v.Name == name
               for v in DB.FilteredElementCollector(doc)
               .OfClass(DB.View).ToElements())


def next_unique_view_name(base_name):
    if not view_name_exists(base_name):
        return base_name
    i = 2
    while True:
        alt = "{} ({})".format(base_name, i)
        if not view_name_exists(alt):
            return alt
        i += 1


def rename_views_on_sheet(sheet, suffix):
    for vp in DB.FilteredElementCollector(doc, sheet.Id).OfClass(DB.Viewport):
        v = doc.GetElement(vp.ViewId)
        if v.ViewType == DB.ViewType.Legend:
            continue
        try:
            new_name = "{}-{}".format(v.Name, sanitize_token(suffix))
            v.Name = next_unique_view_name(new_name)
        except Exception:
            pass



# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

# Step 1: Settings
dlg = SettingsDialog()
dlg.ShowDialog()
if not dlg.result:
    alert_exit("Cancelled.")
source_mode = dlg.result['source']
sheet_mode = dlg.result['sheet_mode']
replace_number = dlg.result['replace_number']
allow_name_param = dlg.result['allow_name_param']

# Step 2: Select sheets
selected = []
if sheet_mode == SHEET_CURRENT:
    v = doc.ActiveView
    if not isinstance(v, DB.ViewSheet):
        alert_exit("Active view is not a sheet.")
    if source_mode == MODE_ASSEMBLY:
        asm = find_assembly_from_sheet(v)
        if not asm:
            alert_exit("No assembly on current sheet.\n"
                       "Use 'View Elements' for non-assembly sheets.")
        selected = [{'sheet': v, 'asm': asm}]
    else:
        selected = [{'sheet': v, 'asm': None}]

elif sheet_mode == SHEET_ASSEMBLY:
    asm_list = _get_assembly_sheet_list()
    if not asm_list:
        alert_exit("No assembly sheets found.")
    labels = ["{} - {}".format(s.SheetNumber, s.Name) for s, a in asm_list]
    label_map = {l: i for i, l in enumerate(labels)}
    dlg2 = MultiPickerDialog(labels,
        "Assembly Sheets ({})".format(len(labels)),
        HELP_SHEET_PICK, "Select")
    dlg2.ShowDialog()
    if not dlg2.result:
        alert_exit("No sheets selected.")
    for label in dlg2.result:
        idx = label_map.get(label)
        if idx is not None:
            s, a = asm_list[idx]
            selected.append({'sheet': s, 'asm': a})

else:
    every_sheet = sorted(all_sheets(), key=lambda s: s.SheetNumber)
    if not every_sheet:
        alert_exit("No sheets in project.")
    labels = ["{} - {}".format(s.SheetNumber, s.Name) for s in every_sheet]
    label_map = {l: i for i, l in enumerate(labels)}
    dlg2 = MultiPickerDialog(labels,
        "All Sheets ({})".format(len(labels)),
        HELP_SHEET_PICK, "Select")
    dlg2.ShowDialog()
    if not dlg2.result:
        alert_exit("No sheets selected.")
    no_asm = []
    for label in dlg2.result:
        idx = label_map.get(label)
        if idx is None:
            continue
        s = every_sheet[idx]
        if source_mode == MODE_ASSEMBLY:
            asm = find_assembly_from_sheet(s)
            if asm:
                selected.append({'sheet': s, 'asm': asm})
            else:
                no_asm.append(label)
        else:
            selected.append({'sheet': s, 'asm': None})
    if no_asm:
        forms.alert("Skipped (no assembly):\n\n{}".format(
            "\n".join(no_asm[:15])))

if not selected:
    alert_exit("No valid sheets selected.")

# Step 3: Discover parameters
first = selected[0]
if source_mode == MODE_ASSEMBLY and first['asm']:
    first_elems = collect_assembly_members(first['asm'])
else:
    first_elems = collect_elements_from_sheet_views(first['sheet'])
if not first_elems:
    alert_exit("No elements found on first sheet.")

all_cat_ids = get_all_filter_cat_ids()
name_to_info = _discover_params_from_elements(first_elems)
if not name_to_info:
    alert_exit("No parameters found on elements.")

# Keep a full copy for name-param choices (before filtering for filterability)
all_name_params = dict(name_to_info)

# Remove params that are known to not work in view filters
for hidden in HIDDEN_PARAMS:
    name_to_info.pop(hidden, None)

# Build the set of ALL filterable parameter EIDs across our categories
all_filterable_eids = set()
try:
    filterable_pids = DB.ParameterFilterUtilities.GetFilterableParametersInCommon(
        doc, all_cat_ids)
    for pid in filterable_pids:
        all_filterable_eids.add(pid.IntegerValue)
except Exception:
    pass
# Also check per-category to catch more parameters
for cat_id in all_cat_ids:
    try:
        single = List[DB.ElementId]()
        single.Add(cat_id)
        pids = DB.ParameterFilterUtilities.GetFilterableParametersInCommon(
            doc, single)
        for pid in pids:
            all_filterable_eids.add(pid.IntegerValue)
    except Exception:
        continue

# Filter name_to_info to only params that are filterable
usable = {}
for pname, pinfo in name_to_info.items():
    # Check if the param itself (or its remap) is in the filterable set
    eid = pinfo.get('eid')
    remap = FILTER_REMAP.get(pname)
    is_filterable = False
    if eid and eid.IntegerValue in all_filterable_eids:
        is_filterable = True
    if remap and remap['eid'].IntegerValue in all_filterable_eids:
        is_filterable = True
    if is_filterable:
        usable[pname] = pinfo

if not usable:
    alert_exit("No filterable parameters found on elements.\n"
               "({} params discovered but none are supported "
               "by Revit view filters.)".format(len(name_to_info)))

dlg3 = SinglePickerDialog(
    sorted(usable.keys()),
    "Filter Parameter  ({} available)".format(len(usable)),
    HELP_FILTER_PARAM, "Select")
dlg3.ShowDialog()
if not dlg3.result:
    alert_exit("No parameter selected.")
picked_param = dlg3.result
param_info = usable[picked_param]

filter_cat_ids = _get_filter_cat_ids_for_param(param_info, all_cat_ids)
if filter_cat_ids.Count == 0:
    # For remapped params, try getting categories for the remapped BIP
    remap = FILTER_REMAP.get(picked_param)
    if remap:
        remap_info = {'eid': remap['eid'], 'name': picked_param}
        filter_cat_ids = _get_filter_cat_ids_for_param(remap_info, all_cat_ids)
    # For other special BIPs, use all categories as fallback
    if filter_cat_ids.Count == 0 and picked_param in SPECIAL_BIPS:
        filter_cat_ids = all_cat_ids
    if filter_cat_ids.Count == 0:
        alert_exit("'{}' is not filterable on any configured category.".format(
            picked_param))

# Step 4: Pick values + name param per sheet (combined dialog)
job_list = []
name_param_used = None   # track for report; set on first confirmed selection

for entry in selected:
    sheet = entry['sheet']
    asm = entry['asm']
    sheet_label = u"{} - {}".format(sheet.SheetNumber, sheet.Name)
    if source_mode == MODE_ASSEMBLY and asm:
        elems = collect_assembly_members(asm)
    else:
        elems = collect_elements_from_sheet_views(sheet)
    if not elems:
        continue
    values = get_unique_param_values(elems, param_info)
    if not values:
        continue

    dlg4 = FilterNamePickerDialog(
        values, elems, param_info,
        all_name_params, picked_param,
        sheet_label,
        allow_name_param=allow_name_param)
    dlg4.ShowDialog()
    if not dlg4.result or not dlg4.result['to_create']:
        continue

    if name_param_used is None:
        name_param_used = dlg4.result['name_param'] or picked_param

    job_list.append((sheet, asm, elems, dlg4.result['to_create']))

if not job_list:
    alert_exit("Nothing to create.")

# Step 5: Execute
try:
    dup_option = DB.SheetDuplicateOption.DuplicateSheetWithViewsAndDetailing
except Exception:
    alert_exit("Requires Revit 2023+.")

all_created = []
all_errors = []
all_filter_log = []

t = DB.Transaction(doc, "{} - Split sheets".format(TOOL_PREFIX))
t.Start()
try:
    for template_sheet, asm, elems, to_create_pairs in job_list:
        for fval, sheet_name_val in to_create_pairs:
            try:
                pf, _ = get_or_create_filter_inverted(
                    param_info, picked_param, fval, filter_cat_ids, elems)
                new_sheet_id = template_sheet.Duplicate(dup_option)
                new_sheet = doc.GetElement(new_sheet_id)
                if replace_number:
                    new_sheet.SheetNumber = next_unique_sheet_number_replace(
                        sheet_name_val)
                else:
                    new_sheet.SheetNumber = next_unique_sheet_number(
                        template_sheet.SheetNumber, sheet_name_val, sep="-")
                for vp in DB.FilteredElementCollector(doc, new_sheet.Id) \
                            .OfClass(DB.Viewport).ToElements():
                    vw = doc.GetElement(vp.ViewId)
                    if REMOVE_VIEW_TEMPLATES:
                        tpl_ok, tpl_msg = remove_view_template_from_view(vw)
                        if not tpl_ok:
                            all_filter_log.append(
                                "[TPL FAIL] {}".format(tpl_msg))
                    flt_ok, flt_msg = apply_filter_to_view(vw, pf)
                    all_filter_log.append("[{}] {}".format(
                        "OK" if flt_ok else "FAIL", flt_msg))
                rename_views_on_sheet(new_sheet, fval)
                all_created.append(new_sheet.SheetNumber)
            except Exception as ex:
                all_errors.append("{} > {}: {}".format(
                    template_sheet.SheetNumber, fval, str(ex)))

    if all_created:
        t.Commit()
    else:
        t.RollBack()
        err_detail = ""
        if all_errors:
            err_detail = "\n\nErrors:\n" + "\n".join(all_errors[:10])
        if all_filter_log:
            fails = [e for e in all_filter_log
                     if "[FAIL]" in e or "[TPL FAIL]" in e]
            if fails:
                err_detail += "\n\nFilter log:\n" + "\n".join(fails[:10])
        alert_exit("No sheets were created." + err_detail)
except Exception as ex:
    if t.HasStarted() and not t.HasEnded():
        t.RollBack()
    alert_exit("Transaction failed:\n{}".format(str(ex)))

# Step 6: Report
lines = []
lines.append("Created {} sheet(s):".format(len(all_created)))
for sn in all_created[:30]:
    lines.append("  {}".format(sn))
if len(all_created) > 30:
    lines.append("  ... (+{} more)".format(len(all_created) - 30))

lines.append("")
if name_param_used and name_param_used != picked_param:
    lines.append("Filter param:  {}".format(picked_param))
    lines.append("Name param:    {}".format(name_param_used))
else:
    lines.append("Parameter: {}".format(picked_param))

fail_count = sum(1 for e in all_filter_log
                 if "[FAIL]" in e or "[TPL FAIL]" in e)
ok_count = sum(1 for e in all_filter_log if "[OK]" in e)
lines.append("")
lines.append("Views: {} OK, {} failed".format(ok_count, fail_count))
if fail_count > 0:
    for entry in all_filter_log:
        if "[FAIL]" in entry or "[TPL FAIL]" in entry:
            lines.append("  {}".format(entry))
if all_errors:
    lines.append("")
    lines.append("Errors ({}):".format(len(all_errors)))
    for e in all_errors[:10]:
        lines.append("  {}".format(e))

forms.alert("\n".join(lines), title=TOOL_TITLE)