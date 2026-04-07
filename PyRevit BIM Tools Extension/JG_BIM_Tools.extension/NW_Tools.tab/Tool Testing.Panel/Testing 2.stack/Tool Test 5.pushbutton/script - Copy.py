# -*- coding: utf-8 -*-
"""
RFA Type Manager (script.py)
============================
pyRevit entry point for the RFA Type Manager tool.

Provides a unified UI for:
    - Exporting all family types + parameter values to Excel for bulk editing
    - Importing edited Excel back into the family (create / update / delete)
    - In-Revit type grid: add, edit, duplicate, delete, multi-select fill-down
    - Parameter management: add/remove family params, set formulas
    - SP Profile integration: install SP Loader profiles directly into this family

File layout (all in this pushbutton folder):
    script.py       - this file (UI + entry point)
    type_core.py    - Revit API backend
    excel_bridge.py - CPython Excel I/O (runs as subprocess)
    bundle.yaml     - pyRevit button metadata
"""

import os
import sys
import json

import clr
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System import Array, String
from System.Drawing import (
    Size, Point, Color, Font as DFont, FontStyle,
    SolidBrush, Brushes
)
from System.Windows.Forms import (
    Application, Form, TabControl, TabPage,
    Panel, SplitContainer,
    Button, Label, TextBox, ComboBox, ComboBoxStyle,
    ListView, ListViewItem, ColumnHeader, View,
    DataGridView, DataGridViewColumn, DataGridViewTextBoxColumn,
    DataGridViewCellEventArgs, DataGridViewSelectionMode,
    DataGridViewAutoSizeColumnsMode, DataGridViewClipboardCopyMode,
    DataGridViewColumnHeadersHeightSizeMode,
    CheckBox, CheckedListBox, RadioButton,
    ToolStrip, ToolStripButton, ToolStripSeparator, ToolStripLabel,
    ToolStripComboBox, ToolStripGripStyle, ToolStripControlHost,
    ContextMenuStrip, ToolStripMenuItem,
    StatusStrip, ToolStripStatusLabel,
    AnchorStyles, DockStyle, Padding,
    DialogResult, FormStartPosition, FormBorderStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    OpenFileDialog, SaveFileDialog,
    Keys, Clipboard, DataObject, MouseButtons,
    ProgressBar, ToolTip,
    BorderStyle, ScrollBars,
    Orientation,
)
from Autodesk.Revit.UI import TaskDialog

import type_core as core


# =============================================================================
# CONSTANTS
# =============================================================================

TOOL_TITLE   = "RFA Type Manager"
BUTTON_DIR   = os.path.dirname(__file__)

# Colors used in the grid
COLOR_FORMULA_BG  = Color.FromArgb(235, 235, 235)
COLOR_FORMULA_FG  = Color.FromArgb(140, 140, 140)
COLOR_HEADER_BG   = Color.FromArgb(47, 61, 78)
COLOR_HEADER_FG   = Color.White
COLOR_ALT_ROW     = Color.FromArgb(245, 248, 252)
COLOR_MODIFIED    = Color.FromArgb(255, 252, 220)
COLOR_NEW_ROW     = Color.FromArgb(235, 255, 235)


# =============================================================================
# UNIT CONVERSION HELPERS
# =============================================================================

def _to_frac_inches(feet_val):
    """
    Converts decimal feet to a fractional inches string.
    e.g. 0.520833... -> '6 1/4"'
    Rounds to nearest 1/16".
    """
    try:
        total_inches = float(feet_val) * 12.0
        neg = total_inches < 0
        total_inches = abs(total_inches)
        whole = int(total_inches)
        frac  = total_inches - whole
        # Round to nearest 1/16
        sixteenths = int(round(frac * 16))
        if sixteenths == 16:
            whole += 1; sixteenths = 0
        if sixteenths == 0:
            result = str(whole) + '"'
        else:
            # Reduce fraction using GCD (no imports needed)
            def _gcd(a, b):
                while b: a, b = b, a % b
                return a
            g = _gcd(sixteenths, 16)
            result = (str(whole) + " " if whole else "") + str(sixteenths // g) + "/" + str(16 // g) + '"'
        return ("-" if neg else "") + result
    except:
        return ""

def _from_frac_inches(s):
    """
    Parses a fractional inches string back to decimal inches.
    Accepts: '6 1/4"', '1-1/2"', '0.500', '6.25', '1 1/2'
    Returns decimal inches (float).
    """
    try:
        s = str(s).strip().rstrip('"').strip()
        if not s:
            return 0.0
        # Try plain decimal first
        try:
            return float(s)
        except:
            pass
        # Handle formats like "1-1/2" or "1 1/2"
        s = s.replace("-", " ")
        parts = s.split()
        total = 0.0
        for part in parts:
            if "/" in part:
                n, d = part.split("/")
                total += float(n) / float(d)
            else:
                total += float(part)
        return total
    except:
        return 0.0


# =============================================================================
# COLUMN PICKER DIALOG
# =============================================================================

class ColumnPickerDialog(Form):
    """
    Shown before Export. Lists all non-formula parameters with checkboxes.
    User selects which columns to include in the Excel export.
    Formula-driven params are listed but permanently unchecked and grayed.
    """

    def __init__(self, param_defs, previously_selected=None):
        Form.__init__(self)
        self.Text            = "Select Export Columns"
        self.Size            = Size(480, 560)
        self.StartPosition   = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False

        self.selected_names  = []   # filled on OK

        lbl = Label()
        lbl.Text     = "Choose which parameters to export as columns.\nFormula-driven params are excluded."
        lbl.Location = Point(12, 12)
        lbl.Size     = Size(440, 36)
        lbl.Font     = DFont("Arial", 9)

        # Search box
        self.txtSearch = TextBox()
        self.txtSearch.Location    = Point(12, 54)
        self.txtSearch.Size        = Size(340, 22)
        self.txtSearch.TextChanged += self._on_filter

        btnAll  = Button(); btnAll.Text  = "All";  btnAll.Size = Size(50, 22); btnAll.Location = Point(358, 54)
        btnNone = Button(); btnNone.Text = "None"; btnNone.Size = Size(50, 22); btnNone.Location = Point(412, 54)
        btnAll.Click  += self._on_all
        btnNone.Click += self._on_none

        self.clb = CheckedListBox()
        self.clb.Location     = Point(12, 82)
        self.clb.Size         = Size(450, 380)
        self.clb.Anchor       = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.clb.CheckOnClick = True

        self._all_params = param_defs   # list of param defs
        # Default: all writable params selected; formula params shown but locked
        self._prev       = set(previously_selected or [p["name"] for p in param_defs if not p.get("has_formula") and not p.get("read_only")])
        self._rebuild_list("")

        btnOK     = Button(); btnOK.Text     = "Export";  btnOK.Size = Size(90, 28)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(90, 28)
        btnOK.Anchor     = AnchorStyles.Bottom | AnchorStyles.Right
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnOK.Location     = Point(self.ClientSize.Width - 194, self.ClientSize.Height - 40)
        btnCancel.Location = Point(self.ClientSize.Width - 100, self.ClientSize.Height - 40)

        def _ok(s, a):
            self.selected_names = [
                str(self.clb.Items[i]).split(u" [formula")[0]
                for i in range(self.clb.Items.Count)
                if self.clb.GetItemChecked(i)
                and u"[formula" not in str(self.clb.Items[i])
            ]
            self.DialogResult = DialogResult.OK
            self.Close()

        btnOK.Click     += _ok
        btnCancel.Click += lambda s, a: self.Close()

        for w in [lbl, self.txtSearch, btnAll, btnNone, self.clb, btnOK, btnCancel]:
            self.Controls.Add(w)

    def _rebuild_list(self, needle):
        self.clb.Items.Clear()
        needle = needle.lower()
        for p in self._all_params:
            name       = p["name"]
            is_formula = p.get("has_formula") or p.get("read_only")
            if needle and needle not in name.lower():
                continue
            if is_formula:
                # Formula params: shown as info, always exported as visible gray cols
                label = name + u" [formula - always included as visible]"
            else:
                label = name
            idx = self.clb.Items.Add(label)
            if is_formula:
                # Show as checked but disabled visually via label
                self.clb.SetItemChecked(idx, True)
            else:
                self.clb.SetItemChecked(idx, name in self._prev)

    def _on_filter(self, s, a):
        self._rebuild_list(self.txtSearch.Text or "")

    def _on_all(self, s, a):
        for i in range(self.clb.Items.Count):
            lbl = str(self.clb.Items[i])
            if "[formula]" not in lbl:
                self.clb.SetItemChecked(i, True)

    def _on_none(self, s, a):
        for i in range(self.clb.Items.Count):
            self.clb.SetItemChecked(i, False)


# =============================================================================
# IMPORT PREVIEW DIALOG
# =============================================================================

class ImportPreviewDialog(Form):
    """
    Shown before applying an imported snapshot. Summarises what will change
    and lets the user configure update / create / delete options.
    """

    def __init__(self, existing_names, incoming_snapshot):
        Form.__init__(self)
        self.Text            = "Import Preview"
        self.Size            = Size(540, 420)
        self.StartPosition   = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False

        incoming_names = set(t["name"] for t in incoming_snapshot.get("types", []))
        existing_set   = set(existing_names)

        new_types     = sorted(incoming_names - existing_set)
        update_types  = sorted(incoming_names & existing_set)
        missing_types = sorted(existing_set   - incoming_names)

        lbl_title = Label()
        lbl_title.Text     = "Review changes before applying:"
        lbl_title.Location = Point(12, 12)
        lbl_title.Size     = Size(500, 20)
        lbl_title.Font     = DFont("Arial", 10, FontStyle.Bold)

        def _stat_label(y, text, color):
            l = Label(); l.Text = text; l.Location = Point(20, y)
            l.Size = Size(490, 18); l.Font = DFont("Arial", 9)
            l.ForeColor = color
            return l

        y = 40
        self.Controls.Add(_stat_label(y, "{} types will be CREATED".format(len(new_types)),     Color.DarkGreen)); y += 22
        self.Controls.Add(_stat_label(y, "{} types will be UPDATED".format(len(update_types)),   Color.DarkBlue));  y += 22
        self.Controls.Add(_stat_label(y, "{} types NOT in Excel (see option below)".format(len(missing_types)), Color.DarkOrange)); y += 28

        # Options
        self.chkUpdate  = CheckBox(); self.chkUpdate.Text  = "Update existing types";         self.chkUpdate.Checked = True
        self.chkCreate  = CheckBox(); self.chkCreate.Text  = "Create new types";              self.chkCreate.Checked = True
        self.chkDelete  = CheckBox(); self.chkDelete.Text  = "Delete types NOT in Excel ({})".format(len(missing_types)); self.chkDelete.Checked = False
        self.chkDelete.ForeColor = Color.DarkRed

        for i, cb in enumerate([self.chkUpdate, self.chkCreate, self.chkDelete]):
            cb.Location = Point(20, y + i * 26); cb.Size = Size(480, 22)
            cb.Font     = DFont("Arial", 9)
            self.Controls.Add(cb)
        y += 3 * 26 + 10

        # Detail list (scrollable)
        detail_lbl = Label(); detail_lbl.Text = "Detail (new types):"; detail_lbl.Location = Point(12, y); detail_lbl.Size = Size(200, 18)
        detail_lbl.Font = DFont("Arial", 8, FontStyle.Bold)
        self.Controls.Add(detail_lbl); y += 20

        detail = TextBox()
        detail.Multiline   = True; detail.ReadOnly = True; detail.ScrollBars = ScrollBars.Vertical
        detail.Location    = Point(12, y); detail.Size = Size(500, 100)
        detail.Font        = DFont("Courier New", 8)
        detail.Text        = "\r\n".join(new_types[:200]) + ("\r\n..." if len(new_types) > 200 else "")
        self.Controls.Add(detail); y += 110

        btnOK     = Button(); btnOK.Text     = "Apply";  btnOK.Size = Size(90, 28)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(90, 28)
        btnOK.Location     = Point(self.ClientSize.Width - 196, y)
        btnCancel.Location = Point(self.ClientSize.Width - 102, y)
        btnOK.Click     += lambda s, a: self._set_result(DialogResult.OK)
        btnCancel.Click += lambda s, a: self._set_result(DialogResult.Cancel)
        for b in [lbl_title, btnOK, btnCancel]: self.Controls.Add(b)

    def _set_result(self, r):
        self.DialogResult = r
        self.Close()

    def get_options(self):
        return {
            "update_existing": bool(self.chkUpdate.Checked),
            "create_new":      bool(self.chkCreate.Checked),
            "delete_missing":  bool(self.chkDelete.Checked),
        }


# =============================================================================
# ADD PARAMETER DIALOG
# =============================================================================

class AddParamDialog(Form):
    """
    Add Family Parameter dialog matching Revit's native layout:
    Name / Discipline / Data Type / Group / Instance+Type / Formula
    """

    def __init__(self):
        Form.__init__(self)
        self.Text            = "Add Family Parameter"
        self.Size            = Size(460, 340)
        self.StartPosition   = FormStartPosition.CenterParent
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox     = False

        self.result_name       = ""
        self.result_type_token = "Length"
        self.result_group      = "Construction"
        self.result_inst       = False
        self.result_formula    = ""
        self._group_label_to_token = {}
        self._dtype_rows       = []   # list of (label, spec_attr, pt_attr)

        def _lbl(text, x, y, w=140):
            l = Label(); l.Text = text; l.Location = Point(x, y); l.Size = Size(w, 18)
            l.Font = DFont("Arial", 9); self.Controls.Add(l)

        # Name
        _lbl("Name:", 12, 12)
        self.txtName = TextBox()
        self.txtName.Location = Point(12, 30); self.txtName.Size = Size(420, 22)
        self.Controls.Add(self.txtName)

        # Discipline
        _lbl("Discipline:", 12, 62)
        self.cboDiscipline = ComboBox()
        self.cboDiscipline.Location = Point(12, 80); self.cboDiscipline.Size = Size(190, 22)
        self.cboDiscipline.DropDownStyle = ComboBoxStyle.DropDownList
        for d in core.PARAM_DISCIPLINES:
            self.cboDiscipline.Items.Add(d)
        self.cboDiscipline.SelectedIndex = 0
        self.cboDiscipline.SelectedIndexChanged += self._on_discipline_changed
        self.Controls.Add(self.cboDiscipline)

        # Data Type
        _lbl("Data Type:", 216, 62)
        self.cboDataType = ComboBox()
        self.cboDataType.Location = Point(216, 80); self.cboDataType.Size = Size(216, 22)
        self.cboDataType.DropDownStyle = ComboBoxStyle.DropDownList
        self.Controls.Add(self.cboDataType)

        # Group
        _lbl("Group parameter under:", 12, 114)
        self.cboGroup = ComboBox()
        self.cboGroup.Location = Point(12, 132); self.cboGroup.Size = Size(420, 22)
        self.cboGroup.DropDownStyle = ComboBoxStyle.DropDownList
        try:
            import binder_core as bc
            _doc = __revit__.ActiveUIDocument.Document
            for (lab, tok) in bc.list_user_assignable_group_choices(_doc):
                self.cboGroup.Items.Add(lab)
                self._group_label_to_token[lab] = tok
            if self.cboGroup.Items.Count == 0:
                raise Exception("empty")
        except:
            fallback = [
                ("Analysis Results","AnalysisResults"),("Constraints","Constraints"),
                ("Construction","Construction"),("Data","Data"),
                ("Dimensions","Geometry"),("Electrical","Electrical"),
                ("Forces","Forces"),("General","General"),("Graphics","Graphics"),
                ("Identity Data","IdentityData"),("IFC Parameters","Ifc"),
                ("Life Safety","LifeSafety"),("Materials and Finishes","Materials"),
                ("Mechanical","Mechanical"),("Mechanical - Flow","MechanicalAirflow"),
                ("Mechanical - Loads","MechanicalLoads"),("Other","OTHER"),
                ("Phasing","Phasing"),("Plumbing","Plumbing"),
                ("Structural","Structural"),("Structural Analysis","StructuralAnalysis"),
                ("Text","Text"),("Title Text","Title"),("Visibility","Visibility"),
            ]
            for lab, tok in fallback:
                self.cboGroup.Items.Add(lab)
                self._group_label_to_token[lab] = tok
        for i in range(self.cboGroup.Items.Count):
            if "Dimensions" in str(self.cboGroup.Items[i]) or "Construction" in str(self.cboGroup.Items[i]):
                self.cboGroup.SelectedIndex = i; break
        if self.cboGroup.SelectedIndex < 0 and self.cboGroup.Items.Count > 0:
            self.cboGroup.SelectedIndex = 0
        self.Controls.Add(self.cboGroup)

        # Instance / Type radio buttons
        self.rbType = RadioButton(); self.rbType.Text = "Type"
        self.rbType.Location = Point(12, 168); self.rbType.Size = Size(80, 22); self.rbType.Checked = True
        self.rbInst = RadioButton(); self.rbInst.Text = "Instance"
        self.rbInst.Location = Point(100, 168); self.rbInst.Size = Size(90, 22)
        self.Controls.Add(self.rbType); self.Controls.Add(self.rbInst)

        # Formula
        _lbl("Formula (optional):", 12, 200)
        self.txtFormula = TextBox()
        self.txtFormula.Location = Point(12, 218); self.txtFormula.Size = Size(420, 22)
        self.txtFormula.Font = DFont("Courier New", 9)
        self.Controls.Add(self.txtFormula)

        btnOK     = Button(); btnOK.Text = "Add";    btnOK.Size = Size(90, 28); btnOK.Location = Point(240, 258)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(90, 28); btnCancel.Location = Point(340, 258)
        btnOK.Click     += self._ok
        btnCancel.Click += lambda s, a: self.Close()
        self.Controls.Add(btnOK); self.Controls.Add(btnCancel)

        # Populate data types for default discipline
        self._on_discipline_changed(None, None)

    def _on_discipline_changed(self, s, a):
        """Repopulate Data Type list when discipline changes."""
        disc = str(self.cboDiscipline.SelectedItem or "Common")
        self._dtype_rows = core.get_data_types_for_discipline(disc)
        prev = str(self.cboDataType.SelectedItem or "")
        self.cboDataType.Items.Clear()
        for (label, spec_attr, pt_attr) in self._dtype_rows:
            self.cboDataType.Items.Add(label)
        # Try to keep same selection, default to Length/first
        restored = False
        for i in range(self.cboDataType.Items.Count):
            if str(self.cboDataType.Items[i]) == prev:
                self.cboDataType.SelectedIndex = i; restored = True; break
        if not restored:
            for i in range(self.cboDataType.Items.Count):
                if str(self.cboDataType.Items[i]) in ("Length", "Number"):
                    self.cboDataType.SelectedIndex = i; break
            if self.cboDataType.SelectedIndex < 0 and self.cboDataType.Items.Count > 0:
                self.cboDataType.SelectedIndex = 0

    def _ok(self, s, a):
        name = (self.txtName.Text or "").strip()
        if not name:
            MessageBox.Show("Enter a parameter name.", TOOL_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        # Resolve data type token from selected label
        sel_idx = self.cboDataType.SelectedIndex
        if sel_idx >= 0 and sel_idx < len(self._dtype_rows):
            label, spec_attr, pt_attr = self._dtype_rows[sel_idx]
            # Use spec_attr as token — maps directly in _PARAM_TYPE_MAP
            self.result_type_token = spec_attr
        else:
            self.result_type_token = "Length"
        selected_label      = str(self.cboGroup.SelectedItem or "Construction")
        self.result_name    = name
        self.result_group   = self._group_label_to_token.get(selected_label, selected_label)
        self.result_inst    = bool(self.rbInst.Checked)
        self.result_formula = (self.txtFormula.Text or "").strip()
        self.DialogResult   = DialogResult.OK
        self.Close()


# =============================================================================
# MAIN FORM
# =============================================================================

class TypeManagerForm(Form):
    """
    Main RFA Type Manager window.
    Tabs: Types Grid | Parameters | SP Profiles
    Bottom toolbar: Export / Import / Apply / status
    """

    def __init__(self, doc):
        Form.__init__(self)
        self.doc         = doc
        self._snapshot   = None   # last snapshot loaded from Revit
        self._dirty      = False  # True if grid has unsaved edits
        self._unit_mode  = "inches"
        self._last_export_params = None  # remember column selection

        family_name = ""
        try:
            family_name = doc.Title.replace(".rfa", "")
        except:
            pass

        self.Text            = "{} — {}".format(TOOL_TITLE, family_name)
        self.Size            = Size(1300, 820)
        self.MinimumSize     = Size(900, 600)
        self.StartPosition   = FormStartPosition.CenterScreen

        # ------------------------------------------------------------------ #
        # TOP TOOLBAR                                                          #
        # ------------------------------------------------------------------ #
        self.toolbar = ToolStrip()
        self.toolbar.Dock = DockStyle.Top

        def _tbtn(text, tooltip, handler, width=None):
            b = ToolStripButton(text)
            b.ToolTipText = tooltip
            if handler is not None:
                b.Click += handler
            if width:
                b.Width = width
            return b

        self.btnRefresh  = _tbtn("⟳ Refresh",       "Reload types from Revit",             self._on_refresh)
        self.btnExport   = _tbtn("↓ Export Excel",   "Export types to Excel for editing",   self._on_export, 110)
        self.btnImport   = _tbtn("↑ Import Excel",   "Import edited Excel back into family", self._on_import, 110)
        self.btnApply    = _tbtn("✔ Apply Changes",  "Write grid edits back to Revit",      self._on_apply,  110)

        sep1 = ToolStripSeparator()
        sep2 = ToolStripSeparator()

        lbl_unit = ToolStripLabel("Units:")
        self.cboUnit = ToolStripComboBox()
        self.cboUnit.Items.Add("Decimal Inches")
        self.cboUnit.Items.Add("Decimal Feet")
        self.cboUnit.Items.Add("Fractional Inches")
        self.cboUnit.SelectedIndex = 0
        self.cboUnit.Width = 130
        self.cboUnit.ToolTipText   = "Display unit for length parameters"
        self.cboUnit.SelectedIndexChanged += self._on_unit_changed

        sep3 = ToolStripSeparator()
        self.btnToggleFormulas = ToolStripButton("Show Formulas")
        self.btnToggleFormulas.ToolTipText = "Toggle formula columns on/off in the grid"
        self.btnToggleFormulas.CheckOnClick = True
        self.btnToggleFormulas.Checked      = True
        self.btnToggleFormulas.CheckedChanged += self._on_toggle_formulas

        sep4 = ToolStripSeparator()
        btnDiag = ToolStripButton("Diagnose")
        btnDiag.ToolTipText = "Show debug info about parameter reading"
        btnDiag.Click += self._on_diagnose

        for item in [self.btnRefresh, sep1, self.btnExport, self.btnImport,
                     self.btnApply, sep2, lbl_unit, self.cboUnit,
                     sep3, self.btnToggleFormulas, sep4, btnDiag]:
            self.toolbar.Items.Add(item)

        # ------------------------------------------------------------------ #
        # TAB CONTROL                                                          #
        # ------------------------------------------------------------------ #
        self.tabs = TabControl()
        self.tabs.Dock = DockStyle.Fill

        self.tab_grid    = TabPage("Types")
        self.tab_params  = TabPage("Parameters")
        self.tab_profiles = TabPage("SP Profiles")

        self.tabs.TabPages.Add(self.tab_grid)
        self.tabs.TabPages.Add(self.tab_params)
        self.tabs.TabPages.Add(self.tab_profiles)

        # ------------------------------------------------------------------ #
        # STATUS BAR                                                           #
        # ------------------------------------------------------------------ #
        self.statusbar = StatusStrip()
        self.statusbar.Dock = DockStyle.Bottom
        self.lblStatus = ToolStripStatusLabel("Ready")
        self.lblStatus.Spring = True
        self.lblTypeCount = ToolStripStatusLabel("")
        self.statusbar.Items.Add(self.lblStatus)
        self.statusbar.Items.Add(self.lblTypeCount)

        # ------------------------------------------------------------------ #
        # ASSEMBLE MAIN LAYOUT                                                 #
        # ------------------------------------------------------------------ #
        self.Controls.Add(self.tabs)
        self.Controls.Add(self.toolbar)
        self.Controls.Add(self.statusbar)

        # Build each tab
        self._build_tab_grid()
        self._build_tab_params()
        self._build_tab_profiles()

        # Load initial snapshot
        self._load_snapshot()

    # ======================================================================= #
    # GRID TAB                                                                 #
    # ======================================================================= #

    def _build_tab_grid(self):
        """Builds the Types DataGridView tab."""
        panel = Panel(); panel.Dock = DockStyle.Fill

        # --- Mini toolbar above grid ---
        mini = ToolStrip(); mini.Dock = DockStyle.Top; mini.GripStyle = ToolStripGripStyle.Hidden

        def _mbtn(text, tip, handler):
            b = ToolStripButton(text)
            b.ToolTipText = tip
            if handler is not None:
                b.Click += handler
            return b

        btn_add_row  = _mbtn("+ Row",       "Add a new empty type row",        self._on_add_row)
        btn_dup_row  = _mbtn("⊕ Duplicate", "Duplicate selected row(s)",       self._on_dup_rows)
        btn_del_row  = _mbtn("✕ Delete",    "Delete selected row(s)",          self._on_del_rows)
        btn_fill_dn  = _mbtn("↓ Fill Down", "Fill selected cell value down",   self._on_fill_down)
        btn_sort     = _mbtn("A↓ Sort",     "Sort types alphabetically",       self._on_sort)

        sep_g = ToolStripSeparator()

        # Fill Series button
        btn_fill_series = _mbtn("∑ Series", "Fill selected column as arithmetic series", self._on_fill_series)

        # Column visibility button
        btn_cols = _mbtn("⊞ Columns", "Show/hide/reorder columns", self._on_manage_columns)

        for item in [btn_add_row, btn_dup_row, btn_del_row, sep_g,
                     btn_fill_dn, btn_sort, btn_fill_series, btn_cols]:
            mini.Items.Add(item)

        # Search bar as a proper Panel (more reliable than ToolStripControlHost)
        search_panel = Panel()
        search_panel.Dock = DockStyle.Top
        search_panel.Height = 26
        search_panel.BackColor = Color.FromArgb(240, 240, 240)

        lbl_s = Label()
        lbl_s.Text = "Search:"
        lbl_s.Location = Point(6, 5); lbl_s.Size = Size(46, 18)
        lbl_s.Font = DFont("Arial", 9)

        self.txtGridFilter = TextBox()
        self.txtGridFilter.Location = Point(56, 3)
        self.txtGridFilter.Size     = Size(260, 20)
        self.txtGridFilter.Font     = DFont("Arial", 9)
        self.txtGridFilter.TextChanged += self._on_grid_filter

        lbl_hint = Label()
        lbl_hint.Text = "(filters type names)"
        lbl_hint.Location = Point(324, 5); lbl_hint.Size = Size(140, 18)
        lbl_hint.Font = DFont("Arial", 8); lbl_hint.ForeColor = Color.Gray

        search_panel.Controls.Add(lbl_s)
        search_panel.Controls.Add(self.txtGridFilter)
        search_panel.Controls.Add(lbl_hint)

        # --- DataGridView ---
        self.grid = DataGridView()
        self.grid.Dock                    = DockStyle.Fill
        self.grid.AllowUserToAddRows      = True
        self.grid.AllowUserToDeleteRows   = True
        self.grid.MultiSelect             = True
        self.grid.SelectionMode           = DataGridViewSelectionMode.RowHeaderSelect
        self.grid.AutoSizeColumnsMode     = getattr(DataGridViewAutoSizeColumnsMode, "None")
        self.grid.RowHeadersWidth              = 24
        self.grid.ColumnHeadersHeightSizeMode  = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self.grid.ColumnHeadersHeight          = 26
        self.grid.AllowUserToOrderColumns      = True   # drag column headers to reorder
        self.grid.AllowUserToResizeColumns     = True   # drag column edges to resize
        self.grid.DefaultCellStyle.Font   = DFont("Arial", 9)
        self.grid.ColumnHeadersDefaultCellStyle.Font       = DFont("Arial", 9, FontStyle.Bold)
        self.grid.ColumnHeadersDefaultCellStyle.BackColor  = COLOR_HEADER_BG
        self.grid.ColumnHeadersDefaultCellStyle.ForeColor  = COLOR_HEADER_FG
        self.grid.RowTemplate.Height      = 18
        self.grid.ClipboardCopyMode       = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
        self.grid.EnableHeadersVisualStyles = False

        # Context menu for right-click operations
        ctx = ContextMenuStrip()

        def _mitem(text, handler):
            m = ToolStripMenuItem(text); m.Click += handler; return m

        ctx.Items.Add(_mitem("Fill Down",                    self._on_fill_down))
        ctx.Items.Add(_mitem("Fill Up",                      self._on_fill_up))
        ctx.Items.Add(ToolStripSeparator())
        ctx.Items.Add(_mitem("Duplicate Row(s)",             self._on_dup_rows))
        ctx.Items.Add(_mitem("Delete Row(s)",                self._on_del_rows))
        ctx.Items.Add(ToolStripSeparator())
        ctx.Items.Add(_mitem("Copy",                         self._on_grid_copy))
        ctx.Items.Add(_mitem("Paste",                        self._on_grid_paste))
        ctx.Items.Add(_mitem("Paste to Selected Rows (same value)", self._on_paste_broadcast))
        ctx.Items.Add(ToolStripSeparator())
        ctx.Items.Add(_mitem("Fill Series...",                self._on_fill_series))
        ctx.Items.Add(_mitem("Fill Right (Ctrl+R)",           self._on_fill_right))
        ctx.Items.Add(ToolStripSeparator())
        ctx.Items.Add(_mitem("Hide This Column",              self._on_hide_col))
        ctx.Items.Add(_mitem("Show/Hide Columns...",          self._on_manage_columns))
        self.grid.ContextMenuStrip = ctx

        # Column header right-click for hide/show
        col_ctx = ContextMenuStrip()
        col_ctx.Items.Add(_mitem("Hide This Column",          self._on_hide_col_hdr))
        col_ctx.Items.Add(_mitem("Show/Hide Columns...",      self._on_manage_columns))
        col_ctx.Items.Add(_mitem("Move Left",                 self._on_col_move_left))
        col_ctx.Items.Add(_mitem("Move Right",                self._on_col_move_right))
        self._col_ctx = col_ctx
        self.grid.ColumnHeaderMouseClick += self._on_col_header_click
        # Allow left-click to sort (default), right-click handled in _on_col_header_click

        # Wire events
        self.grid.CellValueChanged    += self._on_cell_changed
        self.grid.KeyDown             += self._on_grid_keydown
        self.grid.CellBeginEdit       += self._on_cell_begin_edit
        self.grid.CellDoubleClick     += self._on_cell_double_click
        self.grid.DefaultValuesNeeded += self._on_default_values

        panel.Controls.Add(self.grid)
        panel.Controls.Add(search_panel)
        panel.Controls.Add(mini)
        self.tab_grid.Controls.Add(panel)

    # ======================================================================= #
    # PARAMETERS TAB                                                           #
    # ======================================================================= #

    def _build_tab_params(self):
        """Builds the Parameters management tab."""
        split = SplitContainer()
        split.Dock             = DockStyle.Fill
        split.SplitterDistance = 460
        split.Orientation      = Orientation.Vertical

        # LEFT: parameter list
        lbl = Label(); lbl.Text = "Family Parameters"; lbl.Dock = DockStyle.Top
        lbl.Font = DFont("Arial", 9, FontStyle.Bold); lbl.Height = 20
        lbl.Padding = Padding(4, 4, 0, 0)

        self.lvParams = ListView()
        self.lvParams.Dock         = DockStyle.Fill
        self.lvParams.View         = View.Details
        self.lvParams.FullRowSelect = True
        self.lvParams.MultiSelect  = True
        for t, w in [("Name", 220), ("Type", 80), ("I/T", 40), ("Shared", 55), ("Formula", 180)]:
            ch = ColumnHeader(); ch.Text = t; ch.Width = w
            self.lvParams.Columns.Add(ch)

        split.Panel1.Controls.Add(self.lvParams)
        split.Panel1.Controls.Add(lbl)

        # RIGHT: action buttons
        right = Panel(); right.Dock = DockStyle.Fill; right.Padding = Padding(12)

        def _btn(text, tip, handler, y):
            b = Button(); b.Text = text; b.Size = Size(200, 30)
            b.Location = Point(12, y); b.Font = DFont("Arial", 9)
            b.Click += handler
            return b

        y = 12
        btnAddFam   = _btn("+ Add Family Parameter",     "Add a new non-shared parameter", self._on_add_fam_param, y); y += 38
        btnRemove   = _btn("✕ Remove Selected",          "Remove selected parameter(s)",   self._on_remove_param,  y); y += 38
        sep = Label(); sep.Text = ""; sep.Size = Size(200, 10); sep.Location = Point(12, y); y += 18
        btnFormula  = _btn("ƒ Set Formula...",            "Set or clear formula",           self._on_set_formula,   y); y += 38

        lbl_note = Label()
        lbl_note.Text = (
            "To add SHARED parameters,\nuse the SP Profiles tab\nor the SP Loader button."
        )
        lbl_note.Location = Point(12, y + 20)
        lbl_note.Size     = Size(220, 60)
        lbl_note.Font     = DFont("Arial", 8)
        lbl_note.ForeColor = Color.Gray

        for w in [btnAddFam, btnRemove, sep, btnFormula, lbl_note]:
            right.Controls.Add(w)

        split.Panel2.Controls.Add(right)
        self.tab_params.Controls.Add(split)

    # ======================================================================= #
    # SP PROFILES TAB                                                          #
    # ======================================================================= #

    def _build_tab_profiles(self):
        """Builds the SP Profiles tab (reuses binder_core config scanning)."""
        panel = Panel(); panel.Dock = DockStyle.Fill; panel.Padding = Padding(8)

        lbl = Label()
        lbl.Text = "Install a saved SP Loader profile into this family. The profile's shared parameters will be added."
        lbl.Location = Point(8, 8); lbl.Size = Size(860, 32)
        lbl.Font = DFont("Arial", 9)

        self.lvProfiles = ListView()
        self.lvProfiles.View         = View.Details
        self.lvProfiles.FullRowSelect = True
        self.lvProfiles.Location     = Point(8, 46)
        self.lvProfiles.Size         = Size(860, 500)
        self.lvProfiles.Anchor       = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        for t, w in [("Display Name", 320), ("Scope", 80), ("File", 280), ("Modified", 150)]:
            ch = ColumnHeader(); ch.Text = t; ch.Width = w
            self.lvProfiles.Columns.Add(ch)

        btnInstall  = Button(); btnInstall.Text  = "Install Selected Profile"
        btnInstall.Size = Size(180, 30)
        btnInstall.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btnInstall.Click += self._on_install_profile

        btnRefreshP = Button(); btnRefreshP.Text = "Refresh"
        btnRefreshP.Size = Size(80, 30)
        btnRefreshP.Anchor = AnchorStyles.Bottom | AnchorStyles.Left
        btnRefreshP.Click += self._on_refresh_profiles

        btnOpenSPLoader = Button(); btnOpenSPLoader.Text = "Open SP Loader (Create/Edit Profiles)"
        btnOpenSPLoader.Size = Size(260, 30)
        btnOpenSPLoader.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnOpenSPLoader.Click += self._on_open_sp_loader

        def _layout_p(s=None, a=None):
            y = panel.ClientSize.Height - 42
            self.lvProfiles.Height = y - 54
            btnInstall.Location      = Point(8, y)
            btnRefreshP.Location     = Point(196, y)
            btnOpenSPLoader.Location = Point(panel.ClientSize.Width - 268, y)
        panel.Resize += _layout_p
        _layout_p()

        for w in [lbl, self.lvProfiles, btnInstall, btnRefreshP, btnOpenSPLoader]:
            panel.Controls.Add(w)

        self.tab_profiles.Controls.Add(panel)
        self._reload_profiles()

    # ======================================================================= #
    # DATA LOADING                                                              #
    # ======================================================================= #

    def _load_snapshot(self):
        """Reads current family state from Revit and populates all tabs."""
        self._set_status("Loading family types...")
        try:
            self._snapshot = core.snapshot_family(self.doc)
            self._populate_grid()
            self._populate_params()
            n_types  = len(self._snapshot.get("types", []))
            n_params = len(self._snapshot.get("parameters", []))
            self.lblTypeCount.Text = "{} types  |  {} params".format(n_types, n_params)
            self._dirty = False
            self._set_status("Loaded.")
        except Exception as e:
            self._set_status("Load error: " + str(e))
            TaskDialog.Show(TOOL_TITLE, "Failed to load family:\n" + str(e))

    def _populate_grid(self):
        """Rebuilds the DataGridView columns and rows from the current snapshot."""
        self.grid.SuspendLayout()
        self.grid.Columns.Clear()
        self.grid.Rows.Clear()

        params   = self._snapshot.get("parameters", [])
        types    = self._snapshot.get("types", [])

        # Build column list: Type Name first, then params
        # Store column metadata as tag on each column
        col = DataGridViewTextBoxColumn()
        col.Name         = "__type_name__"
        col.HeaderText   = "Type Name"
        col.Width        = 180
        col.Frozen       = True  # Type Name always visible when scrolling
        col.Tag          = None
        self.grid.Columns.Add(col)

        for p in params:
            c = DataGridViewTextBoxColumn()
            c.Name       = p["name"]
            st           = p.get("storage_type", "String")
            unit_hint    = ""
            if st == "Double" and self._unit_mode != "feet":
                unit_hint = " (in)"
            elif st == "Double":
                unit_hint = " (ft)"
            c.HeaderText = p["name"] + unit_hint
            c.Width      = max(90, len(p["name"]) * 8 + 20)
            c.Tag        = p  # full param def stored on column

            if p.get("has_formula") or p.get("read_only"):
                c.ReadOnly                    = False
                c.DefaultCellStyle.BackColor  = COLOR_FORMULA_BG
                c.DefaultCellStyle.ForeColor  = Color.FromArgb(80, 80, 140)
                c.DefaultCellStyle.Font       = DFont("Arial", 8, FontStyle.Italic)
                c.ToolTipText                 = "[formula param — double-click to edit formula]"
                # Respect current toggle state
                try:
                    c.Visible = bool(self.btnToggleFormulas.Checked)
                except:
                    pass

            self.grid.Columns.Add(c)

        # Populate rows
        scale = 12.0 if self._unit_mode == "inches" else 1.0
        for type_row in types:
            row_data = [type_row["name"]]
            values   = type_row.get("values", {})
            for p in params:
                name = p["name"]
                val  = values.get(name)
                if p.get("has_formula") or p.get("read_only"):
                    # Show the formula expression string so user can read/edit it
                    row_data.append("=" + p.get("formula_expr", "") if p.get("formula_expr") else "")
                elif val is None:
                    row_data.append("")
                elif p.get("storage_type") == "Double":
                    try:
                        fv = float(val)
                        if self._unit_mode == "frac_inches":
                            row_data.append(_to_frac_inches(fv))
                        else:
                            row_data.append("{:.4f}".format(fv * scale))
                    except:
                        row_data.append("")
                else:
                    row_data.append(str(val) if val is not None else "")
            self.grid.Rows.Add(Array[object](row_data))

        # Color alternating rows
        for i, row in enumerate(self.grid.Rows):
            try:
                if i % 2 == 1:
                    row.DefaultCellStyle.BackColor = COLOR_ALT_ROW
            except:
                pass

        self.grid.ResumeLayout()

    def _populate_params(self):
        """Rebuilds the parameter ListView from the current snapshot."""
        self.lvParams.Items.Clear()
        for p in self._snapshot.get("parameters", []):
            inst_str    = "Inst" if p.get("is_instance") else "Type"
            shared_str  = "Yes"  if p.get("is_shared")   else "No"
            formula_str = "[formula]" if p.get("has_formula") else ""
            sub = Array[String]([p["name"], p.get("storage_type",""), inst_str, shared_str, formula_str])
            item = ListViewItem(sub)
            if p.get("has_formula"):
                item.ForeColor = Color.Gray
                item.Font      = DFont("Arial", 9, FontStyle.Italic)
            self.lvParams.Items.Add(item)

    # ======================================================================= #
    # GRID HELPERS                                                              #
    # ======================================================================= #

    def _on_grid_filter(self, s, a):
        """Show/hide rows based on Type Name filter."""
        needle = (self.txtGridFilter.Text or "").strip().lower()
        for row in self.grid.Rows:
            try:
                if row.IsNewRow:
                    continue
                name = str(row.Cells[0].Value or "").lower()
                row.Visible = (not needle) or (needle in name)
            except:
                pass

    def _on_cell_changed(self, s, e):
        """Mark row as modified when a cell value changes."""
        try:
            row = self.grid.Rows[e.RowIndex]
            row.DefaultCellStyle.BackColor = COLOR_MODIFIED
            self._dirty = True
        except:
            pass

    def _on_cell_begin_edit(self, s, e):
        """Column index 0 (Type Name) is always editable. Formula cols handled via double-click."""
        pass  # no blocking — all cols editable inline

    def _on_cell_double_click(self, s, e):
        """
        Double-clicking a formula-column cell opens the formula editor dialog
        so the user can read and modify the formula expression directly.
        The same formula applies to ALL types (family params share one formula),
        so we open a single editor and apply it to the parameter globally.
        """
        try:
            if e.ColumnIndex < 1 or e.RowIndex < 0:
                return
            col = self.grid.Columns[e.ColumnIndex]
            p   = col.Tag
            if not p or not (p.get("has_formula") or p.get("read_only")):
                return  # not a formula col — normal edit handles it

            param_name = p["name"]
            current_expr = p.get("formula_expr", "")

            # Simple formula edit dialog
            frm = Form()
            frm.Text            = "Edit Formula: " + param_name
            frm.Size            = Size(520, 160)
            frm.StartPosition   = FormStartPosition.CenterParent
            frm.FormBorderStyle = FormBorderStyle.FixedDialog
            frm.MaximizeBox     = False

            lbl = Label()
            lbl.Text     = "Formula for '" + param_name + "' (applies to all types):"
            lbl.Location = Point(12, 12); lbl.Size = Size(490, 18)
            lbl.Font     = DFont("Arial", 9)

            txt = TextBox()
            txt.Location = Point(12, 34); txt.Size = Size(490, 22)
            txt.Text     = current_expr
            txt.Font     = DFont("Courier New", 9)

            btnOK     = Button(); btnOK.Text = "Apply";   btnOK.Size = Size(90, 26)
            btnClear  = Button(); btnClear.Text = "Clear Formula"; btnClear.Size = Size(110, 26)
            btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(80, 26)
            btnOK.Location     = Point(230, 72)
            btnClear.Location  = Point(328, 72)
            btnCancel.Location = Point(446, 72)

            def _apply(x, y):
                frm.Tag = ("apply", (txt.Text or "").strip())
                frm.DialogResult = DialogResult.OK
                frm.Close()
            def _clear(x, y):
                frm.Tag = ("clear", "")
                frm.DialogResult = DialogResult.OK
                frm.Close()
            btnOK.Click     += _apply
            btnClear.Click  += _clear
            btnCancel.Click += lambda x, y: frm.Close()

            for w in [lbl, txt, btnOK, btnClear, btnCancel]:
                frm.Controls.Add(w)

            if frm.ShowDialog() != DialogResult.OK or frm.Tag is None:
                return

            action, new_expr = frm.Tag
            try:
                core.set_formula(self.doc, param_name, new_expr)
                self._load_snapshot()  # reload to reflect the change
            except Exception as ex:
                TaskDialog.Show(TOOL_TITLE, "Formula error: " + str(ex))
        except Exception as ex:
            TaskDialog.Show(TOOL_TITLE, "Double-click error: " + str(ex))

    def _on_default_values(self, s, e):
        """Set sensible defaults for new rows."""
        try:
            e.Row.Cells[0].Value = "New Type"
        except:
            pass

    def _on_grid_keydown(self, s, e):
        """Keyboard shortcuts: Del = clear cell, Ctrl+D = fill down, Ctrl+C/V = copy/paste."""
        if e.Control and e.KeyCode == Keys.D:
            self._on_fill_down(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.R:
            self._on_fill_right(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.C:
            self._on_grid_copy(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.V:
            self._on_grid_paste(None, None)
            e.Handled = True

    def _selected_row_indices(self):
        """Returns sorted list of selected non-new-row indices."""
        indices = []
        for row in self.grid.SelectedRows:
            if not row.IsNewRow:
                indices.append(row.Index)
        indices.sort()
        return indices

    def _on_add_row(self, s, a):
        """Adds a blank row at the bottom of the grid."""
        self.grid.Rows.Add()
        last = self.grid.Rows.Count - 2  # -2 because of the "new row" placeholder
        if last >= 0:
            self.grid.CurrentCell = self.grid.Rows[last].Cells[0]
        self._dirty = True

    def _on_dup_rows(self, s, a):
        """Duplicates selected rows, appending them at the bottom."""
        indices = self._selected_row_indices()
        if not indices:
            return
        for idx in indices:
            src = self.grid.Rows[idx]
            vals = [src.Cells[c].Value for c in range(self.grid.Columns.Count)]
            # Modify type name to avoid collision
            vals[0] = (str(vals[0] or "") + " Copy") if vals[0] else "Copy"
            self.grid.Rows.Add(Array[object](vals))
        self._dirty = True

    def _on_del_rows(self, s, a):
        """Deletes selected rows after confirmation."""
        indices = self._selected_row_indices()
        if not indices:
            return
        if MessageBox.Show(
            "Delete {} selected type(s)? This will be applied to Revit on Apply.".format(len(indices)),
            TOOL_TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning
        ).ToString() != "OK":
            return
        for idx in sorted(indices, reverse=True):
            try:
                self.grid.Rows.RemoveAt(idx)
            except:
                pass
        self._dirty = True

    def _on_fill_down(self, s, a):
        """
        Fills the value of the topmost selected cell downward through all
        selected rows in the same column.
        """
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        # Get the anchor cell (topmost row, leftmost column in selection)
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex))
        anchor    = sel_cells[0]
        value     = anchor.Value
        col_idx   = anchor.ColumnIndex
        row_indices = sorted(set(c.RowIndex for c in sel_cells))
        for ri in row_indices:
            try:
                cell = self.grid.Rows[ri].Cells[col_idx]
                if not cell.ReadOnly:
                    cell.Value = value
            except:
                pass
        self._dirty = True

    def _on_fill_up(self, s, a):
        """Fills the value of the bottommost selected cell upward."""
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex), reverse=True)
        anchor    = sel_cells[0]
        value     = anchor.Value
        col_idx   = anchor.ColumnIndex
        row_indices = sorted(set(c.RowIndex for c in sel_cells), reverse=True)
        for ri in row_indices:
            try:
                cell = self.grid.Rows[ri].Cells[col_idx]
                if not cell.ReadOnly:
                    cell.Value = value
            except:
                pass
        self._dirty = True

    def _on_sort(self, s, a):
        """Sorts all rows alphabetically by Type Name."""
        rows_data = []
        for row in self.grid.Rows:
            if row.IsNewRow:
                continue
            vals = [row.Cells[c].Value for c in range(self.grid.Columns.Count)]
            rows_data.append(vals)
        rows_data.sort(key=lambda r: str(r[0] or "").lower())
        self.grid.Rows.Clear()
        for vals in rows_data:
            self.grid.Rows.Add(Array[object](vals))

    def _on_grid_copy(self, s, a):
        """Copies selected cells to clipboard in tab-delimited format."""
        try:
            sel_cells = sorted(self.grid.SelectedCells, key=lambda c: (c.RowIndex, c.ColumnIndex))
            if not sel_cells:
                return
            rows = {}
            for cell in sel_cells:
                rows.setdefault(cell.RowIndex, {})[cell.ColumnIndex] = str(cell.Value or "")
            lines = []
            for ri in sorted(rows.keys()):
                row = rows[ri]
                lines.append("\t".join(row[ci] for ci in sorted(row.keys())))
            Clipboard.SetText("\r\n".join(lines))
        except:
            pass

    def _on_grid_paste(self, s, a):
        """
        Pastes clipboard tab-delimited data starting at the currently
        selected cell, extending downward and rightward.
        Excel-style: each clipboard row maps to a grid row, each tab = next column.
        """
        try:
            text = Clipboard.GetText()
            if not text:
                return
            lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            sel   = sorted(self.grid.SelectedCells, key=lambda c: (c.RowIndex, c.ColumnIndex))
            if not sel:
                return
            start_row = sel[0].RowIndex
            start_col = sel[0].ColumnIndex
            for r_off, line in enumerate(lines):
                if not line:
                    continue
                values = line.split("\t")
                row_idx = start_row + r_off
                # Add rows if needed
                while row_idx >= self.grid.Rows.Count - 1:
                    self.grid.Rows.Add()
                for c_off, val in enumerate(values):
                    col_idx = start_col + c_off
                    if col_idx >= self.grid.Columns.Count:
                        break
                    cell = self.grid.Rows[row_idx].Cells[col_idx]
                    if not cell.ReadOnly:
                        cell.Value = val
            self._dirty = True
        except:
            pass

    def _on_paste_broadcast(self, s, a):
        """
        Pastes the clipboard value (single cell) into ALL selected cells
        that share the same column as the first selected cell.
        Useful for setting the same value across many types at once.
        """
        try:
            text = (Clipboard.GetText() or "").split("\t")[0].split("\n")[0].strip()
            sel_cells = list(self.grid.SelectedCells)
            if not sel_cells:
                return
            col_idx = sorted(sel_cells, key=lambda c: c.ColumnIndex)[0].ColumnIndex
            for cell in sel_cells:
                if cell.ColumnIndex == col_idx and not cell.ReadOnly:
                    cell.Value = text
            self._dirty = True
        except:
            pass

    # ======================================================================= #
    # TOOLBAR ACTIONS                                                           #
    # ======================================================================= #

    def _on_refresh(self, s, a):
        if self._dirty:
            if MessageBox.Show(
                "You have unsaved changes. Refresh will discard them. Continue?",
                TOOL_TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning
            ).ToString() != "OK":
                return
        self._load_snapshot()

    def _on_unit_changed(self, s, a):
        """Toggle display unit and rebuild the grid columns."""
        idx = self.cboUnit.SelectedIndex
        if idx == 0:
            self._unit_mode = "inches"
        elif idx == 1:
            self._unit_mode = "feet"
        else:
            self._unit_mode = "frac_inches"
        if self._snapshot:
            self._populate_grid()

    def _on_toggle_formulas(self, s, a):
        """Show or hide formula-driven columns in the grid."""
        show = bool(self.btnToggleFormulas.Checked)
        for col in self.grid.Columns:
            try:
                p = col.Tag
                if p and (p.get("has_formula") or p.get("read_only")):
                    col.Visible = show
            except:
                pass
        self.btnToggleFormulas.Text = "Hide Formulas" if show else "Show Formulas"

    def _on_diagnose(self, s, a):
        """Run diagnostic and show result - helps debug value reading issues."""
        try:
            result = core.diagnose_snapshot(self.doc)
            TaskDialog.Show(TOOL_TITLE + " - Diagnostics", result)
        except Exception as e:
            import traceback
            TaskDialog.Show(TOOL_TITLE + " - Diag Error",
                str(e) + "\n\n" + traceback.format_exc())

    def _on_apply(self, s, a):
        """Reads grid rows and applies them to Revit via type_core.apply_snapshot."""
        if not self._dirty:
            self._set_status("No changes to apply.")
            return
        new_snapshot = self._grid_to_snapshot()
        if new_snapshot is None:
            return

        existing_names = [t["name"] for t in self._snapshot.get("types", [])]
        dlg = ImportPreviewDialog(existing_names, new_snapshot)
        if dlg.ShowDialog() != DialogResult.OK:
            return
        opts = dlg.get_options()

        self._set_status("Applying changes...")
        try:
            result = core.apply_snapshot(self.doc, new_snapshot, opts)
            msg = (
                "Applied successfully.\n\nCreated: {}\nUpdated: {}\nDeleted: {}\nSkipped: {}\nErrors: {}"
            ).format(
                len(result["added"]),
                len(result["updated"]),
                len(result["deleted"]),
                len(result["skipped"]),
                len(result["errors"]),
            )
            if result["errors"]:
                msg += "\n\nFirst errors:\n" + "\n".join(
                    "{} / {}: {}".format(t, p, e) for t, p, e in result["errors"][:10]
                )
            TaskDialog.Show(TOOL_TITLE, msg)
            self._load_snapshot()
            self._dirty = False
        except Exception as e:
            TaskDialog.Show(TOOL_TITLE + " - Error", str(e))
            self._set_status("Apply failed.")

    def _grid_to_snapshot(self):
        """
        Reads the current grid state and constructs a TypeSnapshot dict
        suitable for apply_snapshot or export.

        Returns None if there are validation errors.
        """
        if self._snapshot is None:
            return None

        params  = self._snapshot.get("parameters", [])
        scale   = 12.0 if self._unit_mode == "inches" else 1.0

        # Map column index -> param def (index 0 = Type Name, no param)
        col_param_map = {}
        for ci in range(1, self.grid.Columns.Count):
            col = self.grid.Columns[ci]
            if col.Tag:
                col_param_map[ci] = col.Tag

        type_rows = []
        seen_names = set()
        for row in self.grid.Rows:
            if row.IsNewRow:
                continue
            type_name = str(row.Cells[0].Value or "").strip()
            if not type_name:
                continue
            if type_name in seen_names:
                MessageBox.Show(
                    "Duplicate type name: '{}'. Please fix before applying.".format(type_name),
                    TOOL_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning
                )
                return None
            seen_names.add(type_name)

            values = {}
            for ci, p in col_param_map.items():
                if p.get("has_formula") or p.get("read_only"):
                    continue
                raw = row.Cells[ci].Value
                st  = p.get("storage_type", "String")
                if raw is None or str(raw).strip() == "":
                    values[p["name"]] = None
                    continue
                try:
                    if st == "Double":
                        if self._unit_mode == "frac_inches":
                            values[p["name"]] = _from_frac_inches(str(raw)) / 12.0
                        else:
                            values[p["name"]] = float(str(raw).replace(",", "")) / scale
                    elif st == "Integer":
                        values[p["name"]] = int(str(raw).replace(",", "").split(".")[0])
                    elif st == "ElementId":
                        values[p["name"]] = int(str(raw).split(".")[0])
                    else:
                        values[p["name"]] = str(raw)
                except:
                    values[p["name"]] = None

            type_rows.append({"name": type_name, "values": values})

        snap = dict(self._snapshot)   # shallow copy
        snap["types"] = type_rows
        return snap

    # ======================================================================= #
    # EXPORT / IMPORT                                                           #
    # ======================================================================= #

    # CPython detection removed — excel_bridge.py is pure IronPython

    def _on_export(self, s, a):
        """Exports current snapshot to Excel after column selection."""
        if self._snapshot is None:
            self._set_status("Load a family first.")
            return

        # Column picker
        params = self._snapshot.get("parameters", [])
        picker = ColumnPickerDialog(params, self._last_export_params)
        if picker.ShowDialog() != DialogResult.OK:
            return
        selected = picker.selected_names
        if not selected:
            MessageBox.Show("No columns selected.", TOOL_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        self._last_export_params = selected

        # Destination file
        dlg = SaveFileDialog()
        dlg.Title    = "Export Family Types to Excel"
        dlg.Filter   = "Excel Files (*.xlsx)|*.xlsx"
        dlg.FileName = (self._snapshot.get("family_name") or "types") + "_types.xlsx"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        xlsx_path = dlg.FileName

        snap_to_export = self._grid_to_snapshot() if self._dirty else self._snapshot
        if snap_to_export is None:
            return

        self._set_status("Exporting to Excel...")
        try:
            import excel_bridge as eb
            eb.export_to_xlsx(snap_to_export, xlsx_path,
                              unit_mode=self._unit_mode,
                              selected_params=selected)
            self._set_status("Exported: " + xlsx_path)
            try:
                os.startfile(xlsx_path)
            except:
                pass
        except Exception as e:
            import traceback
            TaskDialog.Show(TOOL_TITLE + " - Export Error",
                str(e) + "\n\n" + traceback.format_exc())
            self._set_status("Export failed.")

    def _on_import(self, s, a):
        """Imports an Excel file and applies the snapshot after user preview."""
        if self._snapshot is None:
            self._set_status("Load a family first.")
            return

        dlg = OpenFileDialog()
        dlg.Title  = "Import Family Types from Excel"
        dlg.Filter = "Excel Files (*.xlsx)|*.xlsx"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        xlsx_path = dlg.FileName

        self._set_status("Reading Excel file...")
        try:
            import excel_bridge as eb
            imported_snapshot = eb.import_from_xlsx(xlsx_path)
        except Exception as e:
            import traceback
            TaskDialog.Show(TOOL_TITLE + " - Import Error",
                str(e) + "\n\n" + traceback.format_exc())
            self._set_status("Import failed.")
            return

        # Preview dialog
        existing_names = [t["name"] for t in self._snapshot.get("types", [])]
        dlg2 = ImportPreviewDialog(existing_names, imported_snapshot)
        if dlg2.ShowDialog() != DialogResult.OK:
            return
        opts = dlg2.get_options()

        self._set_status("Applying import...")
        try:
            result = core.apply_snapshot(self.doc, imported_snapshot, opts)
            msg = (
                "Import applied successfully.\n\nCreated: {}\nUpdated: {}\nDeleted: {}\nSkipped: {}\nErrors: {}"
            ).format(
                len(result["added"]),
                len(result["updated"]),
                len(result["deleted"]),
                len(result["skipped"]),
                len(result["errors"]),
            )
            if result["skipped"]:
                msg += "\n\nSkipped (first 10):\n" + "\n".join(
                    "{}: {}".format(n, r) for n, r in result["skipped"][:10]
                )
            if result["errors"]:
                msg += "\n\nErrors (first 10):\n" + "\n".join(
                    "{} / {}: {}".format(t, p, e) for t, p, e in result["errors"][:10]
                )
            TaskDialog.Show(TOOL_TITLE, msg)
            self._load_snapshot()
        except Exception as e:
            TaskDialog.Show(TOOL_TITLE + " - Error", str(e))
            self._set_status("Apply failed.")

    # ======================================================================= #
    # PARAMETER TAB ACTIONS                                                    #
    # ======================================================================= #

    def _on_add_fam_param(self, s, a):
        dlg = AddParamDialog()
        if dlg.ShowDialog() != DialogResult.OK:
            return
        self._set_status("Adding parameter...")
        try:
            core.add_family_parameter(
                self.doc,
                dlg.result_name,
                dlg.result_type_token,
                dlg.result_group,
                dlg.result_inst,
                dlg.result_formula or None,
            )
            self._load_snapshot()
        except Exception as e:
            TaskDialog.Show(TOOL_TITLE + " - Error", str(e))

    def _on_remove_param(self, s, a):
        selected = [str(item.Text) for item in self.lvParams.SelectedItems]
        if not selected:
            return
        if MessageBox.Show(
            "Remove {} parameter(s)? This cannot be undone.\n\n{}".format(
                len(selected), "\n".join(selected[:10])
            ),
            TOOL_TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning
        ).ToString() != "OK":
            return
        errors = []
        for name in selected:
            try:
                core.remove_parameter(self.doc, name)
            except Exception as e:
                errors.append("{}: {}".format(name, e))
        if errors:
            TaskDialog.Show(TOOL_TITLE, "Some removals failed:\n" + "\n".join(errors))
        self._load_snapshot()

    def _on_set_formula(self, s, a):
        selected = [str(item.Text) for item in self.lvParams.SelectedItems]
        if not selected or len(selected) > 1:
            MessageBox.Show("Select exactly one parameter to edit its formula.",
                            TOOL_TITLE, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        name = selected[0]
        # Find current formula
        current = ""
        for p in self._snapshot.get("parameters", []):
            if p["name"] == name and p.get("has_formula"):
                current = "(formula set)"
                break

        # Simple input dialog
        frm = Form(); frm.Text = "Set Formula: " + name
        frm.Size = Size(480, 150); frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog

        lbl = Label(); lbl.Text = "Formula expression (leave blank to clear):"; lbl.Location = Point(12, 12); lbl.Size = Size(440, 18)
        txt = TextBox(); txt.Location = Point(12, 34); txt.Size = Size(440, 22); txt.Text = current if current != "(formula set)" else ""
        btnOK = Button(); btnOK.Text = "Set"; btnOK.Location = Point(290, 66); btnOK.Size = Size(70, 26)
        btnCl = Button(); btnCl.Text = "Cancel"; btnCl.Location = Point(366, 66); btnCl.Size = Size(70, 26)

        def _ok(x, y): frm.DialogResult = DialogResult.OK; frm.Close()
        btnOK.Click += _ok; btnCl.Click += lambda x, y: frm.Close()
        for w in [lbl, txt, btnOK, btnCl]: frm.Controls.Add(w)

        if frm.ShowDialog() != DialogResult.OK:
            return
        formula = (txt.Text or "").strip()
        try:
            core.set_formula(self.doc, name, formula)
            self._load_snapshot()
        except Exception as e:
            TaskDialog.Show(TOOL_TITLE, str(e))

    # ======================================================================= #
    # SP PROFILES TAB ACTIONS                                                  #
    # ======================================================================= #

    def _reload_profiles(self):
        self.lvProfiles.Items.Clear()
        try:
            import binder_core as bc
            import time as _time

            def _fmt_mtime(path):
                try:
                    return _time.strftime("%Y-%m-%d %H:%M", _time.localtime(os.path.getmtime(path)))
                except:
                    return ""

            def _read_disp(path):
                try:
                    d = bc.load_json_file(path)
                    return d.get("display_name") or os.path.basename(path)
                except:
                    return os.path.basename(path)

            for scope, folder in [("shared", bc.shared_dir()), ("user", bc.user_dir())]:
                if not folder:
                    continue
                try:
                    for fn in sorted(os.listdir(folder)):
                        if not fn.lower().endswith(".json"):
                            continue
                        p   = os.path.join(folder, fn)
                        sub = Array[String]([_read_disp(p), scope, fn, _fmt_mtime(p)])
                        it  = ListViewItem(sub); it.Tag = p
                        self.lvProfiles.Items.Add(it)
                except:
                    pass
        except:
            pass

    def _on_refresh_profiles(self, s, a):
        self._reload_profiles()

    def _on_open_sp_loader(self, s, a):
        """
        Launches the SP Loader Launcher form (from script.py / ui_builder.py
        in the SP Loader pushbutton folder) so the user can create, edit, or
        duplicate profiles without leaving the RFA Type Manager context.
        After it closes, refreshes the profile list.
        """
        try:
            import sys as _sys
            # SP Loader lives in a sibling pushbutton folder — add it to path
            sp_loader_dir = os.path.join(os.path.dirname(BUTTON_DIR), "SP_Loader.pushbutton")
            if not os.path.isdir(sp_loader_dir):
                # Try alternate naming conventions
                for candidate in os.listdir(os.path.dirname(BUTTON_DIR)):
                    if "sp" in candidate.lower() and "loader" in candidate.lower():
                        sp_loader_dir = os.path.join(os.path.dirname(BUTTON_DIR), candidate)
                        break
            if not os.path.isdir(sp_loader_dir):
                TaskDialog.Show(TOOL_TITLE,
                    "Could not find SP Loader folder.\n\nExpected a sibling pushbutton folder containing binder_core.py.\n"
                    + "Path searched: " + os.path.dirname(BUTTON_DIR))
                return
            if sp_loader_dir not in _sys.path:
                _sys.path.insert(0, sp_loader_dir)
            import importlib, binder_core as _bc
            # Set paths so SP Loader's launcher uses the right config dirs
            _bc.set_paths(shared_cfg_dir=os.path.join(sp_loader_dir, "Configs"))
            import ui_builder as _ub
            _ub.show_builder(None)
            self._reload_profiles()
        except Exception as e:
            import traceback
            TaskDialog.Show(TOOL_TITLE + " - SP Loader Error",
                str(e) + "\n\n" + traceback.format_exc())

    def _on_install_profile(self, s, a):
        if self.lvProfiles.SelectedItems.Count == 0:
            return
        path = self.lvProfiles.SelectedItems[0].Tag
        if not path or not os.path.exists(path):
            return
        try:
            import binder_core as bc
            bc.install_from_json(path)
            self._load_snapshot()
        except Exception as e:
            TaskDialog.Show(TOOL_TITLE + " - Error", str(e))

    # ======================================================================= #
    # UTILITIES                                                                 #
    # ======================================================================= #

    # ======================================================================= #
    # COLUMN MANAGEMENT                                                        #
    # ======================================================================= #

    def _on_col_header_click(self, s, e):
        """Right-click on column header shows the column context menu."""
        if e.Button == MouseButtons.Right:
            self._clicked_col_index = e.ColumnIndex
            pt = self.grid.PointToClient(System.Windows.Forms.Cursor.Position)
            self._col_ctx.Show(self.grid, pt)

    def _on_hide_col(self, s, a):
        """Hides the column that was right-clicked in the cell area."""
        try:
            cells = list(self.grid.SelectedCells)
            if cells:
                self.grid.Columns[cells[0].ColumnIndex].Visible = False
        except:
            pass

    def _on_hide_col_hdr(self, s, a):
        """Hides the column that was right-clicked in the header."""
        try:
            if hasattr(self, "_clicked_col_index"):
                self.grid.Columns[self._clicked_col_index].Visible = False
        except:
            pass

    def _on_col_move_left(self, s, a):
        """Moves the right-clicked column one position to the left."""
        try:
            ci = getattr(self, "_clicked_col_index", -1)
            if ci > 1:  # can't move before Type Name column
                col = self.grid.Columns[ci]
                col.DisplayIndex = col.DisplayIndex - 1
        except:
            pass

    def _on_col_move_right(self, s, a):
        """Moves the right-clicked column one position to the right."""
        try:
            ci = getattr(self, "_clicked_col_index", -1)
            col = self.grid.Columns[ci]
            if col.DisplayIndex < self.grid.Columns.Count - 1:
                col.DisplayIndex = col.DisplayIndex + 1
        except:
            pass

    def _on_manage_columns(self, s, a):
        """
        Column manager dialog:
        - CheckedListBox showing all columns sorted by current display order
        - Check/uncheck to show/hide
        - Move Up / Move Down buttons to reorder
        - All/None shortcuts
        Note: drag reorder also works directly on the grid header.
        """
        frm = Form()
        frm.Text = "Show / Hide / Reorder Columns"
        frm.Size = Size(440, 560)
        frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False

        lbl = Label()
        lbl.Text = "Check = visible.  Select + Move Up/Down to reorder.  (You can also drag column headers directly in the grid.)"
        lbl.Location = Point(12, 8); lbl.Size = Size(400, 32)
        lbl.Font = DFont("Arial", 9)

        clb = CheckedListBox()
        clb.Location = Point(12, 46); clb.Size = Size(300, 420)
        clb.CheckOnClick = True
        clb.Font = DFont("Arial", 9)

        # Populate sorted by current DisplayIndex, skip Type Name (col 0 - always visible)
        cols_sorted = sorted(
            [self.grid.Columns[i] for i in range(self.grid.Columns.Count)],
            key=lambda c: c.DisplayIndex
        )
        for col in cols_sorted:
            idx = clb.Items.Add(col.HeaderText)
            clb.SetItemChecked(idx, col.Visible)

        # Move Up / Move Down
        btnUp   = Button(); btnUp.Text   = "▲ Up";   btnUp.Size = Size(100, 28); btnUp.Location   = Point(322, 46)
        btnDown = Button(); btnDown.Text = "▼ Down"; btnDown.Size = Size(100, 28); btnDown.Location = Point(322, 80)
        btnAll  = Button(); btnAll.Text  = "All";    btnAll.Size = Size(48, 24);  btnAll.Location  = Point(322, 130)
        btnNone = Button(); btnNone.Text = "None";   btnNone.Size = Size(48, 24); btnNone.Location = Point(374, 130)

        def _move(delta):
            i = clb.SelectedIndex
            if i < 0:
                return
            j = i + delta
            if j < 0 or j >= clb.Items.Count:
                return
            # Swap items
            item_i = clb.Items[i]; chk_i = clb.GetItemChecked(i)
            item_j = clb.Items[j]; chk_j = clb.GetItemChecked(j)
            clb.Items[i] = item_j; clb.SetItemChecked(i, chk_j)
            clb.Items[j] = item_i; clb.SetItemChecked(j, chk_i)
            clb.SelectedIndex = j

        btnUp.Click   += lambda x, y: _move(-1)
        btnDown.Click += lambda x, y: _move(1)
        btnAll.Click  += lambda x, y: [clb.SetItemChecked(i, True)  for i in range(clb.Items.Count)]
        btnNone.Click += lambda x, y: [clb.SetItemChecked(i, False) for i in range(clb.Items.Count)]

        btnApply  = Button(); btnApply.Text  = "Apply";  btnApply.Size = Size(90, 28)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(90, 28)
        btnApply.Location  = Point(240, 478)
        btnCancel.Location = Point(336, 478)

        def _apply(x, y):
            # Apply visibility and display order
            for i in range(clb.Items.Count):
                header  = str(clb.Items[i])
                visible = clb.GetItemChecked(i)
                for col in [self.grid.Columns[j] for j in range(self.grid.Columns.Count)]:
                    if col.HeaderText == header:
                        col.Visible      = visible
                        col.DisplayIndex = i
                        break
            frm.Close()

        btnApply.Click  += _apply
        btnCancel.Click += lambda x, y: frm.Close()

        for w in [lbl, clb, btnUp, btnDown, btnAll, btnNone, btnApply, btnCancel]:
            frm.Controls.Add(w)
        frm.ShowDialog()

    # ======================================================================= #
    # FILL SERIES                                                               #
    # ======================================================================= #

    def _on_fill_series(self, s, a):
        """
        Fills selected column cells as an arithmetic series.
        Prompts for start value and step.
        """
        sel_cells = sorted(list(self.grid.SelectedCells),
                           key=lambda c: (c.ColumnIndex, c.RowIndex))
        if not sel_cells:
            return
        col_idx = sel_cells[0].ColumnIndex
        row_indices = sorted(set(c.RowIndex for c in sel_cells if not self.grid.Rows[c.RowIndex].IsNewRow))
        if not row_indices:
            return

        # Simple dialog for start + step
        frm = Form()
        frm.Text = "Fill Series"
        frm.Size = Size(280, 160)
        frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False

        def _lbl(text, x, y):
            l = Label(); l.Text = text; l.Location = Point(x, y); l.Size = Size(80, 18)
            l.Font = DFont("Arial", 9); frm.Controls.Add(l)

        _lbl("Start:", 12, 14)
        txt_start = TextBox(); txt_start.Location = Point(100, 12); txt_start.Size = Size(150, 22)
        # Pre-fill with first selected cell value
        try:
            txt_start.Text = str(self.grid.Rows[row_indices[0]].Cells[col_idx].Value or "0")
        except:
            txt_start.Text = "0"
        frm.Controls.Add(txt_start)

        _lbl("Step:", 12, 44)
        txt_step = TextBox(); txt_step.Location = Point(100, 42); txt_step.Size = Size(150, 22)
        txt_step.Text = "1"
        frm.Controls.Add(txt_step)

        btnOK = Button(); btnOK.Text = "Fill"; btnOK.Location = Point(100, 78); btnOK.Size = Size(70, 26)
        btnCl = Button(); btnCl.Text = "Cancel"; btnCl.Location = Point(178, 78); btnCl.Size = Size(70, 26)
        frm.Controls.Add(btnOK); frm.Controls.Add(btnCl)

        def _fill(x, y):
            try:
                start = float(txt_start.Text.replace(",",""))
                step  = float(txt_step.Text.replace(",",""))
                for i, ri in enumerate(row_indices):
                    cell = self.grid.Rows[ri].Cells[col_idx]
                    if not cell.ReadOnly:
                        val = start + i * step
                        # Format based on storage type
                        col = self.grid.Columns[col_idx]
                        p   = col.Tag
                        if p and p.get("storage_type") == "Double":
                            cell.Value = "{:.4f}".format(val)
                        elif p and p.get("storage_type") == "Integer":
                            cell.Value = str(int(round(val)))
                        else:
                            cell.Value = str(val)
                self._dirty = True
            except Exception as ex:
                MessageBox.Show("Fill error: " + str(ex), TOOL_TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            frm.Close()

        btnOK.Click += _fill
        btnCl.Click += lambda x, y: frm.Close()
        frm.ShowDialog()

    def _on_fill_right(self, s, a):
        """Ctrl+R: fills the selected cell value to all selected cells in the same row."""
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex))
        anchor = sel_cells[0]
        value  = anchor.Value
        for cell in sel_cells:
            if cell.RowIndex == anchor.RowIndex and not cell.ReadOnly:
                cell.Value = value
        self._dirty = True

    def _set_status(self, msg):
        try:
            self.lblStatus.Text = msg
            self.statusbar.Refresh()
        except:
            pass


# =============================================================================
# ENTRY POINT
# =============================================================================

import traceback

try:
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show(TOOL_TITLE, "No active document.")
    else:
        doc = uidoc.Document
        if not doc.IsFamilyDocument:
            TaskDialog.Show(TOOL_TITLE, "Open a family (.rfa) document before running this tool.")
        else:
            Application.EnableVisualStyles()
            try:
                frm = TypeManagerForm(doc)
            except Exception as e:
                TaskDialog.Show(TOOL_TITLE + " - Init Error",
                    str(e) + "\n\n" + traceback.format_exc())
                raise
            frm.ShowDialog()
except Exception as e:
    TaskDialog.Show(TOOL_TITLE + " - Startup Error",
        str(e) + "\n\n" + traceback.format_exc())