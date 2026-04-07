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

import System
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

# =============================================================================
# MATH EXPRESSION EVALUATOR  (formula bar = prefix)
# =============================================================================

def _eval_math_expr(expr, fallback=""):
    """
    Evaluates a simple math expression entered with = prefix in the formula bar.
    Safe: only digits, operators +  -  *  /  **  ^  and parentheses allowed.
    Examples:
        =2*6        -> "12"
        =12/4       -> "3"
        =1.5*12     -> "18"
        =2**3       -> "8"
        =2^3        -> "8"  (^ treated as **)
    Returns string result, or fallback value on error.
    """
    import re as _rem
    cleaned = expr.strip().replace("^", "**")
    if not _rem.match(r"^[0-9 .+\-*/()]+$", cleaned):
        return fallback
    try:
        result = eval(cleaned, {"__builtins__": {}}, {})
        fv = float(result)
        if fv == int(fv):
            return str(int(fv))
        return "{:.6g}".format(fv)
    except:
        return fallback


def _parse_length_input(s):
    """
    Parses any Revit-style feet-inches string to decimal feet.
    Accepted formats:
      2'-6"   2'-6 1/2"   2'6"   2.5'   2'-6
      2 6     2 6.5       2 6 1/2
      6"      6           1/2"   0.5
      -2'-6"  2 ft 6 in
    Bare number (no unit marker) = inches, matching Revit convention.
    Returns decimal feet as float, or None if unparseable.
    """
    import re as _re2
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None

    neg = s.startswith("-")
    if neg:
        s = s[1:].strip()

    s = _re2.sub(r"\s*'\s*-\s*", "'-", s)
    s = _re2.sub(r"\s+", " ", s).strip()

    def _res(v):
        return -v if neg else v

    # 1. Explicit feet marker: 2'-6 1/2"  2'6"  2.5'  2'-6
    m = _re2.match(
        r"^(\d+(?:\.\d+)?)'[-\s]?"
        r"(?:(\d+(?:\.\d+)?)(?:\s+(\d+)/(\d+))?\"?)?$", s)
    if m:
        ft   = float(m.group(1))
        inch = float(m.group(2) or 0)
        frac = float(m.group(3) or 0) / float(m.group(4) or 1) if m.group(3) else 0
        return _res(ft + (inch + frac) / 12.0)

    # 2. Inches with " marker: 6"  6 1/2"
    m = _re2.match(r"^(\d+(?:\.\d+)?)(?:\s+(\d+)/(\d+))?\"$", s)
    if m:
        inch = float(m.group(1))
        frac = float(m.group(2)) / float(m.group(3)) if m.group(2) else 0
        return _res((inch + frac) / 12.0)

    # 3. Bare fraction with ": 1/2"
    m = _re2.match(r"^(\d+)/(\d+)\"$", s)
    if m:
        return _res((float(m.group(1)) / float(m.group(2))) / 12.0)

    # 4. Two numbers = feet inches: 2 6  |  2 6.5  |  2 6 1/2
    m = _re2.match(
        r"^(\d+(?:\.\d+)?)\s+(\d+(?:\.\d+)?)(?:\s+(\d+)/(\d+))?$", s)
    if m:
        ft   = float(m.group(1))
        inch = float(m.group(2))
        frac = float(m.group(3)) / float(m.group(4)) if m.group(3) else 0
        return _res(ft + (inch + frac) / 12.0)

    # 5. Text units: 2 ft 6 in
    m = _re2.match(
        r"^(\d+(?:\.\d+)?)\s*(?:ft|feet|')\s*"
        r"(\d+(?:\.\d+)?)\s*(?:in|inch|inches|\")?\s*"
        r"(?:(\d+)/(\d+))?$", s, _re2.IGNORECASE)
    if m:
        ft   = float(m.group(1))
        inch = float(m.group(2))
        frac = float(m.group(3)) / float(m.group(4)) if m.group(3) else 0
        return _res(ft + (inch + frac) / 12.0)

    # 6. Plain decimal fallback: bare number = inches (Revit convention)
    try:
        return _res(float(s) / 12.0)
    except:
        return None


def _to_frac_inches(feet_val):
    """
    Formats decimal feet as a Revit-style fractional inches string.
    e.g. 2.541667 -> 2'-6 1/2"
    Rounds to nearest 1/16".
    """
    try:
        total_in = float(feet_val) * 12.0
        neg      = total_in < 0
        total_in = abs(total_in)
        ft       = int(total_in // 12)
        rem      = total_in - ft * 12
        whole    = int(rem)
        frac     = rem - whole
        sixteenths = int(round(frac * 16))
        if sixteenths == 16:
            whole += 1; sixteenths = 0
        if whole == 12:
            ft += 1; whole = 0

        def _gcd(a, b):
            while b: a, b = b, a % b
            return a

        frac_str = (" " + str(sixteenths // _gcd(sixteenths, 16)) +
                    "/" + str(16 // _gcd(sixteenths, 16))) if sixteenths else ""

        result = (str(ft) + "'-" + str(whole) + frac_str + '"') if ft else (str(whole) + frac_str + '"')
        return ("-" if neg else "") + result
    except:
        return ""


def _from_frac_inches(s):
    """
    Parses any Revit-style input to decimal inches (backward compat wrapper).
    """
    result = _parse_length_input(s)
    return (result * 12.0) if result is not None else 0.0


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
        self.chkCreate  = CheckBox(); self.chkCreate.Text  = "Create new types";              self.chkCreate.Checked = len(new_types) > 0
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
        self._last_export_params = None  # remember column selection
        self._unit_mode  = "inches"      # "inches" | "feet" | "frac_inches"
        self._undo_stack    = []         # list of (ri, ci, old_val, new_val)
        self._redo_stack    = []         # list of (ri, ci, old_val, new_val)
        self._modified_rows = set()      # set of actual_ri that have been edited

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
        sep3 = ToolStripSeparator()
        self.btnToggleFormulas = ToolStripButton("Show Formulas")
        self.btnToggleFormulas.ToolTipText = "Toggle formula columns on/off in the grid"
        self.btnToggleFormulas.CheckOnClick = True
        self.btnToggleFormulas.Checked      = True
        self.btnToggleFormulas.CheckedChanged += self._on_toggle_formulas

        sep4 = ToolStripSeparator()
        self.btnUnits = ToolStripButton(u"⚙ Units...")
        self.btnUnits.ToolTipText = "Set display units for all parameter types"
        self.btnUnits.Click += self._on_units_prefs

        sep5b = ToolStripSeparator()
        btnFindReplace = ToolStripButton("Find & Replace")
        btnFindReplace.ToolTipText = "Find and replace values across the grid (Ctrl+H)"
        btnFindReplace.Click += self._on_find_replace

        sep6 = ToolStripSeparator()
        self.btnUndo = ToolStripButton(u"↶ Undo")
        self.btnUndo.ToolTipText = "Undo last grid edit (Ctrl+Z)"
        self.btnUndo.Click += self._on_undo
        self.btnRedo = ToolStripButton(u"↷ Redo")
        self.btnRedo.ToolTipText = "Redo (Ctrl+Y)"
        self.btnRedo.Click += self._on_redo

        for item in [self.btnRefresh, sep1, self.btnExport, self.btnImport,
                     self.btnApply, sep2, self.btnUndo, self.btnRedo,
                     sep3, self.btnToggleFormulas, sep5b, btnFindReplace,
                     sep4, self.btnUnits]:
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
        self.lblColStats  = ToolStripStatusLabel("")   # column stats: count/min/max/sum/avg
        self.lblColStats.ForeColor = Color.FromArgb(60, 60, 120)
        self.statusbar.Items.Add(self.lblStatus)
        self.statusbar.Items.Add(self.lblTypeCount)
        self.statusbar.Items.Add(self.lblColStats)

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
        self.grid.AllowUserToAddRows      = False  # managed manually in virtual mode
        self.grid.AllowUserToDeleteRows   = True
        self.grid.MultiSelect             = True
        self.grid.SelectionMode           = DataGridViewSelectionMode.RowHeaderSelect
        self.grid.AutoSizeColumnsMode     = getattr(DataGridViewAutoSizeColumnsMode, "None")
        self.grid.RowHeadersWidth              = 24
        self.grid.ColumnHeadersHeightSizeMode  = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        self.grid.ColumnHeadersHeight          = 26
        self.grid.AllowUserToOrderColumns      = True   # drag column headers to reorder
        self.grid.AllowUserToResizeColumns     = True   # drag column edges to resize
        self.grid.VirtualMode                  = True
        self.grid.AllowUserToAddRows           = False
        self.grid.ShowCellToolTips             = False  # no hover preview popups
        # DoubleBuffered is read-only on DataGridView — enable via reflection
        try:
            import System.Reflection as _ref
            _prop = type(self.grid).GetProperty(
                "DoubleBuffered",
                _ref.BindingFlags.NonPublic | _ref.BindingFlags.Instance)
            if _prop:
                _prop.SetValue(self.grid, True, None)
        except:
            pass  # older .NET: flicker suppressed by VirtualMode alone
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
        ctx.Items.Add(ToolStripSeparator())
        ctx.Items.Add(_mitem(u"▼ Filter This Column...",  self._on_filter_cell_col))
        ctx.Items.Add(_mitem("Clear This Column Filter",      self._on_clear_cell_col_filter))
        ctx.Items.Add(_mitem("Clear All Filters",             self._on_clear_all_filters))
        self.grid.ContextMenuStrip = ctx

        # Column header right-click for hide/show
        col_ctx = ContextMenuStrip()
        col_ctx.Items.Add(_mitem(u"▼ Filter This Column (Ctrl+Click)...", self._on_filter_hdr))
        col_ctx.Items.Add(_mitem("Clear This Column Filter",  self._on_clear_col_filter))
        col_ctx.Items.Add(ToolStripSeparator())
        col_ctx.Items.Add(_mitem("Hide This Column",          self._on_hide_col_hdr))
        col_ctx.Items.Add(_mitem("Show/Hide Columns...",      self._on_manage_columns))
        col_ctx.Items.Add(ToolStripSeparator())
        col_ctx.Items.Add(_mitem("Unit Preferences...",       self._on_units_prefs))
        col_ctx.Items.Add(ToolStripSeparator())
        col_ctx.Items.Add(_mitem("Move Left",                 self._on_col_move_left))
        col_ctx.Items.Add(_mitem("Move Right",                self._on_col_move_right))
        self._col_ctx = col_ctx
        self.grid.ColumnHeaderMouseClick += self._on_col_header_click

        # Wire events
        self.grid.CellValueNeeded     += self._on_cell_value_needed
        self.grid.CellValuePushed     += self._on_cell_value_pushed
        self.grid.KeyDown             += self._on_grid_keydown
        self.grid.CellBeginEdit       += self._on_cell_begin_edit
        self.grid.CellDoubleClick     += self._on_cell_double_click
        self.grid.RowPrePaint         += self._on_row_pre_paint
        self.grid.CellMouseDown       += self._on_grid_cell_mouse_down  # track right-clicked cell
        # Row drag-to-reorder
        self.grid.MouseDown           += self._on_row_drag_start
        self.grid.MouseMove           += self._on_row_drag_move
        self.grid.MouseUp             += self._on_row_drag_end
        self._row_drag_source = None
        self._ctx_row_index   = -1     # row index of last right-clicked cell
        self._ctx_col_index   = -1     # col index of last right-clicked cell

        # Formula bar: shows current cell value, allows editing long strings
        formula_panel = Panel()
        formula_panel.Dock   = DockStyle.Top
        formula_panel.Height = 26
        formula_panel.BackColor = Color.FromArgb(248, 248, 248)

        lbl_fx = Label()
        lbl_fx.Text = "fx "
        lbl_fx.Location = Point(4, 5); lbl_fx.Size = Size(22, 18)
        lbl_fx.Font = DFont("Arial", 9, FontStyle.Bold)
        lbl_fx.ForeColor = Color.FromArgb(47, 61, 78)

        self.txtFormulaBar = TextBox()
        self.txtFormulaBar.Location = Point(30, 3)
        self.txtFormulaBar.Size     = Size(800, 20)
        self.txtFormulaBar.Font     = DFont("Courier New", 9)
        self.txtFormulaBar.Anchor   = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
        self.txtFormulaBar.ReadOnly = True  # becomes editable when cell selected

        self.chkAutoApply = CheckBox()
        self.chkAutoApply.Text     = "Auto-apply"
        self.chkAutoApply.Location = Point(formula_panel.Width - 110, 4)
        self.chkAutoApply.Size     = Size(100, 18)
        self.chkAutoApply.Font     = DFont("Arial", 8)
        self.chkAutoApply.Anchor   = AnchorStyles.Right | AnchorStyles.Top
        self.chkAutoApply.Checked  = False
        # ToolTip via component (CheckBox has no .ToolTip property in WinForms)
        _tt = ToolTip()
        _tt.SetToolTip(self.chkAutoApply, "When on, each cell edit writes to Revit immediately")

        formula_panel.Controls.Add(lbl_fx)
        formula_panel.Controls.Add(self.txtFormulaBar)
        formula_panel.Controls.Add(self.chkAutoApply)

        # Wire formula bar to grid selection
        self.grid.SelectionChanged   += self._on_selection_changed
        self.txtFormulaBar.KeyDown   += self._on_formula_bar_keydown
        self.txtFormulaBar.LostFocus += self._on_formula_bar_commit

        panel.Controls.Add(self.grid)
        panel.Controls.Add(formula_panel)
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
        """
        Reads current family state FROM REVIT and rebuilds the grid.
        Only call this when Revit state has actually changed (Apply, Import,
        Add/Remove param, Refresh button). For unit-pref changes or column
        toggling use _refresh_grid() instead.
        """
        self._set_status("Loading from Revit...")
        try:
            self._snapshot = core.snapshot_family(self.doc)
            self._fp_by_name = {}
            try:
                fm = self.doc.FamilyManager
                for fp in fm.Parameters:
                    try:
                        self._fp_by_name[fp.Definition.Name] = fp
                    except:
                        pass
            except:
                pass
            self._refresh_grid()
            self._dirty = False
        except Exception as e:
            self._set_status("Load error: " + str(e))
            TaskDialog.Show(TOOL_TITLE, "Failed to load family:\n" + str(e))

    def _refresh_grid(self):
        """
        Rebuilds the grid display from the existing snapshot — no Revit API call.
        Safe to call after unit pref changes, column visibility changes, etc.
        Preserves unsaved edits in _grid_modified.
        """
        if self._snapshot is None:
            return
        # Snapshot all session state BEFORE rebuild
        saved_modified   = dict(self._grid_modified) if hasattr(self, "_grid_modified") else {}
        saved_undo       = list(self._undo_stack)     if hasattr(self, "_undo_stack")     else []
        saved_redo       = list(self._redo_stack)     if hasattr(self, "_redo_stack")     else []
        saved_filters    = dict(self._col_filters)    if hasattr(self, "_col_filters")    else {}

        # Snapshot column visibility by param NAME (stable across rebuilds)
        # Headers change with units e.g. "Weight (lbs)" -> "Weight (kg)"
        # but col.Name is always the raw param name
        saved_col_vis = {}
        saved_col_order = {}
        for i in range(self.grid.Columns.Count):
            col = self.grid.Columns[i]
            saved_col_vis[col.Name]   = col.Visible
            saved_col_order[col.Name] = col.DisplayIndex

        self._populate_grid()
        self._populate_params()

        # Restore column visibility (by name, not header text)
        if saved_col_vis:
            for i in range(self.grid.Columns.Count):
                col = self.grid.Columns[i]
                if col.Name in saved_col_vis:
                    col.Visible = saved_col_vis[col.Name]

        # Restore all other session state
        if hasattr(self, "_grid_modified"):
            self._grid_modified.update(saved_modified)
        if hasattr(self, "_undo_stack"):
            self._undo_stack[:] = saved_undo
            self._redo_stack[:] = saved_redo
        if saved_filters:
            self._col_filters = saved_filters

        # Rebuild color cache with restored edits
        self._rebuild_row_colors()

        n_types  = len(self._snapshot.get("types", []))
        n_params = len(self._snapshot.get("parameters", []))
        self.lblTypeCount.Text = "{} types  |  {} params".format(n_types, n_params)

        # Re-apply any active column filters (unit change etc. rebuilds grid_data
        # but should preserve the user's active filters)
        if getattr(self, "_col_filters", None):
            self._apply_col_filters()
        else:
            self._rebuild_row_colors()

        self._set_status("Loaded.")

    def _populate_grid(self):
        """
        Rebuilds the grid in VIRTUAL MODE for performance with large type sets.
        All cell values are pre-computed into self._grid_data (list of lists)
        and served via CellValueNeeded — the grid never holds actual row objects.
        """
        self.grid.Visible = False   # suppress all intermediate repaints
        self.grid.SuspendLayout()
        self.grid.Columns.Clear()
        self.grid.RowCount = 0

        params = self._snapshot.get("parameters", [])
        types  = self._snapshot.get("types",      [])

        # unit_prefs conversion
        try:
            import unit_prefs as up
            up.load_prefs()
        except Exception:
            up = None

        _frac_mode = False
        if up:
            try:
                attr, _, _ = up.get_current_option("Length")
                _frac_mode = attr in ("FractionalInches", "FeetFractionalInches")
            except:
                pass

        # Store params list for CellValueNeeded/CellValuePushed
        self._grid_params = params

        # Type Name column (frozen)
        col0 = DataGridViewTextBoxColumn()
        col0.Name = "__type_name__"; col0.HeaderText = "Type Name"
        col0.Width = 180; col0.Frozen = True; col0.Tag = None
        col0.ReadOnly = False
        self.grid.Columns.Add(col0)

        # Parameter columns
        for p in params:
            c    = DataGridViewTextBoxColumn()
            c.Name = p["name"]
            st   = p.get("storage_type", "String")
            is_f = p.get("has_formula") or p.get("read_only")
            spec = p.get("spec_type_id", "")

            if st == "Double" and not is_f:
                try:
                    lbl = up.get_label(spec) if up and spec else ""
                except:
                    lbl = ""
                c.HeaderText = (p["name"] + " (" + lbl + ")") if lbl else p["name"]
            else:
                c.HeaderText = p["name"]

            c.Width    = max(90, len(p["name"]) * 8 + 20)
            c.Tag      = p
            c.ReadOnly = False

            if is_f:
                c.DefaultCellStyle.BackColor = COLOR_FORMULA_BG
                c.DefaultCellStyle.ForeColor = Color.FromArgb(80, 80, 140)
                c.DefaultCellStyle.Font      = DFont("Arial", 8, FontStyle.Italic)
                c.ToolTipText                = "[formula - double-click to edit]"
                try:
                    c.Visible = bool(self.btnToggleFormulas.Checked)
                except:
                    pass

            self.grid.Columns.Add(c)

        # Pre-compute all display strings into a flat list-of-lists
        # Row 0..N-1 = type data, each inner list = [type_name, val0, val1, ...]
        self._grid_data = []
        self._grid_modified = {}  # (row, col) -> edited value

        for type_row in types:
            row_data = [type_row["name"]]
            values   = type_row.get("values", {})

            for p in params:
                st   = p.get("storage_type", "String")
                val  = values.get(p["name"])
                is_f = p.get("has_formula") or p.get("read_only")
                spec = p.get("spec_type_id", "")

                if is_f:
                    expr = p.get("formula_expr", "")
                    row_data.append("=" + expr if expr else "")
                elif val is None:
                    row_data.append("")
                elif st == "Double":
                    try:
                        fv = float(val)
                        if up and spec:
                            if _frac_mode and up.is_length(spec):
                                row_data.append(_to_frac_inches(fv))
                            else:
                                disp = up.to_display(fv, spec)
                                row_data.append("{:.4f}".format(float(disp)))
                        else:
                            row_data.append("{:.4f}".format(fv))
                    except:
                        row_data.append("")
                elif st == "Integer":
                    # Yes/No params: show Yes/No instead of 1/0
                    if p.get("is_yes_no"):
                        row_data.append("Yes" if val else "No")
                    else:
                        row_data.append(str(val) if val is not None else "")
                elif st == "ElementId":
                    # ElementId -1 = "none assigned" -> show blank
                    try:
                        row_data.append("" if int(val) < 0 else str(val))
                    except:
                        row_data.append("")
                else:
                    row_data.append(str(val) if val is not None else "")

            self._grid_data.append(row_data)

        # Set virtual row count (triggers CellValueNeeded for visible rows only)
        self.grid.RowCount = len(self._grid_data)
        self._row_color_cache = []
        self._grid_modified   = {}
        self._modified_rows   = set()
        self.grid.ResumeLayout()
        self.grid.Visible = True     # re-enable painting
        self.grid.Invalidate()


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
        """Filter by type name — delegates to _apply_col_filters to combine with col filters."""
        self._apply_col_filters()

    def _on_cell_value_needed(self, s, e):
        """Virtual mode: supply the display value for a cell."""
        try:
            ri = e.RowIndex; ci = e.ColumnIndex
            if ri < 0 or ci < 0:
                return
            # Map through filter if active
            actual_ri = ri
            if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
                if ri >= len(self._grid_filter_map):
                    e.Value = ""; return
                actual_ri = self._grid_filter_map[ri]
            # User-edited override
            override = self._grid_modified.get((actual_ri, ci))
            if override is not None:
                e.Value = override
                return
            if actual_ri < len(self._grid_data) and ci < len(self._grid_data[actual_ri]):
                e.Value = self._grid_data[actual_ri][ci]
            else:
                e.Value = ""
        except:
            e.Value = ""

    def _on_cell_value_pushed(self, s, e):
        """Virtual mode: store a user edit, push to undo stack, mark dirty."""
        try:
            ri = e.RowIndex; ci = e.ColumnIndex
            if ri < 0 or ci < 0:
                return
            actual_ri = ri
            if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
                if ri < len(self._grid_filter_map):
                    actual_ri = self._grid_filter_map[ri]
            # Capture old value for undo
            old_val = self._grid_modified.get((actual_ri, ci))
            if old_val is None and hasattr(self, "_grid_data"):
                row = self._grid_data[actual_ri] if actual_ri < len(self._grid_data) else []
                old_val = row[ci] if ci < len(row) else ""
            new_val = e.Value
            # Push to undo stack (clear redo on new edit)
            if hasattr(self, "_undo_stack"):
                self._undo_stack.append((actual_ri, ci, old_val, new_val))
                if len(self._undo_stack) > 200:
                    self._undo_stack.pop(0)
                self._redo_stack = []
            self._grid_modified[(actual_ri, ci)] = new_val
            self._modified_rows.add(actual_ri)
            self._dirty = True
            self._rebuild_row_colors()
            self.grid.InvalidateRow(ri)
        except:
            pass

    def _on_row_pre_paint(self, s, e):
        pass  # row colors applied via _apply_visible_row_styles(), not during paint

    def _rebuild_row_colors(self):
        """
        Rebuilds the row color cache using _modified_rows set — O(rows) only.
        Then applies colors to realized rows via _apply_visible_row_styles().
        """
        try:
            n     = self.grid.RowCount
            cache = []
            fmap  = (self._grid_filter_map
                     if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None
                     else None)
            drag_hover   = getattr(self, "_drag_hover_row", None)
            modified_rows = getattr(self, "_modified_rows", set())

            for display_ri in range(n):
                actual_ri = fmap[display_ri] if fmap and display_ri < len(fmap) else display_ri
                if drag_hover == display_ri:
                    cache.append(Color.FromArgb(180, 220, 255))
                elif actual_ri in modified_rows:
                    cache.append(COLOR_MODIFIED)
                elif actual_ri % 2 == 1:
                    cache.append(COLOR_ALT_ROW)
                else:
                    cache.append(Color.White)
            self._row_color_cache = cache
        except:
            self._row_color_cache = []
        self._apply_visible_row_styles()

    def _apply_visible_row_styles(self):
        """Apply cached row colors to realized row objects (called after rebuild)."""
        try:
            cache = getattr(self, "_row_color_cache", [])
            for ri in range(min(self.grid.Rows.Count, len(cache))):
                self.grid.Rows[ri].DefaultCellStyle.BackColor = cache[ri]
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
        elif e.Control and e.KeyCode == Keys.H:
            self._on_find_replace(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.Z:
            self._on_undo(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.Y:
            self._on_redo(None, None)
            e.Handled = True
        elif e.KeyCode == Keys.Escape:
            try:
                self.txtFormulaBar.ReadOnly = True
            except:
                pass
        elif e.Control and e.KeyCode == Keys.C:
            self._on_grid_copy(None, None)
            e.Handled = True
        elif e.Control and e.KeyCode == Keys.V:
            self._on_grid_paste(None, None)
            e.Handled = True

    def _selected_row_indices(self):
        """Returns sorted list of selected row indices (virtual mode)."""
        indices = sorted(set(
            row.Index for row in self.grid.SelectedRows
        ))
        return indices

    def _actual_row(self, display_ri):
        """Maps a display row index to actual _grid_data index via filter map."""
        if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
            if display_ri < len(self._grid_filter_map):
                return self._grid_filter_map[display_ri]
        return display_ri

    def _on_add_row(self, s, a):
        """Adds a blank row at the bottom of the grid (virtual mode)."""
        if not hasattr(self, "_grid_data"):
            return
        n_cols = self.grid.Columns.Count
        self._grid_data.append(["New Type"] + [""] * (n_cols - 1))
        self._grid_filter_map = None  # clear filter
        self.grid.RowCount = len(self._grid_data)
        # Select and scroll to new row
        last = len(self._grid_data) - 1
        self.grid.CurrentCell = self.grid.Rows[last].Cells[0]
        self._dirty = True

    def _on_dup_rows(self, s, a):
        """Duplicates selected rows, appending them (virtual mode)."""
        if not hasattr(self, "_grid_data"):
            return
        indices = self._selected_row_indices()
        if not indices:
            return
        for di in indices:
            actual = self._actual_row(di)
            if actual < len(self._grid_data):
                src = list(self._grid_data[actual])
                # Apply any edits to the source row
                for ci in range(self.grid.Columns.Count):
                    ov = self._grid_modified.get((actual, ci))
                    if ov is not None:
                        src[ci] = ov
                src[0] = (str(src[0] or "") + " Copy") if src[0] else "Copy"
                self._grid_data.append(src)
        self._grid_filter_map = None
        self.grid.RowCount = len(self._grid_data)
        self._dirty = True

    def _on_del_rows(self, s, a):
        """Deletes selected rows after confirmation (virtual mode)."""
        if not hasattr(self, "_grid_data"):
            return
        indices = self._selected_row_indices()
        if not indices:
            return
        if MessageBox.Show(
            "Delete {} selected type(s)? This will be applied to Revit on Apply.".format(len(indices)),
            TOOL_TITLE, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning
        ).ToString() != "OK":
            return
        # Map display indices to actual indices, then delete highest-first
        actual_indices = sorted(
            set(self._actual_row(di) for di in indices), reverse=True
        )
        for ai in actual_indices:
            if ai < len(self._grid_data):
                del self._grid_data[ai]
        # Clean up modified cells for deleted rows
        self._grid_modified = {
            k: v for k, v in self._grid_modified.items()
            if k[0] not in actual_indices
        }
        self._grid_filter_map = None
        self.grid.RowCount = len(self._grid_data)
        self._dirty = True

    def _on_fill_down(self, s, a):
        """Fills topmost selected cell value down through selected rows (virtual mode)."""
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex))
        anchor  = sel_cells[0]
        value   = anchor.Value
        col_idx = anchor.ColumnIndex
        for cell in sel_cells:
            actual_ri = self._actual_row(cell.RowIndex)
            self._grid_modified[(actual_ri, col_idx)] = value
            self._modified_rows.add(actual_ri)
        self._dirty = True
        self.grid.Invalidate()

    def _on_fill_up(self, s, a):
        """Fills bottommost selected cell value up through selected rows (virtual mode)."""
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex), reverse=True)
        anchor  = sel_cells[0]
        value   = anchor.Value
        col_idx = anchor.ColumnIndex
        for cell in sel_cells:
            actual_ri = self._actual_row(cell.RowIndex)
            self._grid_modified[(actual_ri, col_idx)] = value
            self._modified_rows.add(actual_ri)
        self._dirty = True
        self.grid.Invalidate()

    def _on_sort(self, s, a):
        """Sorts _grid_data alphabetically by Type Name (virtual mode)."""
        if not hasattr(self, "_grid_data"):
            return
        # Merge pending edits into data before sorting
        for (ri, ci), val in self._grid_modified.items():
            if ri < len(self._grid_data) and ci < len(self._grid_data[ri]):
                self._grid_data[ri][ci] = val
        self._grid_modified = {}
        self._grid_data.sort(key=lambda r: str(r[0] or "").lower())
        self._grid_filter_map = None
        self.grid.RowCount = len(self._grid_data)
        self.grid.Invalidate()

    def _on_grid_copy(self, s, a):
        """Copies selected cells to clipboard (virtual mode — reads _grid_modified + _grid_data)."""
        try:
            sel_cells = sorted(self.grid.SelectedCells, key=lambda c: (c.RowIndex, c.ColumnIndex))
            if not sel_cells:
                return
            rows = {}
            for cell in sel_cells:
                actual_ri = self._actual_row(cell.RowIndex)
                ov = self._grid_modified.get((actual_ri, cell.ColumnIndex))
                if ov is not None:
                    val = str(ov)
                elif actual_ri < len(self._grid_data) and cell.ColumnIndex < len(self._grid_data[actual_ri]):
                    val = str(self._grid_data[actual_ri][cell.ColumnIndex] or "")
                else:
                    val = ""
                rows.setdefault(cell.RowIndex, {})[cell.ColumnIndex] = val
            lines = []
            for ri in sorted(rows.keys()):
                row = rows[ri]
                lines.append("\t".join(row[ci] for ci in sorted(row.keys())))
            Clipboard.SetText("\r\n".join(lines))
        except:
            pass

    def _on_grid_paste(self, s, a):
        """
        Pastes clipboard data at selected cell.
        Handles:
          - Excel copy (tab-separated columns, CR+LF rows)
          - Internal grid copy
          - Single values broadcast across selection
        """
        try:
            text = Clipboard.GetText()
            if not text:
                return
            # Normalize line endings and strip trailing empty row Excel adds
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            if text.endswith("\n"):
                text = text[:-1]
            clip_lines = text.split("\n")
            if not clip_lines:
                return

            sel = sorted(self.grid.SelectedCells, key=lambda c: (c.RowIndex, c.ColumnIndex))
            if not sel:
                return
            start_row = sel[0].RowIndex
            start_col = sel[0].ColumnIndex

            # If pasting a single value onto a multi-cell selection, broadcast it
            if len(clip_lines) == 1 and "\t" not in clip_lines[0] and len(sel) > 1:
                val = clip_lines[0]
                for cell in sel:
                    actual_ri = self._actual_row(cell.RowIndex)
                    if cell.ColumnIndex > 0:  # don't overwrite Type Name accidentally
                        self._grid_modified[(actual_ri, cell.ColumnIndex)] = val
                self._dirty = True
                self._rebuild_row_colors()
                self.grid.Refresh()
                self._set_status("Pasted to {} cells.".format(len(sel)))
                return

            # Standard paste: row-by-row
            rows_pasted = 0
            for r_off, line in enumerate(clip_lines):
                vals = line.split("\t")
                display_ri = start_row + r_off
                # Extend _grid_data if needed
                while display_ri >= len(self._grid_data):
                    n_cols = self.grid.Columns.Count
                    self._grid_data.append(["New Type"] + [""] * (n_cols - 1))
                actual_ri = self._actual_row(display_ri)
                for c_off, val in enumerate(vals):
                    col_idx = start_col + c_off
                    if col_idx >= self.grid.Columns.Count:
                        break
                    self._grid_modified[(actual_ri, col_idx)] = val
                    self._modified_rows.add(actual_ri)
                rows_pasted += 1

            self.grid.RowCount = len(self._grid_data)
            self._dirty = True
            self._rebuild_row_colors()
            self.grid.Invalidate()
            self._set_status("Pasted {} row(s), {} col(s).".format(
                rows_pasted, len(clip_lines[0].split("\t"))))
        except Exception as ex:
            self._set_status("Paste error: " + str(ex))

    def _on_paste_broadcast(self, s, a):
        """Paste clipboard value to all selected cells in same column (virtual mode)."""
        try:
            text = (Clipboard.GetText() or "").split("\t")[0].split("\n")[0].strip()
            sel_cells = list(self.grid.SelectedCells)
            if not sel_cells:
                return
            col_idx = sorted(sel_cells, key=lambda c: c.ColumnIndex)[0].ColumnIndex
            for cell in sel_cells:
                if cell.ColumnIndex == col_idx:
                    actual_ri = self._actual_row(cell.RowIndex)
                    self._grid_modified[(actual_ri, col_idx)] = text
                    self._modified_rows.add(actual_ri)
            self._dirty = True
            self.grid.Invalidate()
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

        n_types = len(new_snapshot.get("types", []))
        self._set_status("Applying {} types...".format(n_types))
        self.statusbar.Refresh()
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
        Builds a TypeSnapshot from the virtual grid state.
        Reads _grid_data for base values and _grid_modified for user edits.
        Only includes values that have actually been changed (comparing
        display string before/after) to minimize fm.Set() calls on apply.
        Returns None on validation error.
        """
        if self._snapshot is None:
            return None
        if not hasattr(self, "_grid_data"):
            return None

        # Build col index -> param def map (respects current DisplayIndex order)
        col_param_map = {}
        for ci in range(1, self.grid.Columns.Count):
            col = self.grid.Columns[ci]
            if col.Tag:
                col_param_map[ci] = col.Tag

        type_rows = []
        seen_names = set()
        n_rows = len(self._grid_data)

        for ri in range(n_rows):
            # Get type name — may have been edited
            name_override = self._grid_modified.get((ri, 0))
            type_name = str(name_override if name_override is not None
                           else self._grid_data[ri][0]).strip()
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

                # Only process cells the user actually edited
                if (ri, ci) not in self._grid_modified:
                    continue

                raw = self._grid_modified[(ri, ci)]
                self._modified_rows.add(ri)
                if raw is None or str(raw).strip() == "":
                    continue

                st = p.get("storage_type", "String")
                try:
                    if st == "Double":
                        raw_str = str(raw).replace(",", "").strip()
                        values[p["name"]] = self._parse_double(p, raw_str, 1.0)
                    elif st == "Integer":
                        sv = str(raw).strip().lower()
                        if sv in ("yes", "true"):
                            values[p["name"]] = 1
                        elif sv in ("no", "false"):
                            values[p["name"]] = 0
                        else:
                            try:
                                values[p["name"]] = int(str(raw).replace(",","").split(".")[0])
                            except:
                                pass
                    elif st == "ElementId":
                        values[p["name"]] = int(str(raw).split(".")[0])
                    else:
                        values[p["name"]] = str(raw)
                except:
                    pass

            type_rows.append({"name": type_name, "values": values})

        snap = dict(self._snapshot)
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
                              unit_mode="inches",  # excel_bridge uses unit_prefs per-param for non-length
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

        n_types = len(imported_snapshot.get("types", []))
        self._set_status("Importing {} types...".format(n_types))
        self.statusbar.Refresh()
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
        """Right-click = column context menu. Left-click = sort."""
        if e.Button == MouseButtons.Right:
            # Set BOTH index vars — header menu uses _clicked_col_index,
            # cell-area handlers use _ctx_col_index. Keep them in sync.
            self._clicked_col_index = e.ColumnIndex
            self._ctx_col_index     = e.ColumnIndex
            pt = self.grid.PointToClient(System.Windows.Forms.Cursor.Position)
            self._col_ctx.Show(self.grid, pt)

    def _on_hide_col(self, s, a):
        """Hides the column that was right-clicked in the cell area."""
        try:
            ci = getattr(self, "_ctx_col_index", -1)
            if ci >= 0:
                self.grid.Columns[ci].Visible = False
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
            if ci < 1:
                return
            col = self.grid.Columns[ci]
            if col.Frozen:
                return
            new_idx = col.DisplayIndex - 1
            if new_idx >= 1:  # can't go before frozen col (DisplayIndex 0)
                col.DisplayIndex = new_idx
        except:
            pass

    def _on_col_move_right(self, s, a):
        """Moves the right-clicked column one position to the right."""
        try:
            ci = getattr(self, "_clicked_col_index", -1)
            if ci < 0:
                return
            col = self.grid.Columns[ci]
            if col.Frozen:
                return
            if col.DisplayIndex < self.grid.Columns.Count - 1:
                col.DisplayIndex = col.DisplayIndex + 1
        except:
            pass

    def _on_filter_hdr(self, s, a):
        """Show filter dropdown for right-clicked column header."""
        ci = getattr(self, "_ctx_col_index", getattr(self, "_clicked_col_index", -1))
        if ci >= 0:
            self._show_col_filter(ci, None)

    def _on_clear_col_filter(self, s, a):
        """Clear filter on the right-clicked column header."""
        ci = getattr(self, "_clicked_col_index", -1)
        if ci >= 0 and hasattr(self, "_col_filters"):
            self._col_filters.pop(ci, None)
            self._apply_col_filters()

    def _on_filter_cell_col(self, s, a):
        """Show filter dropdown for the column of the right-clicked cell."""
        try:
            ci = getattr(self, "_ctx_col_index", -1)
            if ci < 0 and self.grid.CurrentCell is not None:
                ci = self.grid.CurrentCell.ColumnIndex
            if ci <= 0:
                self.txtGridFilter.Focus()
                return
            self._show_col_filter(ci, None)
        except Exception as ex:
            self._set_status("Filter error: " + str(ex))

    def _on_clear_cell_col_filter(self, s, a):
        """Clear filter on the column of the right-clicked cell."""
        try:
            ci = getattr(self, "_ctx_col_index", -1)
            if ci < 0 and self.grid.CurrentCell is not None:
                ci = self.grid.CurrentCell.ColumnIndex
            if ci >= 0 and hasattr(self, "_col_filters"):
                self._col_filters.pop(ci, None)
                self._apply_col_filters()
        except:
            pass

    def _on_clear_all_filters(self, s, a):
        """Clear all active column filters."""
        if hasattr(self, "_col_filters"):
            self._col_filters.clear()
        self._apply_col_filters()

    def _on_units_prefs(self, s, a):
        """Open the unit preferences dialog. Refreshes grid if user applies changes."""
        try:
            import unit_prefs_dialog as upd
            changed = upd.show_prefs_dialog(self)
            if changed and self._snapshot:
                self._refresh_grid()
        except Exception as ex:
            import traceback
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(TOOL_TITLE, "Units dialog error:\n" + str(ex) + "\n\n" + traceback.format_exc())

    def _on_manage_columns(self, s, a):
        """
        Column manager: filter, check/uncheck, move up/down.
        Uses a separate checked-state dict so filtering never loses state.
        """
        frm = Form()
        frm.Text = "Show / Hide / Reorder Columns"
        frm.Size = Size(460, 580)
        frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False

        # Build full ordered list sorted by current DisplayIndex
        cols_ordered = sorted(
            [self.grid.Columns[i] for i in range(self.grid.Columns.Count)],
            key=lambda c: c.DisplayIndex
        )

        def _base_hdr(hdr):
            """Strip unit suffix e.g. ' (in)' for stable key."""
            h = hdr
            if " (" in h and h.endswith(")"):
                h = h[:h.rfind(" (")]
            return h

        # State and order use BASE names (without unit suffix) for stability
        _state = {}
        for col in cols_ordered:
            _state[_base_hdr(col.HeaderText)] = col.Visible

        _order = [_base_hdr(col.HeaderText) for col in cols_ordered]

        lbl = Label()
        lbl.Text = "Check = visible. Select row + Up/Down to reorder."
        lbl.Location = Point(12, 8); lbl.Size = Size(420, 18)
        lbl.Font = DFont("Arial", 9)

        lbl_srch = Label(); lbl_srch.Text = "Filter:"
        lbl_srch.Location = Point(12, 32); lbl_srch.Size = Size(40, 18)
        lbl_srch.Font = DFont("Arial", 9)

        txt_srch = TextBox()
        txt_srch.Location = Point(56, 30); txt_srch.Size = Size(240, 20)
        txt_srch.Font = DFont("Arial", 9)

        clb = CheckedListBox()
        clb.Location = Point(12, 56); clb.Size = Size(300, 430)
        clb.CheckOnClick = True
        clb.Font = DFont("Arial", 9)

        def _refresh_list():
            """Rebuild clb from _order/_state, filtered by search text."""
            needle = (txt_srch.Text or "").strip().lower()
            clb.ItemCheck -= _on_item_check   # detach while rebuilding
            clb.Items.Clear()
            for header in _order:
                if needle and needle not in header.lower():
                    continue
                idx = clb.Items.Add(header)
                clb.SetItemChecked(idx, _state.get(header, True))
            clb.ItemCheck += _on_item_check   # reattach

        def _on_item_check(s2, e2):
            """Sync clb check state back to _state dict immediately."""
            try:
                header = str(clb.Items[e2.Index])
                # e2.NewValue is the state AFTER the click
                from System.Windows.Forms import CheckState
                _state[header] = (e2.NewValue == CheckState.Checked)
            except:
                pass

        clb.ItemCheck += _on_item_check
        txt_srch.TextChanged += lambda s2, a2: _refresh_list()
        _refresh_list()

        btnUp   = Button(); btnUp.Text   = "Up";   btnUp.Size = Size(100, 28); btnUp.Location   = Point(322, 56)
        btnDown = Button(); btnDown.Text = "Down"; btnDown.Size = Size(100, 28); btnDown.Location = Point(322, 90)
        btnAll  = Button(); btnAll.Text  = "All";  btnAll.Size = Size(48, 24);  btnAll.Location  = Point(322, 140)
        btnNone = Button(); btnNone.Text = "None"; btnNone.Size = Size(48, 24); btnNone.Location = Point(374, 140)

        def _move(delta):
            """Move selected item in _order list, then refresh."""
            sel_text = str(clb.SelectedItem) if clb.SelectedIndex >= 0 else None
            if not sel_text:
                return
            # Find in master _order list
            try:
                i = _order.index(sel_text)
            except ValueError:
                return
            j = i + delta
            # Clamp: can't move before index 1 (index 0 = Type Name frozen)
            if j < 1 or j >= len(_order):
                return
            _order[i], _order[j] = _order[j], _order[i]
            _refresh_list()
            # Re-select moved item
            needle = (txt_srch.Text or "").strip().lower()
            for k in range(clb.Items.Count):
                if str(clb.Items[k]) == sel_text:
                    clb.SelectedIndex = k
                    break

        def _set_all(checked):
            for h in _state:
                _state[h] = checked
            _refresh_list()

        btnUp.Click   += lambda s2, a2: _move(-1)
        btnDown.Click += lambda s2, a2: _move(1)
        btnAll.Click  += lambda s2, a2: _set_all(True)
        btnNone.Click += lambda s2, a2: _set_all(False)

        btnApply  = Button(); btnApply.Text  = "Apply";  btnApply.Size = Size(90, 28)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Size = Size(90, 28)
        btnApply.Location  = Point(240, 496)
        btnCancel.Location = Point(336, 496)

        def _apply(s2, a2):
            """Apply visibility from _state and display order from _order."""
            display_idx = 1  # start after frozen Type Name (index 0)
            for header in _order:
                # Match by stored header text (exact match from when dialog opened)
                for col in [self.grid.Columns[j] for j in range(self.grid.Columns.Count)]:
                    if col.Frozen:
                        continue
                    # Match against current header OR base name (strip unit suffix)
                    cur_hdr = col.HeaderText
                    # Strip unit hint e.g. " (in)" or " (lbs)" for comparison
                    base_hdr = cur_hdr
                    if " (" in base_hdr and base_hdr.endswith(")"):
                        base_hdr = base_hdr[:base_hdr.rfind(" (")]
                    if cur_hdr == header or base_hdr == header:
                        visible = _state.get(header, True)
                        col.Visible = visible
                        if visible:
                            # Only set DisplayIndex for visible columns
                            try:
                                col.DisplayIndex = display_idx
                            except:
                                pass
                            display_idx += 1
                        break
            frm.Close()

        btnApply.Click  += _apply
        btnCancel.Click += lambda s2, a2: frm.Close()

        for w in [lbl, lbl_srch, txt_srch, clb,
                  btnUp, btnDown, btnAll, btnNone, btnApply, btnCancel]:
            frm.Controls.Add(w)
        frm.ShowDialog()


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
                col = self.grid.Columns[col_idx]
                p   = col.Tag
                for i, ri in enumerate(row_indices):
                    val = start + i * step
                    actual_ri = self._actual_row(ri)
                    if p and p.get("storage_type") == "Double":
                        disp_val = "{:.4f}".format(val)
                    elif p and p.get("storage_type") == "Integer":
                        disp_val = str(int(round(val)))
                    else:
                        disp_val = str(val)
                    self._grid_modified[(actual_ri, col_idx)] = disp_val
                    self._modified_rows.add(actual_ri)
                self._dirty = True
                self.grid.Invalidate()
            except Exception as ex:
                MessageBox.Show("Fill error: " + str(ex), TOOL_TITLE,
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)
            frm.Close()

        btnOK.Click += _fill
        btnCl.Click += lambda x, y: frm.Close()
        frm.ShowDialog()

    def _on_fill_right(self, s, a):
        """Ctrl+R: fills selected cell value across row (virtual mode)."""
        sel_cells = list(self.grid.SelectedCells)
        if not sel_cells:
            return
        sel_cells.sort(key=lambda c: (c.RowIndex, c.ColumnIndex))
        anchor    = sel_cells[0]
        value     = anchor.Value
        anchor_ri = anchor.RowIndex
        actual_ri = self._actual_row(anchor_ri)
        for cell in sel_cells:
            if cell.RowIndex == anchor_ri:
                self._grid_modified[(actual_ri, cell.ColumnIndex)] = value
        self._dirty = True
        self.grid.InvalidateRow(anchor_ri)

    def _parse_double(self, p, raw_str, ignored_scale):
        """
        Converts a display string to Revit internal units.
        For length params: accepts all Revit-style input (2'-6", 2 6, 6", etc.)
        For other params: passes display value through unit_prefs.to_internal.
        """
        if not raw_str or str(raw_str).strip() == "":
            return 0.0
        spec = p.get("spec_type_id", "")
        s    = str(raw_str).strip()
        try:
            import unit_prefs as up
            if up.is_length(spec):
                # Use full Revit-style parser -- handles all formats
                feet = _parse_length_input(s)
                if feet is not None:
                    return up.to_internal(feet, spec)
                # Fallback: plain decimal through unit_prefs
                return up.to_internal(float(s.replace(",", "")), spec)
            else:
                # Non-length: plain decimal through unit_prefs
                return up.to_internal(float(s.replace(",", "")), spec)
        except:
            try:
                return float(s.replace(",", ""))
            except:
                return 0.0

    # ======================================================================= #
    # FORMULA BAR                                                              #
    # ======================================================================= #

    def _on_selection_changed(self, s, a):
        """Update formula bar when the selected cell changes."""
        try:
            cells = list(self.grid.SelectedCells)
            if not cells:
                self.txtFormulaBar.Text = ""
                self.txtFormulaBar.ReadOnly = True
                return
            cell = sorted(cells, key=lambda c: (c.RowIndex, c.ColumnIndex))[0]
            ri   = cell.RowIndex
            ci   = cell.ColumnIndex

            # Get current value (modified or base)
            actual_ri = ri
            if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
                if ri < len(self._grid_filter_map):
                    actual_ri = self._grid_filter_map[ri]
            ov = self._grid_modified.get((actual_ri, ci))
            if ov is not None:
                val = str(ov)
            elif hasattr(self, "_grid_data") and actual_ri < len(self._grid_data) and ci < len(self._grid_data[actual_ri]):
                val = str(self._grid_data[actual_ri][ci] or "")
            else:
                val = ""

            self.txtFormulaBar.Text = val

            # Make editable unless it's a formula param
            col = self.grid.Columns[ci] if ci < self.grid.Columns.Count else None
            p   = col.Tag if col else None
            is_formula = p and (p.get("has_formula") or p.get("read_only"))
            self.txtFormulaBar.ReadOnly = bool(is_formula)

            # Update column statistics in status bar
            self._update_col_stats(ci)

        except:
            self.txtFormulaBar.Text = ""

    def _on_formula_bar_keydown(self, s, e):
        """Enter in formula bar commits the value to the current cell."""
        if e.KeyCode == Keys.Return:
            self._on_formula_bar_commit(None, None)
            e.Handled = True

    def _on_formula_bar_commit(self, s, a):
        """
        Commit formula bar value back to the current selected cell.
        Supports basic math: type =2*6+1 or =12/3 to evaluate before storing.
        Supports Revit dimension input: 2'-6", 2 6, etc.
        """
        try:
            if self.txtFormulaBar.ReadOnly:
                return
            cells = list(self.grid.SelectedCells)
            if not cells:
                return
            cell = sorted(cells, key=lambda c: (c.RowIndex, c.ColumnIndex))[0]
            ri   = cell.RowIndex
            ci   = cell.ColumnIndex
            actual_ri = ri
            if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
                if ri < len(self._grid_filter_map):
                    actual_ri = self._grid_filter_map[ri]

            raw_val = self.txtFormulaBar.Text.strip()

            # Math expression: starts with = and contains only safe math chars
            if raw_val.startswith("="):
                expr = raw_val[1:].strip()
                new_val = _eval_math_expr(expr, raw_val)
            else:
                new_val = raw_val

            # Update formula bar to show the evaluated result
            self.txtFormulaBar.Text = new_val

            self._grid_modified[(actual_ri, ci)] = new_val
            self._modified_rows.add(actual_ri)
            self._dirty = True
            self._rebuild_row_colors()
            self.grid.InvalidateRow(ri)

            if self.chkAutoApply.Checked:
                self._auto_apply_cell(actual_ri, ci, new_val)
        except:
            pass

    def _auto_apply_cell(self, actual_ri, ci, display_val):
        """
        Instantly writes a single cell edit to Revit without the preview dialog.
        Creates its own named transaction so Revit's undo stack shows a useful name.
        """
        try:
            if self._snapshot is None or not hasattr(self, "_grid_data"):
                return
            # Get type name
            tn_override = self._grid_modified.get((actual_ri, 0))
            type_name   = str(tn_override if tn_override is not None
                              else self._grid_data[actual_ri][0]).strip()

            # Get param def from column
            col = self.grid.Columns[ci] if ci < self.grid.Columns.Count else None
            p   = col.Tag if col else None
            if not p or p.get("has_formula") or p.get("read_only"):
                return

            # Convert display value to internal
            st = p.get("storage_type", "String")
            if st == "Double":
                raw_str   = str(display_val).replace(",", "").strip()
                int_val   = self._parse_double(p, raw_str, 1.0)
            elif st == "Integer":
                sv = str(display_val).strip().lower()
                int_val   = 1 if sv in ("yes","true") else (0 if sv in ("no","false") else int(sv.split(".")[0]))
            else:
                int_val   = str(display_val)

            # Find the FamilyType and FamilyParameter
            fm = self.doc.FamilyManager
            ft = None
            for t in fm.Types:
                try:
                    if t.Name == type_name:
                        ft = t; break
                except:
                    pass
            if ft is None:
                return

            fp = None
            for param in fm.Parameters:
                try:
                    if param.Definition.Name == p["name"]:
                        fp = param; break
                except:
                    pass
            if fp is None:
                return

            # Write in a named transaction
            from Autodesk.Revit.DB import Transaction, StorageType
            tx_name = "Edit {} -> {}".format(p["name"], type_name)
            t = Transaction(self.doc, tx_name)
            t.Start()
            try:
                # Placeholder pattern if ft is active type
                active_type = None
                try:
                    active_type = fm.CurrentType
                except:
                    pass
                temp_ft = None
                if active_type and active_type.Name == type_name:
                    try:
                        temp_ft = fm.NewType("__AUTOAPPLY_TEMP__")
                        fm.CurrentType = temp_ft
                    except:
                        temp_ft = None

                fm.CurrentType = ft
                core._write_param_value(fm, fp, int_val)

                if temp_ft is not None:
                    try:
                        fm.CurrentType = temp_ft
                        fm.DeleteCurrentType()
                        fm.CurrentType = ft
                    except:
                        pass

                t.Commit()
                self._set_status("Applied: {} = {} on {}".format(p["name"], display_val, type_name))
            except Exception as ex:
                t.RollBack()
                self._set_status("Auto-apply failed: " + str(ex))
        except Exception as ex:
            self._set_status("Auto-apply error: " + str(ex))

    # ======================================================================= #
    # FIND & REPLACE                                                            #
    # ======================================================================= #

    def _on_find_replace(self, s, a):
        """Find & Replace dialog — searches all visible, non-formula cells."""
        if not hasattr(self, "_grid_data"):
            return

        frm = Form()
        frm.Text = "Find & Replace"
        frm.Size = Size(460, 280)
        frm.StartPosition = FormStartPosition.CenterParent
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False

        def _lbl(text, x, y):
            l = Label(); l.Text = text
            l.Location = Point(x, y); l.Size = Size(90, 18)
            l.Font = DFont("Arial", 9); frm.Controls.Add(l)

        _lbl("Find:", 12, 16)
        txt_find = TextBox()
        txt_find.Location = Point(106, 14); txt_find.Size = Size(320, 22)
        txt_find.Font = DFont("Arial", 9)
        frm.Controls.Add(txt_find)

        _lbl("Replace with:", 12, 46)
        txt_replace = TextBox()
        txt_replace.Location = Point(106, 44); txt_replace.Size = Size(320, 22)
        txt_replace.Font = DFont("Arial", 9)
        frm.Controls.Add(txt_replace)

        # Options
        chk_case = CheckBox(); chk_case.Text = "Match case"
        chk_case.Location = Point(106, 76); chk_case.Size = Size(120, 20)
        chk_case.Font = DFont("Arial", 9)
        frm.Controls.Add(chk_case)

        chk_whole = CheckBox(); chk_whole.Text = "Whole cell only"
        chk_whole.Location = Point(236, 76); chk_whole.Size = Size(130, 20)
        chk_whole.Font = DFont("Arial", 9)
        frm.Controls.Add(chk_whole)

        # Scope: selected column only vs all columns
        chk_col = CheckBox(); chk_col.Text = "Selected column only"
        chk_col.Location = Point(106, 100); chk_col.Size = Size(160, 20)
        chk_col.Font = DFont("Arial", 9)
        frm.Controls.Add(chk_col)

        lbl_result = Label()
        lbl_result.Location = Point(12, 134); lbl_result.Size = Size(420, 18)
        lbl_result.Font = DFont("Arial", 9); lbl_result.ForeColor = Color.DarkGreen
        frm.Controls.Add(lbl_result)

        btnReplace = Button(); btnReplace.Text = "Replace All"
        btnReplace.Size = Size(110, 28); btnReplace.Location = Point(106, 160)
        btnFindNext = Button(); btnFindNext.Text = "Find Next"
        btnFindNext.Size = Size(100, 28); btnFindNext.Location = Point(224, 160)
        btnClose = Button(); btnClose.Text = "Close"
        btnClose.Size = Size(80, 28); btnClose.Location = Point(332, 160)
        btnClose.Click += lambda sv, av: frm.Close()
        frm.Controls.Add(btnReplace); frm.Controls.Add(btnFindNext); frm.Controls.Add(btnClose)

        # Track find position for Find Next
        _find_pos = [0, 0]  # [row, col]

        def _get_scope_cols():
            """Returns list of column indices to search."""
            sel_cells = list(self.grid.SelectedCells)
            sel_cols  = set(c.ColumnIndex for c in sel_cells)
            if chk_col.Checked and sel_cols:
                return [ci for ci in sel_cols if ci > 0]
            return list(range(1, self.grid.Columns.Count))

        def _cell_matches(val, needle):
            if val is None:
                return False
            sv = str(val)
            nd = needle
            if not chk_case.Checked:
                sv = sv.lower(); nd = nd.lower()
            return sv == nd if chk_whole.Checked else nd in sv

        def _find_next(s2, a2):
            needle = txt_find.Text
            if not needle:
                return
            scope_cols = _get_scope_cols()
            n_rows = len(self._grid_data)
            start_r, start_c_idx = _find_pos[0], _find_pos[1]
            # Flatten search space
            cells = [(ri, ci) for ri in range(n_rows) for ci in scope_cols]
            # Rotate from current position
            start_flat = next((i for i, (r, c) in enumerate(cells)
                               if r > start_r or (r == start_r and c >= start_c_idx)), 0)
            search_order = cells[start_flat:] + cells[:start_flat]
            for ri, ci in search_order:
                ov = self._grid_modified.get((ri, ci))
                val = ov if ov is not None else (self._grid_data[ri][ci] if ci < len(self._grid_data[ri]) else "")
                if _cell_matches(val, needle):
                    # Scroll to and select this cell
                    display_ri = ri
                    if hasattr(self, "_grid_filter_map") and self._grid_filter_map:
                        if ri not in self._grid_filter_map:
                            continue
                        display_ri = self._grid_filter_map.index(ri)
                    try:
                        self.grid.CurrentCell = self.grid.Rows[display_ri].Cells[ci]
                        self.grid.FirstDisplayedScrollingRowIndex = max(0, display_ri - 3)
                    except:
                        pass
                    _find_pos[0] = ri
                    _find_pos[1] = ci + 1
                    lbl_result.Text = "Found at row: " + str(ri + 1)
                    lbl_result.ForeColor = Color.DarkGreen
                    return
            lbl_result.Text = "Not found."
            lbl_result.ForeColor = Color.Red
            _find_pos[0] = 0; _find_pos[1] = 0

        def _replace_all(s2, a2):
            needle  = txt_find.Text
            replace = txt_replace.Text
            if not needle:
                return
            scope_cols = _get_scope_cols()
            count = 0
            for ri in range(len(self._grid_data)):
                for ci in scope_cols:
                    col = self.grid.Columns[ci] if ci < self.grid.Columns.Count else None
                    p2  = col.Tag if col else None
                    if p2 and (p2.get("has_formula") or p2.get("read_only")):
                        continue
                    ov  = self._grid_modified.get((ri, ci))
                    val = ov if ov is not None else (self._grid_data[ri][ci] if ci < len(self._grid_data[ri]) else "")
                    sv  = str(val or "")
                    nd  = needle
                    if not chk_case.Checked:
                        match = nd.lower() in sv.lower() if not chk_whole.Checked else nd.lower() == sv.lower()
                    else:
                        match = nd in sv if not chk_whole.Checked else nd == sv
                    if match:
                        if chk_whole.Checked:
                            new_val = replace
                        else:
                            if chk_case.Checked:
                                new_val = sv.replace(needle, replace)
                            else:
                                # Case-insensitive replace
                                import re as _re3
                                new_val = _re3.sub(_re3.escape(needle), replace, sv, flags=_re3.IGNORECASE)
                        self._grid_modified[(ri, ci)] = new_val
                        self._modified_rows.add(ri)
                        count += 1
            if count:
                self._dirty = True
                self.grid.Invalidate()
            lbl_result.ForeColor = Color.DarkGreen if count else Color.Red
            lbl_result.Text = "Replaced {} cell(s).".format(count) if count else "Not found."

        btnReplace.Click  += _replace_all
        btnFindNext.Click += _find_next
        txt_find.KeyDown  += lambda sv, ev: _find_next(None, None) if ev.KeyCode == Keys.Return else None
        frm.ShowDialog()

    # ======================================================================= #
    # COLUMN STATISTICS                                                        #
    # ======================================================================= #

    def _update_col_stats(self, col_index):
        """
        Computes and displays statistics for the selected column in the status bar.
        Shows: Count | Numeric: Min / Max / Sum / Avg
        """
        try:
            if not hasattr(self, "_grid_data") or col_index < 1:
                self.lblColStats.Text = ""
                return
            col = self.grid.Columns[col_index] if col_index < self.grid.Columns.Count else None
            p   = col.Tag if col else None

            n_rows  = self.grid.RowCount
            values  = []
            strings = []
            for display_ri in range(n_rows):
                actual_ri = display_ri
                if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
                    if display_ri < len(self._grid_filter_map):
                        actual_ri = self._grid_filter_map[display_ri]
                    else:
                        continue
                ov  = self._grid_modified.get((actual_ri, col_index))
                row = self._grid_data[actual_ri] if actual_ri < len(self._grid_data) else []
                raw = ov if ov is not None else (row[col_index] if col_index < len(row) else "")
                sv  = str(raw or "").strip()
                if sv:
                    strings.append(sv)
                    try:
                        values.append(float(sv.replace(",", "")))
                    except:
                        pass

            count = len(strings)
            if not count:
                self.lblColStats.Text = ""
                return

            col_name = col.HeaderText if col else ""
            if values and len(values) == count:
                # All numeric
                mn  = min(values)
                mx  = max(values)
                sm  = sum(values)
                avg = sm / len(values)
                self.lblColStats.Text = (
                    u"{}: Count={} | Min={:.4f} | Max={:.4f} | Sum={:.4f} | Avg={:.4f}"
                ).format(col_name, count, mn, mx, sm, avg)
            elif values:
                # Mixed
                mn  = min(values)
                mx  = max(values)
                self.lblColStats.Text = (
                    u"{}: Count={} | Numeric={} | Min={:.4f} | Max={:.4f}"
                ).format(col_name, count, len(values), mn, mx)
            else:
                # All strings — show count and unique count
                unique = len(set(strings))
                self.lblColStats.Text = (
                    u"{}: Count={} | Unique={}"
                ).format(col_name, count, unique)

            self.statusbar.Refresh()
        except:
            self.lblColStats.Text = ""

    # ======================================================================= #
    # UNDO / REDO                                                              #
    # ======================================================================= #

    def _on_undo(self, s, a):
        """Undo the last grid edit."""
        if not hasattr(self, "_undo_stack") or not self._undo_stack:
            self._set_status("Nothing to undo.")
            return
        ri, ci, old_val, new_val = self._undo_stack.pop()
        self._redo_stack.append((ri, ci, old_val, new_val))
        # Restore old value
        if old_val is None or str(old_val).strip() == "":
            self._grid_modified.pop((ri, ci), None)
        else:
            self._grid_modified[(ri, ci)] = old_val
        # Update _modified_rows: keep row marked if it still has other edits
        row_still_edited = any(
            k[0] == ri for k in self._grid_modified
        )
        if row_still_edited:
            self._modified_rows.add(ri)
        else:
            self._modified_rows.discard(ri)
            self._modified_rows.add(ri)
        # Find display row and invalidate
        display_ri = ri
        if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
            if ri in self._grid_filter_map:
                display_ri = self._grid_filter_map.index(ri)
        try:
            self.grid.InvalidateRow(display_ri)
        except:
            pass
        self.grid.Refresh()
        self._set_status("Undone. ({} left)".format(len(self._undo_stack)))
        # Update formula bar if this cell is selected
        try:
            cells = list(self.grid.SelectedCells)
            if cells:
                self._on_selection_changed(None, None)
        except:
            pass

    def _on_redo(self, s, a):
        """Redo the last undone edit."""
        if not hasattr(self, "_redo_stack") or not self._redo_stack:
            self._set_status("Nothing to redo.")
            return
        ri, ci, old_val, new_val = self._redo_stack.pop()
        self._undo_stack.append((ri, ci, old_val, new_val))
        self._grid_modified[(ri, ci)] = new_val
        self._modified_rows.add(ri)
        display_ri = ri
        if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None:
            if ri in self._grid_filter_map:
                display_ri = self._grid_filter_map.index(ri)
        try:
            self.grid.InvalidateRow(display_ri)
        except:
            pass
        self.grid.Refresh()
        self._set_status("Redone. ({} undos left)".format(len(self._undo_stack)))
        try:
            cells = list(self.grid.SelectedCells)
            if cells:
                self._on_selection_changed(None, None)
        except:
            pass

    # ======================================================================= #
    # AUTOFILTER (column header dropdowns)                                     #
    # ======================================================================= #

    def _show_col_filter(self, col_index, screen_pt):
        """
        Shows a filter dropdown for the given column at screen_pt.
        Builds a checklist of unique values; unchecking hides those rows.
        Works alongside the text search filter — both can be active.
        """
        if not hasattr(self, "_grid_data"):
            return

        col = self.grid.Columns[col_index] if col_index < self.grid.Columns.Count else None
        if not col:
            return

        # Collect unique values for this column across ALL data (not just filtered)
        seen = {}
        for ri, row in enumerate(self._grid_data):
            ov  = self._grid_modified.get((ri, col_index))
            val = ov if ov is not None else (row[col_index] if col_index < len(row) else "")
            sv  = str(val or "").strip()
            if sv not in seen:
                seen[sv] = 0
            seen[sv] += 1

        if not seen:
            return

        # Init col_filters if needed
        if not hasattr(self, "_col_filters"):
            self._col_filters = {}  # col_index -> set of HIDDEN values

        hidden = self._col_filters.get(col_index, set())

        # Build dropdown form
        frm = Form()
        frm.Text            = "Filter: " + col.HeaderText
        frm.Size            = Size(340, 580)
        frm.MinimumSize     = Size(300, 480)
        frm.StartPosition   = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.Sizable
        frm.MaximizeBox     = False
        frm.TopMost         = True

        lbl = Label(); lbl.Text = "Check = show rows with this value"
        lbl.Location = Point(8, 6); lbl.Size = Size(310, 16)
        lbl.Font = DFont("Arial", 8)
        lbl.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        # Search within filter dropdown
        txt = TextBox(); txt.Location = Point(8, 24); txt.Size = Size(310, 22)
        txt.Font = DFont("Arial", 9)
        txt.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        clb = CheckedListBox()
        clb.Location = Point(8, 50)
        clb.Size     = Size(310, 420)
        clb.Anchor   = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        clb.CheckOnClick = True
        clb.Font = DFont("Arial", 9)

        _all_items = sorted(seen.keys(), key=lambda x: (x == "", x.lower()))
        _state = {v: (v not in hidden) for v in _all_items}

        def _rebuild(needle=""):
            clb.Items.Clear()
            for v in _all_items:
                if needle and needle.lower() not in v.lower():
                    continue
                lbl_txt = "(blank)" if v == "" else "{} ({})".format(v, seen[v])
                idx = clb.Items.Add(lbl_txt)
                clb.SetItemChecked(idx, _state.get(v, True))

        def _item_check(sv, ev):
            try:
                raw_lbl = str(clb.Items[ev.Index])
                # Extract the original value from "value (count)" format
                orig_v = "(blank)" if raw_lbl.startswith("(blank)") else raw_lbl.rsplit(" (", 1)[0]
                if orig_v == "(blank)":
                    orig_v = ""
                from System.Windows.Forms import CheckState
                _state[orig_v] = (ev.NewValue == CheckState.Checked)
            except:
                pass

        clb.ItemCheck += _item_check
        txt.TextChanged += lambda sv, av: _rebuild(txt.Text or "")
        _rebuild()

        btnAll   = Button(); btnAll.Text   = "All";          btnAll.Size   = Size(64, 26)
        btnNone  = Button(); btnNone.Text  = "None";         btnNone.Size  = Size(64, 26)
        btnOK    = Button(); btnOK.Text    = "Apply";        btnOK.Size    = Size(84, 28)
        btnCl    = Button(); btnCl.Text    = "Clear Filter"; btnCl.Size    = Size(110, 28)
        btnClose = Button(); btnClose.Text = "Cancel";       btnClose.Size = Size(84, 28)

        # No anchors — _layout_btns owns all positions/sizes
        def _layout_btns(sv=None, av=None):
            h = frm.ClientSize.Height
            w = frm.ClientSize.Width
            # Checklist fills middle
            clb.Location = Point(8, 50)
            clb.Size     = Size(w - 16, max(60, h - 112))
            # Bottom row 1: All / None / Apply
            y1 = h - 60
            btnAll.Location   = Point(8,      y1)
            btnNone.Location  = Point(78,     y1)
            btnOK.Location    = Point(w - 92, y1)
            # Bottom row 2: Clear Filter / Cancel
            y2 = h - 30
            btnCl.Location    = Point(8,      y2)
            btnClose.Location = Point(w - 92, y2)

        frm.Resize += _layout_btns

        btnAll.Click  += lambda sv, av: [_state.update({v: True  for v in _all_items}), _rebuild(txt.Text or "")]
        btnNone.Click += lambda sv, av: [_state.update({v: False for v in _all_items}), _rebuild(txt.Text or "")]

        def _apply(sv, av):
            new_hidden = set(v for v, show in _state.items() if not show)
            if new_hidden:
                self._col_filters[col_index] = new_hidden
            else:
                self._col_filters.pop(col_index, None)
            self._apply_col_filters()
            frm.Close()

        def _clear(sv, av):
            self._col_filters.pop(col_index, None)
            self._apply_col_filters()
            frm.Close()

        btnOK.Click    += _apply
        btnCl.Click    += _clear
        btnClose.Click += lambda sv, av: frm.Close()

        for w in [lbl, txt, clb, btnAll, btnNone, btnOK, btnCl, btnClose]:
            frm.Controls.Add(w)
        _layout_btns()   # initial layout before form is shown
        frm.ShowDialog()

    def _apply_col_filters(self):
        """
        Rebuilds _grid_filter_map combining text search AND column filters.
        """
        if not hasattr(self, "_grid_data"):
            return
        needle = (self.txtGridFilter.Text or "").strip().lower()
        col_filters = getattr(self, "_col_filters", {})

        if not needle and not col_filters:
            self._grid_filter_map = None
            self.grid.RowCount = len(self._grid_data)
            self.grid.Invalidate()
            return

        fmap = []
        for i, row in enumerate(self._grid_data):
            # Text search on type name (col 0)
            if needle and needle not in str(row[0] or "").lower():
                continue
            # Column filters
            excluded = False
            for ci, hidden_vals in col_filters.items():
                ov  = self._grid_modified.get((i, ci))
                val = ov if ov is not None else (row[ci] if ci < len(row) else "")
                sv  = str(val or "").strip()
                if sv in hidden_vals or ("" in hidden_vals and not sv):
                    excluded = True
                    break
            if not excluded:
                fmap.append(i)

        self._grid_filter_map = fmap if (needle or col_filters) else None
        self.grid.RowCount = len(fmap) if self._grid_filter_map is not None else len(self._grid_data)
        self._rebuild_row_colors()
        self.grid.Refresh()

        # Update column header ▼ indicators
        self._update_filter_indicators()

    def _update_filter_indicators(self):
        """Marks filtered columns with ▼ indicator in header text."""
        if not hasattr(self, "_col_filters"):
            return
        import re as _re_ind
        for i in range(self.grid.Columns.Count):
            col = self.grid.Columns[i]
            p   = col.Tag
            if p is None:
                continue
            # Strip any existing ▼ marker
            base_name = _re_ind.sub(r"\s*▼+$", "", col.HeaderText)
            if i in self._col_filters and self._col_filters[i]:
                col.HeaderText = base_name + u" ▼"
            else:
                col.HeaderText = base_name

    # ======================================================================= #
    # ROW DRAG-TO-REORDER                                                      #
    # ======================================================================= #

    def _on_grid_cell_mouse_down(self, s, e):
        """
        Fires before the context menu opens.
        Records exactly which cell was right-clicked so filter/hide
        handlers use the correct column instead of the stale CurrentCell.
        """
        try:
            if e.Button == MouseButtons.Right and e.RowIndex >= 0 and e.ColumnIndex >= 0:
                self._ctx_row_index = e.RowIndex
                self._ctx_col_index = e.ColumnIndex
                try:
                    self.grid.CurrentCell = self.grid.Rows[e.RowIndex].Cells[e.ColumnIndex]
                    self.grid.ClearSelection()
                    self.grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Selected = True
                except:
                    pass
        except:
            pass

    def _on_row_drag_start(self, s, e):
        """Start tracking a row drag when left-button pressed on row header."""
        try:
            from System.Windows.Forms import MouseButtons as MB
            if e.Button != MB.Left:
                return
            hit = self.grid.HitTest(e.X, e.Y)
            # Only drag from the row header (left narrow strip)
            from System.Windows.Forms import DataGridViewHitTestType
            if (hit.Type == DataGridViewHitTestType.RowHeader and
                    hit.RowIndex >= 0):
                self._row_drag_source = hit.RowIndex
                self.grid.Cursor = System.Windows.Forms.Cursors.SizeNS
        except:
            pass

    def _on_row_drag_move(self, s, e):
        """Drag highlight — invalidate only the two changed rows (prev + new)."""
        try:
            if self._row_drag_source is None:
                return
            hit = self.grid.HitTest(e.X, e.Y)
            from System.Windows.Forms import DataGridViewHitTestType
            new_hover = -1
            if hit.Type in (DataGridViewHitTestType.RowHeader,
                            DataGridViewHitTestType.Cell):
                if hit.RowIndex >= 0 and hit.RowIndex != self._row_drag_source:
                    new_hover = hit.RowIndex
            prev_hover = getattr(self, "_drag_hover_row", -1)
            if new_hover != prev_hover:
                self._drag_hover_row = new_hover if new_hover >= 0 else None
                # Update cache only for the two affected rows
                cache = getattr(self, "_row_color_cache", [])
                drag_hover = self._drag_hover_row
                for ri in [prev_hover, new_hover]:
                    if ri < 0 or ri >= len(cache):
                        continue
                    fmap = (self._grid_filter_map
                            if hasattr(self, "_grid_filter_map") and self._grid_filter_map is not None
                            else None)
                    actual_ri = fmap[ri] if fmap and ri < len(fmap) else ri
                    if drag_hover == ri:
                        cache[ri] = Color.FromArgb(180, 220, 255)
                    elif actual_ri in getattr(self, "_modified_rows", set()):
                        cache[ri] = COLOR_MODIFIED
                    elif actual_ri % 2 == 1:
                        cache[ri] = COLOR_ALT_ROW
                    else:
                        cache[ri] = Color.White
                    try:
                        self.grid.Rows[ri].DefaultCellStyle.BackColor = cache[ri]
                        self.grid.InvalidateRow(ri)
                    except:
                        pass
        except:
            pass

    def _on_row_drag_end(self, s, e):
        """Complete the row drag — reorder _grid_data."""
        try:
            src = self._row_drag_source
            self._row_drag_source = None
            self._drag_hover_row  = None
            self.grid.Cursor = System.Windows.Forms.Cursors.Default
            if src is None:
                return

            hit = self.grid.HitTest(e.X, e.Y)
            from System.Windows.Forms import DataGridViewHitTestType
            if hit.Type not in (DataGridViewHitTestType.RowHeader,
                                DataGridViewHitTestType.Cell):
                self.grid.Invalidate()
                return

            dst = hit.RowIndex
            if dst < 0 or dst == src or not hasattr(self, "_grid_data"):
                self.grid.Invalidate()
                return

            # Map display indices through filter map
            actual_src = self._actual_row(src)
            actual_dst = self._actual_row(dst)

            # Move the row in _grid_data
            row = self._grid_data.pop(actual_src)
            self._grid_data.insert(actual_dst, row)

            # Remap _grid_modified keys since row indices shifted
            new_modified = {}
            for (ri, ci), val in self._grid_modified.items():
                if ri == actual_src:
                    new_modified[(actual_dst, ci)] = val
                elif actual_src < actual_dst and actual_src < ri <= actual_dst:
                    new_modified[(ri - 1, ci)] = val
                elif actual_dst < actual_src and actual_dst <= ri < actual_src:
                    new_modified[(ri + 1, ci)] = val
                else:
                    new_modified[(ri, ci)] = val
            self._grid_modified = new_modified

            # Rebuild undo stack (best-effort — clear it on reorder to avoid confusion)
            if hasattr(self, "_undo_stack"):
                self._undo_stack = []
                self._redo_stack = []

            self._grid_filter_map = None
            self.grid.RowCount = len(self._grid_data)
            self._dirty = True
            self._rebuild_row_colors()
            self.grid.Invalidate()
            self._set_status("Row moved. Hit Apply Changes to save order to Revit.")
        except Exception as ex:
            self._row_drag_source = None
            self._drag_hover_row  = None
            self.grid.Cursor = System.Windows.Forms.Cursors.Default
            self._rebuild_row_colors()
            self.grid.Invalidate()

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