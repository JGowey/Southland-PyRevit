# -*- coding: utf-8 -*-
"""
SP Loader - Config Editor UI (ui_builder.py)
=============================================
WinForms UI for creating and editing parameter config JSON files.

Contains:
    - ParamRow:     Data model for a single parameter entry (serializes to/from JSON)
    - ParamEditor:  Detail editor for one parameter (SP picker, group, categories)
    - BuilderForm:  Config-level editor (display name, parameter list, save/load)
    - show_builder: Entry point called by script.py Launcher

Imported by:
    - script.py (lazy import — only when Create/Edit is clicked)

Depends on:
    - binder_core (paths, SP file access, group enumeration)
"""

import os
import json
import getpass
import clr

# .NET UI
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
from System import Array, String
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Application, Form, Label, TextBox, Button, CheckBox, ComboBox, ComboBoxStyle,
    ListView, ListViewItem, ColumnHeader, AnchorStyles, View,
    DialogResult, FormStartPosition, OpenFileDialog, TreeView, TreeNode,
    SplitContainer, RadioButton, CheckedListBox, Keys
)

# Local backend
import binder_core as core


# =============================================================================
# PARAMETER DATA MODEL
# =============================================================================

class ParamRow(object):
    """
    Data model for one parameter entry in a config.
    Maps 1:1 with the "parameters" array items in the JSON config file.
    """
    def __init__(self, name="", group_key="Construction", rfa_is_instance=True, rvt_is_instance=True, allow_vary=True, cats=None, per_spf=""):
        self.name = name or ""
        self.group_key = group_key or "Construction"
        self.rfa_is_instance = bool(rfa_is_instance)
        self.rvt_is_instance = bool(rvt_is_instance)
        self.allow_vary = bool(allow_vary)
        self.cats = list(cats or [])
        self.per_spf = per_spf or ""

    def to_dict(self):
        """Serialize to JSON-ready dict."""
        d = {
            "name": self.name,
            "source": "shared",
            "group_key": self.group_key,
            "rfa_is_instance": bool(self.rfa_is_instance),
            "rvt_is_instance": bool(self.rvt_is_instance),
            "rvt_allow_vary_between_groups": bool(self.allow_vary),
            "rvt_categories": list(self.cats)
        }
        if self.per_spf and os.path.exists(self.per_spf):
            d["shared_parameter_file"] = self.per_spf
        return d

    @staticmethod
    def from_dict(p):
        """Deserialize from a JSON dict."""
        return ParamRow(
            name=(p.get("name") or "").strip(),
            group_key=p.get("group_key", "Construction"),
            rfa_is_instance=bool(p.get("rfa_is_instance", True)),
            rvt_is_instance=bool(p.get("rvt_is_instance", True)),
            allow_vary=bool(p.get("rvt_allow_vary_between_groups", True)),
            cats=p.get("rvt_categories", []) or [],
            per_spf=p.get("shared_parameter_file", "") or ""
        )


# =============================================================================
# COMMON HELPERS
# =============================================================================

# Static fallback list — used only if the dynamic Revit API enumeration fails
GROUP_KEY_CHOICES = [
    "Analysis Results", "Analytical Alignment", "Analytical Model", "Constraints",
    "Construction", "Data", "Dimensions", "Division Geometry", "Electrical",
    "Electrical - Circuiting", "Electrical - Lighting", "Electrical - Loads",
    "Electrical Analysis", "Electrical Engineering", "Energy Analysis",
    "Fire Protection", "Forces", "General", "Graphics", "Green Building Properties",
    "Identity Data", "IFC Parameters", "Layers", "Life Safety",
    "Materials and Finishes", "Mechanical", "Mechanical - Flow", "Mechanical - Loads",
    "Model Properties", "Moments", "Other", "Overall Legend", "Phasing",
    "Photometrics", "Plumbing", "Primary End", "Rebar Set", "Releases / Member Forces",
    "Secondary End", "Segments and Fittings", "Set", "Slab Shape Edit",
    "Structural", "Structural Analysis", "Text", "Title Text", "Visibility"
]

def _doc():
    """Returns the active Revit document."""
    return __revit__.ActiveUIDocument.Document

def _category_names():
    """Returns a sorted list of all category names in the active document."""
    names = []
    try:
        for c in _doc().Settings.Categories:
            try: names.append(c.Name)
            except: pass
    except: pass
    names.sort(key=lambda s: s.lower())
    return names


# =============================================================================
# PARAMETER EDITOR FORM
# =============================================================================
# Edits a single ParamRow: pick parameter from SP tree, set group key,
# instance/type, vary, and select RVT categories.

class ParamEditor(Form):
    """Detail editor for one parameter entry."""

    def __init__(self, top_spf, row):
        Form.__init__(self)
        self.Text = "Parameter Editor"
        self.Size = Size(1190, 800)
        self.StartPosition = FormStartPosition.CenterScreen
        self.row = row or ParamRow()
        self.top_spf = top_spf or ""
        self._all_cat_names = _category_names()
        self._cat_checked = set(self.row.cats or [])

        # Guard flags: prevent CheckedListBox from toggling on scroll/keyboard
        # _programmatic_check bypasses the guard during _refresh_categories
        self._block_itemcheck = False
        self._last_input_mouse = False
        self._programmatic_check = False

        # --- Per-parameter SP file chooser ---
        lblSp = Label(Text="Per-parameter SP file (optional):")
        lblSp.Location = Point(12, 12); lblSp.Size = Size(250, 18)

        self.cboSp = ComboBox()
        self.cboSp.Location = Point(12, 32); self.cboSp.Size = Size(900, 22)
        self.cboSp.DropDownStyle = ComboBoxStyle.DropDown
        self.cboSp.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        btnBrowseSp = Button(Text="Browse...")
        btnBrowseSp.Location = Point(920, 31); btnBrowseSp.Size = Size(100, 24)
        btnBrowseSp.Anchor = AnchorStyles.Top | AnchorStyles.Right
        btnBrowseSp.Click += self._on_browse_sp

        for p in self._sp_candidates():
            self.cboSp.Items.Add(p)
        if self.row.per_spf:
            self.cboSp.Text = self.row.per_spf
        elif self.top_spf:
            self.cboSp.Text = self.top_spf

        # SP resolution note (anchored to bottom)
        lblNote = Label(Text="SP file resolution: Per-parameter > Top-level (Builder) > Revit current")
        lblNote.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        lblNote.Location = Point(12, self.ClientSize.Height - 30)
        lblNote.Size = Size(560, 16)
        def _relayout_note(_s=None, _a=None):
            lblNote.Location = Point(12, self.ClientSize.Height - 30)
        self.Resize += _relayout_note

        # --- Split panel: left = SP tree, right = settings + categories ---
        split = SplitContainer()
        split.Location = Point(12, 60)
        split.Size = Size(1128, 650)
        split.SplitterDistance = 450
        split.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.split = split

        # LEFT PANEL: Shared Parameter tree with filter and view mode
        lblLeft = Label(Text="Parameters in selected SP file:")
        lblLeft.Location = Point(6, 6); lblLeft.Size = Size(220, 18)
        split.Panel1.Controls.Add(lblLeft)

        self.txtParamFilter = TextBox()
        self.txtParamFilter.Location = Point(6, 26); self.txtParamFilter.Size = Size(440, 22)
        self.txtParamFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.txtParamFilter.TextChanged += self._refresh_tree
        split.Panel1.Controls.Add(self.txtParamFilter)

        self.rbGrouped = RadioButton(Text="Grouped")
        self.rbGrouped.Location = Point(6, 52); self.rbGrouped.Size = Size(80, 22)
        self.rbGrouped.Checked = True
        self.rbGrouped.CheckedChanged += self._refresh_tree
        split.Panel1.Controls.Add(self.rbGrouped)

        self.rbAZ = RadioButton(Text="A-Z")
        self.rbAZ.Location = Point(92, 52); self.rbAZ.Size = Size(60, 22)
        self.rbAZ.CheckedChanged += self._refresh_tree
        split.Panel1.Controls.Add(self.rbAZ)

        self.tv = TreeView()
        self.tv.Location = Point(6, 78); self.tv.Size = Size(454, 516)
        self.tv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.tv.AfterSelect += self._on_tree_select
        split.Panel1.Controls.Add(self.tv)

        # RIGHT PANEL: Selected parameter name, group, instance/type, categories
        lblSel = Label(Text="Selected parameter:")
        lblSel.Location = Point(6, 6); lblSel.Size = Size(140, 18)
        split.Panel2.Controls.Add(lblSel)

        self.txtSel = TextBox()
        self.txtSel.Location = Point(6, 26); self.txtSel.Size = Size(610, 22)
        self.txtSel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        split.Panel2.Controls.Add(self.txtSel)

        lblGroup = Label(Text="Group Key:")
        lblGroup.Location = Point(6, 56); lblGroup.Size = Size(120, 18)
        split.Panel2.Controls.Add(lblGroup)

        self.cboGroup = ComboBox()
        self.cboGroup.Location = Point(6, 76); self.cboGroup.Size = Size(260, 22)
        self.cboGroup.DropDownStyle = ComboBoxStyle.DropDownList
        # Populate from Revit API; store stable tokens in config (not UI labels)
        self._group_label_to_token = {}
        self._group_token_to_label = {}
        try:
            doc = __revit__.ActiveUIDocument.Document
            for (lab, tok) in core.list_user_assignable_group_choices(doc):
                self.cboGroup.Items.Add(lab)
                self._group_label_to_token[lab] = tok
                self._group_token_to_label[tok] = lab
        except:
            # Fallback to static list if API enumeration fails
            for g in GROUP_KEY_CHOICES:
                self.cboGroup.Items.Add(g)
                self._group_label_to_token[g] = g
                self._group_token_to_label[g] = g
        split.Panel2.Controls.Add(self.cboGroup)

        self.chkRfa = CheckBox(Text="RFA: Instance (unchecked = Type)")
        self.chkRfa.Location = Point(280, 76); self.chkRfa.Size = Size(220, 22)
        split.Panel2.Controls.Add(self.chkRfa)

        self.chkRvt = CheckBox(Text="RVT: Instance (unchecked = Type)")
        self.chkRvt.Location = Point(280, 100); self.chkRvt.Size = Size(220, 22)
        split.Panel2.Controls.Add(self.chkRvt)

        self.chkVary = CheckBox(Text="RVT: Allow values to vary by group instance")
        self.chkVary.Location = Point(280, 126); self.chkVary.Size = Size(320, 22)
        split.Panel2.Controls.Add(self.chkVary)

        lblCats = Label(Text="RVT Categories (check to include)")
        lblCats.Location = Point(6, 128); lblCats.Size = Size(260, 18)
        split.Panel2.Controls.Add(lblCats)

        self.chkShowOnly = CheckBox(Text="Show only selected")
        self.chkShowOnly.Location = Point(6, 152); self.chkShowOnly.Size = Size(150, 22)
        self.chkShowOnly.CheckedChanged += self._refresh_categories
        split.Panel2.Controls.Add(self.chkShowOnly)

        self.txtCatFilter = TextBox()
        self.txtCatFilter.Location = Point(6, 176)
        self.txtCatFilter.Size = Size(624, 22)
        self.txtCatFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.txtCatFilter.TextChanged += self._refresh_categories
        split.Panel2.Controls.Add(self.txtCatFilter)

        self.clbCats = CheckedListBox()
        self.clbCats.Location = Point(6, 202)
        self.clbCats.Size = Size(624, 392)
        self.clbCats.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.clbCats.CheckOnClick = False
        # Wire guards: only mouse click or Space key should toggle checks
        self.clbCats.MouseDown += self._on_cats_mouse_down
        self.clbCats.MouseUp += self._on_cats_mouse_up
        self.clbCats.KeyDown += self._on_cats_key_down
        self.clbCats.KeyUp += self._on_cats_key_up
        self.clbCats.ItemCheck += self._on_cat_check
        split.Panel2.Controls.Add(self.clbCats)

        # Add top-level controls to form
        self.Controls.Add(lblSp)
        self.Controls.Add(self.cboSp)
        self.Controls.Add(btnBrowseSp)
        self.Controls.Add(split)
        self.Controls.Add(lblNote)

        # Bottom buttons
        btnSave = Button(Text="Save")
        btnSave.Size = Size(100, 28)
        btnSave.Location = Point(self.Width - 300, self.Height - 80)
        btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnSave.Click += self._on_ok
        self.Controls.Add(btnSave)

        btnCancel = Button(Text="Cancel")
        btnCancel.Size = Size(100, 28)
        btnCancel.Location = Point(self.Width - 190, self.Height - 80)
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnCancel.Click += self._on_cancel
        self.Controls.Add(btnCancel)

        # Initialize form state from the row
        self._apply_row(self.row)
        self._refresh_tree(None, None)
        self._refresh_categories(None, None)

    # --- Category check guards ---
    # CheckedListBox fires ItemCheck on scroll, keyboard nav, and programmatic
    # SetItemChecked. These guards ensure only mouse clicks and Space key
    # actually toggle the check state.

    def _on_cats_mouse_down(self, *_):
        self._last_input_mouse = True
        self._block_itemcheck = False
    def _on_cats_mouse_up(self, *_):
        self._last_input_mouse = False
    def _on_cats_key_down(self, _s, e):
        self._last_input_mouse = False
        if e.KeyCode == Keys.Space:
            self._block_itemcheck = True
    def _on_cats_key_up(self, _s, e):
        if e.KeyCode == Keys.Space:
            self._block_itemcheck = False

    # --- SP file helpers ---

    def _sp_candidates(self):
        """Returns candidate SP file paths for the dropdown."""
        items = []
        if self.row.per_spf: items.append(self.row.per_spf)
        if self.top_spf and (self.top_spf not in items): items.append(self.top_spf)
        try:
            sp = __revit__.Application.SharedParametersFilename
            if sp and (sp not in items): items.append(sp)
        except: pass
        return items

    def _param_filter(self):
        """Returns the current parameter filter text (lowercased)."""
        return (self.txtParamFilter.Text or "").strip().lower()

    def _cat_filter(self):
        """Returns the current category filter text (lowercased)."""
        return (self.txtCatFilter.Text or "").strip().lower()

    def _apply_row(self, row):
        """Populates form controls from a ParamRow."""
        if row.name: self.txtSel.Text = row.name
        try:
            idx = 0
            for i in range(self.cboGroup.Items.Count):
                if self.cboGroup.Items[i] == (self._group_token_to_label.get(row.group_key, row.group_key) or "Construction"):
                    idx = i; break
            self.cboGroup.SelectedIndex = idx
        except: pass
        self.chkRfa.Checked = bool(row.rfa_is_instance)
        self.chkRvt.Checked = bool(row.rvt_is_instance)
        self.chkVary.Checked = bool(row.allow_vary)

    # --- SP Tree (left panel) ---

    def _refresh_tree(self, _s, _a):
        """Rebuilds the SP parameter tree from the selected SP file."""
        self.tv.BeginUpdate()
        try:
            self.tv.Nodes.Clear()
            sp = (self.cboSp.Text or "").strip()
            spf = core.open_shared_parameter_file(sp) if sp else None
            if not spf:
                self.tv.Nodes.Add("No shared parameter file available").EnsureVisible()
                return
            needle = self._param_filter()
            def _match(nm): return (not needle) or (needle in (nm or "").lower())
            if self.rbAZ.Checked:
                names = []
                for grp in spf.Groups:
                    for d in grp.Definitions:
                        try:
                            if _match(d.Name): names.append(d.Name)
                        except: pass
                for nm in sorted(set(names), key=lambda x: x.lower()):
                    self.tv.Nodes.Add(TreeNode(nm))
            else:
                for grp in spf.Groups:
                    try:
                        gnode = TreeNode(grp.Name); added = False
                        for d in grp.Definitions:
                            try:
                                if _match(d.Name):
                                    gnode.Nodes.Add(TreeNode(d.Name)); added = True
                            except: pass
                        if added: self.tv.Nodes.Add(gnode)
                    except: pass
            self.tv.ExpandAll()
        finally:
            self.tv.EndUpdate()

    def _on_tree_select(self, _s, args):
        """When a leaf node is selected in the SP tree, fill the parameter name."""
        node = args.Node
        if node is None: return
        if node.Nodes is not None and node.Nodes.Count > 0: return
        self.txtSel.Text = node.Text or ""

    # --- Category list (right panel) ---

    def _refresh_categories(self, _s, _a):
        """
        Rebuilds the category checklist, preserving check state from _cat_checked.
        Uses _programmatic_check flag to bypass the ItemCheck guard during
        SetItemChecked calls (otherwise the guard reverts them).
        """
        current_checked = set(self._cat_checked or [])
        needle = self._cat_filter()
        show_only = bool(self.chkShowOnly.Checked)
        def _match(nm):
            if show_only and nm not in current_checked: return False
            if needle: return needle in (nm or "").lower()
            return True
        self._programmatic_check = True
        self.clbCats.BeginUpdate()
        try:
            self.clbCats.Items.Clear()
            for nm in self._all_cat_names:
                if not _match(nm): continue
                idx = self.clbCats.Items.Add(nm)
                if nm in current_checked: self.clbCats.SetItemChecked(idx, True)
        finally:
            self.clbCats.EndUpdate()
            self._programmatic_check = False

    def _on_cat_check(self, _s, e):
        """
        ItemCheck handler for the category list.
        Allows programmatic SetItemChecked through via _programmatic_check flag.
        Blocks spurious toggles from scroll/keyboard (only mouse click allowed).
        Keeps _cat_checked set in sync with the UI.
        """
        if getattr(self, '_programmatic_check', False):
            return
        if (self._block_itemcheck or not self._last_input_mouse):
            e.NewValue = e.CurrentValue; return
        try:
            name = self.clbCats.Items[e.Index]
            if e.NewValue.ToString() == "Checked":
                self._cat_checked.add(name)
            else:
                if name in self._cat_checked: self._cat_checked.remove(name)
        except: pass

    # --- Top bar events ---

    def _on_browse_sp(self, _s, _a):
        """Browse for a per-parameter SP file."""
        dlg = OpenFileDialog(); dlg.Filter = "Shared Parameters (*.txt)|*.txt|All Files (*.*)|*.*"
        if self.cboSp.Text and os.path.exists(self.cboSp.Text):
            try:
                dlg.InitialDirectory = os.path.dirname(self.cboSp.Text)
                dlg.FileName = os.path.basename(self.cboSp.Text)
            except: pass
        if dlg.ShowDialog() == DialogResult.OK:
            self.cboSp.Text = dlg.FileName
            self._refresh_tree(None, None)

    # --- Bottom buttons ---

    def _on_cancel(self, _s, _a):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def _on_ok(self, _s, _a):
        """Validates and saves the row, then closes with OK result."""
        name = (self.txtSel.Text or "").strip()
        if not name:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Param Editor", "Pick a parameter from the tree (or type its name).")
            return
        self.row.name = name
        self.row.group_key = self._group_label_to_token.get(self.cboGroup.SelectedItem, self.cboGroup.SelectedItem) or "Construction"
        self.row.rfa_is_instance = bool(self.chkRfa.Checked)
        self.row.rvt_is_instance = bool(self.chkRvt.Checked)
        self.row.allow_vary = bool(self.chkVary.Checked)
        self.row.cats = sorted(self._cat_checked, key=lambda s: s.lower())
        self.row.per_spf = (self.cboSp.Text or "").strip()
        self.DialogResult = DialogResult.OK
        self.Close()


# =============================================================================
# CONFIG BUILDER FORM
# =============================================================================
# Top-level editor for a complete parameter config JSON file.
# Shows a list of ParamRows and lets you add/edit/duplicate/remove them.

class BuilderForm(Form):
    """Config-level editor: display name, scope, parameter list, save/load."""

    def __init__(self, json_path=None):
        Form.__init__(self)
        self.Text = "Shared Parameter JSON Builder"
        self.Size = Size(1120, 740)
        self.StartPosition = FormStartPosition.CenterScreen
        self.rows = []
        self.current_path = json_path

        lblDN = Label(Text="Display name:")
        lblDN.Location = Point(12, 12); lblDN.Size = Size(120, 18)
        self.txtDN = TextBox(); self.txtDN.Location = Point(12, 32); self.txtDN.Size = Size(320, 22)

        lblScope = Label(Text="Save to:")
        lblScope.Location = Point(340, 12); lblScope.Size = Size(80, 18)
        self.cboScope = ComboBox(); self.cboScope.Location = Point(340, 32); self.cboScope.Size = Size(140, 22)
        self.cboScope.DropDownStyle = ComboBoxStyle.DropDownList
        if core.can_edit_shared():
            self.cboScope.Items.Add("shared")
            self.cboScope.Items.Add("user")
            self.cboScope.SelectedIndex = 0
        else:
            self.cboScope.Items.Add("user")
            self.cboScope.SelectedIndex = 0

        lblTop = Label(Text="Top-level Shared Parameters file (optional):")
        lblTop.Location = Point(500, 12); lblTop.Size = Size(360, 18)
        self.txtTop = TextBox(); self.txtTop.Location = Point(500, 32); self.txtTop.Size = Size(460, 22)
        self.txtTop.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        btnBrowseTop = Button(Text="Browse...")
        btnBrowseTop.Location = Point(970, 31); btnBrowseTop.Size = Size(90, 24)
        btnBrowseTop.Anchor = AnchorStyles.Top | AnchorStyles.Right
        btnBrowseTop.Click += self._on_browse_top

        lblNote = Label(Text="SP file resolution: Per-parameter > Top-level (Builder) > Revit current")
        lblNote.Location = Point(12, 58); lblNote.Size = Size(520, 16)

        # Parameter list
        self.lv = ListView(); self.lv.View = View.Details; self.lv.FullRowSelect = True
        self.lv.Location = Point(12, 80); self.lv.Size = Size(1086, 560)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        for t, w in [("Name", 360), ("Group", 220), ("RFA Inst", 80), ("RVT Inst", 80), ("Vary", 60), ("Cats", 60), ("Per-Param SPF", 220)]:
            ch = ColumnHeader(); ch.Text = t; ch.Width = w; self.lv.Columns.Add(ch)

        # Action buttons
        btnAdd = Button(Text="Add Param"); btnAdd.Location = Point(12, 650); btnAdd.Size = Size(100, 28)
        btnAdd.Anchor = AnchorStyles.Bottom | AnchorStyles.Left; btnAdd.Click += self._on_add
        btnEdit = Button(Text="Edit Selected"); btnEdit.Location = Point(118, 650); btnEdit.Size = Size(110, 28)
        btnEdit.Anchor = AnchorStyles.Bottom | AnchorStyles.Left; btnEdit.Click += self._on_edit
        btnDup = Button(Text="Duplicate Selected"); btnDup.Location = Point(234, 650); btnDup.Size = Size(140, 28)
        btnDup.Anchor = AnchorStyles.Bottom | AnchorStyles.Left; btnDup.Click += self._on_dup
        btnRemove = Button(Text="Remove Selected"); btnRemove.Location = Point(380, 650); btnRemove.Size = Size(140, 28)
        btnRemove.Anchor = AnchorStyles.Bottom | AnchorStyles.Left; btnRemove.Click += self._on_remove
        btnClear = Button(Text="Clear List"); btnClear.Location = Point(526, 650); btnClear.Size = Size(100, 28)
        btnClear.Anchor = AnchorStyles.Bottom | AnchorStyles.Left; btnClear.Click += self._on_clear

        btnSave = Button(Text="Save"); btnSave.Location = Point(self.Width - 232, 650); btnSave.Size = Size(100, 28)
        btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right; btnSave.Click += self._on_save
        btnCancel = Button(Text="Cancel"); btnCancel.Location = Point(self.Width - 122, 650); btnCancel.Size = Size(100, 28)
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnCancel.Click += (lambda _s,_a: self.Close())

        for w in [lblDN, self.txtDN, lblScope, self.cboScope, lblTop, self.txtTop, btnBrowseTop, lblNote,
                  self.lv, btnAdd, btnEdit, btnDup, btnRemove, btnClear, btnSave, btnCancel]:
            self.Controls.Add(w)

        # Pre-fill the top-level SP path from Revit's current setting
        try:
            sp = __revit__.Application.SharedParametersFilename
            if sp and os.path.exists(sp):
                self.txtTop.Text = sp
        except: pass

        # Load existing config if editing
        if json_path and os.path.exists(json_path):
            try: self._load_from_config(json_path)
            except Exception as e:
                from Autodesk.Revit.UI import TaskDialog
                TaskDialog.Show("Builder - Load Error", str(e))

        self._refresh()

    # --- Builder helpers ---

    def _on_browse_top(self, s, a):
        """Browse for a top-level Shared Parameters file."""
        dlg = OpenFileDialog(); dlg.Filter = "Shared Parameters (*.txt)|*.txt|All Files (*.*)|*.*"
        if self.txtTop.Text and os.path.exists(self.txtTop.Text):
            try:
                dlg.InitialDirectory = os.path.dirname(self.txtTop.Text)
                dlg.FileName = os.path.basename(self.txtTop.Text)
            except: pass
        if dlg.ShowDialog() == DialogResult.OK:
            self.txtTop.Text = dlg.FileName

    def _load_from_config(self, path):
        """Loads an existing JSON config into the form."""
        data = core.load_json_file(path)
        self.txtDN.Text = data.get("display_name", "")
        self.txtTop.Text = data.get("shared_parameter_file", "") or self.txtTop.Text
        self.rows = [ParamRow.from_dict(p) for p in (data.get("parameters", []) or [])]
        # Detect scope from file path for display purposes
        try:
            p = os.path.normcase(os.path.normpath(path))
            sd = os.path.normcase(os.path.normpath(core.shared_dir() or "")); ud = os.path.normcase(os.path.normpath(core.user_dir() or ""))
            scope = "shared" if sd and p.startswith(sd) else ("user" if ud and p.startswith(ud) else None)
            if scope is not None:
                for i in range(self.cboScope.Items.Count):
                    if self.cboScope.Items[i] == scope:
                        self.cboScope.SelectedIndex = i; break
        except: pass

    def _refresh(self):
        """Refreshes the parameter ListView from self.rows."""
        self.lv.Items.Clear()
        for r in self.rows:
            sub = Array[String]([
                r.name or "",
                r.group_key or "",
                "Yes" if r.rfa_is_instance else "No",
                "Yes" if r.rvt_is_instance else "No",
                "Yes" if r.allow_vary else "No",
                str(len(r.cats)),
                os.path.basename(r.per_spf) if r.per_spf else ""
            ])
            self.lv.Items.Add(ListViewItem(sub))

    def _sel_index(self):
        """Returns the selected row index, or -1 if nothing selected."""
        return int(self.lv.SelectedIndices[0]) if self.lv.SelectedIndices.Count else -1

    def _on_add(self, s, a):
        """Opens ParamEditor to add a new parameter."""
        ed = ParamEditor(self.txtTop.Text.strip() or "", None)
        if ed.ShowDialog() == DialogResult.OK:
            self.rows.append(ed.row); self._refresh()

    def _on_edit(self, s, a):
        """Opens ParamEditor to edit the selected parameter."""
        idx = self._sel_index()
        if idx < 0:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Builder", "Select a parameter to edit."); return
        ed = ParamEditor(self.txtTop.Text.strip() or "", self.rows[idx])
        if ed.ShowDialog() == DialogResult.OK:
            self.rows[idx] = ed.row; self._refresh()

    def _on_dup(self, s, a):
        """Duplicates the selected parameter into a new ParamEditor."""
        idx = self._sel_index()
        if idx < 0:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Builder", "Select a parameter to duplicate."); return
        src = self.rows[idx]
        dup = ParamRow(src.name, src.group_key, src.rfa_is_instance, src.rvt_is_instance, src.allow_vary, list(src.cats), src.per_spf)
        ed = ParamEditor(self.txtTop.Text.strip() or "", dup)
        if ed.ShowDialog() == DialogResult.OK:
            self.rows.append(ed.row); self._refresh()

    def _on_remove(self, s, a):
        """Removes the selected parameter from the list."""
        idx = self._sel_index()
        if idx >= 0:
            del self.rows[idx]; self._refresh()

    def _on_clear(self, s, a):
        """Clears all parameters from the list."""
        self.rows = []; self._refresh()

    def _on_save(self, s, a):
        """Validates and saves the config to a JSON file."""
        disp = (self.txtDN.Text or "").strip()
        if not disp:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Builder", "Enter a display name."); return
        js = {"display_name": disp, "parameters": [r.to_dict() for r in self.rows]}
        top_sp = self.txtTop.Text.strip()
        if top_sp:
            js["shared_parameter_file"] = top_sp
        try:
            js["version"] = int(js.get("version", 0)) + 1
            js["last_modified_by"] = getpass.getuser()
            import datetime as _dt
            js["last_modified_on"] = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except: pass
        scope = self.cboScope.SelectedItem or ("shared" if core.can_edit_shared() else "user")
        folder = core.shared_dir() if scope == "shared" else core.user_dir()
        name = core.slugify(disp) + ".json"
        target_path = os.path.join(folder, name)
        # Write to temp file first, then rename (atomic save)
        tmp = target_path + ".tmp"
        with open(tmp, "w") as fp:
            json.dump(js, fp, indent=2)
        if os.path.exists(target_path):
            try: os.remove(target_path)
            except: pass
        os.rename(tmp, target_path)
        self.Close()


# =============================================================================
# ENTRY POINT
# =============================================================================

def show_builder(selected_json_path=None):
    """Opens the BuilderForm. Called by script.py Launcher on Create/Edit."""
    Application.EnableVisualStyles()
    frm = BuilderForm(selected_json_path)
    frm.ShowDialog()
