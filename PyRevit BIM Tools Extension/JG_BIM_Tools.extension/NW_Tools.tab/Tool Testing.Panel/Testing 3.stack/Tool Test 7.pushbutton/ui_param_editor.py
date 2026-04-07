# -*- coding: utf-8 -*-
import os
import clr

# --- .NET UI
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
from System import Array, String
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Application, Form, Label, TextBox, Button, CheckBox, ComboBox,
    ComboBoxStyle, ListBox, SelectionMode, AnchorStyles,
    DialogResult, FormStartPosition, OpenFileDialog, TreeView, TreeNode,
    SplitContainer, RadioButton, CheckedListBox, Keys
)

# --- Revit + helpers
clr.AddReference("RevitAPI")
import Autodesk.Revit.DB as DB
import binder_core as core

GROUP_KEY_CHOICES = [
    "Analysis Results", "Analytical Alignment", "Analytical Model", "Constraints",
    "Construction", "Data", "Dimensions", "Division Geometry",
    "Electrical", "Electrical - Circuiting", "Electrical - Lighting",
    "Electrical - Loads", "Electrical Analysis", "Electrical Engineering",
    "Energy Analysis", "Fire Protection", "Forces", "General", "Graphics",
    "Green Building Properties", "Identity Data", "IFC Parameters", "Layers",
    "Life Safety", "Materials and Finishes", "Mechanical",
    "Mechanical - Flow", "Mechanical - Loads", "Model Properties", "Moments",
    "Other", "Overall Legend", "Phasing", "Photometrics", "Plumbing",
    "Primary End", "Rebar Set", "Releases / Member Forces", "Secondary End",
    "Segments and Fittings", "Set", "Slab Shape Edit", "Structural",
    "Structural Analysis", "Text", "Title Text", "Visibility"
]

class ParamRow(object):
    def __init__(self, name="", group_key="Construction",
                 rfa_is_instance=True, rvt_is_instance=True, allow_vary=True,
                 cats=None, per_spf=""):
        self.name = name
        self.group_key = group_key or "Construction"
        self.rfa_is_instance = bool(rfa_is_instance)
        self.rvt_is_instance = bool(rvt_is_instance)
        self.allow_vary = bool(allow_vary)
        self.cats = list(cats or [])
        self.per_spf = per_spf or ""
    def to_dict(self):
        d = {
            "name": self.name, "source": "shared", "group_key": self.group_key,
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
        return ParamRow(
            name=(p.get("name") or "").strip(),
            group_key=p.get("group_key", "Construction"),
            rfa_is_instance=bool(p.get("rfa_is_instance", True)),
            rvt_is_instance=bool(p.get("rvt_is_instance", True)),
            allow_vary=bool(p.get("rvt_allow_vary_between_groups", True)),
            cats=p.get("rvt_categories", []) or [],
            per_spf=p.get("shared_parameter_file", "") or ""
        )

def _doc():
    return __revit__.ActiveUIDocument.Document

def _category_names():
    names = []
    try:
        for c in _doc().Settings.Categories:
            try: names.append(c.Name)
            except: pass
    except: pass
    names.sort(key=lambda s: s.lower())
    return names

def _sp_candidates(row_per_spf, top_spf):
    cands = []
    if row_per_spf: cands.append(row_per_spf)
    if top_spf and top_spf not in cands: cands.append(top_spf)
    try:
        cur = __revit__.Application.SharedParametersFilename
        if cur and cur not in cands: cands.append(cur)
    except: pass
    return cands

class ParamEditor(Form):
    def __init__(self, top_spf, row):
        Form.__init__(self)
        self.Text = "Parameter Editor"
        self.Size = Size(1120, 760)
        self.StartPosition = FormStartPosition.CenterScreen

        self.row = row or ParamRow()
        self.top_spf = top_spf or ""
        self._all_cat_names = _category_names()

        # guard: only mouse click (or space) toggles categories
        self._block_itemcheck = False
        self._last_input_mouse = False

        # ----- Top SP selector -----
        lblSp = Label(Text="Per-parameter SP file (optional):")
        lblSp.Location = Point(12, 14); lblSp.Size = Size(260, 18)

        self.cboSp = ComboBox()
        self.cboSp.Location = Point(12, 34); self.cboSp.Size = Size(980, 22)
        self.cboSp.DropDownStyle = ComboBoxStyle.DropDown
        self.cboSp.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        btnBrowseSp = Button(Text="Browse...")
        btnBrowseSp.Location = Point(1000, 33); btnBrowseSp.Size = Size(90, 24)
        btnBrowseSp.Anchor = AnchorStyles.Top | AnchorStyles.Right
        btnBrowseSp.Click += self._on_browse_sp

        for p in _sp_candidates(self.row.per_spf, self.top_spf):
            self.cboSp.Items.Add(p)
        if self.row.per_spf:
            self.cboSp.Text = self.row.per_spf
        elif self.top_spf:
            self.cboSp.Text = self.top_spf

        # NOTE label pinned to bottom now
        lblNote = Label(Text="SP file resolution: Per-parameter > Top-level (Builder) > Revit current")
        lblNote.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        lblNote.Location = Point(12, self.ClientSize.Height - 28)
        lblNote.Size = Size(520, 16)
        def _relayout_note(_s=None, _a=None):
            lblNote.Location = Point(12, self.ClientSize.Height - 28)
        self.Resize += _relayout_note

        # ----- Splitter -----
        split = SplitContainer()
        split.Location = Point(12, 84)
        split.Size = Size(1078, 620)
        split.SplitterDistance = 470
        split.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.split = split

        # LEFT header + filter
        leftHdr = Label(Text="Parameters in selected SP file:")
        leftHdr.Location = Point(6, 6); leftHdr.Size = Size(220, 18)

        self.txtParamFilter = TextBox()
        self.txtParamFilter.Location = Point(6, 26); self.txtParamFilter.Size = Size(450, 22)
        self.txtParamFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.txtParamFilter.TextChanged += self._refresh_sources

        self.rbGrouped = RadioButton(Text="Grouped")
        self.rbGrouped.Location = Point(6, 52); self.rbGrouped.Checked = True
        self.rbAZ = RadioButton(Text="A-Z")
        self.rbAZ.Location = Point(92, 52)
        self.rbGrouped.CheckedChanged += self._on_source_mode_changed
        self.rbAZ.CheckedChanged += self._on_source_mode_changed

        for w in [leftHdr, self.txtParamFilter, self.rbGrouped, self.rbAZ]:
            split.Panel1.Controls.Add(w)

        self.tv = TreeView()
        self.tv.Location = Point(6, 78); self.tv.Size = Size(450, 580)
        self.tv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.tv.AfterSelect += self._on_tree_select
        split.Panel1.Controls.Add(self.tv)

        self.lstAZ = ListBox()
        self.lstAZ.Location = Point(6, 78); self.lstAZ.Size = Size(450, 580)
        self.lstAZ.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.lstAZ.Visible = False
        self.lstAZ.SelectionMode = SelectionMode.One
        self.lstAZ.SelectedIndexChanged += self._on_az_select
        split.Panel1.Controls.Add(self.lstAZ)

        # RIGHT side
        lblSel = Label(Text="Selected parameter:")
        lblSel.Location = Point(6, 6); lblSel.Size = Size(130, 18)
        split.Panel2.Controls.Add(lblSel)

        self.txtSel = TextBox()
        self.txtSel.Location = Point(6, 26); self.txtSel.Size = Size(580, 22)
        self.txtSel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        split.Panel2.Controls.Add(self.txtSel)

        lblGroup = Label(Text="Group Key:")
        lblGroup.Location = Point(6, 56); lblGroup.Size = Size(120, 18)
        split.Panel2.Controls.Add(lblGroup)

        self.cboGroup = ComboBox()
        self.cboGroup.Location = Point(6, 76); self.cboGroup.Size = Size(240, 22)
        self.cboGroup.DropDownStyle = ComboBoxStyle.DropDownList
        # Populate group dropdown dynamically from Revit API, store stable tokens in config
        self._group_label_to_token = {}
        self._group_token_to_label = {}
        try:
            doc = __revit__.ActiveUIDocument.Document
            for (lab, tok) in binder_core.list_user_assignable_group_choices(doc):
                self.cboGroup.Items.Add(lab)
                self._group_label_to_token[lab] = tok
                self._group_token_to_label[tok] = lab
        except:
            # Fallback to static labels if API enumeration fails
            for g in GROUP_KEY_CHOICES:
                self.cboGroup.Items.Add(g)
                self._group_label_to_token[g] = g
                self._group_token_to_label[g] = g
        split.Panel2.Controls.Add(self.cboGroup)

        self.chkRfa = CheckBox(Text="RFA: Instance (unchecked = Type)")
        self.chkRfa.Location = Point(260, 76); self.chkRfa.Size = Size(240, 22)
        split.Panel2.Controls.Add(self.chkRfa)

        self.chkRvt = CheckBox(Text="RVT: Instance (unchecked = Type)")
        self.chkRvt.Location = Point(260, 100); self.chkRvt.Size = Size(240, 22)
        split.Panel2.Controls.Add(self.chkRvt)

        self.chkVary = CheckBox(Text="RVT: Allow values to vary by group instance")
        self.chkVary.Location = Point(6, 100); self.chkVary.Size = Size(320, 22)
        split.Panel2.Controls.Add(self.chkVary)

        lblCats = Label(Text="RVT Categories (check to include)")
        lblCats.Location = Point(6, 130); lblCats.Size = Size(260, 18)
        split.Panel2.Controls.Add(lblCats)

        self.chkCatsOnly = CheckBox(Text="Show only selected")
        self.chkCatsOnly.Location = Point(6, 152); self.chkCatsOnly.Size = Size(160, 18)
        self.chkCatsOnly.CheckedChanged += self._rebuild_cats
        split.Panel2.Controls.Add(self.chkCatsOnly)

        # Filter moved UNDER the label/“show only selected”
        self.txtCatFilter = TextBox()
        self.txtCatFilter.Location = Point(6, 176)
        self.txtCatFilter.Size = Size(620, 22)
        self.txtCatFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
        self.txtCatFilter.TextChanged += self._rebuild_cats
        split.Panel2.Controls.Add(self.txtCatFilter)

        # Categories list BELOW the filter
        self.clbCats = CheckedListBox()
        self.clbCats.Location = Point(6, 202)
        self.clbCats.Size = Size(620, 456 - 52)   # keep overall panel height similar
        self.clbCats.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        self.clbCats.CheckOnClick = False
        self.clbCats.MouseDown += self._on_cats_mouse_down
        self.clbCats.MouseUp += self._on_cats_mouse_up
        self.clbCats.KeyDown += self._on_cats_key_down
        self.clbCats.KeyUp += self._on_cats_key_up
        self.clbCats.ItemCheck += self._on_item_check_guarded
        split.Panel2.Controls.Add(self.clbCats)

        # Bottom buttons (unchanged)
        btnSave = Button(Text="Save")
        btnSave.Location = Point(920, 714); btnSave.Size = Size(90, 28)
        btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnSave.Click += self._on_ok

        btnCancel = Button(Text="Cancel")
        btnCancel.Location = Point(1010, 714); btnCancel.Size = Size(90, 28)
        btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        btnCancel.Click += self._on_cancel

        for w in [lblSp, self.cboSp, btnBrowseSp, split, lblNote, btnSave, btnCancel]:
            self.Controls.Add(w)

        # Init
        self.cboSp.TextChanged += self._refresh_sources
        self._apply_row(self.row)
        self._refresh_sources(None, None)
        self._rebuild_cats(None, None)

    # ==== guards ====
    def _on_cats_mouse_down(self, *_): self._last_input_mouse = True; self._block_itemcheck = False
    def _on_cats_mouse_up(self, *_):   self._last_input_mouse = False
    def _on_cats_key_down(self, _s, e):
        self._last_input_mouse = False
        if e.KeyCode == Keys.Space: self._block_itemcheck = True
    def _on_cats_key_up(self, _s, e):
        if e.KeyCode == Keys.Space: self._block_itemcheck = False

    # ==== helpers ====
    def _apply_row(self, row):
        if row.name: self.txtSel.Text = row.name
        try:
            idx = 0
            for i in range(self.cboGroup.Items.Count):
                if self.cboGroup.Items[i].lower() == (self._group_token_to_label.get(row.group_key, row.group_key) or "Construction").lower():
                    idx = i; break
            self.cboGroup.SelectedIndex = idx
        except: pass
        self.chkRfa.Checked = bool(row.rfa_is_instance)
        self.chkRvt.Checked = bool(row.rvt_is_instance)
        self.chkVary.Checked = bool(row.allow_vary)
        self._row_cats_set = dict((c, True) for c in (row.cats or []))

    def _effective_sp(self):
        t = (self.cboSp.Text or "").strip()
        if t: return t
        if self.top_spf: return self.top_spf
        try:
            sp = __revit__.Application.SharedParametersFilename
            if sp: return sp
        except: pass
        return None

    def _refresh_sources(self, *_):
        spf = core.open_shared_parameter_file(self._effective_sp())
        needle = (self.txtParamFilter.Text or "").strip().lower()
        def _match(nm): return (not needle) or (needle in (nm or "").lower())

        self.tv.Nodes.Clear()
        if spf:
            for grp in spf.Groups:
                try:
                    gnode = TreeNode(grp.Name); added = False
                    for d in grp.Definitions:
                        try:
                            if _match(d.Name):
                                gnode.Nodes.Add(d.Name); added = True
                        except: pass
                    if added: self.tv.Nodes.Add(gnode)
                except: pass
        else:
            self.tv.Nodes.Add("No shared parameter file available")
        self.tv.ExpandAll()

        self.lstAZ.Items.Clear()
        if spf:
            names = []
            for grp in spf.Groups:
                for d in grp.Definitions:
                    try:
                        if _match(d.Name): names.append(d.Name)
                    except: pass
            for n in sorted(set(names), key=lambda s: s.lower()):
                self.lstAZ.Items.Add(n)

        self._apply_source_mode()

    def _apply_source_mode(self):
        az = bool(self.rbAZ.Checked)
        self.lstAZ.Visible = az
        self.tv.Visible = not az

    def _on_source_mode_changed(self, *_): self._apply_source_mode()
    def _on_tree_select(self, _s, args):
        node = args.Node
        if node is None: return
        if node.Nodes is not None and node.Nodes.Count > 0: return
        self.txtSel.Text = node.Text or ""
    def _on_az_select(self, _s, _a):
        if self.lstAZ.SelectedIndex >= 0:
            self.txtSel.Text = self.lstAZ.SelectedItem or ""

    def _on_browse_sp(self, _s, _a):
        dlg = OpenFileDialog(); dlg.Filter = "Shared Parameters (*.txt)|*.txt|All Files (*.*)|*.*"
        if self.cboSp.Text and os.path.exists(self.cboSp.Text):
            try:
                dlg.InitialDirectory = os.path.dirname(self.cboSp.Text)
                dlg.FileName = os.path.basename(self.cboSp.Text)
            except: pass
        if dlg.ShowDialog() == DialogResult.OK:
            self.cboSp.Text = dlg.FileName  # triggers refresh

    def _current_checked_categories(self):
        checked = []
        for i in range(self.clbCats.Items.Count):
            if self.clbCats.GetItemChecked(i): checked.append(self.clbCats.Items[i])
        return checked

    def _rebuild_cats(self, *_):
        preserve = dict((c, True) for c in self._current_checked_categories())
        if hasattr(self, "_row_cats_set"): preserve.update(self._row_cats_set)

        needle = (self.txtCatFilter.Text or "").strip().lower()
        show_only = False  # this editor has no separate 'show only selected' box
        def _match(nm):
            if show_only and nm not in preserve: return False
            if needle: return needle in (nm or "").lower()
            return True

        self.clbCats.Items.Clear()
        for nm in self._all_cat_names:
            if not _match(nm): continue
            idx = self.clbCats.Items.Add(nm)
            if nm in preserve: self.clbCats.SetItemChecked(idx, True)

    def _on_item_check_guarded(self, _s, e):
        if (self._block_itemcheck or not self._last_input_mouse):
            e.NewValue = e.CurrentValue

    # ----- buttons -----
    def _on_cancel(self, _s, _a):
        self.DialogResult = DialogResult.Cancel
        self.Close()

def _on_test(self, s, a):
    try:
        # build a minimal config from current UI state
        js = {"display_name": "(param editor test)", "parameters": [r.to_dict() for r in self.rows]}
        doc = __revit__.ActiveUIDocument.Document
        res = core.validate_config(js, doc)
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Param Binder - Test", core.format_validation_results(res))
    except Exception as e:
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Param Binder - Test Error", str(e))


    def _on_ok(self, _s, _a):
        name = (self.txtSel.Text or "").strip()
        if not name:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show("Param Editor", "Pick a parameter from the left.")
            return
        self.row.name = name
        self.row.group_key = self._group_label_to_token.get(self.cboGroup.SelectedItem, self.cboGroup.SelectedItem) or "Construction"
        self.row.rfa_is_instance = bool(self.chkRfa.Checked)
        self.row.rvt_is_instance = bool(self.chkRvt.Checked)
        self.row.allow_vary = bool(self.chkVary.Checked)
        self.row.cats = self._current_checked_categories()
        self.row.per_spf = (self.cboSp.Text or "").strip()
        self.DialogResult = DialogResult.OK
        self.Close()

def edit_param(top_spf, row=None):
    Application.EnableVisualStyles()
    frm = ParamEditor(top_spf, row or ParamRow())
    return (frm.ShowDialog() == DialogResult.OK, frm.row)