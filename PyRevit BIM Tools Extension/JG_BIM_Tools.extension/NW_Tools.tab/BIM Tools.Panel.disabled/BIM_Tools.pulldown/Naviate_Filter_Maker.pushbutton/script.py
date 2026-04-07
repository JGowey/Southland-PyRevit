# -*- coding: utf-8 -*-
# Build Naviate Filter Import JSON from Revit Filterable Categories
# Adds live text filter for categories
# IronPython 2.7 / pyRevit / Revit 2022+
# Last updated: 2025-08-28

import sys
import os
import clr
import json

# Revit API
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import Category, ParameterFilterUtilities

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

# .NET / Windows
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.Windows.Forms')
from System import Environment
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Form, Label, TextBox, CheckedListBox, Button, DialogResult,
    SaveFileDialog, AnchorStyles, FormStartPosition, CheckState
)

# ----------------------------- Helpers ---------------------------------------
def current_doc():
    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show("Filter Import JSON", "No active document.")
        sys.exit(0)
    return uidoc.Document

def get_filterable_category_names(doc):
    cat_ids = []
    try:
        all_ids = ParameterFilterUtilities.GetAllFilterableCategories()
        for eid in all_ids:
            cat_ids.append(eid)
    except:
        for c in doc.Settings.Categories:
            try:
                if ParameterFilterUtilities.IsFilterableCategoryId(c.Id):
                    cat_ids.append(c.Id)
            except:
                pass

    names = []
    seen = set()
    for cid in cat_ids:
        cat = Category.GetCategory(doc, cid)
        if cat is not None:
            n = cat.Name
            if n not in seen:
                names.append(n)
                seen.add(n)
    names.sort(key=lambda s: s.lower())
    return names

def ensure_default_dir():
    # %APPDATA%\pyRevit\JG_Tools_Settings
    try:
        base = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        target = os.path.join(base, "pyRevit", "JG_Tools_Settings")
        if not os.path.isdir(target):
            os.makedirs(target)
        return target
    except:
        return None

def save_text_file(default_name, content):
    sfd = SaveFileDialog()
    sfd.Title = "Save Filter Import JSON"
    sfd.Filter = "JSON File (*.json)|*.json|All Files (*.*)|*.*"
    sfd.FileName = default_name
    init = ensure_default_dir()
    if init:
        sfd.InitialDirectory = init
    if sfd.ShowDialog() == DialogResult.OK:
        path = sfd.FileName
        try:
            import System
            import System.IO
            sw = System.IO.StreamWriter(path, False, System.Text.UTF8Encoding(False))  # UTF-8 (no BOM)
            sw.Write(content)
            sw.Close()
            return path
        except Exception as ex:
            TaskDialog.Show("Save Error", str(ex))
            return None
    return None

def build_json_payload(condition_name, filter_name, selected_categories, all_categories):
    # Build child nodes for each category with Checked flag per user selection
    child_nodes = []
    selset = set(selected_categories)
    for cat in all_categories:
        child_nodes.append({
            "Text": cat,
            "FormattedQuantity": "(0)",
            "RootQuantity": "0",
            "ElementIds": [],
            "Checked": (cat in selset),
            "IsSelected": False,
            "IsRoot": False,
            "Nodes": []
        })

    categories_root = {
        "Text": "Categories",
        "FormattedQuantity": "(0)",
        "RootQuantity": "100",
        "ElementIds": [],
        "Checked": False,
        "IsSelected": False,
        "IsRoot": False,
        "Nodes": child_nodes
    }

    payload = {
        "OptionGuid":"141463ff-a33d-4013-9b02-9aa8a8ec9baf",
        "OldOptionGuid":"8b08b74c-bf03-4bea-a709-ba76e47f93de",
        "OptionsName":"DataFlowOptions",
        "ConditionList":[
            {
                "Name": condition_name,
                "FilterName": filter_name,
                "NaviateFilter":{
                    "SelectionMode":1,
                    "SelectElementsInView":True,
                    "SelectAllElements":False,
                    "SelectFamilyInstances":False,
                    "SelectModelElements":False,
                    "SelectBy3DZone":False,
                    "UsePreselect":False,
                    "UseOnlyCommonParameters":False,
                    "UseLimitByParameter":False,
                    "LimitParameter":{
                        "BuiltInName":"ELEM_FAMILY_AND_TYPE_PARAM",
                        "StorageType":"",
                        "IsBuiltIn":True,
                        "IsBool":False,
                        "BoolValue":False,
                        "Name":"",
                        "Filter":"",
                        "Selected":False,
                        "IsChecked":False,
                        "IdInt":-1
                    },
                    "LimitParametersToInstance":True,
                    "LimitParametersToType":False,
                    "Nodes":[categories_root],
                    "ParamGuidsByParamNameAndValue":{},
                    "SelectedParameters":[],
                    "InvertedSelection":False,
                    "Zone3DName":"Scope Box 1",
                    "SettingWithAOB":True,
                    "Name": filter_name,
                    "Filter":"",
                    "Selected":False,
                    "IsChecked":False,
                    "IdInt":-1
                },
                "NaviateFilterValid":True,
                "TargetParameterValid":True,
                "PowershellScriptValid":False,
                "NaviateFilterIdTag":{"Name": filter_name},
                "PowershellScriptContent":"",
                "IsValidCondition":False,
                "GroupName":"<Default>",
                "ExecutionIndex":9999,
                "ExecutionType":1,
                "IsActive":True,
                "Description":"Used for Filter imports only. Delete after import.",
                "Status":0,
                "Group":{"IsEditing":False,"Name":"<Default>","PreviousName":""},
                "ConditionResultList":[],
                "Filter":"",
                "Selected":True,
                "IsChecked":True,
                "IdInt":-1
            }
        ],
        "ConditionGroupList":[
            {"IsEditing":False,"Name":"<Default>","PreviousName":""},
            {"IsEditing":False,"Name":"BOM Reporting","PreviousName":""}
        ]
    }

    # Compact JSON like your sample
    return json.dumps(payload, ensure_ascii=False, separators=(',', ':'))

# ------------------------------- UI ------------------------------------------
class FilterJsonForm(Form):
    def __init__(self, all_categories):
        self.Text = "Build Filter Import JSON"
        self.StartPosition = FormStartPosition.CenterScreen
        self.ClientSize = Size(780, 700)
        self.MinimizeBox = False
        self.MaximizeBox = False

        # Labels
        self.lblCond = Label()
        self.lblCond.Text = "Condition Name:"
        self.lblCond.Location = Point(12, 12)
        self.lblCond.AutoSize = True
        self.Controls.Add(self.lblCond)

        self.lblFilter = Label()
        self.lblFilter.Text = "Filter Name:"
        self.lblFilter.Location = Point(12, 60)
        self.lblFilter.AutoSize = True
        self.Controls.Add(self.lblFilter)

        self.lblCatFilter = Label()
        self.lblCatFilter.Text = "Category filter (type to narrow):"
        self.lblCatFilter.Location = Point(12, 108)
        self.lblCatFilter.AutoSize = True
        self.Controls.Add(self.lblCatFilter)

        self.lblCats = Label()
        self.lblCats.Text = "Filterable Categories (check to include):"
        self.lblCats.Location = Point(12, 156)
        self.lblCats.AutoSize = True
        self.Controls.Add(self.lblCats)

        # TextBoxes
        self.txtCond = TextBox()
        self.txtCond.Size = Size(590, 22)
        self.txtCond.Location = Point(12, 30)
        self.txtCond.Text = "Filter Import"
        self.Controls.Add(self.txtCond)

        self.txtFilter = TextBox()
        self.txtFilter.Size = Size(590, 22)
        self.txtFilter.Location = Point(12, 78)
        self.txtFilter.Text = "Multi Category in View"
        self.Controls.Add(self.txtFilter)

        self.txtCatFilter = TextBox()
        self.txtCatFilter.Size = Size(590, 22)
        self.txtCatFilter.Location = Point(12, 126)
        self.txtCatFilter.TextChanged += self.on_filter_changed
        self.Controls.Add(self.txtCatFilter)

        # CheckedListBox
        self.clb = CheckedListBox()
        self.clb.Location = Point(12, 176)
        self.clb.Size = Size(756, 456)
        self.clb.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
        self.clb.CheckOnClick = True
        self.clb.ItemCheck += self.on_item_check
        self.Controls.Add(self.clb)

        # Buttons
        self.btnSelectAll = Button()
        self.btnSelectAll.Text = "Select All (filtered)"
        self.btnSelectAll.Location = Point(12, 648)
        self.btnSelectAll.Size = Size(120, 30) 
        self.btnSelectAll.Click += self.on_select_all
        self.Controls.Add(self.btnSelectAll)

        self.btnClearAll = Button()
        self.btnClearAll.Text = "Clear All (filtered)"
        self.btnClearAll.Location = Point(150, 648)
        self.btnClearAll.Size = Size(120, 30)  
        self.btnClearAll.Click += self.on_clear_all
        self.Controls.Add(self.btnClearAll)

        self.btnSave = Button()
        self.btnSave.Text = "Save JSON"
        self.btnSave.Location = Point(552, 648)
        self.btnSave.Size = Size(90, 30) 
        self.btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnSave.Click += self.on_save
        self.Controls.Add(self.btnSave)

        self.btnCancel = Button()
        self.btnCancel.Text = "Cancel"
        self.btnCancel.Location = Point(654, 648)
        self.btnCancel.Size = Size(90, 30) 
        self.btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
        self.btnCancel.Click += self.on_cancel
        self.Controls.Add(self.btnCancel)

        # Data
        self._all_categories = list(all_categories)   # ordered list
        self._checked = set()                         # names checked (global, not just visible)
        self._suppress_itemcheck = False

        # Initial populate
        self._refresh_list("")

    # --------- internal helpers ----------
    def _refresh_list(self, filter_text):
        """Rebuild the visible list from the full set, preserving check state."""
        filt = (filter_text or "").strip().lower()
        self._suppress_itemcheck = True
        try:
            self.clb.Items.Clear()
            for name in self._all_categories:
                if (not filt) or (name.lower().find(filt) >= 0):
                    idx = self.clb.Items.Add(name)
                    self.clb.SetItemChecked(idx, name in self._checked)
        finally:
            self._suppress_itemcheck = False

    def _visible_names(self):
        names = []
        for i in range(self.clb.Items.Count):
            names.append(self.clb.Items[i])
        return names

    # --------- events ----------
    def on_filter_changed(self, sender, args):
        self._refresh_list(self.txtCatFilter.Text)

    def on_item_check(self, sender, e):
        if self._suppress_itemcheck:
            return
        name = self.clb.Items[e.Index]
        if e.NewValue == CheckState.Checked:
            self._checked.add(name)
        else:
            if name in self._checked:
                self._checked.remove(name)

    def on_select_all(self, sender, args):
        vis = self._visible_names()
        for name in vis:
            self._checked.add(name)
        self._suppress_itemcheck = True
        try:
            for i in range(self.clb.Items.Count):
                self.clb.SetItemChecked(i, True)
        finally:
            self._suppress_itemcheck = False

    def on_clear_all(self, sender, args):
        vis = self._visible_names()
        for name in vis:
            if name in self._checked:
                self._checked.remove(name)
        self._suppress_itemcheck = True
        try:
            for i in range(self.clb.Items.Count):
                self.clb.SetItemChecked(i, False)
        finally:
            self._suppress_itemcheck = False

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    def on_save(self, sender, args):
        cond_name = self.txtCond.Text.strip()
        filt_name = self.txtFilter.Text.strip()
        if not cond_name:
            TaskDialog.Show("Filter Import JSON", "Please enter a Condition Name.")
            return
        if not filt_name:
            TaskDialog.Show("Filter Import JSON", "Please enter a Filter Name.")
            return

        # Selected = all globally checked, returned in original order
        selected = [n for n in self._all_categories if n in self._checked]

        json_text = build_json_payload(cond_name, filt_name, selected, self._all_categories)
        saved_path = save_text_file("FilterImport.json", json_text)
        if saved_path:
            TaskDialog.Show("Filter Import JSON", "Saved JSON to:\n{0}".format(saved_path))
            self.DialogResult = DialogResult.OK
            self.Close()

# ------------------------------- Main ----------------------------------------
def main():
    doc = current_doc()
    names = get_filterable_category_names(doc)

    # No print view / no pyRevit console output

    form = FilterJsonForm(names)
    form.ShowDialog()

if __name__ == "__main__":
    main()
