# -*- coding: utf-8 -*-
"""
Assembly Export to PCF
Replicates and extends SI_Tools AssemblyExportPCFCommandService.

Settings Dialog (modal):
  - Export Mode: Full Assembly / Selection Set
  - Hide elements after process (checkbox, either mode)
  - Naming: Auto-increment (.1, .2) / Custom name (Selection Set only)
  - Group Export by Parameter: batch export grouped by any Revit parameter

Both modes: folder up front -> Process/Multiple/Close loop -> export each
Process immediately -> Close -> restore hidden -> report.

Full Assembly: selects AssemblyInstances, exports members, AssemblyTypeName
  as filename.

Selection Set: selects FabricationParts.
  - If parts are in an assembly: uses assembly name, naming settings apply
    PIPELINE-REFERENCE and SPOOL-IDENTIFIER updated to match file name
  - If parts are NOT in an assembly: mandatory name prompt (forced)
  - Ghost assembly cleanup: detects assemblies Revit auto-creates during
    ExportToPCF on loose parts, deletes them after export
  - PIPELINE-REFERENCE post-processing: replaces Revit's auto-generated
    pipeline reference with the user-provided name

Group Export: batch export by parameter grouping.
  - Source: current selection or all parts in active view
  - Column picker: choose which parameters to display
  - Group by any parameter (Material, Comments, etc.)
  - Each unique value = separate PCF file
  - Auto-name (group value) or custom name each group
  - PIPELINE-REFERENCE and SPOOL-IDENTIFIER set to the group/file name

Run inside Revit via pyRevit.
"""

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Fabrication import FabricationUtils
from Autodesk.Revit.UI.Selection import ISelectionFilter
import os
import datetime
import time

import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import (
    Application, Form, Button, CheckBox, RadioButton,
    GroupBox, Label, TextBox,
    FormStartPosition, FormBorderStyle,
    ComboBox, ComboBoxStyle, ListBox, SelectionMode,
    DataGridView, DataGridViewTextBoxColumn,
    DataGridViewCheckBoxColumn,
    DataGridViewSelectionMode, DataGridViewColumnHeadersHeightSizeMode,
    DockStyle, Panel, Padding
)
from System.Drawing import Point, Size, Font, FontStyle, Color as DrawColor
from System.Collections.Generic import List
from Autodesk.Revit.DB import Color as RevitColor


# ═══════════════════════════════════════════════════════════════════════════════
#  HELP DIALOG (reusable ? button popup)
# ═══════════════════════════════════════════════════════════════════════════════

class HelpDialog(Form):
    """Reusable help popup triggered by ? buttons."""
    def __init__(self, title, help_text):
        self.Text = title
        self.ClientSize = Size(440, 0)  # height set dynamically
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True

        lbl = Label()
        lbl.Text = help_text
        lbl.Location = Point(15, 15)
        lbl.AutoSize = True
        lbl.MaximumSize = Size(410, 0)

        self.Controls.Add(lbl)

        # Force layout to get label height
        lbl.PerformLayout()
        content_height = lbl.PreferredHeight + 70

        self.ClientSize = Size(440, content_height)

        btn_ok = Button()
        btn_ok.Text = "OK"
        btn_ok.Size = Size(80, 28)
        btn_ok.Location = Point(
            (440 - 80) // 2, content_height - 42)
        btn_ok.Click += lambda s, e: self.Close()

        self.Controls.Add(btn_ok)


def make_help_button(parent, x, y, title, help_text):
    """Create a small ? button at (x,y) on parent that opens a HelpDialog."""
    btn = Button()
    btn.Text = "?"
    btn.Size = Size(24, 24)
    btn.Location = Point(x, y)
    btn.Font = Font("Segoe UI", 8, FontStyle.Bold)

    def on_click(sender, args):
        dlg = HelpDialog(title, help_text)
        try:
            dlg.ShowDialog()
        finally:
            dlg.Dispose()

    btn.Click += on_click
    parent.Controls.Add(btn)
    return btn

output = script.get_output()
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView


# ═══════════════════════════════════════════════════════════════════════════════
#  SETTINGS DIALOG (modal)
# ═══════════════════════════════════════════════════════════════════════════════

class SettingsDialog(Form):
    def __init__(self):
        self.Text = "PCF Export Settings"
        self.ClientSize = Size(310, 330)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        grp_mode = GroupBox()
        grp_mode.Text = "Export Mode"
        grp_mode.Location = Point(12, 10)
        grp_mode.Size = Size(285, 70)

        self.rb_full = RadioButton()
        self.rb_full.Text = "Full Assembly"
        self.rb_full.Location = Point(15, 22)
        self.rb_full.AutoSize = True
        self.rb_full.Checked = True
        self.rb_full.CheckedChanged += self.on_mode_changed

        self.rb_selection = RadioButton()
        self.rb_selection.Text = "Selection Set"
        self.rb_selection.Location = Point(15, 44)
        self.rb_selection.AutoSize = True
        self.rb_selection.CheckedChanged += self.on_mode_changed

        grp_mode.Controls.Add(self.rb_full)
        grp_mode.Controls.Add(self.rb_selection)

        self.chk_hide = CheckBox()
        self.chk_hide.Text = "Hide elements after process"
        self.chk_hide.Location = Point(15, 90)
        self.chk_hide.AutoSize = True

        self.grp_naming = GroupBox()
        self.grp_naming.Text = "Naming (Selection Set)"
        self.grp_naming.Location = Point(12, 118)
        self.grp_naming.Size = Size(285, 70)
        self.grp_naming.Enabled = False

        self.rb_auto = RadioButton()
        self.rb_auto.Text = "Auto-increment  (.1, .2, .3 ...)"
        self.rb_auto.Location = Point(15, 22)
        self.rb_auto.AutoSize = True
        self.rb_auto.Checked = True

        self.rb_custom = RadioButton()
        self.rb_custom.Text = "Custom name  (prompt each time)"
        self.rb_custom.Location = Point(15, 44)
        self.rb_custom.AutoSize = True

        self.grp_naming.Controls.Add(self.rb_auto)
        self.grp_naming.Controls.Add(self.rb_custom)

        self.btn_filter = Button()
        self.btn_filter.Text = "Group Export by Parameter..."
        self.btn_filter.Location = Point(12, 200)
        self.btn_filter.Size = Size(285, 30)
        self.btn_filter.Click += self.on_filter

        self.btn_start = Button()
        self.btn_start.Text = "Start"
        self.btn_start.Location = Point(120, 290)
        self.btn_start.Size = Size(80, 30)
        self.btn_start.Click += self.on_start

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(210, 290)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel

        self.Controls.Add(grp_mode)
        self.Controls.Add(self.chk_hide)
        self.Controls.Add(self.grp_naming)
        self.Controls.Add(self.btn_filter)
        self.Controls.Add(self.btn_start)
        self.Controls.Add(self.btn_cancel)

        # ── Help buttons ──
        make_help_button(self, 275, 10, "PCF Export - Tool Overview",
            "WORKFLOW:\n"
            "1. Choose an Export Mode:\n"
            "     Full Assembly - pick assemblies in Revit,\n"
            "       each exports as its own PCF file named\n"
            "       by the Assembly Type Name.\n"
            "     Selection Set - pick fabrication parts,\n"
            "       exports them as a PCF. Parts can be in\n"
            "       an assembly or loose (name required).\n"
            "2. Optionally check 'Hide elements after process'\n"
            "     to fade out exported parts in the view.\n"
            "3. For Selection Set, choose a naming convention.\n"
            "4. Click Start, pick a folder, then use the\n"
            "     Process/Multiple/Close dialog to export.\n"
            "5. Or click 'Group Export by Parameter' to batch\n"
            "     export parts grouped by any Revit parameter.\n\n"
            "POST-PROCESSING (all modes):\n"
            "- ITEM-NUMBER is inserted into the PCF from\n"
            "  each part's ItemNumber via GUID matching.\n"
            "- 'STD' schedule references become 'SCH STD'.\n"
            "- PIPELINE-REFERENCE and SPOOL-IDENTIFIER\n"
            "  are set to match the exported file name.\n"
            "- Ghost assemblies created by Revit during\n"
            "  ExportToPCF on loose parts are auto-deleted.\n\n"
            "KNOWN LIMITATIONS:\n"
            "- Requires Revit 2023+ with Fabrication Parts.\n"
            "- Parts must have valid PartGuid for ITEM-NUMBER.\n"
            "- CP_PCF Report Data parameter is written for\n"
            "  Full Assembly mode only.")

        make_help_button(grp_mode, 257, 0, "Export Mode",
            "Full Assembly:\n"
            "  You will pick AssemblyInstance elements in Revit.\n"
            "  Only assemblies can be highlighted/selected.\n"
            "  Each assembly's members are exported as one PCF.\n"
            "  File name = Assembly Type Name.\n"
            "  The CP_PCF Report Data parameter is stamped.\n\n"
            "Selection Set:\n"
            "  You will select FabricationParts in Revit.\n"
            "  If parts are in an assembly, that name is used.\n"
            "  If parts are NOT in an assembly, you must\n"
            "  provide a name (mandatory prompt).\n"
            "  Ghost assemblies are auto-cleaned after export.")

        make_help_button(self.grp_naming, 257, 0, "Naming (Selection Set)",
            "Auto-increment (.1, .2, .3 ...):\n"
            "  Each export gets AssemblyName.1, .2, .3, etc.\n"
            "  Counters reset each time the tool runs.\n\n"
            "Custom name (prompt each time):\n"
            "  A dialog appears before each export letting\n"
            "  you type a custom file name. A default is\n"
            "  suggested based on the assembly name + counter.")

        self.result = None

    def on_mode_changed(self, sender, args):
        self.grp_naming.Enabled = self.rb_selection.Checked

    def on_filter(self, sender, args):
        self.result = "filter"
        self.Close()

    def on_start(self, sender, args):
        self.result = "start"
        self.Close()

    def on_cancel(self, sender, args):
        self.result = "cancel"
        self.Close()


# ═══════════════════════════════════════════════════════════════════════════════
#  PROCESS DIALOG (non-modal)
# ═══════════════════════════════════════════════════════════════════════════════

class ProcessDialog(Form):
    def __init__(self, mode="full_assembly"):
        self.Text = "PCF Export"
        self.ClientSize = Size(285, 40)
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.StartPosition = FormStartPosition.CenterScreen
        self.TopMost = True

        self.chk_multiple = CheckBox()
        self.chk_multiple.Text = "Multiple"
        self.chk_multiple.Checked = True
        self.chk_multiple.Location = Point(8, 9)
        self.chk_multiple.AutoSize = True

        self.btn_process = Button()
        self.btn_process.Text = "Process"
        self.btn_process.Location = Point(100, 6)
        self.btn_process.Size = Size(70, 28)
        self.btn_process.Click += self.on_process

        self.btn_close = Button()
        self.btn_close.Text = "Close"
        self.btn_close.Location = Point(176, 6)
        self.btn_close.Size = Size(70, 28)
        self.btn_close.Click += self.on_close

        self.Controls.Add(self.chk_multiple)
        self.Controls.Add(self.btn_process)
        self.Controls.Add(self.btn_close)

        help_text = ProcessDialog.get_help_text(mode)
        make_help_button(self, 253, 8, "Process Dialog", help_text)

        self.user_action = None
        self._is_open = False

    @staticmethod
    def get_help_text(mode):
        """Return mode-specific help text for the process dialog."""
        if mode == "full_assembly":
            return (
                "PROCESS DIALOG - Full Assembly Mode:\n\n"
                "1. Click 'Process' — Revit's pick mode activates.\n"
                "     Only AssemblyInstance elements can be selected.\n"
                "     Pick one or more assemblies, then click Finish\n"
                "     in Revit's options bar.\n"
                "2. Each picked assembly is exported as a PCF file\n"
                "     using the Assembly Type Name.\n"
                "3. Already-exported assemblies are skipped (no dupes).\n"
                "4. If 'Multiple' is checked, the dialog stays open\n"
                "     so you can pick and process more assemblies.\n"
                "5. Click 'Close' when done. Hidden elements\n"
                "     are restored and a report is printed.")
        else:
            return (
                "PROCESS DIALOG - Selection Set Mode:\n\n"
                "1. Select FabricationParts in Revit using\n"
                "     any standard selection method.\n"
                "2. Click 'Process' to export the current selection.\n"
                "3. If parts are in an assembly, the assembly name\n"
                "     is used for the PCF file name.\n"
                "   If parts are NOT in an assembly, a name dialog\n"
                "     appears (name is mandatory).\n"
                "4. If 'Multiple' is checked, the dialog stays open\n"
                "     so you can select and process more sets.\n"
                "5. Click 'Close' when done. Hidden elements\n"
                "     are restored and a report is printed.")

    def on_process(self, sender, args):
        self.user_action = "process"
        self._is_open = False
        self.Hide()

    def on_close(self, sender, args):
        self.user_action = "close"
        self._is_open = False
        self.Hide()

    def OnFormClosing(self, e):
        if self.user_action is None:
            self.user_action = "close"
        self._is_open = False
        Form.OnFormClosing(self, e)

    def show_and_wait(self):
        self.user_action = None
        self._is_open = True
        self.Show()
        while self._is_open:
            Application.DoEvents()
            time.sleep(0.05)
        return self.user_action


# ═══════════════════════════════════════════════════════════════════════════════
#  NAME DIALOGS
# ═══════════════════════════════════════════════════════════════════════════════

class NameDialog(Form):
    """Optional custom name dialog (for Selection Set custom naming)."""
    def __init__(self, default_name):
        self.Text = "PCF File Name"
        self.ClientSize = Size(350, 75)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True

        lbl = Label()
        lbl.Text = "File name:"
        lbl.Location = Point(10, 14)
        lbl.AutoSize = True

        self.txt_name = TextBox()
        self.txt_name.Text = default_name
        self.txt_name.Location = Point(75, 11)
        self.txt_name.Size = Size(190, 22)

        lbl_ext = Label()
        lbl_ext.Text = ".pcf"
        lbl_ext.Location = Point(268, 14)
        lbl_ext.AutoSize = True

        self.btn_save = Button()
        self.btn_save.Text = "Save"
        self.btn_save.Location = Point(180, 42)
        self.btn_save.Size = Size(75, 28)
        self.btn_save.Click += self.on_save

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(262, 42)
        self.btn_cancel.Size = Size(75, 28)
        self.btn_cancel.Click += self.on_cancel

        self.Controls.Add(lbl)
        self.Controls.Add(self.txt_name)
        self.Controls.Add(lbl_ext)
        self.Controls.Add(self.btn_save)
        self.Controls.Add(self.btn_cancel)

        self.file_name = None

        make_help_button(self, 318, 8, "PCF File Name",
            "Enter a custom name for this PCF export.\n\n"
            "The .pcf extension is added automatically.\n"
            "A default name is suggested based on the\n"
            "assembly name and export counter.\n\n"
            "Click Save to use the name, or Cancel\n"
            "to skip this export and pick again.")

    def on_save(self, sender, args):
        name = self.txt_name.Text.strip()
        if name:
            self.file_name = name
        self.Close()

    def on_cancel(self, sender, args):
        self.file_name = None
        self.Close()


class MandatoryNameDialog(Form):
    """Mandatory name dialog for unassembled parts. No cancel — must provide name."""
    def __init__(self):
        self.Text = "Name Required"
        self.ClientSize = Size(380, 95)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.TopMost = True
        self.ControlBox = False  # no X button

        lbl_warn = Label()
        lbl_warn.Text = "Selected parts are not in an assembly. Enter a name:"
        lbl_warn.Location = Point(10, 10)
        lbl_warn.AutoSize = True

        lbl_name = Label()
        lbl_name.Text = "Name:"
        lbl_name.Location = Point(10, 42)
        lbl_name.AutoSize = True

        self.txt_name = TextBox()
        self.txt_name.Text = ""
        self.txt_name.Location = Point(55, 39)
        self.txt_name.Size = Size(230, 22)

        lbl_ext = Label()
        lbl_ext.Text = ".pcf"
        lbl_ext.Location = Point(288, 42)
        lbl_ext.AutoSize = True

        self.btn_ok = Button()
        self.btn_ok.Text = "OK"
        self.btn_ok.Location = Point(210, 65)
        self.btn_ok.Size = Size(75, 28)
        self.btn_ok.Click += self.on_ok

        self.btn_skip = Button()
        self.btn_skip.Text = "Skip"
        self.btn_skip.Location = Point(292, 65)
        self.btn_skip.Size = Size(75, 28)
        self.btn_skip.Click += self.on_skip

        self.Controls.Add(lbl_warn)
        self.Controls.Add(lbl_name)
        self.Controls.Add(self.txt_name)
        self.Controls.Add(lbl_ext)
        self.Controls.Add(self.btn_ok)
        self.Controls.Add(self.btn_skip)

        self.file_name = None
        self.skipped = False

        make_help_button(self, 348, 8, "Name Required",
            "The selected parts are NOT inside an assembly,\n"
            "so there is no automatic name to use.\n\n"
            "You must provide a name for the PCF file.\n"
            "This name is also used for:\n"
            "  - PIPELINE-REFERENCE in the PCF\n"
            "  - SPOOL-IDENTIFIER in the PCF\n\n"
            "Click OK to export with the given name.\n"
            "Click Skip to skip this set and pick again.")

    def on_ok(self, sender, args):
        name = self.txt_name.Text.strip()
        if not name:
            forms.alert("A name is required to export unassembled parts.")
            return
        self.file_name = name
        self.Close()

    def on_skip(self, sender, args):
        self.skipped = True
        self.Close()


# ═══════════════════════════════════════════════════════════════════════════════
#  GROUP EXPORT — Collect parts, group by parameter(s), export each as own PCF
# ═══════════════════════════════════════════════════════════════════════════════

class FabPartSelectionFilter(ISelectionFilter):
    """Filter that only allows FabricationPart selection in Revit picker."""
    def AllowElement(self, element):
        return isinstance(element, FabricationPart)
    def AllowReference(self, reference, position):
        return False


class AssemblySelectionFilter(ISelectionFilter):
    """Filter that only allows AssemblyInstance selection in Revit picker."""
    def AllowElement(self, element):
        return isinstance(element, AssemblyInstance)
    def AllowReference(self, reference, position):
        return False


def collect_fab_parts(source):
    """Collect FabricationParts from 'selection', 'view', or 'pick'.
    Returns list of FabricationPart elements."""
    if source == "pick":
        try:
            from Autodesk.Revit.UI.Selection import ObjectType
            sel_filter = FabPartSelectionFilter()
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element, sel_filter,
                "Select Fabrication Parts, then click Finish")
            parts = []
            for ref in refs:
                el = doc.GetElement(ref.ElementId)
                if isinstance(el, FabricationPart):
                    parts.append(el)
            return parts
        except Exception:
            # User cancelled the pick
            return []
    elif source == "selection":
        sel_ids = list(uidoc.Selection.GetElementIds())
        parts = []
        for eid in sel_ids:
            el = doc.GetElement(eid)
            if isinstance(el, FabricationPart):
                parts.append(el)
        return parts
    else:
        collector = FilteredElementCollector(doc, active_view.Id) \
            .OfClass(FabricationPart) \
            .WhereElementIsNotElementType()
        return list(collector)


def get_param_value_string(part, param_name):
    """Get a parameter's display value as string from a FabricationPart."""
    if param_name == "ItemNumber":
        try:
            return part.ItemNumber or ""
        except Exception:
            return ""
    if param_name == "ServiceName":
        try:
            return part.ServiceName or ""
        except Exception:
            return ""
    param = part.LookupParameter(param_name)
    if param is None or not param.HasValue:
        return ""
    if param.StorageType == StorageType.String:
        return param.AsString() or ""
    elif param.StorageType == StorageType.Double:
        return param.AsValueString() or str(param.AsDouble())
    elif param.StorageType == StorageType.Integer:
        return param.AsValueString() or str(param.AsInteger())
    elif param.StorageType == StorageType.ElementId:
        return param.AsValueString() or str(
            param.AsElementId().IntegerValue)
    return ""


def get_all_param_names(parts):
    """Get sorted list of all parameter names across parts."""
    names = set()
    for part in parts:
        for param in part.Parameters:
            names.add(param.Definition.Name)
    names.add("ItemNumber")
    names.add("ServiceName")
    return sorted(names)


def sanitize_filename(name):
    """Remove characters not allowed in filenames."""
    for ch in '/\\:*?"<>|':
        name = name.replace(ch, "-" if ch in "/\\:" else "")
    return name.strip() or "unnamed"


class SourcePickerDialog(Form):
    """Pick where parts come from: current selection, active view, or pick."""
    def __init__(self):
        self.Text = "Group Export - Part Source"
        self.ClientSize = Size(310, 155)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        lbl = Label()
        lbl.Text = "Collect FabricationParts from:"
        lbl.Location = Point(15, 10)
        lbl.AutoSize = True

        self.rb_selection = RadioButton()
        self.rb_selection.Text = "Current Revit selection"
        self.rb_selection.Location = Point(20, 32)
        self.rb_selection.Size = Size(270, 20)
        self.rb_selection.Checked = True

        self.rb_view = RadioButton()
        self.rb_view.Text = "All parts in active view"
        self.rb_view.Location = Point(20, 54)
        self.rb_view.Size = Size(270, 20)

        self.rb_pick = RadioButton()
        self.rb_pick.Text = "Pick elements in Revit (Fab Parts only)"
        self.rb_pick.Location = Point(20, 76)
        self.rb_pick.Size = Size(270, 20)

        self.btn_ok = Button()
        self.btn_ok.Text = "OK"
        self.btn_ok.Location = Point(130, 110)
        self.btn_ok.Size = Size(75, 28)
        self.btn_ok.Click += self.on_ok

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(215, 110)
        self.btn_cancel.Size = Size(75, 28)
        self.btn_cancel.Click += self.on_cancel

        self.Controls.Add(lbl)
        self.Controls.Add(self.rb_selection)
        self.Controls.Add(self.rb_view)
        self.Controls.Add(self.rb_pick)
        self.Controls.Add(self.btn_ok)
        self.Controls.Add(self.btn_cancel)
        self.source = None

        make_help_button(self, 275, 5, "Part Source",
            "Choose where to collect FabricationParts from:\n\n"
            "Current Revit selection:\n"
            "  Uses whatever is already selected in Revit.\n"
            "  Select parts before opening this tool.\n\n"
            "All parts in active view:\n"
            "  Collects every FabricationPart visible in\n"
            "  the current view. Good for full-view exports.\n\n"
            "Pick elements in Revit:\n"
            "  Opens Revit's pick mode. Only FabricationParts\n"
            "  can be highlighted/selected. Pick parts, then\n"
            "  click Finish in Revit's options bar.")

    def on_ok(self, sender, args):
        if self.rb_pick.Checked:
            self.source = "pick"
        elif self.rb_view.Checked:
            self.source = "view"
        else:
            self.source = "selection"
        self.Close()

    def on_cancel(self, sender, args):
        self.source = None
        self.Close()


class ColumnPickerDialog(Form):
    """Add/Remove display columns + toggle Group-by.
    Search filter for available params, up/down ordering."""
    def __init__(self, all_param_names, current_columns, current_group_params):
        self.Text = "Columns & Grouping"
        self.ClientSize = Size(660, 490)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition = FormStartPosition.CenterScreen
        self.MaximizeBox = False
        self.MinimizeBox = False

        self._all = all_param_names
        self._selected_list = list(current_columns)
        self._current_groups = set(current_group_params)

        # ── Left: Search + Available list ──
        lbl_avail = Label()
        lbl_avail.Text = "Available Parameters:"
        lbl_avail.Location = Point(10, 8)
        lbl_avail.AutoSize = True

        lbl_search = Label()
        lbl_search.Text = "Search:"
        lbl_search.Location = Point(10, 30)
        lbl_search.AutoSize = True

        self.txt_search = TextBox()
        self.txt_search.Location = Point(58, 27)
        self.txt_search.Size = Size(172, 22)
        self.txt_search.TextChanged += self.on_search_changed

        self.lst_avail = ListBox()
        self.lst_avail.Location = Point(10, 52)
        self.lst_avail.Size = Size(220, 350)
        self.lst_avail.SelectionMode = SelectionMode.MultiExtended

        # ── Center: Add/Remove buttons ──
        self.btn_add = Button()
        self.btn_add.Text = "Add >>"
        self.btn_add.Location = Point(240, 160)
        self.btn_add.Size = Size(80, 28)
        self.btn_add.Click += self.on_add

        self.btn_remove = Button()
        self.btn_remove.Text = "<< Remove"
        self.btn_remove.Location = Point(240, 195)
        self.btn_remove.Size = Size(80, 28)
        self.btn_remove.Click += self.on_remove

        # ── Right: Selected grid ──
        lbl_sel = Label()
        lbl_sel.Text = "Display Columns (check 'Group' to group by):"
        lbl_sel.Location = Point(330, 8)
        lbl_sel.AutoSize = True

        self.grid_sel = DataGridView()
        self.grid_sel.Location = Point(330, 28)
        self.grid_sel.Size = Size(270, 374)
        self.grid_sel.AllowUserToAddRows = False
        self.grid_sel.AllowUserToDeleteRows = False
        self.grid_sel.ColumnHeadersHeightSizeMode = \
            DataGridViewColumnHeadersHeightSizeMode.AutoSize
        self.grid_sel.SelectionMode = \
            DataGridViewSelectionMode.FullRowSelect
        self.grid_sel.RowHeadersVisible = False

        col_name = DataGridViewTextBoxColumn()
        col_name.HeaderText = "Parameter"
        col_name.Name = "param"
        col_name.Width = 200
        col_name.ReadOnly = True
        self.grid_sel.Columns.Add(col_name)

        col_grp = DataGridViewCheckBoxColumn()
        col_grp.HeaderText = "Group"
        col_grp.Name = "group"
        col_grp.Width = 50
        self.grid_sel.Columns.Add(col_grp)

        # ── Up/Down buttons ──
        self.btn_up = Button()
        self.btn_up.Text = u"\u25B2"
        self.btn_up.Location = Point(610, 160)
        self.btn_up.Size = Size(36, 32)
        self.btn_up.Click += self.on_move_up

        self.btn_down = Button()
        self.btn_down.Text = u"\u25BC"
        self.btn_down.Location = Point(610, 198)
        self.btn_down.Size = Size(36, 32)
        self.btn_down.Click += self.on_move_down

        # ── Bottom buttons ──
        self.btn_ok = Button()
        self.btn_ok.Text = "OK"
        self.btn_ok.Location = Point(490, 450)
        self.btn_ok.Size = Size(75, 28)
        self.btn_ok.Click += self.on_ok

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(575, 450)
        self.btn_cancel.Size = Size(75, 28)
        self.btn_cancel.Click += self.on_cancel

        self.Controls.Add(lbl_avail)
        self.Controls.Add(lbl_search)
        self.Controls.Add(self.txt_search)
        self.Controls.Add(self.lst_avail)
        self.Controls.Add(self.btn_add)
        self.Controls.Add(self.btn_remove)
        self.Controls.Add(lbl_sel)
        self.Controls.Add(self.grid_sel)
        self.Controls.Add(self.btn_up)
        self.Controls.Add(self.btn_down)
        self.Controls.Add(self.btn_ok)
        self.Controls.Add(self.btn_cancel)

        self._refresh_lists()

        self.chosen_columns = None
        self.chosen_group_params = None

        make_help_button(self, 625, 8, "Columns & Grouping",
            "COLUMNS:\n"
            "  Choose which parameters to display in the\n"
            "  Group Export grid as preview columns.\n"
            "  Use Search to filter the available list.\n"
            "  Add >> moves selected params to the display.\n"
            "  << Remove takes them out.\n"
            "  Use the arrow buttons to reorder columns.\n\n"
            "GROUPING:\n"
            "  Check the 'Group' checkbox next to any\n"
            "  parameter to group parts by that value.\n"
            "  Each unique value (or combination) becomes\n"
            "  a separate PCF file.\n\n"
            "  Example: Group by 'Comments' — all parts\n"
            "  with Comments='CHWS' export as one PCF,\n"
            "  parts with 'CHWR' as another, etc.\n\n"
            "  Multiple group params create compound keys:\n"
            "  e.g. grouping by Service + Comments gives\n"
            "  'Domestic Cold Water-PIPE1' as a group.")

    def on_search_changed(self, sender, args):
        """Filter the available list by search text."""
        filt = self.txt_search.Text.strip().lower()
        selected_set = set(self._selected_list)
        self.lst_avail.Items.Clear()
        for p in self._all:
            if p in selected_set:
                continue
            if filt and filt not in p.lower():
                continue
            self.lst_avail.Items.Add(p)

    def _refresh_lists(self, restore_sel_index=-1):
        """Rebuild both lists preserving _selected_list order."""
        selected_set = set(self._selected_list)
        filt = self.txt_search.Text.strip().lower() if hasattr(self, 'txt_search') else ""

        self.lst_avail.Items.Clear()
        for p in self._all:
            if p in selected_set:
                continue
            if filt and filt not in p.lower():
                continue
            self.lst_avail.Items.Add(p)

        self.grid_sel.Rows.Clear()
        for p in self._selected_list:
            is_grp = p in self._current_groups
            self.grid_sel.Rows.Add(p, is_grp)

        # Restore selection in grid
        if 0 <= restore_sel_index < self.grid_sel.Rows.Count:
            self.grid_sel.ClearSelection()
            self.grid_sel.Rows[restore_sel_index].Selected = True
            self.grid_sel.CurrentCell = \
                self.grid_sel.Rows[restore_sel_index].Cells[0]

    def _save_group_checks(self):
        """Read group checkboxes from grid back into _current_groups."""
        self._current_groups = set()
        for i in range(self.grid_sel.Rows.Count):
            row = self.grid_sel.Rows[i]
            if row.IsNewRow:
                continue
            gval = row.Cells["group"].Value
            if gval is True or str(gval) == "True":
                self._current_groups.add(str(row.Cells["param"].Value))

    def on_add(self, sender, args):
        to_add = []
        for i in self.lst_avail.SelectedIndices:
            to_add.append(self.lst_avail.Items[i])
        if not to_add:
            return
        self._save_group_checks()
        for p in to_add:
            if p not in self._selected_list:
                self._selected_list.append(p)
        self._refresh_lists()

    def on_remove(self, sender, args):
        to_remove = []
        for i in range(self.grid_sel.Rows.Count):
            row = self.grid_sel.Rows[i]
            if row.IsNewRow:
                continue
            if row.Selected:
                to_remove.append(str(row.Cells["param"].Value))
        self._save_group_checks()
        for p in to_remove:
            if p in self._selected_list:
                self._selected_list.remove(p)
            self._current_groups.discard(p)
        self._refresh_lists()

    def _get_selected_row_index(self):
        """Get the index of the first selected row in grid_sel."""
        for i in range(self.grid_sel.Rows.Count):
            if self.grid_sel.Rows[i].Selected:
                return i
        return -1

    def on_move_up(self, sender, args):
        idx = self._get_selected_row_index()
        if idx <= 0:
            return
        self._save_group_checks()
        self._selected_list[idx], self._selected_list[idx - 1] = \
            self._selected_list[idx - 1], self._selected_list[idx]
        self._refresh_lists(restore_sel_index=idx - 1)

    def on_move_down(self, sender, args):
        idx = self._get_selected_row_index()
        if idx < 0 or idx >= len(self._selected_list) - 1:
            return
        self._save_group_checks()
        self._selected_list[idx], self._selected_list[idx + 1] = \
            self._selected_list[idx + 1], self._selected_list[idx]
        self._refresh_lists(restore_sel_index=idx + 1)

    def on_ok(self, sender, args):
        self._save_group_checks()
        if not self._selected_list:
            forms.alert("Select at least one column.")
            return
        self.chosen_columns = list(self._selected_list)
        self.chosen_group_params = [
            p for p in self._selected_list
            if p in self._current_groups
        ]
        self.Close()

    def on_cancel(self, sender, args):
        self.chosen_columns = None
        self.chosen_group_params = None
        self.Close()


class GroupExportDialog(Form):
    """Group Export dialog.
    Top: Naming mode radios, Prefix field, Columns & Grouping button.
    Grid: grouped rows with Export checkbox and editable PCF Name.
    Naming modes:
      - Prefix-Auto: Prefix-GroupValue  (default)
      - Assembly-Auto: AssemblyName-GroupValue
      - Custom: edit each name directly in the grid
    Bottom: Check All / Uncheck All, Refresh Names, Export buttons.
    """
    def __init__(self, parts, all_param_names):
        self.Text = "Group Export by Parameter"
        self.ClientSize = Size(960, 640)
        self.FormBorderStyle = FormBorderStyle.Sizable
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimizeBox = False

        self._parts = parts
        self._all_param_names = all_param_names
        self._groups = {}
        self._group_order = []
        self._group_params = []

        # Try to detect assembly name from first part
        self._assembly_name = ""
        for p in parts:
            try:
                aid = p.AssemblyInstanceId
                if aid and aid != ElementId.InvalidElementId:
                    assy = doc.GetElement(aid)
                    if isinstance(assy, AssemblyInstance):
                        self._assembly_name = assy.AssemblyTypeName
                        break
            except Exception:
                pass

        # Default display columns
        self._display_columns = []
        defaults = [
            "Fabrication Service Name", "Comments"
        ]
        for d in defaults:
            if d in all_param_names:
                self._display_columns.append(d)

        # ── Top panel (two rows: naming mode + prefix/info) ──
        top_panel = Panel()
        top_panel.Height = 72
        top_panel.Dock = DockStyle.Top
        top_panel.Padding = Padding(8, 6, 8, 2)

        # Row 1: Naming mode radios
        lbl_naming = Label()
        lbl_naming.Text = "Naming:"
        lbl_naming.Location = Point(10, 8)
        lbl_naming.AutoSize = True

        self.rb_prefix = RadioButton()
        self.rb_prefix.Text = "Prefix - Auto"
        self.rb_prefix.Location = Point(65, 6)
        self.rb_prefix.AutoSize = True
        self.rb_prefix.Checked = True
        self.rb_prefix.CheckedChanged += self.on_naming_changed

        self.rb_assembly = RadioButton()
        self.rb_assembly.Text = "Assembly - Auto"
        self.rb_assembly.Location = Point(180, 6)
        self.rb_assembly.AutoSize = True
        self.rb_assembly.CheckedChanged += self.on_naming_changed

        self.rb_custom = RadioButton()
        self.rb_custom.Text = "Custom - Type Name in PCF Name Field"
        self.rb_custom.Location = Point(315, 6)
        self.rb_custom.AutoSize = True
        self.rb_custom.CheckedChanged += self.on_naming_changed

        self.btn_columns = Button()
        self.btn_columns.Text = "Columns && Grouping..."
        self.btn_columns.Location = Point(770, 3)
        self.btn_columns.Size = Size(170, 26)
        self.btn_columns.Click += self.on_edit_columns

        # Row 2: Prefix field + info
        lbl_prefix = Label()
        lbl_prefix.Text = "Prefix:"
        lbl_prefix.Location = Point(10, 40)
        lbl_prefix.AutoSize = True

        self.txt_prefix = TextBox()
        self.txt_prefix.Location = Point(50, 37)
        self.txt_prefix.Size = Size(180, 22)
        self.txt_prefix.TextChanged += self.on_prefix_changed

        self.lbl_assy = Label()
        assy_display = self._assembly_name if self._assembly_name \
            else "(no assembly found)"
        self.lbl_assy.Text = "Assembly: {}".format(assy_display)
        self.lbl_assy.Location = Point(245, 40)
        self.lbl_assy.AutoSize = True

        self.lbl_info = Label()
        self.lbl_info.Text = "{} parts — open Columns & Grouping to set groups".format(
            len(parts))
        self.lbl_info.Location = Point(550, 40)
        self.lbl_info.AutoSize = True

        top_panel.Controls.Add(lbl_naming)
        top_panel.Controls.Add(self.rb_prefix)
        top_panel.Controls.Add(self.rb_assembly)
        top_panel.Controls.Add(self.rb_custom)
        top_panel.Controls.Add(self.btn_columns)
        top_panel.Controls.Add(lbl_prefix)
        top_panel.Controls.Add(self.txt_prefix)
        top_panel.Controls.Add(self.lbl_assy)
        top_panel.Controls.Add(self.lbl_info)

        # ── Bottom panel ──
        bot_panel = Panel()
        bot_panel.Height = 45
        bot_panel.Dock = DockStyle.Bottom
        bot_panel.Padding = Padding(8, 4, 8, 8)

        self.btn_check_all = Button()
        self.btn_check_all.Text = "Check All"
        self.btn_check_all.Location = Point(10, 8)
        self.btn_check_all.Size = Size(80, 28)
        self.btn_check_all.Click += self.on_check_all

        self.btn_uncheck_all = Button()
        self.btn_uncheck_all.Text = "Uncheck All"
        self.btn_uncheck_all.Location = Point(95, 8)
        self.btn_uncheck_all.Size = Size(85, 28)
        self.btn_uncheck_all.Click += self.on_uncheck_all

        self.btn_refresh = Button()
        self.btn_refresh.Text = "Refresh Names"
        self.btn_refresh.Location = Point(200, 8)
        self.btn_refresh.Size = Size(110, 28)
        self.btn_refresh.Click += self.on_refresh_names

        self.btn_export_checked = Button()
        self.btn_export_checked.Text = "Export Checked Groups"
        self.btn_export_checked.Location = Point(620, 8)
        self.btn_export_checked.Size = Size(160, 28)
        self.btn_export_checked.Click += self.on_export_checked

        self.btn_export_all = Button()
        self.btn_export_all.Text = "Export All Groups"
        self.btn_export_all.Location = Point(790, 8)
        self.btn_export_all.Size = Size(140, 28)
        self.btn_export_all.Click += self.on_export_all

        bot_panel.Controls.Add(self.btn_check_all)
        bot_panel.Controls.Add(self.btn_uncheck_all)
        bot_panel.Controls.Add(self.btn_refresh)
        bot_panel.Controls.Add(self.btn_export_checked)
        bot_panel.Controls.Add(self.btn_export_all)

        # ── Grid ──
        self.grid = DataGridView()
        self.grid.Dock = DockStyle.Fill
        self.grid.AllowUserToAddRows = False
        self.grid.AllowUserToDeleteRows = False
        self.grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        self.grid.ColumnHeadersHeightSizeMode = \
            DataGridViewColumnHeadersHeightSizeMode.AutoSize

        self._show_empty_prompt()

        self.Controls.Add(self.grid)
        self.Controls.Add(bot_panel)
        self.Controls.Add(top_panel)

        self.export_plan = None
        self._grid_ready = False  # flag to avoid refresh before grid built

        # ── Help buttons ──
        # Main ? on the top panel (tool-level help)
        make_help_button(top_panel, 745, 3, "Group Export - Full Guide",
            "GROUP EXPORT BY PARAMETER\n"
            "═══════════════════════════════\n\n"
            "PURPOSE:\n"
            "Batch-export FabricationParts into separate PCF files,\n"
            "one per unique parameter value (or combination).\n\n"
            "STEP-BY-STEP:\n"
            "1. You already chose a part source (selection, view,\n"
            "     or pick) on the previous screen.\n"
            "2. Click 'Columns & Grouping' to choose which\n"
            "     parameters to display and which to group by.\n"
            "     Check the 'Group' box next to the parameter(s)\n"
            "     you want to split exports on.\n"
            "3. The grid populates with one row per unique group.\n"
            "     Each row shows: Export checkbox, PCF Name,\n"
            "     part count, group values, and preview columns.\n"
            "4. Choose a Naming mode (top row):\n"
            "     Prefix-Auto: Prefix + group value\n"
            "     Assembly-Auto: Assembly name + group value\n"
            "     Custom: type each name in the PCF Name column\n"
            "5. Check/uncheck rows to control which groups export.\n"
            "6. Click 'Export Checked Groups' or 'Export All Groups'.\n"
            "7. Pick an output folder. Each group exports as a\n"
            "     separate PCF file.\n\n"
            "POST-PROCESSING (automatic):\n"
            "- ITEM-NUMBER inserted via GUID matching.\n"
            "- STD replaced with SCH STD.\n"
            "- PIPELINE-REFERENCE and SPOOL-IDENTIFIER set\n"
            "  to the PCF file name for each group.\n"
            "- Ghost assemblies auto-cleaned.\n\n"
            "BUTTONS:\n"
            "- Check All / Uncheck All: toggle export checkboxes.\n"
            "- Refresh Names: recalculate PCF names from the\n"
            "  current naming mode and prefix.\n"
            "- Columns & Grouping: change display/group params.\n\n"
            "TIPS:\n"
            "- Group by 'Comments' or 'Fabrication Service Name'\n"
            "  to split by system type.\n"
            "- You can group by multiple params for finer splits.\n"
            "- Custom mode lets you manually name each PCF.")

    def _auto_refresh_names(self):
        """Auto-refresh names if grid has rows."""
        if not self._grid_ready or not self._group_order:
            return
        for i in range(self.grid.Rows.Count):
            row = self.grid.Rows[i]
            if row.IsNewRow:
                continue
            if i < len(self._group_order):
                gk = self._group_order[i]
                row.Cells["_pcf_name"].Value = self._build_auto_name(gk)

    def on_naming_changed(self, sender, args):
        self._auto_refresh_names()

    def on_prefix_changed(self, sender, args):
        if self.rb_prefix.Checked:
            self._auto_refresh_names()

    def _show_empty_prompt(self):
        self.grid.Columns.Clear()
        self.grid.Rows.Clear()
        col = DataGridViewTextBoxColumn()
        col.HeaderText = ""
        col.Width = 920
        col.ReadOnly = True
        self.grid.Columns.Add(col)
        self.grid.Rows.Add(
            "Click 'Columns & Grouping' to choose display columns "
            "and check which parameters to group by.")

    def _build_auto_name(self, group_key):
        """Build name based on current naming mode."""
        if self.rb_custom.Checked:
            # Custom mode — just return group key, user will edit
            return sanitize_filename(group_key)
        elif self.rb_assembly.Checked:
            base = self._assembly_name if self._assembly_name else "Assembly"
            return sanitize_filename("{}-{}".format(base, group_key))
        else:
            # Prefix mode (default)
            prefix = self.txt_prefix.Text.strip()
            if prefix:
                return sanitize_filename("{}-{}".format(prefix, group_key))
            return sanitize_filename(group_key)

    def _rebuild_grid(self):
        self.grid.Columns.Clear()
        self.grid.Rows.Clear()
        self._grid_ready = False

        if not self._group_params:
            self._show_empty_prompt()
            return

        # Build groups
        self._groups = {}
        for part in self._parts:
            key_parts = []
            for gp in self._group_params:
                val = get_param_value_string(part, gp)
                key_parts.append(val if val else "(empty)")
            group_key = "-".join(key_parts)
            if group_key not in self._groups:
                self._groups[group_key] = []
            self._groups[group_key].append(part)

        self._group_order = sorted(self._groups.keys())

        # Hide row headers for cleaner look
        self.grid.RowHeadersVisible = False

        # Columns: Export, PCF Name, Parts, group params, preview cols
        chk_col = DataGridViewCheckBoxColumn()
        chk_col.HeaderText = "Export"
        chk_col.Name = "_export"
        chk_col.Width = 50
        self.grid.Columns.Add(chk_col)

        name_col = DataGridViewTextBoxColumn()
        name_col.HeaderText = "PCF Name"
        name_col.Name = "_pcf_name"
        name_col.Width = 230
        self.grid.Columns.Add(name_col)

        cnt_col = DataGridViewTextBoxColumn()
        cnt_col.HeaderText = "Parts"
        cnt_col.Name = "_count"
        cnt_col.Width = 50
        cnt_col.ReadOnly = True
        self.grid.Columns.Add(cnt_col)

        for gp in self._group_params:
            gc = DataGridViewTextBoxColumn()
            gc.HeaderText = gp
            gc.Name = "_grp_" + gp
            gc.ReadOnly = True
            self.grid.Columns.Add(gc)

        preview = [c for c in self._display_columns
                   if c not in self._group_params][:6]
        for cn in preview:
            tc = DataGridViewTextBoxColumn()
            tc.HeaderText = cn
            tc.Name = cn
            tc.ReadOnly = True
            self.grid.Columns.Add(tc)

        # Populate
        for gk in self._group_order:
            parts_in = self._groups[gk]
            first = parts_in[0]
            auto_name = self._build_auto_name(gk)

            row_vals = [True, auto_name, str(len(parts_in))]
            key_vals = gk.split("-") if len(self._group_params) > 1 \
                else [gk]
            while len(key_vals) < len(self._group_params):
                key_vals.append("")
            for kv in key_vals:
                row_vals.append(kv)
            for cn in preview:
                row_vals.append(get_param_value_string(first, cn))
            self.grid.Rows.Add(*row_vals)

        # Auto-size all columns except Export and PCF Name
        for i in range(self.grid.Columns.Count):
            col = self.grid.Columns[i]
            if col.Name in ("_export", "_pcf_name"):
                continue
            # Auto-size by scanning cell content
            max_w = len(col.HeaderText) * 7 + 20
            for r in range(self.grid.Rows.Count):
                cell_val = self.grid.Rows[r].Cells[i].Value
                if cell_val:
                    w = len(str(cell_val)) * 7 + 16
                    if w > max_w:
                        max_w = w
            col.Width = min(max_w, 300)

        self._grid_ready = True
        self.lbl_info.Text = "{} parts in {} groups (by {})".format(
            len(self._parts), len(self._groups),
            ", ".join(self._group_params))

    def on_refresh_names(self, sender, args):
        """Manual refresh button."""
        self._auto_refresh_names()

    def on_edit_columns(self, sender, args):
        cpd = ColumnPickerDialog(
            self._all_param_names,
            self._display_columns,
            self._group_params)
        try:
            cpd.ShowDialog()
            if cpd.chosen_columns is not None:
                self._display_columns = cpd.chosen_columns
                self._group_params = cpd.chosen_group_params or []
                self._rebuild_grid()
        finally:
            cpd.Dispose()

    def on_check_all(self, sender, args):
        for i in range(self.grid.Rows.Count):
            row = self.grid.Rows[i]
            if not row.IsNewRow:
                row.Cells["_export"].Value = True

    def on_uncheck_all(self, sender, args):
        for i in range(self.grid.Rows.Count):
            row = self.grid.Rows[i]
            if not row.IsNewRow:
                row.Cells["_export"].Value = False

    def _collect_export_rows(self, only_checked):
        results = []
        for i in range(self.grid.Rows.Count):
            row = self.grid.Rows[i]
            if row.IsNewRow:
                continue
            if i >= len(self._group_order):
                continue
            if only_checked:
                cell_val = row.Cells["_export"].Value
                if not (cell_val is True or str(cell_val) == "True"):
                    continue
            pcf_name = str(row.Cells["_pcf_name"].Value or "").strip()
            gk = self._group_order[i]
            if not pcf_name:
                pcf_name = self._build_auto_name(gk)
            results.append((pcf_name, gk))
        return results

    def on_export_checked(self, sender, args):
        if not self._groups:
            forms.alert(
                "No groups defined.\n"
                "Click 'Columns & Grouping' and check parameters "
                "to group by.")
            return
        rows = self._collect_export_rows(only_checked=True)
        if not rows:
            forms.alert("No groups checked for export.")
            return
        self.export_plan = {
            "exports": [(name, self._groups[gk]) for name, gk in rows],
        }
        self.Close()

    def on_export_all(self, sender, args):
        if not self._groups:
            forms.alert(
                "No groups defined.\n"
                "Click 'Columns & Grouping' and check parameters "
                "to group by.")
            return
        rows = self._collect_export_rows(only_checked=False)
        self.export_plan = {
            "exports": [(name, self._groups[gk]) for name, gk in rows],
        }
        self.Close()


def run_group_export(export_plan, do_hide):
    """Execute grouped PCF export.
    export_plan = {
        'exports': [(pcf_name_str, [FabricationPart, ...]), ...]
    }
    Returns report_lines list.
    """
    report_lines = []
    all_hidden_ids = []
    total_ghosts = 0

    folder = forms.pick_folder(title="Select Folder for PCF Export")
    if not folder:
        return None

    for file_base, parts in export_plan["exports"]:
        fab_ids = [p.Id for p in parts]
        safe_name = sanitize_filename(file_base)

        filename = os.path.join(folder, safe_name + ".pcf")
        pipeline_override = safe_name

        # Snapshot for ghost detection
        assemblies_before = get_all_assembly_ids()

        # Export
        FabricationUtils.ExportToPCF(
            doc, List[ElementId](fab_ids), filename)

        # Ghost cleanup
        assemblies_after = get_all_assembly_ids()
        ghosts = delete_ghost_assemblies(
            assemblies_before, assemblies_after)
        total_ghosts += ghosts

        if not os.path.exists(filename):
            report_lines.append(
                "FAILED: PCF not created for '{}'".format(safe_name))
            continue

        # Post-process
        guid_map = build_guid_to_itemnumber(fab_ids)
        items_ins, std_rep = postprocess_pcf(
            filename, guid_map, pipeline_override)

        report_lines.append("Export: {}".format(safe_name))
        report_lines.append("  Parts: {}".format(len(fab_ids)))
        report_lines.append(
            "  ITEM-NUMBER inserted: {}".format(items_ins))
        report_lines.append(
            "  STD -> SCH STD: {}".format(std_rep))
        report_lines.append(
            "  PIPELINE-REFERENCE / SPOOL-IDENTIFIER: {}".format(
                pipeline_override))
        if ghosts > 0:
            report_lines.append(
                "  Ghost assemblies cleaned: {}".format(ghosts))
        report_lines.append("  Output: {}".format(filename))
        report_lines.append("")

        if do_hide:
            hide_exported_elements(fab_ids)
            all_hidden_ids.extend(fab_ids)

    if all_hidden_ids:
        restore_exported_elements(all_hidden_ids)

    if total_ghosts > 0:
        report_lines.append(
            "Total ghost assemblies cleaned: {}".format(total_ghosts))
        report_lines.append("")

    return report_lines


# ═══════════════════════════════════════════════════════════════════════════════
#  HIDE / RESTORE
# ═══════════════════════════════════════════════════════════════════════════════

def hide_exported_elements(element_ids):
    if not element_ids:
        return
    id_list = List[ElementId](element_ids)
    t = Transaction(doc, "Hide Exported Parts")
    t.Start()
    try:
        ogs = OverrideGraphicSettings()
        ogs.SetHalftone(True)
        ogs.SetSurfaceTransparency(100)
        white = RevitColor(255, 255, 255)
        ogs.SetProjectionLineColor(white)
        ogs.SetCutLineColor(white)
        ogs.SetProjectionLineWeight(1)
        ogs.SetCutLineWeight(1)
        for eid in element_ids:
            active_view.SetElementOverrides(eid, ogs)
        active_view.HideElements(id_list)
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()


def restore_exported_elements(element_ids):
    if not element_ids:
        return
    id_list = List[ElementId](element_ids)
    default_ogs = OverrideGraphicSettings()
    t = Transaction(doc, "Restore Exported Parts")
    t.Start()
    try:
        active_view.UnhideElements(id_list)
        for eid in element_ids:
            active_view.SetElementOverrides(eid, default_ogs)
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()


# ═══════════════════════════════════════════════════════════════════════════════
#  GHOST ASSEMBLY DETECTION + CLEANUP
# ═══════════════════════════════════════════════════════════════════════════════

def get_all_assembly_ids():
    """Return set of all current AssemblyInstance IntegerValue ids in doc."""
    collector = FilteredElementCollector(doc) \
        .OfClass(AssemblyInstance) \
        .WhereElementIsNotElementType()
    return set(e.Id.IntegerValue for e in collector)


def delete_ghost_assemblies(before_ids, after_ids):
    """Delete any assemblies that appeared between before and after snapshots.
    Returns count of deleted assemblies."""
    new_ids = after_ids - before_ids
    if not new_ids:
        return 0

    count = 0
    t = Transaction(doc, "Remove Ghost Assemblies")
    t.Start()
    try:
        for int_id in new_ids:
            eid = ElementId(int_id)
            el = doc.GetElement(eid)
            if el and isinstance(el, AssemblyInstance):
                doc.Delete(eid)
                count += 1
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
    return count


# ═══════════════════════════════════════════════════════════════════════════════
#  POST-PROCESSING
# ═══════════════════════════════════════════════════════════════════════════════

def build_guid_to_itemnumber(element_ids):
    guid_map = {}
    for eid in element_ids:
        el = doc.GetElement(eid)
        if isinstance(el, FabricationPart):
            guid_str = el.PartGuid.ToString()
            item_num = el.ItemNumber
            guid_map[guid_str] = item_num if item_num else ""
    return guid_map


def postprocess_pcf(filename, guid_to_item, pipeline_name_override=None):
    """
    Post-process a single PCF file.
    - Insert ITEM-NUMBER after ITEM-CODE matched via GUID
    - Replace standalone STD with SCH STD
    - If pipeline_name_override provided, replace all PIPELINE-REFERENCE
      and SPOOL-IDENTIFIER values with it
    Returns (items_inserted, std_replaced).
    """
    if not os.path.exists(filename):
        return (0, 0)

    dir_name = os.path.dirname(filename)
    base_name = os.path.splitext(os.path.basename(filename))[0]
    original = os.path.join(dir_name, base_name + "_Original.pcf")

    if os.path.exists(original):
        os.remove(original)
    os.rename(filename, original)

    # First pass: build ITEM-CODE -> GUID -> ItemNumber
    item_and_guid = []
    line_number = 0
    line_item_code_number = 0
    line_item_code = None

    with open(original, "r") as f:
        for raw_line in f:
            line_number += 1
            line = raw_line.rstrip("\r\n")
            if "ITEM-CODE" in line:
                line_item_code_number = line_number
                line_item_code = line
            if "UNIQUE-COMPONENT-IDENTIFIER" in line:
                item_guid = line.replace(
                    "UNIQUE-COMPONENT-IDENTIFIER", ""
                ).strip()
                item_num = guid_to_item.get(item_guid, "None")
                item_and_guid.append({
                    "line_number": line_item_code_number,
                    "item_code": line_item_code,
                    "guid": item_guid,
                    "item_number": item_num
                })
                line_item_code_number = 0
                line_item_code = None

    # Second pass: rewrite
    items_inserted = 0
    std_replaced = 0
    new_line_number = 0

    with open(original, "r") as reader, open(filename, "w") as writer:
        for raw_line in reader:
            line = raw_line.rstrip("\r\n")
            new_line_number += 1

            has_item_code = "ITEM-CODE" in line
            words = line.split()
            has_std = (not has_item_code
                       and any(w.upper() == "STD" for w in words))

            # ── PIPELINE-REFERENCE override ──
            if (pipeline_name_override
                    and line.startswith("PIPELINE-REFERENCE ")):
                writer.write(
                    "PIPELINE-REFERENCE "
                    + pipeline_name_override + "\n"
                )
                continue

            # ── SPOOL-IDENTIFIER override ──
            if (pipeline_name_override
                    and "SPOOL-IDENTIFIER" in line):
                indent = line[:len(line) - len(line.lstrip())]
                writer.write(
                    indent + "SPOOL-IDENTIFIER  "
                    + pipeline_name_override + "\n"
                )
                continue

            # ── Normal lines ──
            if not has_item_code and not has_std:
                writer.write(line + "\n")

            # ── ITEM-CODE lines ──
            if has_item_code:
                last_item = "IsLast"
                for entry in item_and_guid:
                    if entry["line_number"] == new_line_number:
                        last_item = None
                        writer.write(line + "\n")
                        writer.write(
                            "\tITEM-NUMBER  "
                            + entry["item_number"] + "\n"
                        )
                        items_inserted += 1
                if last_item == "IsLast":
                    writer.write(line + "\n")

            # ── STD replacement ──
            if has_std:
                replaced = False
                for word in words:
                    if word.upper() == "STD":
                        writer.write(
                            line.replace(word, "SCH STD", 1) + "\n"
                        )
                        std_replaced += 1
                        replaced = True
                        break
                if not replaced:
                    writer.write(line + "\n")

    try:
        os.remove(original)
    except Exception:
        pass

    return (items_inserted, std_replaced)


def write_pcf_report_param(assembly, assembly_name):
    timestamp = datetime.datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
    param_value = "PCF Report - {} - {}.pcf;".format(
        timestamp, assembly_name
    )
    try:
        t = Transaction(doc, "Write Parameter Values")
        t.Start()
        param = assembly.LookupParameter("CP_PCF Report Data")
        if param and not param.IsReadOnly:
            param.Set(param_value)
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()


def get_assembly_name_from_parts(element_ids):
    """Check if ALL parts share the same assembly. Returns name or None."""
    assembly_name = None
    for eid in element_ids:
        el = doc.GetElement(eid)
        if isinstance(el, FabricationPart):
            try:
                assy_id = el.AssemblyInstanceId
                if assy_id and assy_id != ElementId.InvalidElementId:
                    assy = doc.GetElement(assy_id)
                    if isinstance(assy, AssemblyInstance):
                        if assembly_name is None:
                            assembly_name = assy.AssemblyTypeName
                        # All parts should be from same assembly
                        # but return whatever first one gives us
                        return assembly_name
            except Exception:
                pass
    return None


def check_parts_in_assembly(element_ids):
    """
    Check assembly status of selected parts.
    Returns (is_assembled, assembly_name).
    is_assembled = True if ANY part is in an assembly.
    """
    for eid in element_ids:
        el = doc.GetElement(eid)
        if isinstance(el, FabricationPart):
            try:
                assy_id = el.AssemblyInstanceId
                if assy_id and assy_id != ElementId.InvalidElementId:
                    assy = doc.GetElement(assy_id)
                    if isinstance(assy, AssemblyInstance):
                        return (True, assy.AssemblyTypeName)
            except Exception:
                pass
    return (False, None)


# ═══════════════════════════════════════════════════════════════════════════════
#  MODE A — FULL ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════

def run_full_assembly(do_hide):
    report_lines = []
    all_hidden_ids = []
    exported_set = set()

    folder = forms.pick_folder(title="Select Folder for PCF Export")
    if not folder:
        return None

    dlg = ProcessDialog(mode="full_assembly")
    try:
        while True:
            action = dlg.show_and_wait()

            if action == "close":
                break

            # Pick assemblies only using filter
            try:
                from Autodesk.Revit.UI.Selection import ObjectType
                assy_filter = AssemblySelectionFilter()
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element, assy_filter,
                    "Select Assembly Instances, then click Finish"
                    "  (only assemblies can be highlighted)")
                assembly_ids = []
                for ref in refs:
                    assembly_ids.append(ref.ElementId)
            except Exception as ex:
                # Revit throws OperationCanceledException when user presses Esc
                # Only break if user cancelled — otherwise continue the loop
                ex_type = type(ex).__name__
                if "Cancel" in ex_type or "cancel" in str(ex).lower():
                    break
                # For other exceptions, alert and let user try again
                forms.alert(
                    "Selection error: {}\n"
                    "Click Process to try again.".format(str(ex)))
                continue

            if not assembly_ids:
                forms.alert(
                    "No Assembly Instances selected."
                )
                continue

            for eid in assembly_ids:
                if eid.IntegerValue in exported_set:
                    continue
                exported_set.add(eid.IntegerValue)

                el = doc.GetElement(eid)
                if not isinstance(el, AssemblyInstance):
                    continue
                assembly = el
                member_ids = list(assembly.GetMemberIds())
                assembly_name = assembly.AssemblyTypeName
                filename = os.path.join(folder, assembly_name + ".pcf")

                FabricationUtils.ExportToPCF(doc, member_ids, filename)
                write_pcf_report_param(assembly, assembly_name)

                if not os.path.exists(filename):
                    report_lines.append(
                        "FAILED: PCF not created for {}".format(
                            assembly_name)
                    )
                    continue

                guid_map = build_guid_to_itemnumber(member_ids)
                items_ins, std_rep = postprocess_pcf(filename, guid_map)

                fab_count = sum(
                    1 for mid in member_ids
                    if isinstance(doc.GetElement(mid), FabricationPart)
                )
                report_lines.append("Assembly: {}".format(assembly_name))
                report_lines.append(
                    "  FabricationParts: {}".format(fab_count)
                )
                report_lines.append(
                    "  ITEM-NUMBER inserted: {}".format(items_ins)
                )
                report_lines.append(
                    "  STD -> SCH STD: {}".format(std_rep)
                )
                report_lines.append("  Output: {}".format(filename))
                report_lines.append("")

                if do_hide:
                    hide_ids = [eid]
                    hide_ids.extend(member_ids)
                    hide_exported_elements(hide_ids)
                    all_hidden_ids.extend(hide_ids)

            if not dlg.chk_multiple.Checked:
                break

    finally:
        dlg.Close()
        dlg.Dispose()
        if all_hidden_ids:
            restore_exported_elements(all_hidden_ids)

    return report_lines


# ═══════════════════════════════════════════════════════════════════════════════
#  MODE B — SELECTION SET
# ═══════════════════════════════════════════════════════════════════════════════

def run_selection_set(do_hide, naming_mode):
    report_lines = []
    all_hidden_ids = []
    total_ghosts_cleaned = 0

    folder = forms.pick_folder(title="Select Folder for PCF Export")
    if not folder:
        return None

    export_counter = 0

    dlg = ProcessDialog(mode="selection_set")
    try:
        while True:
            action = dlg.show_and_wait()

            if action == "close":
                break

            # Grab FabricationParts
            sel_ids = list(uidoc.Selection.GetElementIds())
            fab_ids = []
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, FabricationPart):
                    fab_ids.append(eid)

            if not fab_ids:
                forms.alert(
                    "No Fabrication Parts in selection.\n"
                    "Select parts, then click Process."
                )
                continue

            # Check if parts are in an assembly
            is_assembled, assembly_name = check_parts_in_assembly(fab_ids)
            pipeline_override = None

            if is_assembled:
                # Parts are in an assembly — use normal naming
                export_counter += 1

                if naming_mode == "auto":
                    file_base = "{}.{}".format(
                        assembly_name, export_counter
                    )
                    filename = os.path.join(folder, file_base + ".pcf")
                    pipeline_override = file_base
                else:
                    default = "{}_{}".format(
                        assembly_name, export_counter
                    )
                    name_dlg = NameDialog(default)
                    try:
                        name_dlg.ShowDialog()
                        if not name_dlg.file_name:
                            export_counter -= 1
                            continue
                        file_base = name_dlg.file_name
                        filename = os.path.join(
                            folder, file_base + ".pcf"
                        )
                        pipeline_override = file_base
                    finally:
                        name_dlg.Dispose()

            else:
                # Parts NOT in an assembly — mandatory name prompt
                mnd = MandatoryNameDialog()
                try:
                    mnd.ShowDialog()
                    if mnd.skipped or not mnd.file_name:
                        continue
                    file_base = mnd.file_name
                    pipeline_override = mnd.file_name
                    filename = os.path.join(folder, file_base + ".pcf")
                finally:
                    mnd.Dispose()

            # Snapshot assemblies before export
            assemblies_before = get_all_assembly_ids()

            # Export
            FabricationUtils.ExportToPCF(
                doc, List[ElementId](fab_ids), filename
            )

            # Detect and clean ghost assemblies
            assemblies_after = get_all_assembly_ids()
            ghosts = delete_ghost_assemblies(
                assemblies_before, assemblies_after
            )
            total_ghosts_cleaned += ghosts

            if not os.path.exists(filename):
                report_lines.append(
                    "FAILED: PCF not created for {}".format(file_base)
                )
                continue

            # Post-process (with pipeline override for unassembled parts)
            guid_map = build_guid_to_itemnumber(fab_ids)
            items_ins, std_rep = postprocess_pcf(
                filename, guid_map, pipeline_override
            )

            report_lines.append("Export: {}".format(file_base))
            report_lines.append("  Parts: {}".format(len(fab_ids)))
            if not is_assembled:
                report_lines.append("  Source: unassembled parts")
            else:
                report_lines.append(
                    "  Source: {}".format(assembly_name)
                )
            report_lines.append(
                "  ITEM-NUMBER inserted: {}".format(items_ins)
            )
            report_lines.append("  STD -> SCH STD: {}".format(std_rep))
            if pipeline_override:
                report_lines.append(
                    "  PIPELINE-REFERENCE / SPOOL-IDENTIFIER set to: {}".format(
                        pipeline_override)
                )
            if ghosts > 0:
                report_lines.append(
                    "  Ghost assemblies cleaned: {}".format(ghosts)
                )
            report_lines.append("  Output: {}".format(filename))
            report_lines.append("")

            # Hide processed parts
            if do_hide:
                hide_exported_elements(fab_ids)
                all_hidden_ids.extend(fab_ids)

            if not dlg.chk_multiple.Checked:
                break

    finally:
        dlg.Close()
        dlg.Dispose()
        if all_hidden_ids:
            restore_exported_elements(all_hidden_ids)

    # Summary of ghost cleanup
    if total_ghosts_cleaned > 0:
        report_lines.append(
            "Total ghost assemblies cleaned up: {}".format(
                total_ghosts_cleaned)
        )
        report_lines.append("")

    return report_lines


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

while True:
    settings = SettingsDialog()
    try:
        settings.ShowDialog()
        if settings.result == "cancel" or settings.result is None:
            script.exit()

        if settings.result == "filter":
            do_hide_grp = settings.chk_hide.Checked
            settings.Dispose()

            # Step 1: Source picker
            sp = SourcePickerDialog()
            try:
                sp.ShowDialog()
                if not sp.source:
                    continue  # back to settings
                source = sp.source
            finally:
                sp.Dispose()

            parts = collect_fab_parts(source)
            if not parts:
                if source == "pick":
                    msg = "No Fabrication Parts were picked."
                elif source == "selection":
                    msg = "No FabricationParts in the current selection."
                else:
                    msg = "No FabricationParts in the active view."
                forms.alert(msg)
                continue

            # Step 2: Group Export dialog (columns managed inside)
            all_params = get_all_param_names(parts)
            ged = GroupExportDialog(parts, all_params)
            try:
                ged.ShowDialog()
                if not ged.export_plan:
                    continue  # back to settings
                export_plan = ged.export_plan
            finally:
                ged.Dispose()

            # Step 3: Execute group export
            report = run_group_export(export_plan, do_hide_grp)
            script.exit()

        # result == "start" — normal export modes
        mode_full = settings.rb_full.Checked
        do_hide = settings.chk_hide.Checked
        naming_auto = settings.rb_auto.Checked
    finally:
        settings.Dispose()

    break  # got "start", proceed to normal export

if mode_full:
    report = run_full_assembly(do_hide)
else:
    naming_mode = "auto" if naming_auto else "custom"
    report = run_selection_set(do_hide, naming_mode)
