# -*- coding: utf-8 -*-
"""
Assembly Export to PCF
Replicates and extends SI_Tools AssemblyExportPCFCommandService.

Settings Dialog (modal):
  - Export Mode: Full Assembly / Selection Set
  - Hide elements after process (checkbox, either mode)
  - Naming: Auto-increment (.1, .2) / Custom name (Selection Set only)

Both modes: folder up front -> Process/Multiple/Close loop -> export each
Process immediately -> Close -> restore hidden -> report.

Full Assembly: selects AssemblyInstances, exports members, AssemblyTypeName
  as filename.

Selection Set: selects FabricationParts.
  - If parts are in an assembly: uses assembly name, naming settings apply
  - If parts are NOT in an assembly: mandatory name prompt (forced)
  - Ghost assembly cleanup: detects assemblies Revit auto-creates during
    ExportToPCF on loose parts, deletes them after export
  - PIPELINE-REFERENCE post-processing: replaces Revit's auto-generated
    pipeline reference with the user-provided name for unassembled parts

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
    FormStartPosition, FormBorderStyle
)
from System.Drawing import Point, Size
from System.Collections.Generic import List
from Autodesk.Revit.DB import Color as RevitColor

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
        self.ClientSize = Size(310, 290)
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

        self.btn_start = Button()
        self.btn_start.Text = "Start"
        self.btn_start.Location = Point(120, 250)
        self.btn_start.Size = Size(80, 30)
        self.btn_start.Click += self.on_start

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(210, 250)
        self.btn_cancel.Size = Size(80, 30)
        self.btn_cancel.Click += self.on_cancel

        self.Controls.Add(grp_mode)
        self.Controls.Add(self.chk_hide)
        self.Controls.Add(self.grp_naming)
        self.Controls.Add(self.btn_start)
        self.Controls.Add(self.btn_cancel)

        self.result = None

    def on_mode_changed(self, sender, args):
        self.grp_naming.Enabled = self.rb_selection.Checked

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
    def __init__(self):
        self.Text = "PCF Export"
        self.ClientSize = Size(255, 40)
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

        self.user_action = None
        self._is_open = False

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

    dlg = ProcessDialog()
    try:
        while True:
            action = dlg.show_and_wait()

            if action == "close":
                break

            sel_ids = list(uidoc.Selection.GetElementIds())
            assembly_ids = []
            non_assembly_count = 0

            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, AssemblyInstance):
                    assembly_ids.append(eid)
                else:
                    non_assembly_count += 1

            if not assembly_ids:
                forms.alert(
                    "No Assembly Instances in selection.\n"
                    "Only assemblies can be exported in Full Assembly mode."
                )
                continue

            if non_assembly_count > 0:
                forms.alert(
                    "{} non-assembly element(s) ignored.\n"
                    "{} assembly(ies) will be exported.".format(
                        non_assembly_count, len(assembly_ids)
                    )
                )

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

    dlg = ProcessDialog()
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
                    "  PIPELINE-REFERENCE set to: {}".format(
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

settings = SettingsDialog()
try:
    settings.ShowDialog()
    if settings.result != "start":
        script.exit()

    mode_full = settings.rb_full.Checked
    do_hide = settings.chk_hide.Checked
    naming_auto = settings.rb_auto.Checked
finally:
    settings.Dispose()

if mode_full:
    report = run_full_assembly(do_hide)
else:
    naming_mode = "auto" if naming_auto else "custom"
    report = run_selection_set(do_hide, naming_mode)

if report:
    output.print_md("## PCF Export Report")
    for line in report:
        print(line)
    forms.alert("PCF Export Complete. Check pyRevit output window.")
else:
    forms.alert("No exports were processed.")
