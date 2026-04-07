# -*- coding: utf-8 -*-
"""
Family Parameter Mapper for pyRevit / Revit Family Editor

Purpose
-------
Copy text values between family parameters or between a parameter and Type Name.

Design goals
------------
- Clean, compact UI with minimal dead space
- Safe pipeline-style execution
- Well-commented and easy to maintain
- Supports long parameter names better than prior versions
- Includes a help button with a workflow guide

Supported targets
-----------------
- Type Name
- Writable text parameters only

Run context
-----------
This script must be run inside the Revit Family Editor on an RFA document.
"""

import clr
import traceback

clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from System.Drawing import Point, Size
from System.Windows.Forms import (
    Form, Label, ListBox, Button, CheckBox, TextBox, GroupBox,
    DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon,
    FormStartPosition, FormBorderStyle, SelectionMode
)
from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog


# ============================================================
# REVIT CONTEXT
# ============================================================
uidoc = __revit__.ActiveUIDocument
if uidoc is None:
    raise Exception("No active Revit document.")

doc = uidoc.Document
if not doc.IsFamilyDocument:
    raise Exception("Run this script in the Family Editor (RFA), not in a project (RVT).")

fam_mgr = doc.FamilyManager


# ============================================================
# CONSTANTS
# ============================================================
TYPE_NAME_KEY = "__TYPE_NAME__"
WINDOW_WIDTH = 980
WINDOW_HEIGHT = 710


# ============================================================
# BASIC HELPERS
# ============================================================
def safe_str(value):
    """Return a guaranteed string."""
    if value is None:
        return ""
    return str(value)


def is_blank(value):
    """True if value is None or whitespace-only."""
    return safe_str(value).strip() == ""


def build_help_text():
    """
    Return the workflow/help text shown in the separate pop-up.
    """
    lines = []
    lines.append("Family Parameter Mapper - Help")
    lines.append("")
    lines.append("Workflow")
    lines.append("1. Pick the source on the left.")
    lines.append("2. Pick the target on the right.")
    lines.append("3. Set any options you want.")
    lines.append("4. Click Run.")
    lines.append("")
    lines.append("What the controls do")
    lines.append("- Show instance parameters: includes instance parameters in both lists.")
    lines.append("- Source filter: narrows the left list.")
    lines.append("- Target filter: narrows the right list.")
    lines.append("- Swap: swaps the current source and target selections.")
    lines.append("- Clear Filters: clears both filter boxes.")
    lines.append("- Only fill blank target values: skips targets that already have text.")
    lines.append("- Skip blank source values: skips empty source values.")
    lines.append("- Trim spaces: removes leading and trailing spaces before writing.")
    lines.append("- Use Find / Replace: applies text replacement before writing.")
    lines.append("")
    lines.append("How Find / Replace works")
    lines.append("- It runs after Trim spaces.")
    lines.append("- It is a simple literal text replace.")
    lines.append("- It replaces all matches.")
    lines.append("- It is case-sensitive.")
    lines.append("")
    lines.append("Type Name notes")
    lines.append("- You can push values into Type Name.")
    lines.append("- Blank names are skipped.")
    lines.append("- Duplicate type names are skipped.")
    lines.append("")
    lines.append("Supported targets")
    lines.append("- Type Name")
    lines.append("- Writable text parameters")
    lines.append("")
    lines.append("Tip")
    lines.append("For most cleanup work, leave Skip blank source values and Trim spaces on.")
    return "\n".join(lines)


# ============================================================
# PARAMETER DISCOVERY
# ============================================================
def build_mapping_items(show_instance_parameters):
    """
    Build the selectable source/target items shown in the UI.

    Returns a list of dictionaries with consistent fields so the rest
    of the pipeline can stay simple and predictable.
    """
    items = [{
        "label": "Type Name",
        "kind": "type_name",
        "name": TYPE_NAME_KEY,
        "param": None,
        "is_instance": False
    }]

    for fam_param in fam_mgr.Parameters:
        try:
            param_name = fam_param.Definition.Name
        except:
            continue

        try:
            is_instance = fam_param.IsInstance
        except:
            is_instance = False

        if (not show_instance_parameters) and is_instance:
            continue

        label = "{} [{}]".format(param_name, "Instance" if is_instance else "Type")

        items.append({
            "label": label,
            "kind": "family_param",
            "name": param_name,
            "param": fam_param,
            "is_instance": is_instance
        })

    # Keep Type Name at the top and sort everything else alphabetically.
    type_name_items = [x for x in items if x["kind"] == "type_name"]
    parameter_items = [x for x in items if x["kind"] != "type_name"]
    parameter_items.sort(key=lambda x: x["label"].lower())

    return type_name_items + parameter_items


def filter_items(items, filter_text):
    """Filter items by text typed into the search boxes."""
    needle = safe_str(filter_text).strip().lower()
    if needle == "":
        return list(items)

    filtered = []
    for item in items:
        if needle in item["label"].lower():
            filtered.append(item)
    return filtered


def find_item_by_label(items, label):
    """Resolve the selected UI label back to its backing item."""
    for item in items:
        if item["label"] == label:
            return item
    return None


# ============================================================
# VALUE READ / WRITE
# ============================================================
def can_write_to_target(item):
    """
    Allowed targets:
    - Type Name
    - Writable text parameters
    """
    if item["kind"] == "type_name":
        return True

    fam_param = item["param"]
    if fam_param is None:
        return False

    try:
        if fam_param.Formula:
            return False
    except:
        pass

    try:
        return fam_param.StorageType == DB.StorageType.String
    except:
        return False


def read_item_value(item, family_type):
    """
    Read a source value for the currently active family type.

    Always returns either a string or None.
    """
    if item["kind"] == "type_name":
        try:
            return family_type.Name
        except:
            return None

    fam_param = item["param"]
    if fam_param is None:
        return None

    try:
        value = fam_mgr.CurrentType.AsString(fam_param)
        if value is not None:
            return value
    except:
        pass

    try:
        value = fam_mgr.CurrentType.AsValueString(fam_param)
        if value is not None:
            return value
    except:
        pass

    return None


def get_existing_type_names():
    """Collect current family type names for rename collision checks."""
    names = []
    for family_type in fam_mgr.Types:
        try:
            names.append(family_type.Name)
        except:
            pass
    return names


def can_rename_type(old_name, new_name):
    """
    Validate a target family type name before rename.
    """
    if is_blank(new_name):
        return False, "New type name is blank."

    for existing_name in get_existing_type_names():
        if existing_name == new_name and existing_name != old_name:
            return False, "Type name already exists."

    return True, None


def transform_value(raw_value, trim_spaces, use_find_replace, find_text, replace_text):
    """
    Apply the value transformation pipeline in one place.

    Pipeline order:
    1. Convert to safe string
    2. Trim leading/trailing spaces if requested
    3. Run literal find/replace if requested
    """
    transformed = safe_str(raw_value)

    if trim_spaces:
        transformed = transformed.strip()

    if use_find_replace and find_text != "":
        transformed = transformed.replace(find_text, replace_text)

    return transformed


def write_item_value(target_item, current_type_name, value_to_write):
    """
    Write the prepared value to the selected target.

    Returns:
        (success_bool, reason_string_or_none)
    """
    if target_item["kind"] == "type_name":
        new_name = safe_str(value_to_write)
        can_rename, reason = can_rename_type(current_type_name, new_name)
        if not can_rename:
            return False, reason

        try:
            fam_mgr.RenameCurrentType(new_name)
            return True, None
        except Exception as ex:
            return False, str(ex)

    fam_param = target_item["param"]
    if fam_param is None:
        return False, "Target parameter missing."

    if not can_write_to_target(target_item):
        return False, "Target must be Type Name or writable text parameter."

    try:
        fam_mgr.Set(fam_param, safe_str(value_to_write))
        return True, None
    except Exception as ex:
        return False, str(ex)


# ============================================================
# SUMMARY TEXT
# ============================================================
# Confirmation popup intentionally removed.
# The tool now runs immediately after validation when the user clicks Run.

def build_result_text(source_item, target_item, settings, updated_count, skipped_count, error_count):
    """
    Final compact summary. No line-by-line preview spam.
    """
    lines = []
    lines.append("Finished.")
    lines.append("")
    lines.append("Source: {}".format(source_item["label"]))
    lines.append("Target: {}".format(target_item["label"]))
    lines.append("")
    lines.append("Only fill blank targets: {}".format("Yes" if settings["blank_only"] else "No"))
    lines.append("Skip blank source values: {}".format("Yes" if settings["skip_blank_source"] else "No"))
    lines.append("Trim spaces: {}".format("Yes" if settings["trim_spaces"] else "No"))
    lines.append("Find / Replace: {}".format("Yes" if settings["do_find_replace"] else "No"))

    if settings["do_find_replace"]:
        lines.append("Find: {}".format(settings["find_text"]))
        lines.append("Replace: {}".format(settings["replace_text"]))

    lines.append("")
    lines.append("Updated: {}".format(updated_count))
    lines.append("Skipped: {}".format(skipped_count))
    lines.append("Errors: {}".format(error_count))
    return "\n".join(lines)


# ============================================================
# UI
# ============================================================
class ParamMapForm(Form):
    """
    Compact side-by-side selector UI.

    Key layout choices:
    - Wide source/target columns for long parameter names
    - Narrow center column for only the most-used actions
    - One help button opens a separate workflow pop-up
    - Extra bottom spacing so Run/Cancel are not clipped
    """
    def __init__(self):
        Form.__init__(self)

        self.all_items = []
        self.filtered_source_items = []
        self.filtered_target_items = []
        self.result = None

        self._build_window()
        self._build_header_help()
        self._build_source_target_group()
        self._build_options_group()
        self._build_action_buttons()
        self.reload_item_lists()

    # ----------------------------
    # UI construction
    # ----------------------------
    def _build_window(self):
        self.Text = "Family Parameter Mapper"
        self.Width = WINDOW_WIDTH
        self.Height = WINDOW_HEIGHT
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimizeBox = False
        self.MaximizeBox = False
        self.FormBorderStyle = FormBorderStyle.FixedSingle

    def _build_header_help(self):
        self.btn_help = Button()
        self.btn_help.Text = "?"
        self.btn_help.Location = Point(918, 10)
        self.btn_help.Size = Size(34, 28)
        self.btn_help.Click += self.on_help
        self.Controls.Add(self.btn_help)

    def _build_source_target_group(self):
        self.grp_map = GroupBox()
        self.grp_map.Text = "Source and Target"
        self.grp_map.Location = Point(12, 38)
        self.grp_map.Size = Size(940, 430)
        self.Controls.Add(self.grp_map)

        self.chk_show_instance = CheckBox()
        self.chk_show_instance.Text = "Show instance parameters"
        self.chk_show_instance.Location = Point(18, 28)
        self.chk_show_instance.Size = Size(210, 24)
        self.chk_show_instance.Checked = True
        self.chk_show_instance.CheckedChanged += self.on_filter_inputs_changed
        self.grp_map.Controls.Add(self.chk_show_instance)

        # Left side
        self.lbl_source = Label()
        self.lbl_source.Text = "Copy FROM"
        self.lbl_source.Location = Point(18, 62)
        self.lbl_source.Size = Size(120, 20)
        self.grp_map.Controls.Add(self.lbl_source)

        self.txt_source_filter = TextBox()
        self.txt_source_filter.Location = Point(18, 86)
        self.txt_source_filter.Size = Size(390, 26)
        self.txt_source_filter.TextChanged += self.on_filter_inputs_changed
        self.grp_map.Controls.Add(self.txt_source_filter)

        self.lst_source = ListBox()
        self.lst_source.Location = Point(18, 118)
        self.lst_source.Size = Size(390, 250)
        self.lst_source.SelectionMode = SelectionMode.One
        self.lst_source.HorizontalScrollbar = True
        self.lst_source.SelectedIndexChanged += self.on_source_selected
        self.grp_map.Controls.Add(self.lst_source)

        self.lbl_source_selected = Label()
        self.lbl_source_selected.Text = "Selected: "
        self.lbl_source_selected.Location = Point(18, 377)
        self.lbl_source_selected.Size = Size(390, 24)
        self.grp_map.Controls.Add(self.lbl_source_selected)

        # Center controls
        self.btn_swap = Button()
        self.btn_swap.Text = "Swap"
        self.btn_swap.Location = Point(430, 165)
        self.btn_swap.Size = Size(80, 32)
        self.btn_swap.Click += self.on_swap
        self.grp_map.Controls.Add(self.btn_swap)

        self.btn_clear_filters = Button()
        self.btn_clear_filters.Text = "Clear Filters"
        self.btn_clear_filters.Location = Point(422, 210)
        self.btn_clear_filters.Size = Size(96, 32)
        self.btn_clear_filters.Click += self.on_clear_filters
        self.grp_map.Controls.Add(self.btn_clear_filters)

        # Right side
        self.lbl_target = Label()
        self.lbl_target.Text = "Paste TO"
        self.lbl_target.Location = Point(532, 62)
        self.lbl_target.Size = Size(120, 20)
        self.grp_map.Controls.Add(self.lbl_target)

        self.txt_target_filter = TextBox()
        self.txt_target_filter.Location = Point(532, 86)
        self.txt_target_filter.Size = Size(390, 26)
        self.txt_target_filter.TextChanged += self.on_filter_inputs_changed
        self.grp_map.Controls.Add(self.txt_target_filter)

        self.lst_target = ListBox()
        self.lst_target.Location = Point(532, 118)
        self.lst_target.Size = Size(390, 250)
        self.lst_target.SelectionMode = SelectionMode.One
        self.lst_target.HorizontalScrollbar = True
        self.lst_target.SelectedIndexChanged += self.on_target_selected
        self.grp_map.Controls.Add(self.lst_target)

        self.lbl_target_selected = Label()
        self.lbl_target_selected.Text = "Selected: "
        self.lbl_target_selected.Location = Point(532, 377)
        self.lbl_target_selected.Size = Size(390, 24)
        self.grp_map.Controls.Add(self.lbl_target_selected)

    def _build_options_group(self):
        self.grp_options = GroupBox()
        self.grp_options.Text = "Options"
        self.grp_options.Location = Point(12, 474)
        self.grp_options.Size = Size(940, 120)
        self.Controls.Add(self.grp_options)

        self.chk_blank_only = CheckBox()
        self.chk_blank_only.Text = "Only fill blank target values"
        self.chk_blank_only.Location = Point(18, 28)
        self.chk_blank_only.Size = Size(220, 24)
        self.grp_options.Controls.Add(self.chk_blank_only)

        self.chk_skip_blank_source = CheckBox()
        self.chk_skip_blank_source.Text = "Skip blank source values"
        self.chk_skip_blank_source.Location = Point(260, 28)
        self.chk_skip_blank_source.Size = Size(190, 24)
        self.chk_skip_blank_source.Checked = True
        self.grp_options.Controls.Add(self.chk_skip_blank_source)

        self.chk_trim_spaces = CheckBox()
        self.chk_trim_spaces.Text = "Trim spaces"
        self.chk_trim_spaces.Location = Point(472, 28)
        self.chk_trim_spaces.Size = Size(120, 24)
        self.chk_trim_spaces.Checked = True
        self.grp_options.Controls.Add(self.chk_trim_spaces)

        self.chk_find_replace = CheckBox()
        self.chk_find_replace.Text = "Use Find / Replace"
        self.chk_find_replace.Location = Point(620, 28)
        self.chk_find_replace.Size = Size(155, 24)
        self.chk_find_replace.CheckedChanged += self.on_find_replace_changed
        self.grp_options.Controls.Add(self.chk_find_replace)

        self.lbl_find = Label()
        self.lbl_find.Text = "Find:"
        self.lbl_find.Location = Point(18, 66)
        self.lbl_find.Size = Size(42, 20)
        self.grp_options.Controls.Add(self.lbl_find)

        self.txt_find = TextBox()
        self.txt_find.Location = Point(62, 62)
        self.txt_find.Size = Size(360, 26)
        self.txt_find.Enabled = False
        self.grp_options.Controls.Add(self.txt_find)

        self.lbl_replace = Label()
        self.lbl_replace.Text = "Replace:"
        self.lbl_replace.Location = Point(445, 66)
        self.lbl_replace.Size = Size(58, 20)
        self.grp_options.Controls.Add(self.lbl_replace)

        self.txt_replace = TextBox()
        self.txt_replace.Location = Point(508, 62)
        self.txt_replace.Size = Size(396, 26)
        self.txt_replace.Enabled = False
        self.grp_options.Controls.Add(self.txt_replace)

    def _build_action_buttons(self):
        self.btn_run = Button()
        self.btn_run.Text = "Run"
        self.btn_run.Location = Point(786, 606)
        self.btn_run.Size = Size(80, 32)
        self.btn_run.Click += self.on_run
        self.Controls.Add(self.btn_run)

        self.btn_cancel = Button()
        self.btn_cancel.Text = "Cancel"
        self.btn_cancel.Location = Point(872, 606)
        self.btn_cancel.Size = Size(80, 32)
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)

    # ----------------------------
    # UI event handlers
    # ----------------------------
    def on_help(self, sender, args):
        MessageBox.Show(
            build_help_text(),
            "Family Parameter Mapper Help",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )

    def on_filter_inputs_changed(self, sender, args):
        self.reload_item_lists()

    def on_find_replace_changed(self, sender, args):
        enabled = self.chk_find_replace.Checked
        self.txt_find.Enabled = enabled
        self.txt_replace.Enabled = enabled

    def on_source_selected(self, sender, args):
        selected = self.get_selected_label(self.lst_source)
        self.lbl_source_selected.Text = "Selected: {}".format(selected if selected else "")

    def on_target_selected(self, sender, args):
        selected = self.get_selected_label(self.lst_target)
        self.lbl_target_selected.Text = "Selected: {}".format(selected if selected else "")

    def on_swap(self, sender, args):
        source_label = self.get_selected_label(self.lst_source)
        target_label = self.get_selected_label(self.lst_target)

        if not source_label or not target_label:
            return

        self.select_listbox_value(self.lst_source, target_label)
        self.select_listbox_value(self.lst_target, source_label)

    def on_clear_filters(self, sender, args):
        self.txt_source_filter.Text = ""
        self.txt_target_filter.Text = ""

    def on_run(self, sender, args):
        source_label = self.get_selected_label(self.lst_source)
        target_label = self.get_selected_label(self.lst_target)

        if not source_label or not target_label:
            MessageBox.Show("Select both source and target.", "Missing Selection")
            return

        if source_label == target_label:
            MessageBox.Show("Source and target cannot be the same.", "Invalid Selection")
            return

        self.result = {
            "source_label": source_label,
            "target_label": target_label,
            "blank_only": self.chk_blank_only.Checked,
            "skip_blank_source": self.chk_skip_blank_source.Checked,
            "trim_spaces": self.chk_trim_spaces.Checked,
            "do_find_replace": self.chk_find_replace.Checked,
            "find_text": safe_str(self.txt_find.Text),
            "replace_text": safe_str(self.txt_replace.Text)
        }

        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    # ----------------------------
    # UI data population helpers
    # ----------------------------
    def reload_item_lists(self):
        """
        Refresh available source/target lists while preserving current selections.
        """
        old_source_label = self.get_selected_label(self.lst_source)
        old_target_label = self.get_selected_label(self.lst_target)

        self.all_items = build_mapping_items(self.chk_show_instance.Checked)
        self.filtered_source_items = filter_items(self.all_items, self.txt_source_filter.Text)
        self.filtered_target_items = filter_items(self.all_items, self.txt_target_filter.Text)

        self.populate_listbox(self.lst_source, self.filtered_source_items)
        self.populate_listbox(self.lst_target, self.filtered_target_items)

        self.restore_listbox_selection(self.lst_source, old_source_label, 0)
        self.restore_listbox_selection(self.lst_target, old_target_label, 1)

        self.on_source_selected(None, None)
        self.on_target_selected(None, None)

    def populate_listbox(self, listbox, items):
        """Load labels into a list box."""
        listbox.Items.Clear()
        for item in items:
            listbox.Items.Add(item["label"])

    def restore_listbox_selection(self, listbox, old_label, fallback_index):
        """Restore prior selection if possible, otherwise use a safe default."""
        if listbox.Items.Count == 0:
            return

        if old_label:
            for i in range(listbox.Items.Count):
                if safe_str(listbox.Items[i]) == old_label:
                    listbox.SelectedIndex = i
                    return

        if fallback_index < listbox.Items.Count:
            listbox.SelectedIndex = fallback_index
        else:
            listbox.SelectedIndex = 0

    def get_selected_label(self, listbox):
        """Read selected text from a list box safely."""
        try:
            if listbox.SelectedItem is None:
                return None
            return safe_str(listbox.SelectedItem)
        except:
            return None

    def select_listbox_value(self, listbox, label):
        """Select a given label inside a list box if it exists."""
        for i in range(listbox.Items.Count):
            if safe_str(listbox.Items[i]) == label:
                listbox.SelectedIndex = i
                return


# ============================================================
# EXECUTION PIPELINE
# ============================================================
def collect_user_settings():
    """
    Stage 1: Show the UI and collect user intent.
    """
    form = ParamMapForm()
    result = form.ShowDialog()

    if result != DialogResult.OK or form.result is None:
        return None, None, None

    settings = form.result
    source_item = find_item_by_label(form.all_items, settings["source_label"])
    target_item = find_item_by_label(form.all_items, settings["target_label"])

    return settings, source_item, target_item


def validate_selection(source_item, target_item):
    """
    Stage 2: Validate source/target before any model changes happen.
    """
    if source_item is None or target_item is None:
        TaskDialog.Show("Family Parameter Mapper", "Could not resolve selected parameters.")
        return False

    if not can_write_to_target(target_item):
        TaskDialog.Show(
            "Family Parameter Mapper",
            "Target must be either:\n"
            "- Type Name\n"
            "- A writable text parameter\n\n"
            "Selected target:\n{}".format(target_item["label"])
        )
        return False

    return True


# Stage 3 confirmation intentionally removed.
# The pipeline now proceeds directly from validation to execution.


def execute_mapping(settings, source_item, target_item):
    """
    Stage 4: Execute the copy pipeline across all family types inside one transaction.

    Returns a tuple:
        (updated_count, skipped_count, error_count)
    """
    updated_count = 0
    skipped_count = 0
    error_count = 0

    transaction = DB.Transaction(doc, "Family Parameter Mapper")
    try:
        transaction.Start()

        for family_type in fam_mgr.Types:
            try:
                # Make this family type active before reading/writing.
                fam_mgr.CurrentType = family_type

                current_type_name = safe_str(family_type.Name)
                raw_source_value = read_item_value(source_item, family_type)
                current_target_value = read_item_value(target_item, family_type)

                # Skip empty sources when the user asks for that.
                if settings["skip_blank_source"] and is_blank(raw_source_value):
                    skipped_count += 1
                    continue

                # Apply the text transformation pipeline once in one place.
                prepared_value = transform_value(
                    raw_source_value,
                    settings["trim_spaces"],
                    settings["do_find_replace"],
                    settings["find_text"],
                    settings["replace_text"]
                )

                # Skip non-empty targets when blank-only mode is enabled.
                if settings["blank_only"] and not is_blank(current_target_value):
                    skipped_count += 1
                    continue

                success, reason = write_item_value(target_item, current_type_name, prepared_value)

                if success:
                    updated_count += 1
                else:
                    skipped_count += 1

            except Exception:
                error_count += 1

        transaction.Commit()

    except Exception:
        try:
            transaction.RollBack()
        except:
            pass
        raise

    return updated_count, skipped_count, error_count


def show_results(source_item, target_item, settings, updated_count, skipped_count, error_count):
    """
    Stage 5: Show a compact result summary.
    """
    TaskDialog.Show(
        "Family Parameter Mapper",
        build_result_text(
            source_item,
            target_item,
            settings,
            updated_count,
            skipped_count,
            error_count
        )
    )


def run():
    """
    Main pipeline coordinator.
    """
    try:
        settings, source_item, target_item = collect_user_settings()
        if settings is None:
            return

        if not validate_selection(source_item, target_item):
            return

        updated_count, skipped_count, error_count = execute_mapping(
            settings,
            source_item,
            target_item
        )

        show_results(
            source_item,
            target_item,
            settings,
            updated_count,
            skipped_count,
            error_count
        )

    except Exception:
        TaskDialog.Show("Family Parameter Mapper - Error", traceback.format_exc())


run()
