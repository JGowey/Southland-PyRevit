# -*- coding: utf-8 -*-
# pyRevit: Weld Number TextNote labeler (Plan + ISO/3D) with:
# - Readable UI (Revit TaskDialog command links)
# - Smart options (only show "Fill gaps" when gaps exist)
# - Duplicate detection + resolution workflow
# - Invalid text handling
#
# CHANGE REQUEST:
# - Removed the end "report window" / any popup summary after success.
# - Only popups that remain are decision dialogs and duplicate-warning (only if duplicates still exist).

from pyrevit import revit, DB, script
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult
import re

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
out = script.get_output()

# ----------------------------
# CONFIG
# ----------------------------
TARGET_TYPE_NAME = "Weld Number"   # TextNoteType name (Text Style)
LABEL_PREFIX = "W"                # W1, W2, ...
MATCH_TYPE_MODE = "equals"        # "equals" or "contains"

# Invalid text behavior:
# - "treat_as_blank": assign numbers to invalid text in continue/gaps modes
# - "skip": leave invalid text unchanged unless reset
# - "force_reset": if any invalid exists, only reset allowed
INVALID_TEXT_BEHAVIOR = "treat_as_blank"  # "treat_as_blank" | "skip" | "force_reset"
# ----------------------------


# ----------------------------
# Helpers
# ----------------------------
def bail(msg):
    # Keep errors visible in output pane
    try:
        out.print_md(u"❌ " + msg)
    except:
        pass
    script.exit()

def safe_str(x):
    try:
        return x if x is not None else ""
    except:
        return ""

def normalize_text(s):
    s = safe_str(s)
    s = s.replace(u"\u00A0", " ")
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def is_effectively_blank(s):
    return normalize_text(s) == ""

def get_view_right_up(v):
    try:
        return v.RightDirection, v.UpDirection
    except:
        return DB.XYZ.BasisX, DB.XYZ.BasisY

def dot(a, b):
    return a.X*b.X + a.Y*b.Y + a.Z*b.Z

def view_xy(v, p):
    right, up = get_view_right_up(v)
    return dot(p, right), dot(p, up)

def note_sort_key(tn):
    try:
        p = tn.Coord
    except:
        try:
            loc = tn.Location
            p = loc.Point if loc else DB.XYZ.Zero
        except:
            p = DB.XYZ.Zero
    x, y = view_xy(view, p)
    return (-y, x)

def get_type_elem(tn):
    try:
        return doc.GetElement(tn.GetTypeId())
    except:
        return None

def get_type_name(type_elem):
    if not type_elem:
        return ""
    try:
        p = type_elem.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return safe_str(p.AsString())
    except:
        pass
    try:
        return safe_str(getattr(type_elem, "Name", ""))
    except:
        return ""

def type_name_matches(type_name):
    t = normalize_text(type_name).lower()
    target = normalize_text(TARGET_TYPE_NAME).lower()
    if not t or not target:
        return False
    if MATCH_TYPE_MODE == "contains":
        return target in t
    return t == target

_w_re = re.compile(r"^\s*[Ww]\s*(\d+)\s*$")

def parse_w_number(text):
    m = _w_re.match(safe_str(text))
    if not m:
        return None
    try:
        n = int(m.group(1))
        return n if n > 0 else None
    except:
        return None

def format_w(n):
    return "{}{}".format(LABEL_PREFIX, n)

def taskdialog_choice(title, main_instruction, content, options):
    """
    options: list of tuples [(key, label, description), ...]
    returns: key or None
    """
    td = TaskDialog(title)
    td.MainInstruction = main_instruction
    td.MainContent = content
    td.CommonButtons = TaskDialogCommonButtons.Cancel

    link_ids = [
        TaskDialogCommandLinkId.CommandLink1,
        TaskDialogCommandLinkId.CommandLink2,
        TaskDialogCommandLinkId.CommandLink3,
        TaskDialogCommandLinkId.CommandLink4
    ]
    mapping = {}

    for i, opt in enumerate(options[:4]):
        key, label, desc = opt
        td.AddCommandLink(link_ids[i], label, desc)
        if i == 0:
            mapping[TaskDialogResult.CommandLink1] = key
        elif i == 1:
            mapping[TaskDialogResult.CommandLink2] = key
        elif i == 2:
            mapping[TaskDialogResult.CommandLink3] = key
        elif i == 3:
            mapping[TaskDialogResult.CommandLink4] = key

    res = td.Show()
    return mapping.get(res, None)

def set_selection(elem_ids):
    try:
        from System.Collections.Generic import List
        ids = List[DB.ElementId]()
        for eid in elem_ids:
            ids.Add(eid)
        uidoc.Selection.SetElementIds(ids)
    except:
        pass

def next_unused(start_from, used_set):
    n = start_from
    while n in used_set or n <= 0:
        n += 1
    return n


# ----------------------------
# Collect target TextNotes in active view
# ----------------------------
collector = (DB.FilteredElementCollector(doc, view.Id)
             .OfClass(DB.TextNote)
             .WhereElementIsNotElementType())

all_notes = list(collector)
if not all_notes:
    bail("No TextNotes found in the current view.")

targets = []
for tn in all_notes:
    try:
        te = get_type_elem(tn)
        if type_name_matches(get_type_name(te)):
            targets.append(tn)
    except:
        pass

if not targets:
    bail("No TextNotes of type '{}' found in this view.".format(TARGET_TYPE_NAME))

targets_sorted = sorted(targets, key=note_sort_key)

# ----------------------------
# Classify notes
# ----------------------------
blank_notes = []
numbered_notes = []     # (tn, num) valid W#
invalid_notes = []      # non-blank, not W#

for tn in targets_sorted:
    txt = safe_str(tn.Text)
    if is_effectively_blank(txt):
        blank_notes.append(tn)
        continue
    n = parse_w_number(txt)
    if n is not None:
        numbered_notes.append((tn, n))
    else:
        invalid_notes.append(tn)

existing_nums = [n for (_, n) in numbered_notes]
existing_set = set(existing_nums)
max_existing = max(existing_nums) if existing_nums else 0

# Duplicates
dupes = {}
for tn, n in numbered_notes:
    dupes.setdefault(n, []).append(tn)
duplicate_nums = sorted([n for n, lst in dupes.items() if len(lst) > 1])

# Gaps
gaps = []
if max_existing > 0:
    for k in range(1, max_existing + 1):
        if k not in existing_set:
            gaps.append(k)

# ----------------------------
# Enforce reset-only if configured
# ----------------------------
action = None
if INVALID_TEXT_BEHAVIOR == "force_reset" and invalid_notes:
    action = "reset"

# ----------------------------
# Duplicate workflow
# ----------------------------
dup_reassign_notes = []
keep_fixed_numbers = set()

if duplicate_nums and action is None:
    lines = []
    for n in duplicate_nums:
        lines.append("- {}{} occurs {} times".format(LABEL_PREFIX, n, len(dupes[n])))
    summary = "\n".join(lines)

    dup_choice = taskdialog_choice(
        title="Weld Number Tagging",
        main_instruction="Duplicate weld numbers found",
        content="Duplicates are not allowed in this view.\n\n{}".format(summary),
        options=[
            ("dup_reset", "Renumber EVERYTHING (recommended)",
             "Reassign all Weld Number notes in this view to W1..Wn. Fixes duplicates, gaps, and invalid text."),
            ("dup_autofix", "Auto-fix duplicates only",
             "Keep one of each duplicate number, reassign the extra duplicates to available numbers (fills gaps first)."),
            ("dup_select", "Select duplicate notes and stop",
             "Highlights the duplicate TextNotes so you can manually decide which one keeps the number.")
        ]
    )

    if dup_choice is None:
        script.exit()

    if dup_choice == "dup_select":
        ids = []
        for n in duplicate_nums:
            for tn in dupes[n]:
                ids.append(tn.Id)
        set_selection(ids)
        script.exit()

    if dup_choice == "dup_reset":
        action = "reset"
    else:
        # keep first, reassign extras
        for n in duplicate_nums:
            group = sorted(dupes[n], key=note_sort_key)
            keep_fixed_numbers.add(n)
            for extra in group[1:]:
                dup_reassign_notes.append(extra)

        # unique existing remain used
        for tn, n in numbered_notes:
            if n not in duplicate_nums:
                keep_fixed_numbers.add(n)

# ----------------------------
# Main action selection
# ----------------------------
if action is None:
    if (len(numbered_notes) == 0 and len(invalid_notes) == 0 and not duplicate_nums):
        action = "reset"
    else:
        invalid_as_blank = (INVALID_TEXT_BEHAVIOR == "treat_as_blank")
        has_to_assign = (len(blank_notes) > 0) or (len(dup_reassign_notes) > 0) or (invalid_as_blank and len(invalid_notes) > 0)

        options = []
        options.append(("reset", "Reset all weld tags (W1..Wn)",
                        "Renumber ALL Weld Number notes in this view sequentially."))

        if has_to_assign:
            options.append(("continue", "Continue from highest (fill blanks)",
                            "Leaves existing valid tags alone; assigns new tags starting at (max+1)."))

            if gaps:
                options.append(("gaps", "Fill gaps first, then continue",
                                "Uses missing numbers first, then continues above max."))

        # Include invalid note info inside the dialog text (no extra dialogs)
        inv_msg = ""
        if invalid_notes:
            if INVALID_TEXT_BEHAVIOR == "treat_as_blank":
                inv_msg = "\nInvalid text: {} (will be assigned)".format(len(invalid_notes))
            elif INVALID_TEXT_BEHAVIOR == "skip":
                inv_msg = "\nInvalid text: {} (will be left unchanged unless reset)".format(len(invalid_notes))
            else:
                inv_msg = "\nInvalid text: {} (force_reset)".format(len(invalid_notes))

        action = taskdialog_choice(
            title="Weld Number Tagging",
            main_instruction="Choose how to assign weld numbers",
            content=(
                "View: {}\n"
                "Target type: {}\n\n"
                "Target notes: {}\n"
                "Existing valid W#: {}\n"
                "Blank: {}\n"
                "Highest existing: {}\n"
                "Gaps: {}{}"
            ).format(
                view.Name, TARGET_TYPE_NAME,
                len(targets_sorted),
                len(numbered_notes),
                len(blank_notes),
                max_existing,
                len(gaps),
                inv_msg
            ),
            options=options
        )

        if action is None:
            script.exit()

# ----------------------------
# Build assignments
# ----------------------------
assignments = []

if action == "reset":
    for i, tn in enumerate(targets_sorted, start=1):
        assignments.append((tn, format_w(i)))
else:
    invalid_as_blank = (INVALID_TEXT_BEHAVIOR == "treat_as_blank")

    notes_to_assign = []
    notes_to_assign.extend(sorted(blank_notes, key=note_sort_key))
    notes_to_assign.extend(sorted(dup_reassign_notes, key=note_sort_key))
    if invalid_as_blank:
        notes_to_assign.extend(sorted(invalid_notes, key=note_sort_key))

    used = set(keep_fixed_numbers) if dup_reassign_notes else set(existing_set)

    if action == "continue":
        current_max = max(used) if used else 0
        n = current_max + 1
        for tn in notes_to_assign:
            n = next_unused(n, used)
            assignments.append((tn, format_w(n)))
            used.add(n)
            n += 1

    elif action == "gaps":
        current_max = max(used) if used else 0
        missing = []
        for k in range(1, current_max + 1):
            if k not in used:
                missing.append(k)

        idx = 0
        for tn in notes_to_assign:
            if idx < len(missing):
                n = missing[idx]
                assignments.append((tn, format_w(n)))
                used.add(n)
                idx += 1
            else:
                break

        remaining = notes_to_assign[len(assignments):]
        n = current_max + 1
        for tn in remaining:
            n = next_unused(n, used)
            assignments.append((tn, format_w(n)))
            used.add(n)
            n += 1

if not assignments:
    script.exit()

# ----------------------------
# Apply transaction
# ----------------------------
t = DB.Transaction(doc, "Weld Number Tagging")
t.Start()
try:
    for tn, newtxt in assignments:
        tn.Text = newtxt
    t.Commit()
except Exception:
    if t.HasStarted():
        t.RollBack()
    raise

# ----------------------------
# Post-check duplicates (only warn if still present)
# ----------------------------
post_nums = {}
for tn in targets_sorted:
    n = parse_w_number(tn.Text)
    if n is None:
        continue
    post_nums.setdefault(n, []).append(tn)

post_dupes = sorted([n for n, lst in post_nums.items() if len(lst) > 1])
if post_dupes:
    ids = []
    for n in post_dupes:
        for tn in post_nums[n]:
            ids.append(tn.Id)
    set_selection(ids)
    TaskDialog.Show(
        "Weld Number Tagging",
        "⚠ Duplicate weld numbers still exist.\n\n"
        "I selected the duplicate notes in Revit.\n"
        "Duplicates: {}"
        .format(", ".join([format_w(n) for n in post_dupes]))
    )

# Done. No success popups, no summary window.
