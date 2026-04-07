# -*- coding: utf-8 -*-
"""
Place Weld Notes — Click-to-Place with Auto-Numbering
======================================================
pyRevit button for placing numbered weld callout tags (TextNotes with leaders)
by clicking on pipe/weld points.

Three workflows:
    1. PLACE: Click points to place new W# notes (NO SHIFT = left, SHIFT = right)
    2. AUTO-FILL: Assign W# to existing blank/invalid notes, then place new ones
    3. RENUMBER: Reassign ALL notes in view to W1..Wn (spatial sort), then exit

Startup behavior:
    - No existing weld notes in view -> straight to click-to-place, no dialogs
    - Existing notes found -> show dialog with options (continue / autofill / renumber)
    - Duplicates detected -> offer to select them, renumber, or continue anyway

Scale handling:
    All placement offsets and leader geometry are tuned at BASE_TUNED_SCALE (1/2" = 1'-0")
    and scale proportionally with the active view's scale factor.

Revit 2022+ compatible.
"""

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, SubTransaction, XYZ, Plane, SketchPlane,
    TextNote, TextNoteLeaderTypes, ViewType,
    TextNoteType, FilteredElementCollector, BuiltInParameter,
    TransactionStatus,
    IFailuresPreprocessor, FailureProcessingResult
)
from Autodesk.Revit.UI import (
    TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows.Input import Keyboard, Key
import re

# Worksharing imports (optional — not present in non-workshared environments)
try:
    from Autodesk.Revit.DB import WorksharingUtils, CheckoutStatus
except:
    WorksharingUtils = None
    CheckoutStatus = None

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView


# =============================================================================
# CONFIG
# =============================================================================
# Target TextNoteType to match. "equals" matches exact name; "contains" is substring.
TARGET_TYPE_NAME = "Weld Number"
MATCH_TYPE_MODE = "equals"          # "equals" | "contains"
LABEL_PREFIX = "W"

# Base view scale the offsets were tuned at: 1/2" = 1'-0" => Revit scale 24
BASE_TUNED_SCALE = 24.0

# Text placement offsets (feet, tuned at BASE_TUNED_SCALE)
OFFSET_SIDE_FEET = 1.5             # Horizontal distance from click point to text
OFFSET_UP_FEET   = 1.0             # Vertical distance from click point to text
LEFT_EXTRA_MULT  = 1.15            # Extra multiplier when text goes LEFT

# Leader elbow exit length (feet, tuned at BASE_TUNED_SCALE)
ELBOW_OUT_LEFT_FEET = 0.55         # Base exit length when text is LEFT
RIGHT_EXIT_MULT     = 0.3          # When text is RIGHT, exit = left * this

# Per-side base adjustment to exit length depending on arrow direction
BASE_ADJUST_EXIT_RIGHT_FEET = -0.04
BASE_ADJUST_EXIT_LEFT_FEET  =  0.00

# Extra exit length per digit when elbow exits RIGHT (compensates for wider text)
CHAR_ADD_PER_EXTRA_DIGIT_FEET = 0.09
CHAR_ADD_MAX_FEET             = 0.20

# 3D view must be locked before placing notes
REQUIRE_LOCKED_3D = True

# How to handle non-W# text in existing notes during autofill:
#   "treat_as_blank" = overwrite with a number
#   "skip" = leave as-is (only blank notes get filled)
INVALID_TEXT_BEHAVIOR = "treat_as_blank"  # "treat_as_blank" | "skip"


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def safe_str(x):
    """Returns x as string, or empty string if None/error."""
    try:
        return x if x is not None else ""
    except:
        return ""

def normalize_text(s):
    """Strips whitespace variants (NBSP, newlines, tabs) and collapses runs."""
    s = safe_str(s).replace(u"\u00A0", " ")
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def is_effectively_blank(s):
    """Returns True if string is empty after normalization."""
    return normalize_text(s) == ""

def get_type_name(type_elem):
    """Gets the name of a Revit element type, preferring SYMBOL_NAME_PARAM."""
    if not type_elem:
        return ""
    try:
        p = type_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            return safe_str(p.AsString())
    except:
        pass
    try:
        return safe_str(getattr(type_elem, "Name", ""))
    except:
        return ""

def type_name_matches(type_name):
    """Returns True if type_name matches TARGET_TYPE_NAME per MATCH_TYPE_MODE."""
    t = normalize_text(type_name).lower()
    target = normalize_text(TARGET_TYPE_NAME).lower()
    if not t or not target:
        return False
    if MATCH_TYPE_MODE == "contains":
        return target in t
    return t == target

# Pre-compiled regex for "W123" style labels
_w_re = re.compile(r"^\s*[Ww]\s*(\d+)\s*$")

def parse_w_number(text):
    """Parses 'W123' text into int 123. Returns None if not a valid W# label."""
    m = _w_re.match(safe_str(text))
    if not m:
        return None
    try:
        n = int(m.group(1))
        return n if n > 0 else None
    except:
        return None

def number_digits(n):
    """Returns the number of digits in integer n."""
    try:
        return len(str(int(n)))
    except:
        return 1

def format_w(n):
    """Formats n as a weld label string, e.g. 'W5'."""
    return "{}{}".format(LABEL_PREFIX, n)


# --- View geometry helpers ---

def get_view_right_up(v):
    """Returns the view's (right, up) direction vectors."""
    try:
        return v.RightDirection, v.UpDirection
    except:
        return XYZ.BasisX, XYZ.BasisY

def dot(a, b):
    """Dot product of two XYZ vectors."""
    return a.X*b.X + a.Y*b.Y + a.Z*b.Z

def is_shift_down():
    """Returns True if either Shift key is currently held."""
    return Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)

def set_view_sketchplane_facing_camera(pt, view_dir_vec):
    """Creates a SketchPlane at pt facing the camera (required for 3D text placement)."""
    geom_plane = Plane.CreateByNormalAndOrigin(view_dir_vec, pt)
    sketch = SketchPlane.Create(doc, geom_plane)
    try:
        view.SketchPlane = sketch
    except:
        pass

def get_scale_factor(v):
    """Returns view.Scale / BASE_TUNED_SCALE so all offsets scale with the view."""
    try:
        s = float(v.Scale)
        b = float(BASE_TUNED_SCALE) if BASE_TUNED_SCALE else s
        if b == 0:
            return 1.0
        return s / b
    except:
        return 1.0


# --- Leader geometry helpers ---

def leader_type_for_text_side(text_side_sign):
    """
    Returns the leader attachment type for the given text side.
    Text RIGHT (+1) => leader attaches on LEFT; text LEFT (-1) => leader attaches on RIGHT.
    """
    return TextNoteLeaderTypes.TNLT_STRAIGHT_L if text_side_sign > 0 else TextNoteLeaderTypes.TNLT_STRAIGHT_R

def get_exit_len_for_side(text_side_sign):
    """Returns the base elbow exit length (feet) for the given text side."""
    if text_side_sign > 0:
        return ELBOW_OUT_LEFT_FEET * RIGHT_EXIT_MULT
    return ELBOW_OUT_LEFT_FEET

def clamp(x, a, b):
    """Clamps x to the range [a, b]."""
    return max(a, min(b, x))


# --- UI helpers ---

def taskdialog_choice(title, main_instruction, content, options):
    """
    Shows a TaskDialog with up to 4 command link options.

    Args:
        options: list of (key, label, description) tuples
    Returns:
        The key string of the selected option, or None if cancelled.
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
    """Selects the given element IDs in the Revit UI."""
    try:
        from System.Collections.Generic import List
        from Autodesk.Revit.DB import ElementId
        ids = List[ElementId]()
        for eid in elem_ids:
            ids.Add(eid)
        uidoc.Selection.SetElementIds(ids)
    except:
        pass


# --- Numbering helpers ---

def next_unused(start_from, used_set):
    """Returns the first positive integer >= start_from not in used_set."""
    n = start_from
    while n in used_set or n <= 0:
        n += 1
    return n

def build_gaps(used_set):
    """Returns sorted list of missing integers in 1..max(used_set)."""
    if not used_set:
        return []
    m = max(used_set)
    return [k for k in range(1, m + 1) if k not in used_set]

def sort_key_spatial(tn):
    """
    Sort key for spatial ordering: top-to-bottom, then left-to-right.
    Projects the note's position onto the view's right/up axes.
    """
    try:
        p = tn.Coord
    except:
        try:
            loc = tn.Location
            p = loc.Point if loc else XYZ.Zero
        except:
            p = XYZ.Zero
    vr, vu = get_view_right_up(view)
    x = dot(p, vr)
    y = dot(p, vu)
    return (-y, x)


# --- Failure suppression ---

class _WeldNoteFailuresPreprocessor(IFailuresPreprocessor):
    """Deletes warning-level failures so they don't trigger rollback or UI."""
    def PreprocessFailures(self, failuresAccessor):
        try:
            for fmsg in failuresAccessor.GetFailureMessages():
                try:
                    if fmsg.GetSeverity().ToString().lower() == "warning":
                        failuresAccessor.DeleteWarning(fmsg)
                except:
                    pass
        except:
            pass
        return FailureProcessingResult.Continue


# --- View scan ---

def scan_weld_notes():
    """
    Scans the active view for TextNotes matching TARGET_TYPE_NAME.

    Returns dict with keys:
        targets, blank_notes, invalid_notes, numbered_notes,
        used, max_existing, gaps, dupes, duplicate_nums
    """
    targets = []
    for tn in (FilteredElementCollector(doc, view.Id)
               .OfClass(TextNote)
               .WhereElementIsNotElementType()):
        te = doc.GetElement(tn.GetTypeId())
        if type_name_matches(get_type_name(te)):
            targets.append(tn)

    blank_notes = []
    invalid_notes = []
    numbered_notes = []
    for tn in targets:
        txt = safe_str(tn.Text)
        if is_effectively_blank(txt):
            blank_notes.append(tn)
            continue
        n = parse_w_number(txt)
        if n is None:
            invalid_notes.append(tn)
        else:
            numbered_notes.append((tn, n))

    existing_nums = [n for (_, n) in numbered_notes]
    used = set(existing_nums)
    max_existing = max(existing_nums) if existing_nums else 0
    gaps = build_gaps(used)

    dupes = {}
    for tn, n in numbered_notes:
        dupes.setdefault(n, []).append(tn)
    duplicate_nums = sorted([n for n, lst in dupes.items() if len(lst) > 1])

    return {
        "targets": targets,
        "blank_notes": blank_notes,
        "invalid_notes": invalid_notes,
        "numbered_notes": numbered_notes,
        "used": used,
        "max_existing": max_existing,
        "gaps": gaps,
        "dupes": dupes,
        "duplicate_nums": duplicate_nums,
    }


# =============================================================================
# ACTION: RENUMBER ALL
# =============================================================================

def do_renumber_all():
    """
    Renumbers ALL matching weld notes in the view to W1..Wn using spatial sort.

    Uses a single Transaction with SubTransactions per element so one bad note
    (grouped, workshare-locked, etc.) can't undo the whole batch.
    This is dramatically faster than the per-element Transaction approach.
    """
    # Fresh scan to get current state
    fresh_targets = []
    for tn in (FilteredElementCollector(doc, view.Id)
               .OfClass(TextNote)
               .WhereElementIsNotElementType()):
        te = doc.GetElement(tn.GetTypeId())
        if type_name_matches(get_type_name(te)):
            fresh_targets.append(tn)

    ordered = sorted(fresh_targets, key=sort_key_spatial)
    failed = []
    skipped = []

    t = Transaction(doc, "Weld Note Renumber All")
    opts = t.GetFailureHandlingOptions()
    opts = opts.SetFailuresPreprocessor(_WeldNoteFailuresPreprocessor())
    try:
        opts = opts.SetClearAfterRollback(True)
    except:
        pass
    t.SetFailureHandlingOptions(opts)

    t.Start()
    try:
        for i, tn in enumerate(ordered, start=1):
            # Pre-check: skip grouped elements
            try:
                if tn.GroupId and tn.GroupId.IntegerValue != -1:
                    skipped.append((tn.Id.IntegerValue, "In group"))
                    continue
            except:
                pass

            # Pre-check: skip elements owned by another user (worksharing)
            if WorksharingUtils and CheckoutStatus:
                try:
                    cs = WorksharingUtils.GetCheckoutStatus(doc, tn.Id)
                    if cs == CheckoutStatus.OwnedByOtherUser:
                        skipped.append((tn.Id.IntegerValue, "Owned by other user"))
                        continue
                except:
                    pass

            # SubTransaction lets us roll back just this one element on failure
            sub = SubTransaction(doc)
            sub.Start()
            try:
                was_pinned = False
                try:
                    if tn.Pinned:
                        was_pinned = True
                        tn.Pinned = False
                except:
                    pass

                tn.Text = format_w(i)

                try:
                    if was_pinned:
                        tn.Pinned = True
                except:
                    pass

                sub.Commit()
            except Exception as ee:
                try:
                    sub.RollBack()
                except:
                    pass
                failed.append((tn.Id.IntegerValue, str(ee)))

        status = t.Commit()
        if status != TransactionStatus.Committed:
            failed.append((-1, "Transaction status: {}".format(status)))

    except Exception as e:
        try:
            if t.HasStarted():
                t.RollBack()
        except:
            pass
        TaskDialog.Show("Error", str(e))
        return

    # Refresh the view
    try:
        doc.Regenerate()
    except:
        pass
    try:
        uidoc.RefreshActiveView()
    except:
        pass

    # Report skipped/failed if any
    if failed or skipped:
        lines = []
        if skipped:
            lines.append("Skipped:")
            for _id, reason in skipped[:25]:
                lines.append("  - {} ({})".format(_id, reason))
            if len(skipped) > 25:
                lines.append("  ... {} more".format(len(skipped) - 25))
        if failed:
            if lines:
                lines.append("")
            lines.append("Failed:")
            for _id, reason in failed[:10]:
                lines.append("  - {} ({})".format(_id, reason))
            if len(failed) > 10:
                lines.append("  ... {} more".format(len(failed) - 10))
        TaskDialog.Show("Weld Note Tool", "Renumber complete.\n\n" + "\n".join(lines))


# =============================================================================
# ACTION: AUTO-FILL BLANKS/INVALIDS
# =============================================================================

def do_autofill(state):
    """
    Assigns W# numbers to existing blank (and optionally invalid) notes.
    Fills gaps first, then continues above max.

    Uses a single transaction (fast).
    """
    to_fill = list(state["blank_notes"])
    if INVALID_TEXT_BEHAVIOR == "treat_as_blank":
        to_fill.extend(state["invalid_notes"])

    if not to_fill:
        return

    local_used = set(state["used"])
    local_gaps = build_gaps(local_used)
    current_max = max(local_used) if local_used else 0

    t = Transaction(doc, "Weld Note Autofill")
    opts = t.GetFailureHandlingOptions()
    opts = opts.SetFailuresPreprocessor(_WeldNoteFailuresPreprocessor())
    try:
        opts = opts.SetClearAfterRollback(True)
    except:
        pass
    t.SetFailureHandlingOptions(opts)

    t.Start()
    try:
        # Phase 1: fill gaps
        idx = 0
        for tn in to_fill:
            if idx < len(local_gaps):
                n = local_gaps[idx]
                tn.Text = format_w(n)
                local_used.add(n)
                idx += 1
            else:
                break

        # Phase 2: continue above max for remaining
        remaining = to_fill[idx:]
        n = current_max + 1
        for tn in remaining:
            n = next_unused(n, local_used)
            tn.Text = format_w(n)
            local_used.add(n)
            n += 1

        status = t.Commit()
        if status != TransactionStatus.Committed:
            TaskDialog.Show(
                "Weld Note Tool",
                "Autofill transaction did not commit (status: {}).\n\n"
                "Revit may have rolled it back due to a failure."
                .format(status)
            )

    except Exception as e:
        try:
            if t.HasStarted():
                t.RollBack()
        except:
            pass
        TaskDialog.Show("Error", str(e))


# =============================================================================
# ACTION: PLACEMENT LOOP
# =============================================================================

def do_placement_loop(numbering_mode, state):
    """
    Interactive click-to-place loop for new weld notes.

    NO SHIFT = text placed LEFT of click point
    SHIFT    = text placed RIGHT of click point

    Each click creates a TextNote with a leader pointing to the clicked point.
    Leader geometry (elbow, exit length) is tuned per the CONFIG constants.
    """
    used = set(state["used"])
    gap_queue = list(state["gaps"])
    # Mutable container so the inner closure can update max_existing
    # (IronPython 2.7 does not support 'nonlocal')
    _state = {"max_existing": state["max_existing"]}

    def get_next_number():
        """Gets the next available weld number, filling gaps first in 'gaps' mode."""
        if numbering_mode == "gaps":
            # Consume missing numbers (smallest gaps first), then continue above max
            while gap_queue and gap_queue[0] in used:
                gap_queue.pop(0)
            if gap_queue:
                n = gap_queue.pop(0)
                used.add(n)
                return n

        n = _state["max_existing"] + 1 if _state["max_existing"] > 0 else 1
        while n in used:
            n += 1
        used.add(n)
        _state["max_existing"] = max(_state["max_existing"], n)
        return n

    while True:
        # Wait for user click (ESC breaks the loop)
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.PointOnElement,
                "Click weld/pipe point (ESC to finish) | NO SHIFT=Left | SHIFT=Right"
            )
        except OperationCanceledException:
            break

        view_right, view_up = get_view_right_up(view)
        view_dir = getattr(view, "ViewDirection", XYZ.BasisZ)

        sf = get_scale_factor(view)

        pt_arrow = ref.GlobalPoint
        text_side_sign = 1.0 if is_shift_down() else -1.0  # +1 right, -1 left

        # Compute text position relative to click point
        offset_side = OFFSET_SIDE_FEET * sf
        offset_up   = OFFSET_UP_FEET * sf

        side_offset = offset_side * (LEFT_EXTRA_MULT if text_side_sign < 0 else 1.0)
        pt_text = pt_arrow + (view_right * (side_offset * text_side_sign)) + (view_up * offset_up)

        t = Transaction(doc, "Place Weld Note")
        t.Start()
        try:
            # Set sketch plane facing camera (needed for 3D views)
            set_view_sketchplane_facing_camera(pt_arrow, view_dir)

            weld_num = get_next_number()
            note_text = format_w(weld_num)
            note = TextNote.Create(doc, view.Id, pt_text, note_text, text_type_id)

            leader = note.AddLeader(leader_type_for_text_side(text_side_sign))

            try:
                anchor = leader.Anchor
            except:
                anchor = pt_text

            # Determine which direction the arrow points relative to the text anchor
            arrow_sign = 1.0 if dot(pt_arrow - anchor, view_right) > 0 else -1.0

            # Compute elbow exit length with per-side adjustments
            exit_len = get_exit_len_for_side(text_side_sign) * sf

            if arrow_sign > 0:
                exit_len += BASE_ADJUST_EXIT_RIGHT_FEET * sf
            else:
                exit_len += BASE_ADJUST_EXIT_LEFT_FEET * sf

            # Add per-digit compensation when elbow exits right (wider text = longer exit)
            if arrow_sign > 0:
                d = number_digits(weld_num)
                extra_digits = max(0, d - 1)
                add_len = clamp(extra_digits * (CHAR_ADD_PER_EXTRA_DIGIT_FEET * sf),
                                0.0, CHAR_ADD_MAX_FEET * sf)
                exit_len += add_len

            elbow_pt = anchor + (view_right * (exit_len * arrow_sign))

            # Set leader elbow and end point (order can matter; try both ways)
            ok = False
            try:
                leader.Elbow = elbow_pt
                leader.End = pt_arrow
                ok = True
            except:
                pass
            if not ok:
                leader.End = pt_arrow
                try:
                    leader.Elbow = elbow_pt
                except:
                    pass

            t.Commit()
        except Exception as e:
            if t.HasStarted():
                t.RollBack()
            TaskDialog.Show("Error", str(e))
            break


# =============================================================================
# MAIN EXECUTION
# =============================================================================

# Safety check: 3D views must be locked
if view.ViewType == ViewType.ThreeD and REQUIRE_LOCKED_3D and (not view.IsLocked):
    TaskDialog.Show("Error", "Please LOCK your 3D view first.")
    script.exit()


# --- Resolve TextNoteType ---
text_type_id = None
for tt in FilteredElementCollector(doc).OfClass(TextNoteType).ToElements():
    if type_name_matches(get_type_name(tt)):
        text_type_id = tt.Id
        break

if text_type_id is None:
    existing = []
    try:
        for _tt in FilteredElementCollector(doc).OfClass(TextNoteType).ToElements():
            _n = safe_str(get_type_name(_tt))
            if _n:
                existing.append(_n)
    except:
        pass
    try:
        existing = sorted(set(existing), key=lambda s: s.lower())
    except:
        pass
    msg = (
        "Required Text Note Type not found.\n\n"
        "Please create a Text Note Type named:\n  \"{}\"\n\n"
        "Then run this tool again.\n\n"
        "Existing Text Note Types in this model (for reference):\n- {}"
    ).format(TARGET_TYPE_NAME, "\n- ".join(existing[:25]) if existing else "(could not enumerate)")
    TaskDialog.Show("Missing Text Note Type", msg)
    script.exit()


# --- Scan existing weld notes ---
state = scan_weld_notes()
startup_action = "continue"
numbering_mode = "highest"


# --- Startup UI ---
# No existing notes -> go straight to placement (no dialogs)
if state["targets"]:

    # If duplicates exist, show handling dialog first
    if state["duplicate_nums"]:
        lines = []
        for n in state["duplicate_nums"]:
            lines.append("- {} occurs {} times".format(format_w(n), len(state["dupes"][n])))
        summary = "\n".join(lines)

        dup_choice = taskdialog_choice(
            title="Place Weld Notes",
            main_instruction="Duplicate weld numbers found in this view",
            content="{}\n\nChoose how to proceed:".format(summary),
            options=[
                ("dup_select", "Select duplicates and stop",
                 "Highlights duplicate notes so you can manually fix them, then rerun."),
                ("renumber_all", "Renumber everything in view (stop)",
                 "Renumber ALL weld notes in this view to W1..Wn and EXIT (no placement)."),
                ("continue_anyway", "Continue anyway (place)",
                 "Keep existing tags as-is and proceed to placement.")
            ]
        )
        if dup_choice is None:
            script.exit()
        if dup_choice == "dup_select":
            ids = []
            for n in state["duplicate_nums"]:
                for tn in state["dupes"][n]:
                    ids.append(tn.Id)
            set_selection(ids)
            script.exit()
        if dup_choice == "renumber_all":
            startup_action = "renumber_all"

    # Main 3-option dialog (shown when notes exist, unless renumber already chosen)
    if startup_action != "renumber_all":
        has_fillable = (
            len(state["blank_notes"]) > 0
            or (len(state["invalid_notes"]) > 0 and INVALID_TEXT_BEHAVIOR == "treat_as_blank")
        )

        content = (
            "View: {}\n"
            "Target type: {}\n\n"
            "Matching notes: {}\n"
            "Valid W#: {}\n"
            "Blank: {}\n"
            "Invalid: {}\n"
            "Highest existing: {}\n"
            "Gaps: {}"
        ).format(
            view.Name, TARGET_TYPE_NAME,
            len(state["targets"]),
            len(state["numbered_notes"]),
            len(state["blank_notes"]),
            len(state["invalid_notes"]),
            state["max_existing"],
            len(state["gaps"])
        )

        options = [
            ("continue", "Continue from highest (place)",
             "Start placing new weld notes at (max+1), avoiding used numbers."),
            ("autofill_then_place", "Auto-fill blanks/invalids, then place",
             "Assign numbers to existing blank/invalid notes first, then start placing."),
            ("renumber_all", "Renumber everything in view (stop)",
             "Renumber ALL matching notes to W1..Wn and EXIT (no placement).")
        ]

        if not has_fillable:
            options[1] = ("autofill_then_place", "Auto-fill blanks/invalids, then place",
                          "No blanks/invalids detected; will proceed directly to placement.")

        choice = taskdialog_choice(
            title="Place Weld Notes",
            main_instruction="Choose how to start numbering",
            content=content,
            options=options
        )
        if choice is None:
            script.exit()
        startup_action = choice
        if startup_action == "autofill_then_place":
            numbering_mode = "gaps"


# --- Execute chosen action ---

if startup_action == "renumber_all":
    do_renumber_all()

elif startup_action == "autofill_then_place":
    do_autofill(state)
    state = scan_weld_notes()
    do_placement_loop(numbering_mode, state)

else:
    # "continue" — go straight to placement
    do_placement_loop(numbering_mode, state)
