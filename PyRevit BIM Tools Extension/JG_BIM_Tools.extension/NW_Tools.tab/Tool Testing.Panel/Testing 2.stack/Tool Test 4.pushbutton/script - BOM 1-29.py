# -*- coding: utf-8 -*-
# pyRevit: Place Weld Notes by clicking points
# NO SHIFT = text LEFT, SHIFT = text RIGHT
#
# Scale-safe (tuned-at-base-scale) version:
# - Your original feet-based tuning is preserved at BASE_TUNED_SCALE = 24 (1/2" = 1'-0")
# - When view scale changes, all offsets/leader tuning scale by (view.Scale / BASE_TUNED_SCALE)
#
# Behavior summary:
# - If NO matching Weld Number notes exist in the active view:
#     -> NO dialogs, goes straight into click-to-place (original behavior).
# - If matching notes DO exist:
#     -> Show ONE startup dialog with THREE options:
#         1) Continue from highest (starts placement)
#         2) Auto-fill blanks/invalids, then place (autofill + starts placement)
#         3) Renumber everything in view (renumber + EXIT silently, no placement, no popup)
# - Duplicate detection:
#     -> Optional “Select duplicates and stop” first, else you can choose renumber or proceed.
#
# Notes:
# - Auto-fill assigns W# to existing blank/invalid notes; it uses gaps first, then above max.
# - Renumber assigns ALL matching notes in view to W1..Wn using a stable-ish spatial sort.

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Plane, SketchPlane,
    TextNote, TextNoteLeaderTypes, ElementTypeGroup, ViewType,
    TextNoteType, FilteredElementCollector, BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogCommonButtons, TaskDialogResult
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Windows.Input import Keyboard, Key
import re

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView

# ---------------- CONFIG ----------------
TARGET_TYPE_NAME = "Weld Number"
MATCH_TYPE_MODE = "equals"     # "equals" or "contains"
LABEL_PREFIX = "W"

# Your tuned base view scale: 1/2" = 1'-0" => Revit scale 24
BASE_TUNED_SCALE = 24.0

# Text placement (feet) - tuned at BASE_TUNED_SCALE
OFFSET_SIDE_FEET = 1.5
OFFSET_UP_FEET   = 1.0
LEFT_EXTRA_MULT  = 1.15

# Exit length tuning (feet) (tuned at BASE_TUNED_SCALE)
ELBOW_OUT_LEFT_FEET = 0.55     # left = 0.55 ft
RIGHT_EXIT_MULT     = 0.3      # right = left * 0.3

# Starting condition tweak depending on which way elbow exits (arrow_sign)
BASE_ADJUST_EXIT_RIGHT_FEET = -0.04
BASE_ADJUST_EXIT_LEFT_FEET  =  0.00

# Variable add-on ONLY when elbow exits to the RIGHT (arrow_sign > 0)
CHAR_ADD_PER_EXTRA_DIGIT_FEET = 0.09
CHAR_ADD_MAX_FEET             = 0.20

# 3D SAFETY
REQUIRE_LOCKED_3D = True

# Smart text handling when autofilling existing notes:
# - "treat_as_blank": autofill blank + invalid with numbers
# - "skip": autofill blank only, leave invalid as-is unless renumber-all
INVALID_TEXT_BEHAVIOR = "treat_as_blank"  # "treat_as_blank" | "skip"
# --------------------------------------


def safe_str(x):
    try:
        return x if x is not None else ""
    except:
        return ""

def normalize_text(s):
    s = safe_str(s).replace(u"\u00A0", " ")
    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def is_effectively_blank(s):
    return normalize_text(s) == ""

def get_type_name(type_elem):
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

def number_digits(n):
    try:
        return len(str(int(n)))
    except:
        return 1

def format_w(n):
    return "{}{}".format(LABEL_PREFIX, n)

def get_view_right_up(v):
    try:
        return v.RightDirection, v.UpDirection
    except:
        return XYZ.BasisX, XYZ.BasisY

def dot(a, b):
    return a.X*b.X + a.Y*b.Y + a.Z*b.Z

def is_shift_down():
    return Keyboard.IsKeyDown(Key.LeftShift) or Keyboard.IsKeyDown(Key.RightShift)

def set_view_sketchplane_facing_camera(pt, view_dir_vec):
    geom_plane = Plane.CreateByNormalAndOrigin(view_dir_vec, pt)
    sketch = SketchPlane.Create(doc, geom_plane)
    try:
        view.SketchPlane = sketch
    except:
        pass

def leader_type_for_text_side(text_side_sign):
    # text right => leader attaches left; text left => leader attaches right
    return TextNoteLeaderTypes.TNLT_STRAIGHT_L if text_side_sign > 0 else TextNoteLeaderTypes.TNLT_STRAIGHT_R

def get_exit_len_for_side(text_side_sign):
    # text_side_sign: +1 => text placed RIGHT, -1 => text placed LEFT
    if text_side_sign > 0:
        return ELBOW_OUT_LEFT_FEET * RIGHT_EXIT_MULT
    return ELBOW_OUT_LEFT_FEET

def clamp(x, a, b):
    return max(a, min(b, x))

def get_scale_factor(v):
    try:
        s = float(v.Scale)
        b = float(BASE_TUNED_SCALE) if BASE_TUNED_SCALE else s
        if b == 0:
            return 1.0
        return s / b
    except:
        return 1.0

def taskdialog_choice(title, main_instruction, content, options):
    """
    options: list of tuples [(key, label, description), ...] up to 4
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
        from Autodesk.Revit.DB import ElementId
        ids = List[ElementId]()
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

def build_gaps(used_set):
    if not used_set:
        return []
    m = max(used_set)
    return [k for k in range(1, m + 1) if k not in used_set]


# ----------------------------
# Safety: locked 3D
# ----------------------------
if view.ViewType == ViewType.ThreeD and REQUIRE_LOCKED_3D and (not view.IsLocked):
    TaskDialog.Show("Error", "Please LOCK your 3D view first.")
    script.exit()

# ----------------------------
# Resolve TextNoteType
# ----------------------------
text_type_id = None
for tt in FilteredElementCollector(doc).OfClass(TextNoteType).ToElements():
    if type_name_matches(get_type_name(tt)):
        text_type_id = tt.Id
        break
if text_type_id is None:
    text_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

# ----------------------------
# Scan existing Weld Number notes in THIS view
# ----------------------------
targets = []
for tn in (FilteredElementCollector(doc, view.Id)
           .OfClass(TextNote)
           .WhereElementIsNotElementType()):
    te = doc.GetElement(tn.GetTypeId())
    if type_name_matches(get_type_name(te)):
        targets.append(tn)

blank_notes = []
invalid_notes = []
numbered_notes = []  # (tn, n)

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

dupes = {}
for tn, n in numbered_notes:
    dupes.setdefault(n, []).append(tn)
duplicate_nums = sorted([n for n, lst in dupes.items() if len(lst) > 1])

gaps = build_gaps(used)

gap_queue = list(gaps)

# ----------------------------
# Startup UI logic
# ----------------------------
startup_action = "continue"
NUMBERING_MODE = "highest"  # "highest" or "gaps"
gap_queue = []

# Default: if there are no existing weld notes, go straight into placement
run_place_loop = True

# If there are NO matching notes in view: start placement immediately (no UI)
if not targets:
    startup_action = "continue"
else:
    # If duplicates exist, give a quick handling dialog first
    if duplicate_nums:
        lines = []
        for n in duplicate_nums:
            lines.append("- {} occurs {} times".format(format_w(n), len(dupes[n])))
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
            for n in duplicate_nums:
                for tn in dupes[n]:
                    ids.append(tn.Id)
            set_selection(ids)
            script.exit()
        if dup_choice == "renumber_all":
            startup_action = "renumber_all"
        else:
            startup_action = "continue"

    # Main 3-option screen (always shown when targets exist, unless we already chose renumber_all)
    if startup_action != "renumber_all":
        has_fillable = (len(blank_notes) > 0) or (len(invalid_notes) > 0 and INVALID_TEXT_BEHAVIOR == "treat_as_blank")

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
            len(targets),
            len(numbered_notes),
            len(blank_notes),
            len(invalid_notes),
            max_existing,
            len(gaps)
        )

        options = [
            ("continue", "Continue from highest (place)",
             "Start placing new weld notes at (max+1), avoiding used numbers."),
            ("autofill_then_place", "Auto-fill blanks/invalids, then place",
             "Assign numbers to existing blank/invalid notes first, then start placing."),
            ("renumber_all", "Renumber everything in view (stop)",
             "Renumber ALL matching notes to W1..Wn and EXIT (no placement).")
        ]

        # If there’s nothing to autofill, selecting option 2 would be pointless.
        # We still keep the 3-option screen, but make option 2 describe that it won’t change anything.
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
        run_place_loop = True
        if startup_action == "renumber_all":
            run_place_loop = False
        if startup_action == "autofill_then_place":
            NUMBERING_MODE = "gaps"
        else:
            NUMBERING_MODE = "highest"
# ----------------------------
# Pre-actions (autofill / renumber)
# ----------------------------
def refresh_state():
    global targets, blank_notes, invalid_notes, numbered_notes, used, max_existing, gaps, dupes, duplicate_nums

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


if startup_action in ("renumber_all", "autofill_then_place"):
    # IMPORTANT:
    # - Revit can silently rollback a large transaction on commit if ANY element produces an error/failure.
    # - To make "Renumber All" stick reliably, we commit per-element (SubTransaction) inside a TransactionGroup.
    #   That way, one bad note (grouped, owned by another user, etc.) can't undo the whole renumber operation.
    from Autodesk.Revit.DB import (
        FailureHandlingOptions, IFailuresPreprocessor, FailureProcessingResult,
        TransactionStatus, TransactionGroup, Transaction
    )
    try:
        # Worksharing ownership check (optional; if not workshared, this import will still be ok)
        from Autodesk.Revit.DB import WorksharingUtils, CheckoutStatus
    except:
        WorksharingUtils = None
        CheckoutStatus = None

    class _WeldNoteFailuresPreprocessor(IFailuresPreprocessor):
        def PreprocessFailures(self, failuresAccessor):
            # Delete warnings so they don't trigger rollback / long failure resolution UI.
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

    did_renumber_all = False

    # ---- Renumber All: commit per element so it can't "snap back" ----
    if startup_action == "renumber_all":
        def sort_key(tn):
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

        fresh_targets = []
        for _tn in (FilteredElementCollector(doc, view.Id)
                    .OfClass(TextNote)
                    .WhereElementIsNotElementType()):
            _te = doc.GetElement(_tn.GetTypeId())
            if type_name_matches(get_type_name(_te)):
                fresh_targets.append(_tn)

        ordered = sorted(fresh_targets, key=sort_key)

        failed = []   # list of (id, reason)
        skipped = []  # list of (id, reason)

        tg = TransactionGroup(doc, "Weld Note Renumber All")
        tg.Start()
        try:
            for i, _tn in enumerate(ordered, start=1):
                # Skip obvious "not editable" cases before we even try
                try:
                    if _tn.GroupId and _tn.GroupId.IntegerValue != -1:
                        skipped.append((_tn.Id.IntegerValue, "In group"))
                        continue
                except:
                    pass

                if WorksharingUtils and CheckoutStatus:
                    try:
                        cs = WorksharingUtils.GetCheckoutStatus(doc, _tn.Id)
                        if cs == CheckoutStatus.OwnedByOtherUser:
                            skipped.append((_tn.Id.IntegerValue, "Owned by other user"))
                            continue
                    except:
                        pass

                tx = Transaction(doc, "Renumber Weld Note")
                tx.Start()
                try:
                    # Temporarily unpin if needed
                    was_pinned = False
                    try:
                        if _tn.Pinned:
                            was_pinned = True
                            _tn.Pinned = False
                    except:
                        pass

                    _tn.Text = format_w(i)

                    try:
                        if was_pinned:
                            _tn.Pinned = True
                    except:
                        pass

                    status = tx.Commit()
                    if status != TransactionStatus.Committed:
                        failed.append((_tn.Id.IntegerValue, "Commit status: {}".format(status)))
                except Exception as ee:
                    try:
                        if tx.HasStarted():
                            tx.RollBack()
                    except:
                        pass
                    failed.append((_tn.Id.IntegerValue, str(ee)))

            tg.Assimilate()
            did_renumber_all = True

        except Exception as e:
            try:
                tg.RollBack()
            except:
                pass
            TaskDialog.Show("Error", str(e))
            run_place_loop = False

        # Refresh + report, then stop cleanly (no placement loop)
        try:
            doc.Regenerate()
        except:
            pass
        try:
            uidoc.RefreshActiveView()
        except:
            pass

        if failed or skipped:
            try:
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
            except:
                pass

        run_place_loop = False

    # ---- Autofill at startup: keep as single transaction (fast) ----
    t = Transaction(doc, "Weld Note Startup Numbering")
    t.Start()
    try:
        opts = t.GetFailureHandlingOptions()
        opts = opts.SetFailuresPreprocessor(_WeldNoteFailuresPreprocessor())
        try:
            opts = opts.SetClearAfterRollback(True)
        except:
            pass
        t.SetFailureHandlingOptions(opts)

        # Autofill blanks, and optionally invalid depending on config.
        to_fill = list(blank_notes)
        if INVALID_TEXT_BEHAVIOR == "treat_as_blank":
            to_fill.extend(invalid_notes)

        local_used = set(used)
        local_gaps = build_gaps(local_used)
        current_max = max(local_used) if local_used else 0

        idx = 0
        for tn in to_fill:
            if idx < len(local_gaps):
                n = local_gaps[idx]
                tn.Text = format_w(n)
                local_used.add(n)
                idx += 1
            else:
                break

        remaining = to_fill[idx:]
        n = current_max + 1
        for tn in remaining:
            n = next_unused(n, local_used)
            tn.Text = format_w(n)
            local_used.add(n)
            n += 1

        status = t.Commit()
        try:
            if status != TransactionStatus.Committed:
                TaskDialog.Show(
                    "Weld Note Tool",
                    "Autofill transaction did not commit (status: {}).\n\n"
                    "Revit may have rolled it back due to a failure."
                    .format(status)
                )
        except:
            pass

    except Exception as e:
        try:
            if t.HasStarted():
                t.RollBack()
        except:
            pass
        TaskDialog.Show("Error", str(e))
        run_place_loop = False

# Refresh after any pre-actions


refresh_state()
gap_queue = list(gaps)
# If user chose renumber_all, skip placement loop and end like ESC
if startup_action == "renumber_all":
    run_place_loop = False


# ----------------------------
# Number generator for NEW placements
# ----------------------------
def get_next_number():
    global used, max_existing, gap_queue, NUMBERING_MODE

    if NUMBERING_MODE == "gaps":
        # Consume missing numbers first (smallest gaps), then continue above max.
        while gap_queue and gap_queue[0] in used:
            gap_queue.pop(0)
        if gap_queue:
            n = gap_queue.pop(0)
            used.add(n)
            return n

    n = max_existing + 1 if max_existing > 0 else 1
    while n in used:
        n += 1
    used.add(n)
    max_existing = max(max_existing, n)
    return n


# ----------------------------
# Click loop (placement)
# ----------------------------
if run_place_loop:
    while True:
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
    
        offset_side = OFFSET_SIDE_FEET * sf
        offset_up   = OFFSET_UP_FEET * sf
    
        side_offset = offset_side * (LEFT_EXTRA_MULT if text_side_sign < 0 else 1.0)
        pt_text = pt_arrow + (view_right * (side_offset * text_side_sign)) + (view_up * offset_up)
    
        t = Transaction(doc, "Place Weld Note")
        t.Start()
        try:
            set_view_sketchplane_facing_camera(pt_arrow, view_dir)
    
            weld_num = get_next_number()
            note_text = format_w(weld_num)
            note = TextNote.Create(doc, view.Id, pt_text, note_text, text_type_id)
    
            leader = note.AddLeader(leader_type_for_text_side(text_side_sign))
    
            try:
                anchor = leader.Anchor
            except:
                anchor = pt_text
    
            arrow_sign = 1.0 if dot(pt_arrow - anchor, view_right) > 0 else -1.0
    
            exit_len = get_exit_len_for_side(text_side_sign) * sf
    
            if arrow_sign > 0:
                exit_len += BASE_ADJUST_EXIT_RIGHT_FEET * sf
            else:
                exit_len += BASE_ADJUST_EXIT_LEFT_FEET * sf
    
            if arrow_sign > 0:
                d = number_digits(weld_num)
                extra_digits = max(0, d - 1)
                add_len = clamp(extra_digits * (CHAR_ADD_PER_EXTRA_DIGIT_FEET * sf),
                                0.0, CHAR_ADD_MAX_FEET * sf)
                exit_len += add_len
    
            elbow_pt = anchor + (view_right * (exit_len * arrow_sign))
    
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