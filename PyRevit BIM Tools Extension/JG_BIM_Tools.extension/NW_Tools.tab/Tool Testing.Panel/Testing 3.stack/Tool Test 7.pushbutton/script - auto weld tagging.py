# -*- coding: utf-8 -*-
# pyRevit: Place Weld Notes by clicking points
# NO SHIFT = text LEFT, SHIFT = text RIGHT
#
# Scale-safe (tuned-at-base-scale) version:
# - Your original feet-based tuning is preserved at BASE_TUNED_SCALE = 24 (1/2" = 1'-0")
# - When view scale changes, all offsets/leader tuning scale by (view.Scale / BASE_TUNED_SCALE)
#
# Leader behavior unchanged:
# - leader attaches based on text side
# - arrow_sign determines elbow direction
#
# 3D safety included (requires locked 3D view)

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, XYZ, Plane, SketchPlane,
    TextNote, TextNoteLeaderTypes, ElementTypeGroup, ViewType,
    TextNoteType, FilteredElementCollector, BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog
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
    """
    Scale factor relative to the scale you tuned at.
    Example: tuned at 24.
      - If view.Scale = 24 => factor 1.0 (same as original)
      - If view.Scale = 48 => factor 2.0 (double model distances)
      - If view.Scale = 12 => factor 0.5 (half model distances)
    """
    try:
        s = float(v.Scale)
        b = float(BASE_TUNED_SCALE) if BASE_TUNED_SCALE else s
        if b == 0:
            return 1.0
        return s / b
    except:
        return 1.0


# Safety: locked 3D
if view.ViewType == ViewType.ThreeD and REQUIRE_LOCKED_3D and (not view.IsLocked):
    TaskDialog.Show("Error", "Please LOCK your 3D view first.")
    script.exit()

# Resolve TextNoteType
text_type_id = None
for tt in FilteredElementCollector(doc).OfClass(TextNoteType).ToElements():
    if type_name_matches(get_type_name(tt)):
        text_type_id = tt.Id
        break
if text_type_id is None:
    text_type_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)

# Existing tags in view
existing_nums = []
for tn in (FilteredElementCollector(doc, view.Id)
           .OfClass(TextNote)
           .WhereElementIsNotElementType()):
    te = doc.GetElement(tn.GetTypeId())
    if not type_name_matches(get_type_name(te)):
        continue
    n = parse_w_number(tn.Text)
    if n is not None:
        existing_nums.append(n)

used = set(existing_nums)
max_existing = max(existing_nums) if existing_nums else 0

def get_next_number():
    global used, max_existing
    n = max_existing + 1 if max_existing > 0 else 1
    while n in used:
        n += 1
    used.add(n)
    max_existing = max(max_existing, n)
    return n


# Click loop
while True:
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.PointOnElement,
            "Click weld/pipe point (ESC to finish) | NO SHIFT=Left | SHIFT=Right"
        )
    except OperationCanceledException:
        break

    # Refresh view vectors per pick
    view_right, view_up = get_view_right_up(view)
    view_dir = getattr(view, "ViewDirection", XYZ.BasisZ)

    # Scale factor for THIS view
    sf = get_scale_factor(view)

    pt_arrow = ref.GlobalPoint
    text_side_sign = 1.0 if is_shift_down() else -1.0  # +1 right, -1 left

    # Text placement (scaled from tuned-at-24 values)
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

        # Anchor = actual attachment point on the box
        try:
            anchor = leader.Anchor
        except:
            anchor = pt_text

        # Determine arrow direction along view_right (UNCHANGED)
        arrow_sign = 1.0 if dot(pt_arrow - anchor, view_right) > 0 else -1.0

        # Base tuned exit length (scaled)
        exit_len = get_exit_len_for_side(text_side_sign) * sf

        # Starting condition tweak (scaled)
        if arrow_sign > 0:
            exit_len += BASE_ADJUST_EXIT_RIGHT_FEET * sf
        else:
            exit_len += BASE_ADJUST_EXIT_LEFT_FEET * sf

        # Variable extension for extra digits (scaled)
        if arrow_sign > 0:
            d = number_digits(weld_num)
            extra_digits = max(0, d - 1)
            add_len = clamp(extra_digits * (CHAR_ADD_PER_EXTRA_DIGIT_FEET * sf),
                            0.0, CHAR_ADD_MAX_FEET * sf)
            exit_len += add_len

        # Flat segment out of the box
        elbow_pt = anchor + (view_right * (exit_len * arrow_sign))

        # Apply (order sensitivity)
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
