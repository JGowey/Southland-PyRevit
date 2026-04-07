# -*- coding: utf-8 -*-
"""Assembly Studio - Full assembly lifecycle in one tool.

Combines Assembly Builder and Assembly Sheet Creator into a single workflow:

  BUILD MODE (fabrication elements pre-selected):
    1. Configure assembly name, level, which views to create, and view templates.
    2. Click Create Assembly -- assembly and views are created in Revit.
    3. Views automatically populate the sheet layout canvas.
    4. Drag viewport rectangles to position, set scale/viewport type per view.
    5. Fill titleblock parameters, optionally load a layout preset.
    6. Click Create Sheet.

  EXISTING MODE (assembly pre-selected, or no relevant selection):
    1. Pick an assembly from the dropdown.
    2. Views load automatically in the sheet layout canvas.
    3. Continue from step 4 above.

  Presets save/restore layout (positions, scales, viewport types, parameters)
  for consistent sheets across spools or jobs.

Author: Jeremiah Griffith
Version: 1.0.0
"""

from __future__ import division, print_function
from pyrevit import revit, DB, forms, script
import clr, os, sys, json, math

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Windows import (
    Window, Visibility, Thickness, HorizontalAlignment,
    VerticalAlignment, TextAlignment,
)
from System.Windows.Controls import (
    Canvas, Border, TextBlock, CheckBox, ComboBox, ComboBoxItem,
    StackPanel, Grid, ColumnDefinition, TextBox, Orientation, Label,
    Separator,
)
from System.Windows.Media import (
    SolidColorBrush, Colors, Color, Brushes, DoubleCollection,
)
from System.Windows.Shapes import Line, Rectangle
from System.Collections.Generic import List
import System.Windows

from Autodesk.Revit.DB import (
    AssemblyInstance, AssemblyViewUtils, AssemblyDetailViewOrientation,
    FilteredElementCollector, ViewFamilyType, ViewFamily,
    Transaction, ElementId, Level, View, ViewSheet, ViewSchedule,
    Viewport, ScheduleSheetInstance, XYZ, ElementType,
    BuiltInCategory, BuiltInParameter, StorageType,
    Options as RvtOptions, GeometryInstance,
)
# Alias Revit's Line so it doesn't clash with System.Windows.Shapes.Line
import Autodesk.Revit.DB as RvtDB

doc   = revit.doc
uidoc = revit.uidoc
out   = script.get_output()

XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")

# Presets directory: <extension root>/config/sheet_presets/
EXT_DIR = os.path.dirname(__file__)
for _ in range(5):
    EXT_DIR = os.path.dirname(EXT_DIR)
PRESETS_DIR = os.path.join(EXT_DIR, "config", "sheet_presets")

# Per-titleblock usable area config (persists across sessions)
TB_AREA_CONFIG_PATH = os.path.join(EXT_DIR, "config", "tb_usable_areas.json")


def _load_tb_area_config():
    """Return the full config dict keyed by titleblock type id string."""
    try:
        if os.path.exists(TB_AREA_CONFIG_PATH):
            with open(TB_AREA_CONFIG_PATH, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _save_tb_area_config(tb_type_id, area_dict, margins_dict):
    """Persist usable area + margins for a titleblock type."""
    try:
        if not os.path.exists(os.path.dirname(TB_AREA_CONFIG_PATH)):
            os.makedirs(os.path.dirname(TB_AREA_CONFIG_PATH))
        cfg = _load_tb_area_config()
        cfg[str(tb_type_id.Value)] = {
            "area":    area_dict,    # {x_in, y_in, w_in, h_in}
            "margins": margins_dict, # {margin_l, margin_r, margin_t, margin_b, title_h}
        }
        with open(TB_AREA_CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
    except Exception:
        pass


def _get_saved_tb_area(tb_type_id):
    """Return saved {area, margins} for this titleblock type, or None."""
    cfg = _load_tb_area_config()
    return cfg.get(str(tb_type_id.Value))

# Standard architectural scales: (display_text, scale_factor)
SCALES = [
    ('3" = 1\'-0"',     4),
    ('1-1/2" = 1\'-0"', 8),
    ('1" = 1\'-0"',    12),
    ('3/4" = 1\'-0"',  16),
    ('1/2" = 1\'-0"',  24),
    ('3/8" = 1\'-0"',  32),
    ('1/4" = 1\'-0"',  48),
    ('3/16" = 1\'-0"', 64),
    ('1/8" = 1\'-0"',  96),
]

# Viewport colors keyed by view type
VP_COLORS = {
    "3d":       Color.FromRgb(59,  130, 246),  # blue
    "top":      Color.FromRgb(34,  197, 94),   # green
    "front":    Color.FromRgb(245, 158, 11),   # orange
    "back":     Color.FromRgb(245, 158, 11),
    "left":     Color.FromRgb(245, 158, 11),
    "right":    Color.FromRgb(245, 158, 11),
    "schedule": Color.FromRgb(107, 114, 128),  # gray
    "other":    Color.FromRgb(139, 92,  246),  # purple
}

PX_PER_INCH = 100  # canvas pixels per sheet inch
SNAP_PX     = 25   # 1/4 inch snap grid (PX_PER_INCH / 4)


# ─────────────────────────────────────────────────────────────────────
#  AssemblyDetailViewOrientation enum resolution
#  Revit 2022+ uses ElevationTop / ElevationFront / etc.
# ─────────────────────────────────────────────────────────────────────
def _resolve_orient(primary, fallback):
    for name in (primary, fallback):
        try:
            return getattr(AssemblyDetailViewOrientation, name)
        except AttributeError:
            pass
    return None

ORIENT_TOP   = _resolve_orient("ElevationTop",   "DetailViewTop")
ORIENT_FRONT = _resolve_orient("ElevationFront", "DetailViewFront")
ORIENT_BACK  = _resolve_orient("ElevationBack",  "DetailViewBack")
ORIENT_LEFT  = _resolve_orient("ElevationLeft",  "DetailViewLeft")
ORIENT_RIGHT = _resolve_orient("ElevationRight", "DetailViewRight")


# ─────────────────────────────────────────────────────────────────────
#  DATA HELPERS
# ─────────────────────────────────────────────────────────────────────

def get_selected_elements():
    sel = uidoc.Selection.GetElementIds()
    if not sel:
        return []
    return [doc.GetElement(eid) for eid in sel]


def get_assemblies():
    """Return [(display_name, ElementId), ...] sorted by name."""
    result = []
    for a in (FilteredElementCollector(doc)
              .OfClass(AssemblyInstance)
              .WhereElementIsNotElementType()):
        try:
            name = a.AssemblyTypeName or "Unnamed"
        except Exception:
            name = "Assembly"
        display = "{} (ID {})".format(name, a.Id.Value)
        result.append((display, a.Id))
    return sorted(result, key=lambda x: x[0].lower())


def get_assembly_views(assembly_id):
    """Return [(view_key, view_name, view_id, is_schedule), ...] for an assembly.

    view_key is one of: 3d top front back left right schedule other.
    """
    results = []
    for v in (FilteredElementCollector(doc)
              .OfClass(View)
              .WhereElementIsNotElementType()):
        if v.IsTemplate:
            continue
        try:
            assoc = v.AssociatedAssemblyInstanceId
            if assoc is None or assoc == ElementId.InvalidElementId:
                continue
            if assoc.Value != assembly_id.Value:
                continue
        except Exception:
            continue

        is_sched = isinstance(v, ViewSchedule)
        vname = v.Name or "Unnamed View"
        name_lower = vname.lower()

        if is_sched:
            vkey = "schedule"
        elif "3d" in name_lower or "orthographic" in name_lower:
            vkey = "3d"
        elif "top" in name_lower or "plan" in name_lower:
            vkey = "top"
        elif "front" in name_lower:
            vkey = "front"
        elif "back" in name_lower:
            vkey = "back"
        elif "left" in name_lower:
            vkey = "left"
        elif "right" in name_lower:
            vkey = "right"
        else:
            vkey = "other"

        results.append((vkey, vname, v.Id, is_sched))
    return results


def get_levels():
    levels = list(
        FilteredElementCollector(doc)
        .OfClass(Level)
        .WhereElementIsNotElementType()
        .ToElements()
    )
    return sorted(levels, key=lambda lv: lv.Elevation)


def get_nearest_level(elements):
    """Return the level at or just below the minimum Z of the selection."""
    levels = get_levels()
    if not levels:
        return None
    min_z = None
    for elem in elements:
        try:
            bbox = elem.get_BoundingBox(None)
            if bbox:
                z = bbox.Min.Z
                if min_z is None or z < min_z:
                    min_z = z
        except Exception:
            pass
    if min_z is None:
        return levels[0]
    best = levels[0]
    for lv in levels:
        if lv.Elevation <= min_z:
            best = lv
        else:
            break
    return best


def get_view_templates():
    """Return [(name, ElementId), ...] for non-schedule view templates."""
    result = []
    for v in (FilteredElementCollector(doc)
              .OfClass(View)
              .WhereElementIsNotElementType()):
        if v.IsTemplate and not isinstance(v, ViewSchedule):
            result.append((v.Name, v.Id))
    return sorted(result, key=lambda x: x[0].lower())


def get_schedule_templates():
    """Return [(name, ElementId), ...] for schedule view templates only."""
    result = []
    for v in (FilteredElementCollector(doc)
              .OfClass(ViewSchedule)
              .WhereElementIsNotElementType()):
        if v.IsTemplate:
            result.append((v.Name, v.Id))
    return sorted(result, key=lambda x: x[0].lower())


def get_title_blocks():
    """Return [(display_name, ElementId), ...] sorted."""
    tbs = []
    for elem in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_TitleBlocks)
                 .WhereElementIsElementType()):
        try:
            fam = elem.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString() or "?"
            typ = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "?"
            name = "{}: {}".format(fam, typ)
        except Exception:
            name = getattr(elem, "Name", str(elem.Id))
        tbs.append((name, elem.Id))
    return sorted(tbs, key=lambda x: x[0].lower())


def get_viewport_types():
    """Return [(name, ElementId), ...] for Viewport family types.

    Uses OST_Viewports category filter which is more reliable than
    iterating ElementType and matching FamilyName (FamilyName can be
    empty or inaccessible for system families in IronPython).
    """
    result = []
    try:
        for vt in (FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_Viewports)
                   .WhereElementIsElementType()):
            try:
                name = vt.Name or "Unnamed"
                result.append((name, vt.Id))
            except Exception:
                pass
    except Exception:
        # Fallback: scan all ElementTypes
        for vt in FilteredElementCollector(doc).OfClass(ElementType):
            try:
                fn = ""
                try:    fn = vt.FamilyName or ""
                except: pass
                if fn == "Viewport" or (
                        vt.Category is not None and
                        vt.Category.Id.Value ==
                        int(BuiltInCategory.OST_Viewports)):
                    result.append((vt.Name or "Unnamed", vt.Id))
            except Exception:
                pass
    return sorted(result, key=lambda x: x[0].lower())


def get_titleblock_size(tb_type_id):
    """Return (width_inches, height_inches). Defaults to 17x11."""
    tb_type = doc.GetElement(tb_type_id)
    if tb_type is None:
        return (17.0, 11.0)
    try:
        w = tb_type.get_Parameter(BuiltInParameter.SHEET_WIDTH)
        h = tb_type.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
        if w and h:
            wi = w.AsDouble() * 12.0
            hi = h.AsDouble() * 12.0
            if wi > 0 and hi > 0:
                return (wi, hi)
    except Exception:
        pass
    # Fall back to measuring an existing placed instance
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        try:
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                if tb.GetTypeId().Value == tb_type_id.Value:
                    bbox = tb.get_BoundingBox(sheet)
                    if bbox:
                        wi = (bbox.Max.X - bbox.Min.X) * 12.0
                        hi = (bbox.Max.Y - bbox.Min.Y) * 12.0
                        if wi > 1 and hi > 1:
                            return (wi, hi)
        except Exception:
            pass
    return (17.0, 11.0)


def get_titleblock_params(tb_type_id):
    """Return sorted list of writable string parameter names for a titleblock type."""
    common = ["Drawn By", "Checked By", "Date", "Approved By"]
    for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
        try:
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                if tb.GetTypeId().Value == tb_type_id.Value:
                    params = []
                    for p in tb.GetOrderedParameters():
                        if p.IsReadOnly:
                            continue
                        if p.StorageType == StorageType.String:
                            params.append(p.Definition.Name)
                    if params:
                        return sorted(set(params))
        except Exception:
            pass
    return common


def get_sheet_browser_org_info():
    """Return (param_name, [existing_values]) for the first writable string
    parameter on sheets that looks like a browser organization/folder field.

    Searches existing sheets for parameters whose name contains keywords like
    'folder', 'group', 'discipline', 'package', 'set', or 'area'.
    Falls back to (None, []) if nothing is found.
    """
    FOLDER_KEYWORDS = {"folder", "group", "discipline", "package",
                       "set", "area", "category", "sub"}
    param_name = None
    values = set()

    try:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            for p in sheet.GetOrderedParameters():
                try:
                    if p.IsReadOnly:
                        continue
                    if p.StorageType != StorageType.String:
                        continue
                    nm = p.Definition.Name.lower()
                    if not any(kw in nm for kw in FOLDER_KEYWORDS):
                        continue
                    # Found a candidate
                    if param_name is None:
                        param_name = p.Definition.Name
                    if p.Definition.Name == param_name:
                        val = p.AsString()
                        if val and val.strip():
                            values.add(val.strip())
                except Exception:
                    pass
    except Exception:
        pass

    return (param_name, sorted(values))


def get_titleblock_margins(tb_type_id):
    """Return dict with title block visual zones in inches:
      margin_l, margin_r, margin_t, margin_b  -- print margins
      title_h   -- height of the bottom title strip
      revision_w -- width of the right revision block (0 if none)

    Tries to read known parameters from a placed instance.
    Falls back to conservative typical D-size defaults.
    """
    defaults = {
        "margin_l": 1.5, "margin_r": 0.5,
        "margin_t": 0.5, "margin_b": 0.5,
        "title_h":  2.0, "revision_w": 0.0,
    }
    # Try to read from a placed titleblock instance
    MARGIN_PARAMS = ["Left Margin", "Margin Left", "Binding Margin",
                     "Right Margin", "Margin Right",
                     "Top Margin", "Margin Top",
                     "Bottom Margin", "Margin Bottom",
                     "Title Strip Height", "Title Height", "Title Bar Height"]
    try:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                if tb.GetTypeId().Value != tb_type_id.Value:
                    continue
                # Try to find margin params -- names vary widely by firm
                found = {}
                for p in tb.GetOrderedParameters():
                    try:
                        nm = p.Definition.Name
                        if p.StorageType.ToString() == "Double" and not p.IsReadOnly:
                            val_ft = p.AsDouble()
                            val_in = val_ft * 12.0
                            nm_lo = nm.lower()
                            if "left" in nm_lo and "margin" in nm_lo:
                                found["margin_l"] = val_in
                            elif "right" in nm_lo and "margin" in nm_lo:
                                found["margin_r"] = val_in
                            elif "top" in nm_lo and "margin" in nm_lo:
                                found["margin_t"] = val_in
                            elif "bottom" in nm_lo and "margin" in nm_lo:
                                found["margin_b"] = val_in
                            elif "title" in nm_lo and ("height" in nm_lo or "strip" in nm_lo or "bar" in nm_lo):
                                found["title_h"] = val_in
                    except Exception:
                        pass
                if found:
                    defaults.update(found)
                return defaults
    except Exception:
        pass
    return defaults
    """Return ElementId of a Section-family ViewFamilyType."""
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Section:
            return vft.Id
    return ElementId.InvalidElementId


def next_assembly_name():
    """Auto-generate a unique name like Assembly-001."""
    existing = set()
    for elem in (FilteredElementCollector(doc)
                 .OfClass(AssemblyInstance)
                 .WhereElementIsNotElementType()):
        try:
            existing.add(elem.AssemblyTypeName.lower())
        except Exception:
            pass
    i = 1
    while True:
        name = "Assembly-{:03d}".format(i)
        if name.lower() not in existing:
            return name
        i += 1


def next_sheet_number():
    """Auto-generate next available sheet number like A-001."""
    existing = set()
    for sheet in (FilteredElementCollector(doc)
                  .OfClass(ViewSheet)
                  .WhereElementIsNotElementType()):
        try:
            existing.add(sheet.SheetNumber.lower())
        except Exception:
            pass
    i = 1
    while True:
        num = "A-{:03d}".format(i)
        if num.lower() not in existing:
            return num
        i += 1


# ─────────────────────────────────────────────────────────────────────
#  PRESET I/O
# ─────────────────────────────────────────────────────────────────────

def _ensure_presets_dir():
    if not os.path.exists(PRESETS_DIR):
        os.makedirs(PRESETS_DIR)


def list_presets():
    _ensure_presets_dir()
    return sorted(f[:-5] for f in os.listdir(PRESETS_DIR) if f.endswith(".json"))


def load_preset(name):
    path = os.path.join(PRESETS_DIR, name + ".json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as fp:
            return json.load(fp)
    except Exception:
        return None


def save_preset(name, data):
    _ensure_presets_dir()
    with open(os.path.join(PRESETS_DIR, name + ".json"), "w") as fp:
        json.dump(data, fp, indent=2)


def delete_preset(name):
    path = os.path.join(PRESETS_DIR, name + ".json")
    if os.path.exists(path):
        os.remove(path)


# ─────────────────────────────────────────────────────────────────────
#  ASSEMBLY CREATION
# ─────────────────────────────────────────────────────────────────────

def _try_create_section(assembly_id, orientation, section_type_id):
    """Create a detail section view, trying 4-arg then 3-arg API signature."""
    if orientation is None:
        raise ValueError("Orientation enum value could not be resolved.")
    try:
        return AssemblyViewUtils.CreateDetailSection(
            doc, assembly_id, orientation, section_type_id)
    except Exception:
        return AssemblyViewUtils.CreateDetailSection(
            doc, assembly_id, orientation)


def create_assembly(config):
    """Create the assembly instance and all requested views.

    config keys:
      elements      list[Element]
      name          str
      level_id      ElementId
      views         dict {key: bool}  3d top front back left right schedule
      templates     dict {key: ElementId}  same keys

    Returns (assembly_or_None, [(label, view), ...], [error_str, ...])
    """
    errors = []
    created_views = []
    elem_ids = List[ElementId]([e.Id for e in config["elements"]])
    section_type_id = get_section_view_type_id()

    t = Transaction(doc, "Assembly Studio - Create Assembly")
    t.Start()
    try:
        assembly = AssemblyInstance.Create(doc, elem_ids, config["level_id"])

        # Set name
        try:
            assembly.AssemblyTypeName = config["name"]
        except Exception:
            try:
                p = assembly.get_Parameter(BuiltInParameter.ASSEMBLY_NAME)
                if p and not p.IsReadOnly:
                    p.Set(config["name"])
            except Exception:
                errors.append("Could not set assembly name.")

        vk    = config.get("views", {})
        tmpls = config.get("templates", {})

        def apply_tmpl(view, key):
            tid = tmpls.get(key, ElementId.InvalidElementId)
            if tid and tid != ElementId.InvalidElementId:
                try:
                    view.ViewTemplateId = tid
                except Exception as ex:
                    errors.append("Template skipped for {}: {}".format(key, ex))

        if vk.get("3d"):
            try:
                v = AssemblyViewUtils.Create3DOrthographic(doc, assembly.Id)
                apply_tmpl(v, "3d")
                created_views.append(("3D Orthographic", v))
            except Exception as ex:
                errors.append("3D view: {}".format(ex))

        if vk.get("top"):
            try:
                v = _try_create_section(assembly.Id, ORIENT_TOP, section_type_id)
                apply_tmpl(v, "top")
                created_views.append(("Top (Plan)", v))
            except Exception as ex:
                errors.append("Top view: {}".format(ex))

        if vk.get("front"):
            try:
                v = _try_create_section(assembly.Id, ORIENT_FRONT, section_type_id)
                apply_tmpl(v, "front")
                created_views.append(("Front Elevation", v))
            except Exception as ex:
                errors.append("Front view: {}".format(ex))

        if vk.get("back"):
            try:
                v = _try_create_section(assembly.Id, ORIENT_BACK, section_type_id)
                apply_tmpl(v, "back")
                created_views.append(("Back Elevation", v))
            except Exception as ex:
                errors.append("Back view: {}".format(ex))

        if vk.get("left"):
            try:
                v = _try_create_section(assembly.Id, ORIENT_LEFT, section_type_id)
                apply_tmpl(v, "left")
                created_views.append(("Left Elevation", v))
            except Exception as ex:
                errors.append("Left view: {}".format(ex))

        if vk.get("right"):
            try:
                v = _try_create_section(assembly.Id, ORIENT_RIGHT, section_type_id)
                apply_tmpl(v, "right")
                created_views.append(("Right Elevation", v))
            except Exception as ex:
                errors.append("Right view: {}".format(ex))

        if vk.get("schedule"):
            try:
                v = AssemblyViewUtils.CreatePartList(doc, assembly.Id)
                apply_tmpl(v, "schedule")
                created_views.append(("Parts List", v))
            except Exception as ex:
                errors.append("Parts list: {}".format(ex))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        return None, [], ["Fatal: {}".format(ex)]

    return assembly, created_views, errors


# ─────────────────────────────────────────────────────────────────────
#  SHEET CREATION
# ─────────────────────────────────────────────────────────────────────

def create_sheet(config):
    """Create a ViewSheet with manually placed viewports.

    config keys:
      title_block_id      ElementId
      sheet_number        str
      sheet_name          str
      slots               list of:
                            view_id (ElementId), x (ft), y (ft),
                            scale (int|None), viewport_type_id (ElementId),
                            is_schedule (bool)
      parameters          dict {param_name: value_str}
      destination         "assembly" | "sheets"  (default "sheets")
      assembly_id         ElementId  (required when destination="assembly")
      browser_folder_param  str | None  -- parameter name to set
      browser_folder_value  str | None  -- value to write into that param

    Returns (sheet_or_None, [error_str, ...])
    """
    errors  = []
    dest    = config.get("destination", "sheets")
    asm_id  = config.get("assembly_id")

    t = Transaction(doc, "Assembly Studio - Create Sheet")
    t.Start()
    try:
        # ── Create the sheet ──────────────────────────────────────────
        if dest == "assembly" and asm_id:
            # AssemblyViewUtils.CreateSheet associates the sheet with the
            # assembly so it appears nested under it in the project browser.
            # It auto-places viewports which we immediately delete.
            sheet = AssemblyViewUtils.CreateSheet(
                doc, asm_id, config["title_block_id"])

            # Remove all auto-placed viewports and schedule instances
            auto_vps = list(FilteredElementCollector(doc, sheet.Id)
                            .OfClass(Viewport).ToElements())
            auto_ssi = list(FilteredElementCollector(doc, sheet.Id)
                            .OfClass(ScheduleSheetInstance).ToElements())
            for el in auto_vps + auto_ssi:
                try:
                    doc.Delete(el.Id)
                except Exception:
                    pass
            doc.Regenerate()
        else:
            sheet = ViewSheet.Create(doc, config["title_block_id"])

        sheet.SheetNumber = config["sheet_number"]
        sheet.Name        = config["sheet_name"]

        # ── Place viewports ───────────────────────────────────────────
        for slot in config["slots"]:
            vid = slot.get("view_id")
            if not vid:
                continue
            pos = XYZ(slot["x"], slot["y"], 0)

            if slot.get("is_schedule"):
                try:
                    ScheduleSheetInstance.Create(doc, sheet.Id, vid, pos)
                except Exception as ex:
                    errors.append("Schedule placement: {}".format(ex))
            else:
                try:
                    can_add = Viewport.CanAddViewToSheet(doc, sheet.Id, vid)
                except Exception:
                    can_add = True
                if not can_add:
                    view = doc.GetElement(vid)
                    vname = view.Name if view else str(vid)
                    errors.append("'{}' already on another sheet".format(vname))
                    continue
                try:
                    vp = Viewport.Create(doc, sheet.Id, vid, pos)
                    vp_type_id = slot.get("viewport_type_id")
                    if vp_type_id and vp_type_id != ElementId.InvalidElementId:
                        try:
                            vp.ChangeTypeId(vp_type_id)
                        except Exception as ex:
                            errors.append("VP type: {}".format(ex))
                    scale = slot.get("scale")
                    if scale:
                        view = doc.GetElement(vid)
                        try:
                            view.Scale = int(scale)
                        except Exception as ex:
                            errors.append("Scale: {}".format(ex))
                except Exception as ex:
                    errors.append("Viewport: {}".format(ex))

        # ── Set titleblock parameters ─────────────────────────────────
        params = config.get("parameters", {})
        if params:
            try:
                for tb in (FilteredElementCollector(doc, sheet.Id)
                           .OfCategory(BuiltInCategory.OST_TitleBlocks)
                           .ToElements()):
                    for p in tb.GetOrderedParameters():
                        pname = p.Definition.Name
                        if pname in params and not p.IsReadOnly:
                            val = params[pname]
                            if val:
                                try:
                                    p.Set(val)
                                except Exception as ex:
                                    errors.append("Param '{}': {}".format(pname, ex))
            except Exception as ex:
                errors.append("Titleblock params: {}".format(ex))

        # ── Set browser folder parameter (Sheets destination only) ────
        fp_name  = config.get("browser_folder_param")
        fp_value = config.get("browser_folder_value", "")
        if dest == "sheets" and fp_name and fp_value:
            try:
                p = sheet.LookupParameter(fp_name)
                if p and not p.IsReadOnly and p.StorageType == StorageType.String:
                    p.Set(fp_value)
            except Exception as ex:
                errors.append("Browser folder: {}".format(ex))

        t.Commit()

    except Exception as ex:
        t.RollBack()
        return None, ["Fatal: {}".format(ex)]

    return sheet, errors



def get_titleblock_sheet_origin(tb_type_id):
    """Return (min_x_ft, min_y_ft) -- the bottom-left corner of the placed
    titleblock in Revit sheet coordinates.

    This is the offset needed to convert from 'inches from paper corner'
    to Revit sheet feet.  For most families this is a small negative value
    (the titleblock bleeds slightly past the coordinate origin), but it can
    be zero or positive depending on the family.

    Falls back to (0, 0) if no placed instance is found.
    """
    try:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                try:
                    if tb.GetTypeId().Value != tb_type_id.Value:
                        continue
                    bbox = tb.get_BoundingBox(sheet)
                    if bbox:
                        return (bbox.Min.X, bbox.Min.Y)
                except Exception:
                    pass
    except Exception:
        pass
    return (0.0, 0.0)


def get_titleblock_geometry_lines(tb_type_id, sheet_h_in):
    """Extract the titleblock's actual line geometry from a placed instance.

    Returns list of (x1_px, y1_px, x2_px, y2_px) in canvas pixels.
    Canvas origin is top-left; Y is flipped from Revit's bottom-up sheet coords.
    Returns [] if no placed instance is found.
    """
    lines = []
    try:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            tb_inst = None
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                if tb.GetTypeId().Value == tb_type_id.Value:
                    tb_inst = tb
                    break
            if tb_inst is None:
                continue

            opts = RvtOptions()
            opts.ComputeReferences = False
            geom = tb_inst.get_Geometry(opts)
            if geom is None:
                break

            def _extract(geom_objs):
                for obj in geom_objs:
                    if isinstance(obj, GeometryInstance):
                        _extract(obj.GetInstanceGeometry())
                    elif isinstance(obj, RvtDB.Line):
                        p1 = obj.GetEndPoint(0)
                        p2 = obj.GetEndPoint(1)
                        x1 = p1.X * 12.0 * PX_PER_INCH
                        y1 = (sheet_h_in - p1.Y * 12.0) * PX_PER_INCH
                        x2 = p2.X * 12.0 * PX_PER_INCH
                        y2 = (sheet_h_in - p2.Y * 12.0) * PX_PER_INCH
                        lines.append((x1, y1, x2, y2))
                    elif isinstance(obj, RvtDB.Arc):
                        # Tessellate arcs into short line segments
                        try:
                            pts = obj.Tessellate()
                            for i in range(len(pts) - 1):
                                p1 = pts[i]; p2 = pts[i + 1]
                                lines.append((
                                    p1.X * 12.0 * PX_PER_INCH,
                                    (sheet_h_in - p1.Y * 12.0) * PX_PER_INCH,
                                    p2.X * 12.0 * PX_PER_INCH,
                                    (sheet_h_in - p2.Y * 12.0) * PX_PER_INCH,
                                ))
                        except Exception:
                            pass

            _extract(geom)
            break  # Only need one placed instance

    except Exception:
        pass
    return lines


def get_titleblock_drawing_area_inches(tb_type_id):
    """Infer the usable drawing area by scanning placed viewport bounding boxes
    on sheets that use this titleblock type.

    Returns (x_in, y_in, w_in, h_in) or None.
    This is the most reliable way to determine the printable region --
    we look at where Revit actually places content on real sheets.
    """
    min_x = min_y =  1e9
    max_x = max_y = -1e9
    found = False

    try:
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            has_tb = False
            for tb in (FilteredElementCollector(doc, sheet.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .ToElements()):
                if tb.GetTypeId().Value == tb_type_id.Value:
                    has_tb = True
                    break
            if not has_tb:
                continue

            for vp in (FilteredElementCollector(doc, sheet.Id)
                       .OfClass(Viewport).ToElements()):
                try:
                    bbox = vp.get_BoundingBox(sheet)
                    if bbox:
                        min_x = min(min_x, bbox.Min.X * 12.0)
                        min_y = min(min_y, bbox.Min.Y * 12.0)
                        max_x = max(max_x, bbox.Max.X * 12.0)
                        max_y = max(max_y, bbox.Max.Y * 12.0)
                        found = True
                except Exception:
                    pass

            for ssi in (FilteredElementCollector(doc, sheet.Id)
                        .OfClass(ScheduleSheetInstance).ToElements()):
                try:
                    bbox = ssi.get_BoundingBox(sheet)
                    if bbox:
                        min_x = min(min_x, bbox.Min.X * 12.0)
                        min_y = min(min_y, bbox.Min.Y * 12.0)
                        max_x = max(max_x, bbox.Max.X * 12.0)
                        max_y = max(max_y, bbox.Max.Y * 12.0)
                        found = True
                except Exception:
                    pass

    except Exception:
        pass

    if found and max_x > min_x and max_y > min_y:
        pad = 0.25  # 1/4" padding
        return {
            "x_in": min_x - pad,
            "y_in": min_y - pad,
            "w_in": (max_x - min_x) + pad * 2,
            "h_in": (max_y - min_y) + pad * 2,
        }
    return None


def get_assembly_bbox_inches(assembly_id):
    """Return (x_span_in, y_span_in, z_span_in) of the assembly bounding box."""
    try:
        assembly = doc.GetElement(assembly_id)
        if assembly is None:
            return None
        bbox = assembly.get_BoundingBox(None)
        if bbox:
            return (
                abs(bbox.Max.X - bbox.Min.X) * 12.0,
                abs(bbox.Max.Y - bbox.Min.Y) * 12.0,
                abs(bbox.Max.Z - bbox.Min.Z) * 12.0,
            )
    except Exception:
        pass
    return None


def get_schedule_sheet_size_px(view_id):
    """Return (w_px, h_px) for a schedule by measuring its placed bounding box.

    Searches every sheet for a ScheduleSheetInstance of this ViewSchedule.
    Falls back to a 4" x 3" estimate if none is found.
    """
    try:
        v = doc.GetElement(view_id)
        vid_int = view_id.Value
        for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
            for ssi in (FilteredElementCollector(doc, sheet.Id)
                        .OfClass(ScheduleSheetInstance).ToElements()):
                try:
                    if ssi.ScheduleId.Value == vid_int:
                        bbox = ssi.get_BoundingBox(sheet)
                        if bbox:
                            w_px = abs(bbox.Max.X - bbox.Min.X) * 12.0 * PX_PER_INCH
                            h_px = abs(bbox.Max.Y - bbox.Min.Y) * 12.0 * PX_PER_INCH
                            return (w_px, h_px)
                except Exception:
                    pass
    except Exception:
        pass
    return (4.0 * PX_PER_INCH, 3.0 * PX_PER_INCH)




def get_view_sheet_size_px(view_id, scale_factor, assembly_id=None):
    """Return (w_px, h_px) for a view at the given scale denominator.

    Priority:
      1. Placed viewport bounding box on any existing sheet -- the most
         accurate source, reflects actual rendered size including annotation
         margins. Rescaled from placed scale to requested scale_factor.
      2. View CropBox converted at scale_factor.
      3. Assembly BoundingBox projected per view orientation.
    Returns None if nothing is available.
    """
    if scale_factor <= 0:
        return None

    try:
        v = doc.GetElement(view_id)
        if v is None or isinstance(v, ViewSchedule):
            return None

        vid_int = view_id.Value

        # ── Priority 1: Placed viewport bbox ─────────────────────────
        try:
            for sheet in FilteredElementCollector(doc).OfClass(ViewSheet):
                for vp in (FilteredElementCollector(doc, sheet.Id)
                           .OfClass(Viewport).ToElements()):
                    try:
                        if vp.ViewId.Value != vid_int:
                            continue
                        bbox = vp.get_BoundingBox(sheet)
                        if bbox is None:
                            continue
                        placed_w_in = abs(bbox.Max.X - bbox.Min.X) * 12.0
                        placed_h_in = abs(bbox.Max.Y - bbox.Min.Y) * 12.0
                        if placed_w_in <= 0 or placed_h_in <= 0:
                            continue
                        # Rescale from placed scale to requested scale
                        placed_scale = 1
                        try:    placed_scale = max(1, int(v.Scale))
                        except: pass
                        ratio = float(placed_scale) / float(scale_factor)
                        return (placed_w_in * ratio * PX_PER_INCH,
                                placed_h_in * ratio * PX_PER_INCH)
                    except Exception:
                        pass
        except Exception:
            pass

        # ── Priority 2: CropBox ───────────────────────────────────────
        try:
            if v.CropBoxActive:
                crop = v.CropBox
                w_ft = abs(crop.Max.X - crop.Min.X)
                h_ft = abs(crop.Max.Y - crop.Min.Y)
                if w_ft > 0 and h_ft > 0:
                    return ((w_ft / float(scale_factor)) * 12.0 * PX_PER_INCH,
                            (h_ft / float(scale_factor)) * 12.0 * PX_PER_INCH)
        except Exception:
            pass

        # ── Priority 3: Assembly BoundingBox projection ───────────────
        if assembly_id:
            spans = get_assembly_bbox_inches(assembly_id)
            if spans:
                x_in, y_in, z_in = spans
                name_lo = (v.Name or "").lower()
                if "top" in name_lo or "plan" in name_lo:
                    w_in, h_in = x_in, y_in
                elif "front" in name_lo:
                    w_in, h_in = x_in, z_in
                elif "back" in name_lo:
                    w_in, h_in = x_in, z_in
                elif "left" in name_lo:
                    w_in, h_in = y_in, z_in
                elif "right" in name_lo:
                    w_in, h_in = y_in, z_in
                elif "3d" in name_lo or "orthographic" in name_lo:
                    w_in = math.sqrt(x_in ** 2 + y_in ** 2) * 0.75
                    h_in = z_in * 1.1
                else:
                    w_in, h_in = max(x_in, y_in), z_in
                w_in *= 1.2;  h_in *= 1.2
                return ((w_in / float(scale_factor)) * PX_PER_INCH,
                        (h_in / float(scale_factor)) * PX_PER_INCH)

    except Exception:
        pass
    return None






# ─────────────────────────────────────────────────────────────────────
#  USABLE AREA PICKER WINDOW
# ─────────────────────────────────────────────────────────────────────

class UsableAreaPickerWindow(Window):
    """Modal popup showing the real titleblock geometry.

    Layers (bottom to top):
      1. Actual titleblock line geometry from Revit (curves extracted from
         the placed FamilyInstance -- NOT an image export)
      2. Margin zone tints (title strip, margins, revision block)
      3. Ghost view footprints at selected scales (so you can see fit)
      4. User-drawn usable area rectangle
      5. Sheet border

    Footer contains editable margin inputs so the user can correct zones
    if parameter sniffing missed the right values.

    Result: dict {x_in, y_in, w_in, h_in} in sheet inches, or None.
    """

    def __init__(self, sheet_w, sheet_h, tb_margins,
                 tb_lines=None, view_ghosts=None, suggested_area=None):
        """
        tb_lines:       list of (x1_px, y1_px, x2_px, y2_px) canvas-pixel lines
        view_ghosts:    list of (vname, vkey, w_px, h_px)
        suggested_area: dict {x_in, y_in, w_in, h_in} from viewport scan, or None
        """
        Window.__init__(self)
        self._sheet_w  = sheet_w
        self._sheet_h  = sheet_h
        self._tb_lines = tb_lines or []
        self._view_ghosts = view_ghosts or []
        self._suggested   = suggested_area

        mb = tb_margins or {}
        self._ml  = mb.get("margin_l",   1.5)
        self._mr  = mb.get("margin_r",   0.5)
        self._mt  = mb.get("margin_t",   0.5)
        self._mb  = mb.get("margin_b",   0.5)
        self._th  = mb.get("title_h",    2.0)
        self._rw  = mb.get("revision_w", 0.0)

        self._result     = None
        self._dragging   = False
        self._drag_start = None
        self._area_rect  = None

        self._build_ui()
        self._redraw()

        # If we have a suggested area from viewport scan, pre-draw it
        if self._suggested:
            self._apply_suggested()

    # ── Build UI ────────────────────────────────────────────────────

    def _build_ui(self):
        import System.Windows.Media as WMedia

        self.Title = "Set Usable Area -- Drag to define viewport region"
        self.Width  = 1000
        self.Height = 700
        self.MinWidth  = 800
        self.MinHeight = 560
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(240, 242, 247))
        self.FontFamily = WMedia.FontFamily("Segoe UI")

        outer = Grid()
        for h_val, unit in [
            (50, System.Windows.GridUnitType.Pixel),   # header
            (1,  System.Windows.GridUnitType.Star),    # canvas
            (48, System.Windows.GridUnitType.Pixel),   # margin inputs
            (50, System.Windows.GridUnitType.Pixel),   # footer
        ]:
            rd = System.Windows.Controls.RowDefinition()
            rd.Height = System.Windows.GridLength(h_val, unit)
            outer.RowDefinitions.Add(rd)
        self.Content = outer

        # ── Header ──────────────────────────────────────────────────
        hdr = Border()
        hdr.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
        Grid.SetRow(hdr, 0)
        hdr_g = Grid()
        hdr_g.Margin = Thickness(14, 0, 14, 0)
        hdr.Child = hdr_g

        t1 = TextBlock()
        t1.Text = "Set Usable Area"
        t1.Foreground = Brushes.White
        t1.FontSize = 14
        t1.FontWeight = System.Windows.FontWeights.SemiBold
        t1.VerticalAlignment = VerticalAlignment.Center
        hdr_g.Children.Add(t1)

        hint = TextBlock()
        has_tb  = "Real titleblock drawn" if self._tb_lines else "No placed sheet found -- zone estimate only"
        has_vp  = " | {} views ghosted".format(len(self._view_ghosts)) if self._view_ghosts else ""
        hint.Text = has_tb + has_vp
        hint.Foreground = SolidColorBrush(Color.FromRgb(127, 168, 204))
        hint.FontSize = 10
        hint.VerticalAlignment = VerticalAlignment.Center
        hint.HorizontalAlignment = HorizontalAlignment.Right
        hdr_g.Children.Add(hint)
        outer.Children.Add(hdr)

        # ── Canvas ──────────────────────────────────────────────────
        cb = Border()
        cb.Background = SolidColorBrush(Color.FromRgb(210, 214, 220))
        cb.Margin = Thickness(10, 6, 10, 4)
        Grid.SetRow(cb, 1)

        vb = System.Windows.Controls.Viewbox()
        vb.Stretch = WMedia.Stretch.Uniform
        vb.Margin = Thickness(6)

        cw = self._sheet_w * PX_PER_INCH
        ch = self._sheet_h * PX_PER_INCH
        self._canvas = Canvas()
        self._canvas.Width  = cw
        self._canvas.Height = ch
        self._canvas.Background = Brushes.White
        self._canvas.ClipToBounds = True
        self._canvas.Cursor = System.Windows.Input.Cursors.Cross

        self._canvas.MouseLeftButtonDown += self._on_mouse_down
        self._canvas.MouseMove           += self._on_mouse_move
        self._canvas.MouseLeftButtonUp   += self._on_mouse_up

        vb.Child = self._canvas
        cb.Child = vb
        outer.Children.Add(cb)

        # ── Margin input strip ───────────────────────────────────────
        margin_bar = Border()
        margin_bar.Background = SolidColorBrush(Color.FromRgb(248, 249, 251))
        margin_bar.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
        margin_bar.BorderThickness = Thickness(0, 1, 0, 1)
        Grid.SetRow(margin_bar, 2)

        msp = StackPanel()
        msp.Orientation = Orientation.Horizontal
        msp.VerticalAlignment = VerticalAlignment.Center
        msp.Margin = Thickness(10, 0, 10, 0)
        margin_bar.Child = msp

        def _lbl(t):
            tb = TextBlock()
            tb.Text = t
            tb.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
            tb.FontSize = 11
            tb.VerticalAlignment = VerticalAlignment.Center
            tb.Margin = Thickness(8, 0, 3, 0)
            return tb

        def _field(val, width=48):
            tb = TextBox()
            tb.Text = "{:.2f}".format(val)
            tb.Width = width
            tb.FontSize = 11
            tb.Padding = Thickness(4, 2, 4, 2)
            tb.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
            tb.BorderThickness = Thickness(1)
            tb.VerticalAlignment = VerticalAlignment.Center
            return tb

        header_tb = TextBlock()
        header_tb.Text = "Zone margins (inches):"
        header_tb.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
        header_tb.FontSize = 11
        header_tb.FontWeight = System.Windows.FontWeights.SemiBold
        header_tb.VerticalAlignment = VerticalAlignment.Center
        msp.Children.Add(header_tb)

        msp.Children.Add(_lbl("Left"))
        self._f_ml = _field(self._ml); msp.Children.Add(self._f_ml)
        msp.Children.Add(_lbl("Right"))
        self._f_mr = _field(self._mr); msp.Children.Add(self._f_mr)
        msp.Children.Add(_lbl("Top"))
        self._f_mt = _field(self._mt); msp.Children.Add(self._f_mt)
        msp.Children.Add(_lbl("Bottom"))
        self._f_mb = _field(self._mb); msp.Children.Add(self._f_mb)
        msp.Children.Add(_lbl("Title strip"))
        self._f_th = _field(self._th); msp.Children.Add(self._f_th)

        redraw_btn = System.Windows.Controls.Button()
        redraw_btn.Content = "Redraw Zones"
        redraw_btn.Height = 28
        redraw_btn.Padding = Thickness(10, 0, 10, 0)
        redraw_btn.Background = Brushes.White
        redraw_btn.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
        redraw_btn.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        redraw_btn.BorderThickness = Thickness(1)
        redraw_btn.FontSize = 11
        redraw_btn.Margin = Thickness(12, 0, 6, 0)
        redraw_btn.VerticalAlignment = VerticalAlignment.Center
        redraw_btn.Click += self._on_redraw
        msp.Children.Add(redraw_btn)

        if self._suggested:
            auto_btn = System.Windows.Controls.Button()
            auto_btn.Content = "Auto-fill from existing sheets"
            auto_btn.Height = 28
            auto_btn.Padding = Thickness(10, 0, 10, 0)
            auto_btn.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
            auto_btn.Foreground = Brushes.White
            auto_btn.BorderThickness = Thickness(0)
            auto_btn.FontSize = 11
            auto_btn.Margin = Thickness(0, 0, 0, 0)
            auto_btn.VerticalAlignment = VerticalAlignment.Center
            auto_btn.Click += self._on_auto
            msp.Children.Add(auto_btn)

        outer.Children.Add(margin_bar)

        # ── Footer ──────────────────────────────────────────────────
        footer = Border()
        footer.Background = SolidColorBrush(Color.FromRgb(250, 250, 250))
        footer.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
        footer.BorderThickness = Thickness(0, 1, 0, 0)
        Grid.SetRow(footer, 3)

        fg = Grid()
        fg.Margin = Thickness(14, 0, 14, 0)
        footer.Child = fg

        self._info_lbl = TextBlock()
        self._info_lbl.Text = "Drag on the sheet to define usable area"
        self._info_lbl.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
        self._info_lbl.FontSize = 10
        self._info_lbl.VerticalAlignment = VerticalAlignment.Center
        fg.Children.Add(self._info_lbl)

        bsp = StackPanel()
        bsp.Orientation = Orientation.Horizontal
        bsp.HorizontalAlignment = HorizontalAlignment.Right
        bsp.VerticalAlignment = VerticalAlignment.Center

        cancel_btn = System.Windows.Controls.Button()
        cancel_btn.Content = "Cancel"
        cancel_btn.Width = 90; cancel_btn.Height = 32
        cancel_btn.Background = Brushes.White
        cancel_btn.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
        cancel_btn.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        cancel_btn.BorderThickness = Thickness(1)
        cancel_btn.FontSize = 12; cancel_btn.Margin = Thickness(0, 0, 8, 0)
        cancel_btn.Click += self._on_cancel
        bsp.Children.Add(cancel_btn)

        self._ok_btn = System.Windows.Controls.Button()
        self._ok_btn.Content = "Apply Area"
        self._ok_btn.Width = 110; self._ok_btn.Height = 32
        self._ok_btn.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
        self._ok_btn.Foreground = Brushes.White
        self._ok_btn.BorderThickness = Thickness(0)
        self._ok_btn.FontSize = 12
        self._ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
        self._ok_btn.IsEnabled = False
        self._ok_btn.Click += self._on_ok
        bsp.Children.Add(self._ok_btn)

        fg.Children.Add(bsp)
        outer.Children.Add(footer)

    # ── Drawing ──────────────────────────────────────────────────────

    def _read_margin_fields(self):
        """Parse current margin field values; keep existing on parse error."""
        def _parse(field, current):
            try:    return float(field.Text.strip())
            except: return current
        self._ml = _parse(self._f_ml, self._ml)
        self._mr = _parse(self._f_mr, self._mr)
        self._mt = _parse(self._f_mt, self._mt)
        self._mb = _parse(self._f_mb, self._mb)
        self._th = _parse(self._f_th, self._th)

    def _redraw(self):
        """Clear and repaint the canvas: geometry + zones + ghost views."""
        canvas = self._canvas
        w = canvas.Width
        h = canvas.Height

        # Remove everything except the user's drawn area rect and the sheet border
        to_remove = [c for c in canvas.Children
                     if getattr(c, "Tag", None) not in ("area_rect",)]
        for c in to_remove:
            canvas.Children.Remove(c)

        def _add(e): canvas.Children.Add(e)

        def _rect(x, y, rw, rh, fill=None, stroke=None, thick=0):
            r = Rectangle()
            r.Width = rw; r.Height = rh
            r.IsHitTestVisible = False
            if fill:   r.Fill   = fill
            if stroke: r.Stroke = stroke; r.StrokeThickness = thick
            Canvas.SetLeft(r, x); Canvas.SetTop(r, y)
            _add(r)

        def _line(x1, y1, x2, y2, stroke, thick=1.0):
            ln = Line()
            ln.X1 = x1; ln.Y1 = y1; ln.X2 = x2; ln.Y2 = y2
            ln.Stroke = stroke; ln.StrokeThickness = thick
            ln.IsHitTestVisible = False; _add(ln)

        def _text(t, x, y, size, color, bold=False, wrap=False):
            tb = TextBlock()
            tb.Text = t; tb.FontSize = size
            tb.Foreground = color
            if bold: tb.FontWeight = System.Windows.FontWeights.SemiBold
            if wrap: tb.TextWrapping = System.Windows.TextWrapping.Wrap
            tb.IsHitTestVisible = False
            Canvas.SetLeft(tb, x); Canvas.SetTop(tb, y); _add(tb)

        # ── 1: White background ──────────────────────────────────
        _rect(0, 0, w, h, fill=Brushes.White)

        # ── 2: 1" reference grid ─────────────────────────────────
        grid_c = SolidColorBrush(Color.FromRgb(228, 231, 237))
        for i in range(1, int(w / PX_PER_INCH) + 1):
            _line(i * PX_PER_INCH, 0, i * PX_PER_INCH, h, grid_c, 0.8)
        for i in range(1, int(h / PX_PER_INCH) + 1):
            _line(0, i * PX_PER_INCH, w, i * PX_PER_INCH, grid_c, 0.8)

        # ── 3: Real titleblock geometry ───────────────────────────
        tb_stroke = SolidColorBrush(Color.FromRgb(60, 65, 75))
        for (x1, y1, x2, y2) in self._tb_lines:
            _line(x1, y1, x2, y2, tb_stroke, thick=1.2)

        # ── 4: Zone fills (translucent, drawn OVER geometry) ─────
        ml = self._ml * PX_PER_INCH
        mr = self._mr * PX_PER_INCH
        mt = self._mt * PX_PER_INCH
        mb_px = self._mb * PX_PER_INCH
        th = self._th * PX_PER_INCH

        content_top    = mt
        content_bottom = h - mb_px - th
        content_h      = max(0.0, content_bottom - content_top)
        content_left   = ml
        content_right  = w - mr - (self._rw * PX_PER_INCH)
        content_w      = max(0.0, content_right - content_left)

        alpha = 45
        amber_f = SolidColorBrush(Color.FromArgb(alpha, 200, 160,  60))
        blue_f  = SolidColorBrush(Color.FromArgb(alpha,  80,  80, 200))
        green_f = SolidColorBrush(Color.FromArgb(18,     34, 197,  94))
        amber_l = SolidColorBrush(Color.FromRgb(180, 130,  40))
        blue_l  = SolidColorBrush(Color.FromRgb( 70,  70, 190))
        green_s = SolidColorBrush(Color.FromArgb(130,  34, 197,  94))
        amber_t = SolidColorBrush(Color.FromRgb(140, 100,  20))
        blue_t  = SolidColorBrush(Color.FromRgb( 50,  50, 170))
        green_t = SolidColorBrush(Color.FromArgb(160,  10, 150,  60))

        # Title strip
        _rect(0, content_bottom, w, th + mb_px, fill=amber_f)
        _line(0, content_bottom, w, content_bottom, amber_l, 2.0)
        _text('TITLE STRIP  ({:.2f}")'.format(self._th),
              content_left + 10, content_bottom + 5,
              max(12, min(20, th * 0.22)), amber_t, bold=True)

        # Margins
        if ml > 0:
            _rect(0, mt, ml, content_h, fill=blue_f)
            _line(ml, mt, ml, content_bottom, blue_l, 2.0)
            _text('{:.2f}"'.format(self._ml), 4,
                  content_top + content_h / 2.0 - 10, 14, blue_t)
        if mr > 0:
            _rect(w - mr, mt, mr, content_h, fill=blue_f)
            _line(w - mr, mt, w - mr, content_bottom, blue_l, 2.0)
            _text('{:.2f}"'.format(self._mr), w - mr + 3,
                  content_top + content_h / 2.0 - 10, 14, blue_t)
        if mt > 0:
            _rect(0, 0, w, mt, fill=blue_f)
            _line(0, mt, w, mt, blue_l, 1.5)

        # Drawing area highlight
        if content_w > 0 and content_h > 0:
            _rect(content_left, content_top, content_w, content_h,
                  fill=green_f, stroke=green_s, thick=2)
            _text('DRAWING AREA  {:.2f}" x {:.2f}"'.format(
                      content_w / PX_PER_INCH, content_h / PX_PER_INCH),
                  content_left + content_w / 2.0 - 260,
                  content_top  + content_h / 2.0 - 20,
                  26, green_t, bold=True)

        # ── 5: Ghost view footprints ──────────────────────────────
        if self._view_ghosts:
            ghost_x = content_left + 8
            ghost_y = content_top  + 8
            max_row_h = 0
            for vname, vkey, gw, gh in self._view_ghosts:
                # Clamp ghost to drawing area
                gw = min(gw, content_w - 8)
                gh = min(gh, content_h - 8)
                if gw <= 0 or gh <= 0:
                    continue
                # Wrap to next row
                if ghost_x + gw > content_right - 4:
                    ghost_x = content_left + 8
                    ghost_y += max_row_h + 6
                    max_row_h = 0
                if ghost_y + gh > content_bottom - 4:
                    break

                vc = VP_COLORS.get(vkey, VP_COLORS["other"])
                ghost_fill   = SolidColorBrush(Color.FromArgb(25, vc.R, vc.G, vc.B))
                ghost_stroke = SolidColorBrush(Color.FromArgb(140, vc.R, vc.G, vc.B))
                _rect(ghost_x, ghost_y, gw, gh,
                      fill=ghost_fill, stroke=ghost_stroke, thick=1.5)
                _text(vname, ghost_x + 4, ghost_y + 4, 14,
                      SolidColorBrush(Color.FromArgb(180, vc.R, vc.G, vc.B)))

                max_row_h = max(max_row_h, gh)
                ghost_x  += gw + 6

        # ── 6: Sheet border (always on top) ───────────────────────
        _rect(0, 0, w, h,
              stroke=SolidColorBrush(Color.FromRgb(30, 30, 30)), thick=3)

    def _apply_suggested(self):
        """Pre-draw the auto-detected area rectangle."""
        s = self._suggested
        x  = s["x_in"] * PX_PER_INCH
        y_revit = s["y_in"]
        # Revit Y is up; canvas Y is down
        y  = (self._sheet_h - y_revit - s["h_in"]) * PX_PER_INCH
        rw = s["w_in"] * PX_PER_INCH
        rh = s["h_in"] * PX_PER_INCH
        self._draw_area_rect(x, y, rw, rh)
        self._result = {
            "x_in": x  / PX_PER_INCH,
            "y_in": (self._sheet_h - (y + rh) / PX_PER_INCH),
            "w_in": rw / PX_PER_INCH,
            "h_in": rh / PX_PER_INCH,
        }
        self._ok_btn.IsEnabled = True
        self._info_lbl.Text = (
            'Auto-detected: {:.2f}" x {:.2f}" -- adjust by dragging, then Apply'
            .format(s["w_in"], s["h_in"]))

    def _draw_area_rect(self, x, y, rw, rh):
        """Create or replace the dashed user-area rectangle."""
        if self._area_rect is not None:
            try: self._canvas.Children.Remove(self._area_rect)
            except: pass
        _dash = DoubleCollection()
        _dash.Add(8.0); _dash.Add(4.0)
        rect = Rectangle()
        rect.Stroke = SolidColorBrush(Color.FromRgb(37, 99, 235))
        rect.StrokeThickness = 3
        rect.StrokeDashArray = _dash
        rect.Fill = SolidColorBrush(Color.FromArgb(30, 37, 99, 235))
        rect.Width = rw; rect.Height = rh
        rect.IsHitTestVisible = False
        rect.Tag = "area_rect"
        Canvas.SetLeft(rect, x); Canvas.SetTop(rect, y)
        self._canvas.Children.Add(rect)
        self._area_rect = rect

    # ── Events ───────────────────────────────────────────────────────

    def _on_redraw(self, sender, e):
        self._read_margin_fields()
        self._redraw()
        # Reposition the drawn area rect to match the updated margin boundaries.
        # This is the "snap to drawing area" behaviour -- after adjusting margins
        # the dashed rect should reflect exactly the open drawing region.
        w = self._canvas.Width
        h = self._canvas.Height
        ml  = self._ml * PX_PER_INCH
        mr  = self._mr * PX_PER_INCH
        mt  = self._mt * PX_PER_INCH
        mb  = self._mb * PX_PER_INCH
        th  = self._th * PX_PER_INCH
        rw  = self._rw * PX_PER_INCH
        x   = ml
        y   = mt
        rw_ = max(1.0, w - ml - mr - rw)
        rh  = max(1.0, h - mt - mb - th)
        self._draw_area_rect(x, y, rw_, rh)
        self._result = {
            "x_in": x   / PX_PER_INCH,
            "y_in": (h - y - rh) / PX_PER_INCH,  # canvas Y → Revit bottom-up
            "w_in": rw_ / PX_PER_INCH,
            "h_in": rh  / PX_PER_INCH,
        }
        self._ok_btn.IsEnabled = True
        self._info_lbl.Text = (
            'Zone area: {:.2f}" x {:.2f}" -- adjust by dragging, then Apply'
            .format(rw_ / PX_PER_INCH, rh / PX_PER_INCH)
        )

    def _on_auto(self, sender, e):
        if self._suggested:
            self._apply_suggested()

    def _on_mouse_down(self, sender, e):
        pos = e.GetPosition(self._canvas)
        sx = round(pos.X / SNAP_PX) * SNAP_PX
        sy = round(pos.Y / SNAP_PX) * SNAP_PX
        self._drag_start = (sx, sy)
        self._dragging   = True
        self._result = None
        self._ok_btn.IsEnabled = False
        self._draw_area_rect(sx, sy, 0, 0)
        self._canvas.CaptureMouse()
        e.Handled = True

    def _on_mouse_move(self, sender, e):
        if not self._dragging or self._drag_start is None:
            return
        pos = e.GetPosition(self._canvas)
        ex = round(pos.X / SNAP_PX) * SNAP_PX
        ey = round(pos.Y / SNAP_PX) * SNAP_PX
        sx, sy = self._drag_start
        x  = min(sx, ex); y  = min(sy, ey)
        rw = abs(ex - sx); rh = abs(ey - sy)
        if self._area_rect:
            Canvas.SetLeft(self._area_rect, x)
            Canvas.SetTop(self._area_rect,  y)
            self._area_rect.Width  = rw
            self._area_rect.Height = rh
        self._info_lbl.Text = 'Drawing: {:.2f}" x {:.2f}"'.format(
            rw / PX_PER_INCH, rh / PX_PER_INCH)
        e.Handled = True

    def _on_mouse_up(self, sender, e):
        if not self._dragging:
            return
        self._canvas.ReleaseMouseCapture()
        self._dragging = False
        if self._area_rect:
            x  = Canvas.GetLeft(self._area_rect)
            y  = Canvas.GetTop(self._area_rect)
            rw = self._area_rect.Width
            rh = self._area_rect.Height
            if rw > SNAP_PX and rh > SNAP_PX:
                # Convert canvas Y (top-down) back to Revit sheet inches (bottom-up)
                self._result = {
                    "x_in": x  / PX_PER_INCH,
                    "y_in": self._sheet_h - (y + rh) / PX_PER_INCH,
                    "w_in": rw / PX_PER_INCH,
                    "h_in": rh / PX_PER_INCH,
                }
                self._info_lbl.Text = (
                    'Area: {:.2f}" x {:.2f}" -- click Apply Area'
                    .format(rw / PX_PER_INCH, rh / PX_PER_INCH))
                self._ok_btn.IsEnabled = True
            else:
                self._result = None
        e.Handled = True

    def _on_cancel(self, sender, e):
        self._result = None; self.Close()

    def _on_ok(self, sender, e):
        self.Close()

    @property
    def result(self):
        return self._result


# ─────────────────────────────────────────────────────────────────────
#  ASSEMBLY PROFILE  (full sheet layout + view config, cross-assembly)
# ─────────────────────────────────────────────────────────────────────

PROFILES_DIR = os.path.join(EXT_DIR, "config", "assembly_profiles")


def _ensure_profiles_dir():
    if not os.path.exists(PROFILES_DIR):
        os.makedirs(PROFILES_DIR)


def list_profiles():
    _ensure_profiles_dir()
    return sorted(f[:-5] for f in os.listdir(PROFILES_DIR) if f.endswith(".json"))


def load_profile(name):
    path = os.path.join(PROFILES_DIR, name + ".json")
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as fp:
            return json.load(fp)
    except Exception:
        return None


def save_profile(name, data):
    _ensure_profiles_dir()
    with open(os.path.join(PROFILES_DIR, name + ".json"), "w") as fp:
        json.dump(data, fp, indent=2)


def delete_profile(name):
    path = os.path.join(PROFILES_DIR, name + ".json")
    if os.path.exists(path):
        os.remove(path)


# ─────────────────────────────────────────────────────────────────────
#  PROFILE MAPPER WINDOW
# ─────────────────────────────────────────────────────────────────────

class ProfileMapperWindow(Window):
    """Shown when a profile's view slots don't all match the current
    assembly's views by view_key + name.

    The user assigns each unresolved profile slot to a current view
    (or skips it).  Returns {slot_index: row_dict_or_None}.
    """

    def __init__(self, unresolved_slots, current_view_rows):
        Window.__init__(self)
        self._slots  = unresolved_slots
        self._rows   = current_view_rows
        self._result = None
        self._combos = []

        import System.Windows.Media as WMedia
        self.Title  = "Map Profile Views to Current Assembly"
        self.Width  = 580
        self.Height = min(140 + len(unresolved_slots) * 54 + 60, 640)
        self.MinWidth  = 480
        self.MinHeight = 240
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(240, 242, 247))
        self.FontFamily = WMedia.FontFamily("Segoe UI")

        outer = Grid()
        for h_val, unit in [
            (50, System.Windows.GridUnitType.Pixel),
            (1,  System.Windows.GridUnitType.Star),
            (50, System.Windows.GridUnitType.Pixel),
        ]:
            rd = System.Windows.Controls.RowDefinition()
            rd.Height = System.Windows.GridLength(h_val, unit)
            outer.RowDefinitions.Add(rd)
        self.Content = outer

        # Header
        hdr = Border()
        hdr.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
        Grid.SetRow(hdr, 0)
        hg = Grid(); hg.Margin = Thickness(14, 0, 14, 0)
        hdr.Child = hg
        t1 = TextBlock()
        t1.Text = "Map Profile Views"
        t1.Foreground = Brushes.White
        t1.FontSize = 13
        t1.FontWeight = System.Windows.FontWeights.SemiBold
        t1.VerticalAlignment = VerticalAlignment.Center
        hg.Children.Add(t1)
        t2 = TextBlock()
        t2.Text = "Some saved views don't match -- assign or skip each one"
        t2.Foreground = SolidColorBrush(Color.FromRgb(127, 168, 204))
        t2.FontSize = 10
        t2.VerticalAlignment = VerticalAlignment.Center
        t2.HorizontalAlignment = HorizontalAlignment.Right
        hg.Children.Add(t2)
        outer.Children.Add(hdr)

        # Body
        sv = System.Windows.Controls.ScrollViewer()
        sv.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto)
        sv.Margin = Thickness(14, 10, 14, 6)
        Grid.SetRow(sv, 1)
        body = StackPanel()
        sv.Content = body
        outer.Children.Add(sv)

        # Column headers
        hrow = Grid(); hrow.Margin = Thickness(0, 0, 0, 8)
        for w_val, unit in [
            (190, System.Windows.GridUnitType.Pixel),
            (1,   System.Windows.GridUnitType.Star),
        ]:
            cd = ColumnDefinition()
            cd.Width = System.Windows.GridLength(w_val, unit)
            hrow.ColumnDefinitions.Add(cd)
        for col, txt in enumerate(["Profile slot (no match found)",
                                    "Map to current assembly view"]):
            tb = TextBlock()
            tb.Text = txt; tb.FontSize = 10
            tb.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
            tb.FontWeight = System.Windows.FontWeights.SemiBold
            Grid.SetColumn(tb, col); hrow.Children.Add(tb)
        body.Children.Add(hrow)

        for slot in unresolved_slots:
            rg = Grid(); rg.Margin = Thickness(0, 0, 0, 8)
            for w_val, unit in [
                (190, System.Windows.GridUnitType.Pixel),
                (1,   System.Windows.GridUnitType.Star),
            ]:
                cd = ColumnDefinition()
                cd.Width = System.Windows.GridLength(w_val, unit)
                rg.ColumnDefinitions.Add(cd)

            slot_lbl = TextBlock()
            slot_lbl.Text = "{} [{}]".format(
                slot.get("view_name", "?"), slot.get("view_key", "?"))
            slot_lbl.FontSize = 11
            slot_lbl.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
            slot_lbl.VerticalAlignment = VerticalAlignment.Center
            slot_lbl.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            Grid.SetColumn(slot_lbl, 0); rg.Children.Add(slot_lbl)

            combo = ComboBox()
            combo.FontSize = 11
            combo.Padding = Thickness(4, 3, 4, 3)
            skip = ComboBoxItem(); skip.Content = "-- Skip --"; skip.Tag = None
            combo.Items.Add(skip)
            for r in current_view_rows:
                ci = ComboBoxItem()
                ci.Content = "{} [{}]".format(r["view_name"], r["view_key"])
                ci.Tag = r
                combo.Items.Add(ci)
            combo.SelectedIndex = 0
            # Auto-select best match: same key first, then fuzzy name
            for ci_idx in range(1, combo.Items.Count):
                ci = combo.Items[ci_idx]
                if not (hasattr(ci, "Tag") and ci.Tag):
                    continue
                if ci.Tag.get("view_key") == slot.get("view_key"):
                    combo.SelectedIndex = ci_idx
                    break
            Grid.SetColumn(combo, 1); rg.Children.Add(combo)
            self._combos.append(combo)
            body.Children.Add(rg)

        # Footer
        footer = Border()
        footer.Background = SolidColorBrush(Color.FromRgb(250, 250, 250))
        footer.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
        footer.BorderThickness = Thickness(0, 1, 0, 0)
        Grid.SetRow(footer, 2)
        fg = Grid(); fg.Margin = Thickness(14, 0, 14, 0)
        footer.Child = fg
        bsp = StackPanel()
        bsp.Orientation = Orientation.Horizontal
        bsp.HorizontalAlignment = HorizontalAlignment.Right
        bsp.VerticalAlignment = VerticalAlignment.Center
        for lbl, handler, w, bg, fg_c, bd in [
            ("Cancel",        self._on_cancel, 90,
             Brushes.White,
             SolidColorBrush(Color.FromRgb(55, 65, 81)),
             Thickness(1)),
            ("Apply Mapping", self._on_ok,    130,
             SolidColorBrush(Color.FromRgb(30, 58, 95)),
             Brushes.White,
             Thickness(0)),
        ]:
            btn = System.Windows.Controls.Button()
            btn.Content = lbl; btn.Width = w; btn.Height = 32
            btn.Background = bg; btn.Foreground = fg_c
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
            btn.BorderThickness = bd; btn.FontSize = 12
            btn.Margin = Thickness(0, 0, 8, 0)
            btn.Click += handler
            bsp.Children.Add(btn)
        fg.Children.Add(bsp)
        outer.Children.Add(footer)

    def _on_cancel(self, sender, e):
        self._result = None; self.Close()

    def _on_ok(self, sender, e):
        mapping = {}
        for i, combo in enumerate(self._combos):
            sel = combo.SelectedItem
            mapping[i] = (sel.Tag
                          if (sel and hasattr(sel, "Tag") and sel.Tag is not None)
                          else None)
        self._result = mapping
        self.Close()

    @property
    def result(self):
        return self._result


# ─────────────────────────────────────────────────────────────────────
#  BATCH CREATE WINDOW
# ─────────────────────────────────────────────────────────────────────

class BatchCreateWindow(Window):
    """Select multiple assemblies and configure sheet numbering/naming
    to create one sheet per assembly in a single operation.

    Numbering patterns:
      Sequential   A-001, A-002, A-003 ...
      Base + Suffix  M-100.01, M-100.02 ...  (user sets base + separator)
      Custom prefix  MECH-001, MECH-002 ...

    Returns list of dicts [{assembly_id, sheet_number, sheet_name, tb_params}]
    or None if cancelled.
    """

    def __init__(self, all_assemblies, current_sheet_number,
                 current_sheet_name, current_tb_params):
        Window.__init__(self)
        self._all_assemblies   = all_assemblies  # [(name, ElementId), ...]
        self._base_number      = current_sheet_number
        self._base_name        = current_sheet_name
        self._tb_params        = dict(current_tb_params or {})
        self._result           = None
        self._filter_text      = ""

        import System.Windows.Media as WMedia
        self.Title  = "Batch Create Sheets"
        self.Width  = 740
        self.Height = 640
        self.MinWidth  = 600
        self.MinHeight = 500
        self.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(Color.FromRgb(240, 242, 247))
        self.FontFamily = WMedia.FontFamily("Segoe UI")

        self._build_ui()
        self._repopulate(all_assemblies)

    # ── Build UI ────────────────────────────────────────────────────

    def _build_ui(self):
        import System.Windows.Media as WMedia

        outer = Grid()
        for h_val, unit in [
            (50, System.Windows.GridUnitType.Pixel),   # header
            (1,  System.Windows.GridUnitType.Star),    # body
            (50, System.Windows.GridUnitType.Pixel),   # footer
        ]:
            rd = System.Windows.Controls.RowDefinition()
            rd.Height = System.Windows.GridLength(h_val, unit)
            outer.RowDefinitions.Add(rd)
        self.Content = outer

        # Header
        hdr = Border()
        hdr.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
        Grid.SetRow(hdr, 0)
        hg = Grid(); hg.Margin = Thickness(14, 0, 14, 0)
        hdr.Child = hg
        t1 = TextBlock()
        t1.Text = "Batch Create Sheets"
        t1.Foreground = Brushes.White
        t1.FontSize = 14
        t1.FontWeight = System.Windows.FontWeights.SemiBold
        t1.VerticalAlignment = VerticalAlignment.Center
        hg.Children.Add(t1)
        t2 = TextBlock()
        t2.Text = "Select assemblies -- one sheet per assembly"
        t2.Foreground = SolidColorBrush(Color.FromRgb(127, 168, 204))
        t2.FontSize = 10
        t2.VerticalAlignment = VerticalAlignment.Center
        t2.HorizontalAlignment = HorizontalAlignment.Right
        hg.Children.Add(t2)
        outer.Children.Add(hdr)

        # Body: two columns
        body = Grid()
        body.Margin = Thickness(12, 8, 12, 4)
        Grid.SetRow(body, 1)
        for w_val, unit in [
            (1, System.Windows.GridUnitType.Star),
            (1, System.Windows.GridUnitType.Star),
        ]:
            cd = System.Windows.Controls.ColumnDefinition()
            cd.Width = System.Windows.GridLength(w_val, unit)
            body.ColumnDefinitions.Add(cd)
        outer.Children.Add(body)

        # ── Left: assembly picker ────────────────────────────────────
        left = StackPanel()
        left.Margin = Thickness(0, 0, 6, 0)
        Grid.SetColumn(left, 0)
        body.Children.Add(left)

        hdr_lbl = TextBlock()
        hdr_lbl.Text = "ASSEMBLIES"
        hdr_lbl.Foreground = SolidColorBrush(Color.FromRgb(143, 108, 5))
        hdr_lbl.FontSize = 10; hdr_lbl.FontWeight = System.Windows.FontWeights.Bold
        hdr_lbl.Margin = Thickness(0, 0, 0, 6)
        left.Children.Add(hdr_lbl)

        # Search
        search_g = Grid(); search_g.Margin = Thickness(0, 0, 0, 4)
        self._search = TextBox()
        self._search.FontSize = 11; self._search.Padding = Thickness(6, 4, 6, 4)
        self._search.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        self._search.BorderThickness = Thickness(1)
        self._search.TextChanged += self._on_search
        search_g.Children.Add(self._search)
        hint = TextBlock()
        hint.Text = "Filter assemblies..."
        hint.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
        hint.FontSize = 11; hint.Margin = Thickness(8, 4, 0, 0)
        hint.IsHitTestVisible = False
        hint.Name = "SearchHint"
        search_g.Children.Add(hint)
        self._search_hint = hint
        left.Children.Add(search_g)

        # Select-all row
        sel_row = StackPanel()
        sel_row.Orientation = Orientation.Horizontal
        sel_row.Margin = Thickness(0, 0, 0, 4)
        sel_all_btn = System.Windows.Controls.Button()
        sel_all_btn.Content = "Select All"
        sel_all_btn.Width = 80; sel_all_btn.Height = 24
        sel_all_btn.FontSize = 10
        sel_all_btn.Background = Brushes.White
        sel_all_btn.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        sel_all_btn.BorderThickness = Thickness(1)
        sel_all_btn.Margin = Thickness(0, 0, 6, 0)
        sel_all_btn.Click += self._on_select_all
        sel_row.Children.Add(sel_all_btn)
        clear_btn = System.Windows.Controls.Button()
        clear_btn.Content = "Clear"
        clear_btn.Width = 60; clear_btn.Height = 24
        clear_btn.FontSize = 10
        clear_btn.Background = Brushes.White
        clear_btn.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        clear_btn.BorderThickness = Thickness(1)
        clear_btn.Click += self._on_clear_all
        sel_row.Children.Add(clear_btn)
        self._count_lbl = TextBlock()
        self._count_lbl.Text = "0 selected"
        self._count_lbl.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
        self._count_lbl.FontSize = 10
        self._count_lbl.VerticalAlignment = VerticalAlignment.Center
        self._count_lbl.Margin = Thickness(10, 0, 0, 0)
        sel_row.Children.Add(self._count_lbl)
        left.Children.Add(sel_row)

        # Assembly list (checkboxes)
        list_border = Border()
        list_border.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        list_border.BorderThickness = Thickness(1)
        list_border.CornerRadius = System.Windows.CornerRadius(3)
        sv = System.Windows.Controls.ScrollViewer()
        sv.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto)
        sv.MaxHeight = 340
        self._list_panel = StackPanel()
        sv.Content = self._list_panel
        list_border.Child = sv
        left.Children.Add(list_border)

        # ── Right: naming settings ───────────────────────────────────
        right_sv = System.Windows.Controls.ScrollViewer()
        right_sv.VerticalScrollBarVisibility = (
            System.Windows.Controls.ScrollBarVisibility.Auto)
        right_sv.Margin = Thickness(6, 0, 0, 0)
        Grid.SetColumn(right_sv, 1)
        right = StackPanel()
        right_sv.Content = right
        body.Children.Add(right_sv)

        def _card(title):
            b = Border()
            b.Background = Brushes.White
            b.CornerRadius = System.Windows.CornerRadius(5)
            b.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
            b.BorderThickness = Thickness(1)
            b.Margin = Thickness(0, 0, 0, 8)
            sp = StackPanel(); sp.Margin = Thickness(12, 10, 12, 10)
            b.Child = sp
            hdr = TextBlock()
            hdr.Text = title
            hdr.Foreground = SolidColorBrush(Color.FromRgb(143, 108, 5))
            hdr.FontSize = 10; hdr.FontWeight = System.Windows.FontWeights.Bold
            hdr.Margin = Thickness(0, 0, 0, 8)
            sp.Children.Add(hdr)
            right.Children.Add(b)
            return sp

        def _row(parent, label, ctrl, margin_b=6):
            g = Grid(); g.Margin = Thickness(0, 0, 0, margin_b)
            c0 = System.Windows.Controls.ColumnDefinition()
            c0.Width = System.Windows.GridLength(90)
            c1 = System.Windows.Controls.ColumnDefinition()
            c1.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            g.ColumnDefinitions.Add(c0); g.ColumnDefinitions.Add(c1)
            lbl = TextBlock(); lbl.Text = label
            lbl.FontSize = 11
            lbl.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(lbl, 0); g.Children.Add(lbl)
            Grid.SetColumn(ctrl, 1); g.Children.Add(ctrl)
            parent.Children.Add(g)

        def _tb(val=""):
            t = TextBox(); t.Text = val; t.FontSize = 11
            t.Padding = Thickness(4, 3, 4, 3)
            t.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
            t.BorderThickness = Thickness(1)
            return t

        # NUMBERING card
        num_card = _card("SHEET NUMBERING")

        mode_sp = StackPanel()
        mode_sp.Orientation = Orientation.Horizontal
        mode_sp.Margin = Thickness(0, 0, 0, 10)

        self._rb_seq   = System.Windows.Controls.RadioButton()
        self._rb_seq.Content = "Sequential"
        self._rb_seq.GroupName = "NumMode"
        self._rb_seq.IsChecked = True
        self._rb_seq.FontSize = 11; self._rb_seq.Margin = Thickness(0, 0, 12, 0)
        self._rb_seq.Click += self._on_mode_changed
        mode_sp.Children.Add(self._rb_seq)

        self._rb_suf   = System.Windows.Controls.RadioButton()
        self._rb_suf.Content = "Base + Suffix"
        self._rb_suf.GroupName = "NumMode"
        self._rb_suf.FontSize = 11
        self._rb_suf.Click += self._on_mode_changed
        mode_sp.Children.Add(self._rb_suf)
        num_card.Children.Add(mode_sp)

        # Sequential fields
        self._seq_panel = StackPanel()
        self._base_num_tb = _tb(self._base_number or "A-001")
        _row(self._seq_panel, "Start at", self._base_num_tb)
        hint2 = TextBlock()
        hint2.Text = "Numbers auto-increment. The numeric suffix increments (A-001, A-002...)"
        hint2.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
        hint2.FontSize = 9; hint2.TextWrapping = System.Windows.TextWrapping.Wrap
        self._seq_panel.Children.Add(hint2)
        num_card.Children.Add(self._seq_panel)

        # Base + suffix fields
        self._suf_panel = StackPanel()
        self._suf_panel.Visibility = Visibility.Collapsed
        self._base_suf_tb = _tb(self._base_number or "M-100")
        _row(self._suf_panel, "Base number", self._base_suf_tb)
        self._sep_tb = _tb(".")
        _row(self._suf_panel, "Separator", self._sep_tb)
        self._suf_start_tb = _tb("01")
        _row(self._suf_panel, "Start suffix", self._suf_start_tb)
        self._suf_digits_tb = _tb("2")
        _row(self._suf_panel, "Digits", self._suf_digits_tb)
        hint3 = TextBlock()
        hint3.Text = "Example: M-100.01, M-100.02 ..."
        hint3.Foreground = SolidColorBrush(Color.FromRgb(156, 163, 175))
        hint3.FontSize = 9
        self._suf_example = hint3
        self._suf_panel.Children.Add(hint3)
        for ctrl in (self._base_suf_tb, self._sep_tb,
                     self._suf_start_tb, self._suf_digits_tb):
            ctrl.TextChanged += self._update_suf_example
        num_card.Children.Add(self._suf_panel)

        # NAMING card
        name_card = _card("SHEET NAMING")

        self._rb_name_asm = System.Windows.Controls.RadioButton()
        self._rb_name_asm.Content = "Use assembly name"
        self._rb_name_asm.GroupName = "NameMode"; self._rb_name_asm.IsChecked = True
        self._rb_name_asm.FontSize = 11; self._rb_name_asm.Margin = Thickness(0, 0, 0, 6)
        name_card.Children.Add(self._rb_name_asm)

        self._rb_name_custom = System.Windows.Controls.RadioButton()
        self._rb_name_custom.Content = "Fixed name"
        self._rb_name_custom.GroupName = "NameMode"
        self._rb_name_custom.FontSize = 11; self._rb_name_custom.Margin = Thickness(0, 0, 0, 6)
        name_card.Children.Add(self._rb_name_custom)

        self._custom_name_tb = _tb(self._base_name or "")
        self._custom_name_tb.IsEnabled = False
        self._rb_name_custom.Click += lambda s, e: setattr(
            self._custom_name_tb, "IsEnabled", True)
        self._rb_name_asm.Click += lambda s, e: setattr(
            self._custom_name_tb, "IsEnabled", False)
        name_card.Children.Add(self._custom_name_tb)

        # PREVIEW card
        prev_card = _card("NUMBERING PREVIEW")
        self._preview_lbl = TextBlock()
        self._preview_lbl.FontSize = 10
        self._preview_lbl.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
        self._preview_lbl.TextWrapping = System.Windows.TextWrapping.Wrap
        prev_card.Children.Add(self._preview_lbl)

        # Footer
        footer = Border()
        footer.Background = SolidColorBrush(Color.FromRgb(250, 250, 250))
        footer.BorderBrush = SolidColorBrush(Color.FromRgb(229, 231, 235))
        footer.BorderThickness = Thickness(0, 1, 0, 0)
        Grid.SetRow(footer, 2)
        fg = Grid(); fg.Margin = Thickness(14, 0, 14, 0)
        footer.Child = fg
        bsp = StackPanel()
        bsp.Orientation = Orientation.Horizontal
        bsp.HorizontalAlignment = HorizontalAlignment.Right
        bsp.VerticalAlignment = VerticalAlignment.Center

        cancel_b = System.Windows.Controls.Button()
        cancel_b.Content = "Cancel"; cancel_b.Width = 90; cancel_b.Height = 32
        cancel_b.Background = Brushes.White
        cancel_b.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
        cancel_b.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
        cancel_b.BorderThickness = Thickness(1); cancel_b.FontSize = 12
        cancel_b.Margin = Thickness(0, 0, 8, 0)
        cancel_b.Click += self._on_cancel
        bsp.Children.Add(cancel_b)

        self._ok_btn = System.Windows.Controls.Button()
        self._ok_btn.Content = "Create All Sheets"
        self._ok_btn.Width = 150; self._ok_btn.Height = 32
        self._ok_btn.Background = SolidColorBrush(Color.FromRgb(30, 58, 95))
        self._ok_btn.Foreground = Brushes.White
        self._ok_btn.BorderThickness = Thickness(0); self._ok_btn.FontSize = 12
        self._ok_btn.FontWeight = System.Windows.FontWeights.SemiBold
        self._ok_btn.IsEnabled = False
        self._ok_btn.Click += self._on_ok
        bsp.Children.Add(self._ok_btn)
        fg.Children.Add(bsp)
        outer.Children.Add(footer)

    # ── Assembly list ────────────────────────────────────────────────

    def _repopulate(self, assemblies):
        self._list_panel.Children.Clear()
        for name, aid in assemblies:
            chk = CheckBox()
            chk.Content = name
            chk.FontSize = 11
            chk.Padding = Thickness(4, 3, 4, 3)
            chk.Tag = aid
            chk.Checked   += self._on_check_changed
            chk.Unchecked += self._on_check_changed
            self._list_panel.Children.Add(chk)
        self._update_count()

    def _on_search(self, sender, e):
        text = self._search.Text or ""
        self._search_hint.Visibility = (
            Visibility.Collapsed if text else Visibility.Visible)
        q = text.strip().lower()
        filtered = [(n, aid) for n, aid in self._all_assemblies
                    if not q or q in n.lower()]
        self._repopulate(filtered)

    def _on_select_all(self, sender, e):
        for child in self._list_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = True
        self._update_count()

    def _on_clear_all(self, sender, e):
        for child in self._list_panel.Children:
            if isinstance(child, CheckBox):
                child.IsChecked = False
        self._update_count()

    def _on_check_changed(self, sender, e):
        self._update_count()

    def _update_count(self):
        n = sum(1 for c in self._list_panel.Children
                if isinstance(c, CheckBox) and c.IsChecked == True)
        self._count_lbl.Text = "{} selected".format(n)
        self._ok_btn.IsEnabled = n > 0
        self._refresh_preview(n)

    # ── Mode / preview ───────────────────────────────────────────────

    def _on_mode_changed(self, sender, e):
        is_suf = self._rb_suf.IsChecked == True
        self._seq_panel.Visibility = (
            Visibility.Collapsed if is_suf else Visibility.Visible)
        self._suf_panel.Visibility = (
            Visibility.Visible if is_suf else Visibility.Collapsed)
        self._update_count()

    def _update_suf_example(self, sender, e):
        nums = self._generate_numbers(3)
        self._suf_example.Text = "Example: " + ", ".join(nums) + " ..."
        self._update_count()

    def _generate_numbers(self, count):
        """Generate `count` sheet numbers based on current mode settings."""
        nums = []
        if self._rb_seq.IsChecked == True:
            base = self._base_num_tb.Text.strip()
            # Extract trailing digits from base and increment from there
            import re as _re
            m = _re.search(r'^(.*?)(\d+)$', base)
            if m:
                prefix = m.group(1)
                width  = len(m.group(2))
                start  = int(m.group(2))
                # Collect existing sheet numbers to skip duplicates
                existing = set()
                try:
                    for s in FilteredElementCollector(doc).OfClass(ViewSheet):
                        try: existing.add(s.SheetNumber.lower())
                        except: pass
                except: pass
                n = start
                while len(nums) < count:
                    candidate = "{}{}".format(prefix, str(n).zfill(width))
                    if candidate.lower() not in existing:
                        nums.append(candidate)
                    n += 1
            else:
                for i in range(1, count + 1):
                    nums.append("{}-{:02d}".format(base, i))
        else:
            base  = self._base_suf_tb.Text.strip()
            sep   = self._sep_tb.Text
            try:   start = int(self._suf_start_tb.Text.strip())
            except: start = 1
            try:   digits = max(1, int(self._suf_digits_tb.Text.strip()))
            except: digits = 2
            for i in range(start, start + count):
                nums.append("{}{}{}".format(base, sep, str(i).zfill(digits)))
        return nums

    def _refresh_preview(self, count):
        if count == 0:
            self._preview_lbl.Text = "No assemblies selected."
            return
        sample = self._generate_numbers(min(count, 6))
        if count > 6:
            sample.append("...")
        self._preview_lbl.Text = "\n".join(sample)

    # ── Result ───────────────────────────────────────────────────────

    def _on_cancel(self, sender, e):
        self._result = None; self.Close()

    def _on_ok(self, sender, e):
        selected_ids = []
        for child in self._list_panel.Children:
            if isinstance(child, CheckBox) and child.IsChecked == True:
                selected_ids.append(child.Tag)

        numbers = self._generate_numbers(len(selected_ids))

        rows = []
        for i, aid in enumerate(selected_ids):
            sheet_num  = numbers[i] if i < len(numbers) else "?"
            if self._rb_name_asm.IsChecked == True:
                try:
                    asm = doc.GetElement(aid)
                    sheet_name = asm.AssemblyTypeName or "Assembly"
                except Exception:
                    sheet_name = "Assembly"
            else:
                sheet_name = self._custom_name_tb.Text.strip() or "Assembly"
            rows.append({
                "assembly_id":  aid,
                "sheet_number": sheet_num,
                "sheet_name":   sheet_name,
            })

        self._result = rows
        self.Close()

    @property
    def result(self):
        return self._result


class AssemblyStudioWindow(forms.WPFWindow):

    def __init__(self, xaml_path, build_elements=None, preselected_assembly=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self._result       = None
        self._build_elements = build_elements or []
        self._assembly_id  = None  # set after build or existing selection

        # Canvas drag state (viewport dragging)
        self._dragging    = None
        self._drag_offset = None

        # Assembly list (populated by _init_existing_mode or on first toggle)
        self._assemblies = []

        # Set to True after batch create completes so entry point skips single-sheet path
        self._batch_completed = False

        # Usable area (canvas pixels): dict with x,y,w,h or None
        self._usable_area      = None
        self._usable_area_rect = None   # Rectangle element on canvas

        # View rows and canvas rects
        self._view_rows    = []   # list of row-dicts
        self._canvas_rects = {}   # view_id_int -> Border

        # Titleblock param boxes
        self._param_boxes = {}

        # Sheet dimensions in inches
        self._sheet_w = 17.0
        self._sheet_h = 11.0
        self._tb_margins = None  # populated when titleblock is selected

        # Preload template lists
        self._view_tmpls  = [("(None)", ElementId.InvalidElementId)] + get_view_templates()
        self._sched_tmpls = [("(None)", ElementId.InvalidElementId)] + get_schedule_templates()
        self._vp_types    = [("(Default)", ElementId.InvalidElementId)] + get_viewport_types()

        # Populate title blocks
        self._title_blocks = get_title_blocks()
        for name, tid in self._title_blocks:
            item = ComboBoxItem()
            item.Content = name
            item.Tag = tid
            self.TitleBlockCombo.Items.Add(item)
        if self._title_blocks:
            self.TitleBlockCombo.SelectedIndex = 0
            # Trigger initial titleblock population
            tb_id = self._title_blocks[0][1]
            self._sheet_w, self._sheet_h = get_titleblock_size(tb_id)
            self._tb_margins = get_titleblock_margins(tb_id)
            self.SheetSizeInfo.Text = '{:.0f}" x {:.0f}"'.format(
                self._sheet_w, self._sheet_h)
            self._populate_params(tb_id)
            # Restore saved usable area for this titleblock
            saved = _get_saved_tb_area(tb_id)
            if saved and "area" in saved:
                a   = saved["area"]
                ch  = self._sheet_h * PX_PER_INCH
                h_px = a["h_in"] * PX_PER_INCH
                self._usable_area = {
                    "x": a["x_in"] * PX_PER_INCH,
                    "y": ch - a["y_in"] * PX_PER_INCH - h_px,
                    "w": a["w_in"] * PX_PER_INCH,
                    "h": h_px,
                }
                if "margins" in saved:
                    self._tb_margins = dict(self._tb_margins or {})
                    self._tb_margins.update(saved["margins"])

        # Sheet number
        self.SheetNumberBox.Text = next_sheet_number()

        # Presets + Profiles
        self._refresh_presets()
        self._refresh_profiles()

        # Browser org: discover folder parameter + populate existing values
        self._browser_folder_param, folder_values = get_sheet_browser_org_info()
        for val in folder_values:
            item = ComboBoxItem()
            item.Content = val
            self.SheetFolderCombo.Items.Add(item)
        if self._browser_folder_param:
            self.SheetFolderCombo.ToolTip = (
                "Parameter: \"{}\"".format(self._browser_folder_param))
        else:
            self.SheetFolderCombo.IsEnabled = False
            self.SheetFolderCombo.ToolTip = (
                "No browser organization parameter found on sheets")

        # Draw initial canvas
        self._draw_sheet_outline()

        # Initialize mode
        if self._build_elements:
            self._init_build_mode()
        else:
            self._init_existing_mode(preselected_assembly)

    # ── Mode initialization ──────────────────────────────────────────

    def _init_build_mode(self):
        """Configure window for building a new assembly from selected elements."""
        self.RbCreate.IsChecked = True
        self.RbCreate.IsEnabled = True
        self.BuildSection.Visibility  = Visibility.Visible
        self.ExistingSection.Visibility = Visibility.Collapsed

        # Element summary in header
        cat_counts = {}
        for elem in self._build_elements:
            try:
                cat = elem.Category.Name if elem.Category else "Unknown"
            except Exception:
                cat = "Unknown"
            cat_counts[cat] = cat_counts.get(cat, 0) + 1
        parts = ["{} {}".format(v, k) for k, v in sorted(cat_counts.items())]
        self.HeaderInfo.Text = "{} selected: {}".format(
            len(self._build_elements), ", ".join(parts))

        # Assembly name
        self.AssemblyNameBox.Text = next_assembly_name()

        # Level combo
        self._levels = get_levels()
        for lv in self._levels:
            item = ComboBoxItem()
            item.Content = lv.Name
            item.Tag = lv.Id
            self.LevelCombo.Items.Add(item)
        self.LevelCombo.SelectedIndex = 0
        nearest = get_nearest_level(self._build_elements)
        if nearest:
            for i, lv in enumerate(self._levels):
                if lv.Id.Value == nearest.Id.Value:
                    self.LevelCombo.SelectedIndex = i
                    break

        # View template combos (section views)
        for combo in (self.Tmpl3D, self.TmplTop, self.TmplFront,
                      self.TmplBack, self.TmplLeft, self.TmplRight):
            for name, tid in self._view_tmpls:
                item = ComboBoxItem()
                item.Content = name
                item.Tag = tid
                combo.Items.Add(item)
            combo.SelectedIndex = 0

        # Schedule template combo
        for name, tid in self._sched_tmpls:
            item = ComboBoxItem()
            item.Content = name
            item.Tag = tid
            self.TmplSchedule.Items.Add(item)
        self.TmplSchedule.SelectedIndex = 0

        # Build status row hidden until assembly is created
        self.BuildStatusRow.Visibility = Visibility.Collapsed

    def _init_existing_mode(self, preselected=None):
        """Configure window for using an existing assembly."""
        self.RbExisting.IsChecked = True
        # Disable Create mode toggle if there are no elements to build from
        self.RbCreate.IsEnabled = bool(self._build_elements)

        self.BuildSection.Visibility   = Visibility.Collapsed
        self.ExistingSection.Visibility = Visibility.Visible

        # Load full assembly list into memory -- used by search filter
        self._assemblies = get_assemblies()   # [(display_name, ElementId), ...]
        self._assembly_filter = ""

        # Show placeholder hint
        self.AssemblySearchHint.Visibility = Visibility.Visible
        self.AssemblySearchBox.Text = ""

        # Populate the list initially (all assemblies)
        self._repopulate_assembly_list(self._assemblies)

        if preselected:
            for i, (name, aid) in enumerate(self._assemblies):
                if aid.Value == preselected.Id.Value:
                    self._select_assembly_row(i)
                    break

    # ── Mode toggle handlers (Click fires only on user interaction) ──────

    def RbCreate_Clicked(self, sender, e):
        self.BuildSection.Visibility    = Visibility.Visible
        self.ExistingSection.Visibility = Visibility.Collapsed

    def RbExisting_Clicked(self, sender, e):
        self.BuildSection.Visibility    = Visibility.Collapsed
        self.ExistingSection.Visibility = Visibility.Visible

        # If we started in Build mode, _init_existing_mode was never called.
        # Do it now on first switch -- this populates the list and wires everything.
        if not self._assemblies:
            self._init_existing_mode(None)
            # Auto-select any assembly just built this session
            if self._assembly_id is not None:
                self._select_assembly_row_by_id(self._assembly_id)
            return

        # Already initialised -- just refresh with any newly created assemblies
        fresh = get_assemblies()
        current_ids = {aid.Value for _, aid in self._assemblies}
        added = False
        for name, aid in fresh:
            if aid.Value not in current_ids:
                self._assemblies.append((name, aid))
                added = True
        if added:
            self._assemblies.sort(key=lambda x: x[0].lower())
            self._repopulate_assembly_list(self._assemblies)

        # Auto-select any assembly just built this session
        if self._assembly_id is not None:
            self._select_assembly_row_by_id(self._assembly_id)

    def _repopulate_assembly_list(self, filtered_list):
        """Rebuild the ListBox from a filtered list of (name, aid) tuples."""
        self.AssemblyListBox.Items.Clear()
        for name, aid in filtered_list:
            item = System.Windows.Controls.ListBoxItem()
            item.Content = name
            item.Tag = aid
            item.Padding = Thickness(4, 3, 4, 3)
            self.AssemblyListBox.Items.Add(item)
        count = len(filtered_list)
        total = len(self._assemblies)
        if count == total:
            self.AssemblyInfo.Text = "{} assemblies".format(total)
        else:
            self.AssemblyInfo.Text = "{} of {} shown".format(count, total)

    def _select_assembly_row(self, master_index):
        """Select the row matching master_index in the current filtered list."""
        target_name, target_id = self._assemblies[master_index]
        self._select_assembly_row_by_id(target_id)

    def _select_assembly_row_by_id(self, assembly_id):
        """Select the ListBox row whose Tag matches assembly_id."""
        for i in range(self.AssemblyListBox.Items.Count):
            item = self.AssemblyListBox.Items[i]
            if hasattr(item, "Tag") and item.Tag is not None:
                if item.Tag.Value == assembly_id.Value:
                    self.AssemblyListBox.SelectedIndex = i
                    self.AssemblyListBox.ScrollIntoView(item)
                    return

    def AssemblySearch_Changed(self, sender, e):
        """Filter the assembly list as the user types."""
        text = self.AssemblySearchBox.Text or ""
        # Show/hide placeholder
        self.AssemblySearchHint.Visibility = (
            Visibility.Collapsed if text else Visibility.Visible)
        query = text.strip().lower()
        if not query:
            self._repopulate_assembly_list(self._assemblies)
        else:
            filtered = [(n, aid) for n, aid in self._assemblies
                        if query in n.lower()]
            self._repopulate_assembly_list(filtered)

    def AssemblyListBox_Changed(self, sender, e):
        """Handle selection in the assembly list box."""
        sel = self.AssemblyListBox.SelectedItem
        if sel is None or not hasattr(sel, "Tag") or sel.Tag is None:
            return
        assembly_id = sel.Tag
        self._load_assembly(assembly_id)



    def SheetLocation_Changed(self, sender, e):
        """Toggle folder panel visibility based on selected sheet location."""
        if self.RbLocSheets.IsChecked == True:
            self.LocFolderPanel.Visibility  = Visibility.Visible
            self.LocAssemblyHint.Visibility = Visibility.Collapsed
        else:
            self.LocFolderPanel.Visibility  = Visibility.Collapsed
            self.LocAssemblyHint.Visibility = Visibility.Visible



    def CreateAssembly_Click(self, sender, e):
        name = self.AssemblyNameBox.Text.strip()
        if not name:
            forms.alert("Enter an assembly name.", exitscript=False)
            return
        if self.LevelCombo.SelectedItem is None:
            forms.alert("Select a level.", exitscript=False)
            return

        config = {
            "elements": self._build_elements,
            "name":     name,
            "level_id": self._combo_tag(self.LevelCombo),
            "views": {
                "3d":       self.Chk3D.IsChecked       == True,
                "top":      self.ChkTop.IsChecked       == True,
                "front":    self.ChkFront.IsChecked     == True,
                "back":     self.ChkBack.IsChecked      == True,
                "left":     self.ChkLeft.IsChecked      == True,
                "right":    self.ChkRight.IsChecked     == True,
                "schedule": self.ChkSchedule.IsChecked  == True,
            },
            "templates": {
                "3d":       self._combo_tag(self.Tmpl3D),
                "top":      self._combo_tag(self.TmplTop),
                "front":    self._combo_tag(self.TmplFront),
                "back":     self._combo_tag(self.TmplBack),
                "left":     self._combo_tag(self.TmplLeft),
                "right":    self._combo_tag(self.TmplRight),
                "schedule": self._combo_tag(self.TmplSchedule),
            },
        }

        assembly, created_views, errors = create_assembly(config)

        if assembly is None:
            forms.alert(
                "Assembly creation failed:\n" + "\n".join(errors),
                exitscript=False
            )
            return

        self._assembly_id = assembly.Id

        # Collapse the build form, show status
        self.BuildStatusLabel.Text = "Created: {} ({} views built)".format(
            name, len(created_views))
        self.BuildStatusRow.Visibility = Visibility.Visible
        self.AssemblyNameBox.IsEnabled = False
        self.LevelCombo.IsEnabled      = False
        self.CreateAssemblyBtn.IsEnabled = False

        # Populate views-to-place from the new assembly
        views = get_assembly_views(self._assembly_id)
        self._populate_view_rows(views)

        # Pre-fill sheet name with assembly name
        self.SheetNameBox.Text = name
        self.HeaderInfo.Text = "Assembly: {} | {} views ready to place".format(
            name, len(views))

        if errors:
            forms.alert(
                "Assembly created with warnings:\n" + "\n".join(errors),
                exitscript=False
            )

    # ── Existing assembly selection ──────────────────────────────────

    def _load_assembly(self, assembly_id):
        """Load views and metadata for a selected assembly."""
        if assembly_id is None or assembly_id == ElementId.InvalidElementId:
            self.AssemblyInfo.Text = "Type to search or scroll to select"
            self._clear_view_rows()
            self._assembly_id = None
            return

        self._assembly_id = assembly_id
        assembly = doc.GetElement(assembly_id)
        if assembly is None:
            return

        try:
            count = assembly.GetMemberIds().Count
        except Exception:
            count = 0

        views = get_assembly_views(assembly_id)
        self.AssemblyInfo.Text = "{} elements, {} views available".format(
            count, len(views))

        try:
            aname = assembly.AssemblyTypeName or "Assembly"
        except Exception:
            aname = "Assembly"

        self.SheetNameBox.Text = aname
        self.HeaderInfo.Text = aname
        self._populate_view_rows(views)


    # ── Titleblock selection ─────────────────────────────────────────

    def TitleBlockCombo_Changed(self, sender, e):
        tb_id = self._combo_tag(self.TitleBlockCombo)
        if tb_id == ElementId.InvalidElementId:
            return
        self._sheet_w, self._sheet_h = get_titleblock_size(tb_id)
        self._tb_margins = get_titleblock_margins(tb_id)
        self.SheetSizeInfo.Text = '{:.0f}" x {:.0f}"'.format(
            self._sheet_w, self._sheet_h)
        self._populate_params(tb_id)

        # Restore a previously saved usable area for this titleblock, if any.
        saved = _get_saved_tb_area(tb_id)
        if saved and "area" in saved:
            a   = saved["area"]
            ch  = self._sheet_h * PX_PER_INCH
            h_px = a["h_in"] * PX_PER_INCH
            self._usable_area = {
                "x": a["x_in"] * PX_PER_INCH,
                "y": ch - a["y_in"] * PX_PER_INCH - h_px,
                "w": a["w_in"] * PX_PER_INCH,
                "h": h_px,
            }
            if "margins" in saved:
                self._tb_margins = dict(self._tb_margins or {})
                self._tb_margins.update(saved["margins"])
        else:
            self._usable_area = None
            if self._usable_area_rect is not None:
                try:
                    self._usable_area_rect.Tag = None
                    self.SheetCanvas.Children.Remove(self._usable_area_rect)
                except Exception:
                    pass
                self._usable_area_rect = None

        self._update_usable_area_display()
        self._reposition_all_rects()

    # ── View rows ────────────────────────────────────────────────────

    def _clear_view_rows(self):
        # Remove all viewport Borders from canvas before clearing tracking dict
        for border in list(self._canvas_rects.values()):
            try:
                self.SheetCanvas.Children.Remove(border)
            except Exception:
                pass
        self.ViewRowsPanel.Children.Clear()
        self._view_rows    = []
        self._canvas_rects = {}
        self.NoViewsLabel.Visibility = Visibility.Visible
        self._draw_sheet_outline()

    def _populate_view_rows(self, views):
        # Remove all existing viewport Borders from canvas before rebuild
        for border in list(self._canvas_rects.values()):
            try:
                self.SheetCanvas.Children.Remove(border)
            except Exception:
                pass
        self.ViewRowsPanel.Children.Clear()
        self._view_rows    = []
        self._canvas_rects = {}
        self._draw_sheet_outline()

        if not views:
            self.NoViewsLabel.Visibility = Visibility.Visible
            return
        self.NoViewsLabel.Visibility = Visibility.Collapsed

        for vkey, vname, vid, is_sched in views:
            row_data = {
                "view_key":    vkey,
                "view_name":   vname,
                "view_id":     vid,
                "is_schedule": is_sched,
                "checkbox":    None,
                "scale_combo": None,
                "vp_type_combo": None,
            }

            row_grid = Grid()
            row_grid.Margin = Thickness(0, 2, 0, 2)

            col0 = ColumnDefinition()
            col0.Width = System.Windows.GridLength(24)
            col1 = ColumnDefinition()
            col1.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            col2 = ColumnDefinition()
            col2.Width = System.Windows.GridLength(70)
            col3 = ColumnDefinition()
            col3.Width = System.Windows.GridLength(90)
            row_grid.ColumnDefinitions.Add(col0)
            row_grid.ColumnDefinitions.Add(col1)
            row_grid.ColumnDefinitions.Add(col2)
            row_grid.ColumnDefinitions.Add(col3)

            chk = CheckBox()
            chk.IsChecked = True
            chk.VerticalAlignment = VerticalAlignment.Center
            chk.Tag = vid
            Grid.SetColumn(chk, 0)
            row_grid.Children.Add(chk)
            row_data["checkbox"] = chk

            lbl = TextBlock()
            lbl.Text = vname
            lbl.FontSize = 11
            lbl.Foreground = SolidColorBrush(Color.FromRgb(55, 65, 81))
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            Grid.SetColumn(lbl, 1)
            row_grid.Children.Add(lbl)

            if not is_sched:
                sc = ComboBox()
                sc.FontSize = 10
                sc.Padding = Thickness(2, 1, 2, 1)
                for stext, sfactor in SCALES:
                    si = ComboBoxItem()
                    si.Content = stext
                    si.Tag = sfactor
                    sc.Items.Add(si)

                # Default to the view's actual current scale if it matches
                # a known scale in our list; fall back to 1/4"=1'-0" (index 6)
                default_idx = 6
                try:
                    v_obj = doc.GetElement(vid)
                    actual = v_obj.Scale
                    for i in range(sc.Items.Count):
                        si = sc.Items[i]
                        if hasattr(si, "Tag") and si.Tag == actual:
                            default_idx = i
                            break
                except Exception:
                    pass
                sc.SelectedIndex = default_idx

                Grid.SetColumn(sc, 2)
                row_grid.Children.Add(sc)
                row_data["scale_combo"] = sc

                vpt = ComboBox()
                vpt.FontSize = 10
                vpt.Padding = Thickness(2, 1, 2, 1)
                for tname, tid in self._vp_types:
                    vi = ComboBoxItem()
                    vi.Content = tname
                    vi.Tag = tid
                    vpt.Items.Add(vi)
                vpt.SelectedIndex = 0
                Grid.SetColumn(vpt, 3)
                row_grid.Children.Add(vpt)
                row_data["vp_type_combo"] = vpt

            self.ViewRowsPanel.Children.Add(row_grid)
            self._view_rows.append(row_data)

            chk.Checked   += self._view_checked
            chk.Unchecked += self._view_unchecked

            # Wire scale change AFTER row_data is appended so handler can find it
            if row_data.get("scale_combo"):
                row_data["scale_combo"].SelectionChanged += self._on_scale_changed

            self._add_canvas_rect(row_data)

    def _view_checked(self, sender, e):
        vid = sender.Tag
        for row in self._view_rows:
            if row["view_id"].Value == vid.Value:
                self._add_canvas_rect(row)
                break

    def _view_unchecked(self, sender, e):
        vid = sender.Tag
        key = vid.Value
        if key in self._canvas_rects:
            self.SheetCanvas.Children.Remove(self._canvas_rects[key])
            del self._canvas_rects[key]

    # ── Canvas ───────────────────────────────────────────────────────

    def _draw_sheet_outline(self):
        """Redraw canvas background: fill, 1/4-inch grid, titleblock zones, border.

        Viewport Borders are REMOVED from the canvas, the background stack is
        painted fresh, then viewport Borders are re-appended last so they sit
        on top of (higher z-order than) the white fill and zone overlays.
        Every element added here has IsHitTestVisible = False so the Borders
        can still receive mouse events.
        """
        canvas = self.SheetCanvas

        # 1. Pull out all viewport Borders so we can put them back on top.
        vp_borders = []
        non_vp = []
        for child in canvas.Children:
            try:
                tag = child.Tag
            except Exception:
                tag = None
            if tag == "viewport":
                vp_borders.append(child)
            else:
                non_vp.append(child)

        # 2. Remove everything.
        canvas.Children.Clear()

        w = self._sheet_w * PX_PER_INCH
        h = self._sheet_h * PX_PER_INCH
        canvas.Width  = w
        canvas.Height = h

        def _r(x, y, rw, rh, fill=None, stroke=None, thick=0):
            r = Rectangle()
            r.Width = rw; r.Height = rh
            r.IsHitTestVisible = False
            if fill:   r.Fill   = fill
            if stroke: r.Stroke = stroke;  r.StrokeThickness = thick
            Canvas.SetLeft(r, x); Canvas.SetTop(r, y)
            canvas.Children.Add(r)

        def _l(x1, y1, x2, y2, stroke, thick=1.0):
            ln = Line()
            ln.X1 = x1; ln.Y1 = y1; ln.X2 = x2; ln.Y2 = y2
            ln.Stroke = stroke; ln.StrokeThickness = thick
            ln.IsHitTestVisible = False
            canvas.Children.Add(ln)

        def _t(text, x, y, size, color):
            tb = TextBlock()
            tb.Text = text; tb.FontSize = size
            tb.Foreground = color
            tb.IsHitTestVisible = False
            Canvas.SetLeft(tb, x); Canvas.SetTop(tb, y)
            canvas.Children.Add(tb)

        # ── 3a. White fill ───────────────────────────────────────────
        _r(0, 0, w, h, fill=Brushes.White)

        # ── 3b. 1/4" snap grid ───────────────────────────────────────
        minor = SolidColorBrush(Color.FromRgb(236, 239, 244))
        major = SolidColorBrush(Color.FromRgb(209, 213, 219))
        for i in range(1, int(w / SNAP_PX) + 1):
            _l(i*SNAP_PX, 0, i*SNAP_PX, h,
               major if (i % 4 == 0) else minor,
               1.5 if (i % 4 == 0) else 0.5)
        for i in range(1, int(h / SNAP_PX) + 1):
            _l(0, i*SNAP_PX, w, i*SNAP_PX,
               major if (i % 4 == 0) else minor,
               1.5 if (i % 4 == 0) else 0.5)

        # ── 3c. Titleblock zone overlays ─────────────────────────────
        mb = self._tb_margins or {
            "margin_l": 1.5, "margin_r": 0.5,
            "margin_t": 0.5, "margin_b": 0.5,
            "title_h":  2.0, "revision_w": 0.0,
        }
        ml      = mb["margin_l"]        * PX_PER_INCH
        mr      = mb["margin_r"]        * PX_PER_INCH
        mt      = mb["margin_t"]        * PX_PER_INCH
        mb_px   = mb["margin_b"]        * PX_PER_INCH
        title_h = mb["title_h"]         * PX_PER_INCH
        rev_w   = mb.get("revision_w", 0.0) * PX_PER_INCH

        content_bottom = h - mb_px - title_h
        content_h      = max(0.0, content_bottom - mt)

        amber_fill = SolidColorBrush(Color.FromArgb(30, 180, 140, 60))
        blue_fill  = SolidColorBrush(Color.FromArgb(20, 100, 100, 200))
        red_fill   = SolidColorBrush(Color.FromArgb(20, 200, 100, 100))
        amber_line = SolidColorBrush(Color.FromRgb(180, 140,  60))
        blue_line  = SolidColorBrush(Color.FromRgb(100, 100, 200))
        amber_text = SolidColorBrush(Color.FromRgb(140, 100,  20))

        _r(0, content_bottom, w, title_h + mb_px, fill=amber_fill)
        _l(0, content_bottom, w, content_bottom, amber_line, 1.5)
        _t('TITLE STRIP  ({:.2f}")'.format(mb["title_h"]),
           ml + 8, content_bottom + 4,
           max(10, min(16, title_h * 0.25)), amber_text)
        _r(0, mt, ml, content_h, fill=blue_fill)
        _l(ml, mt, ml, content_bottom, blue_line, 1.5)
        if mr > 0:
            _r(w - mr, mt, mr, content_h, fill=blue_fill)
            _l(w - mr, mt, w - mr, content_bottom, blue_line, 1.5)
        if mt > 0:
            _r(0, 0, w, mt, fill=blue_fill)
            _l(0, mt, w, mt, blue_line, 1.5)
        if rev_w > 0:
            _r(w - mr - rev_w, mt, rev_w, content_h, fill=red_fill)
            _l(w - mr - rev_w, mt, w - mr - rev_w, content_bottom, blue_line, 1.5)

        # ── 3d. Sheet border ─────────────────────────────────────────
        _r(0, 0, w, h,
           stroke=SolidColorBrush(Color.FromRgb(55, 65, 81)), thick=2)

        # ── 4. Re-attach viewport Borders on top ─────────────────────
        for b in vp_borders:
            canvas.Children.Add(b)


    # ── Define Usable Area (popup) ───────────────────────────────────

    def DefineArea_Click(self, sender, e):
        """Gather titleblock geometry + view footprints, open the picker,
        pre-populate from saved config if available, and save on Apply.
        """
        if not self._title_blocks:
            forms.alert("Select a titleblock first.", exitscript=False)
            return

        tb_id = self._combo_tag(self.TitleBlockCombo)
        if tb_id == ElementId.InvalidElementId:
            forms.alert("Select a titleblock first.", exitscript=False)
            return

        # Check for a previously saved area for this titleblock
        saved = _get_saved_tb_area(tb_id)

        # If we have a saved config, override tb_margins with saved margins
        # so the picker opens with the exact settings the user last used.
        tb_margins = self._tb_margins or {}
        if saved and "margins" in saved:
            tb_margins = dict(tb_margins)
            tb_margins.update(saved["margins"])

        # Saved area takes priority over the live viewport scan
        suggested = None
        if saved and "area" in saved:
            suggested = saved["area"]
        else:
            suggested = get_titleblock_drawing_area_inches(tb_id)

        # Extract real titleblock geometry lines from a placed instance
        tb_lines = get_titleblock_geometry_lines(tb_id, self._sheet_h)

        # Build ghost view footprints from current view rows + selected scales
        view_ghosts = []
        for row in self._view_rows:
            if row.get("checkbox") and row["checkbox"].IsChecked != True:
                continue
            vid  = row["view_id"]
            vkey = row["view_key"]
            name = row["view_name"]

            if row["is_schedule"]:
                size = get_schedule_sheet_size_px(vid)
                if size:
                    view_ghosts.append((name, vkey, size[0], size[1]))
            else:
                scale_factor = 48
                sc = row.get("scale_combo")
                if sc and sc.SelectedItem and hasattr(sc.SelectedItem, "Tag"):
                    try: scale_factor = int(sc.SelectedItem.Tag)
                    except: pass
                size = get_view_sheet_size_px(vid, scale_factor, self._assembly_id)
                if size:
                    view_ghosts.append((name, vkey, size[0], size[1]))

        picker = UsableAreaPickerWindow(
            self._sheet_w, self._sheet_h,
            tb_margins,
            tb_lines=tb_lines,
            view_ghosts=view_ghosts,
            suggested_area=suggested,
        )
        picker.ShowDialog()
        result = picker.result
        if result is None:
            return

        # Convert result (Revit bottom-up Y) to canvas pixels (top-down Y)
        ch  = self._sheet_h * PX_PER_INCH
        h_px = result["h_in"] * PX_PER_INCH
        self._usable_area = {
            "x": result["x_in"] * PX_PER_INCH,
            "y": ch - result["y_in"] * PX_PER_INCH - h_px,
            "w": result["w_in"] * PX_PER_INCH,
            "h": h_px,
        }

        # Persist: save area + whatever margins were active in the picker
        final_margins = {
            "margin_l":   picker._ml,
            "margin_r":   picker._mr,
            "margin_t":   picker._mt,
            "margin_b":   picker._mb,
            "title_h":    picker._th,
            "revision_w": picker._rw,
        }
        _save_tb_area_config(tb_id, result, final_margins)

        # Update the main window display
        self._update_usable_area_display()



    def ClearArea_Click(self, sender, e):
        """Remove the defined usable area."""
        self._usable_area = None
        if self._usable_area_rect is not None:
            try:
                self.SheetCanvas.Children.Remove(self._usable_area_rect)
            except Exception:
                pass
            self._usable_area_rect = None
        self.AreaInfoLabel.Text = "No usable area defined (full sheet used)"
        self.ClearAreaBtn.Visibility = Visibility.Collapsed

    def _update_usable_area_display(self):
        """Redraw background (which re-stacks viewport Borders on top),
        then append the dashed usable-area rect as the final child so it
        sits above both background and viewport Borders.
        """
        # Drop any old usable-area rect
        if self._usable_area_rect is not None:
            try:
                self._usable_area_rect.Tag = None
                self.SheetCanvas.Children.Remove(self._usable_area_rect)
            except Exception:
                pass
            self._usable_area_rect = None

        # Repaint background (also re-appends viewport Borders on top)
        self._draw_sheet_outline()

        if self._usable_area is None:
            self.AreaInfoLabel.Text = "No usable area defined (full sheet used)"
            self.ClearAreaBtn.Visibility = Visibility.Collapsed
            return

        # Append the dashed rect LAST -- on top of everything
        ua = self._usable_area
        _dash = DoubleCollection()
        _dash.Add(6.0); _dash.Add(3.0)
        rect = Rectangle()
        rect.Width  = ua["w"]
        rect.Height = ua["h"]
        rect.Stroke = SolidColorBrush(Color.FromRgb(37, 99, 235))
        rect.StrokeThickness = 3
        rect.StrokeDashArray = _dash
        rect.Fill = SolidColorBrush(Color.FromArgb(25, 37, 99, 235))
        rect.Tag  = "usable_area"
        rect.IsHitTestVisible = False
        Canvas.SetLeft(rect, ua["x"])
        Canvas.SetTop(rect,  ua["y"])
        self.SheetCanvas.Children.Add(rect)
        self._usable_area_rect = rect

        w_in = ua["w"] / PX_PER_INCH
        h_in = ua["h"] / PX_PER_INCH
        self.AreaInfoLabel.Text = 'Usable area: {:.2f}" x {:.2f}"'.format(w_in, h_in)
        self.ClearAreaBtn.Visibility = Visibility.Visible
        self.SheetCanvas.UpdateLayout()

    def _on_scale_changed(self, sender, e):
        """Rebuild the canvas rect for a view when its scale combo changes.

        Keeps the current top-left position so the user doesn't lose placement,
        but resizes the rectangle to reflect the view's new printed footprint.
        """
        for row in self._view_rows:
            if row.get("scale_combo") is not sender:
                continue
            vid_int = row["view_id"].Value
            old_border = self._canvas_rects.get(vid_int)
            if old_border is None:
                break
            # Remember current top-left so we restore it after resize
            old_x = Canvas.GetLeft(old_border)
            old_y = Canvas.GetTop(old_border)
            try:
                self.SheetCanvas.Children.Remove(old_border)
            except Exception:
                pass
            del self._canvas_rects[vid_int]
            # Re-add at new size
            self._add_canvas_rect(row)
            new_border = self._canvas_rects.get(vid_int)
            if new_border:
                cw = self.SheetCanvas.Width
                ch = self.SheetCanvas.Height
                x = max(0, min(old_x, cw - new_border.Width))
                y = max(0, min(old_y, ch - new_border.Height))
                Canvas.SetLeft(new_border, x)
                Canvas.SetTop(new_border, y)
            break

    def _add_canvas_rect(self, row_data):
        """Add a draggable viewport rectangle to the canvas.

        Size is calculated from the view's actual CropBox at the selected
        scale (scale combo value). Falls back to proportional sizing when
        the crop box is unavailable (schedules, uncropped 3D views, etc.).
        """
        vkey  = row_data["view_key"]
        vid   = row_data["view_id"]
        vname = row_data["view_name"]
        key   = vid.Value

        if key in self._canvas_rects:
            return

        color = VP_COLORS.get(vkey, VP_COLORS["other"])
        brush = SolidColorBrush(color)
        bg    = Color.FromArgb(40, color.R, color.G, color.B)

        # Reference area for fallback proportional sizing
        if self._usable_area:
            ref_w    = self._usable_area["w"]
            ref_h    = self._usable_area["h"]
            origin_x = self._usable_area["x"]
            origin_y = self._usable_area["y"]
        else:
            ref_w    = self.SheetCanvas.Width
            ref_h    = self.SheetCanvas.Height
            origin_x = 0
            origin_y = 0

        # Attempt accurate sizing
        rect_w = rect_h = None
        if row_data["is_schedule"]:
            size = get_schedule_sheet_size_px(vid)
            if size:
                rect_w, rect_h = size
        else:
            scale_factor = 48
            sc = row_data.get("scale_combo")
            if sc and sc.SelectedItem and hasattr(sc.SelectedItem, "Tag"):
                try: scale_factor = int(sc.SelectedItem.Tag)
                except: pass
            size = get_view_sheet_size_px(vid, scale_factor, self._assembly_id)
            if size:
                rect_w, rect_h = size

        # Snap to grid; fall back to proportional if size query returned nothing
        if rect_w and rect_h:
            rect_w = max(SNAP_PX, round(rect_w / SNAP_PX) * SNAP_PX)
            rect_h = max(SNAP_PX, round(rect_h / SNAP_PX) * SNAP_PX)
        else:
            # Proportional fallback relative to reference area
            ratio_w = 0.50 if row_data["is_schedule"] else 0.45
            ratio_h = 0.30 if row_data["is_schedule"] else 0.40
            rect_w = max(SNAP_PX, round(ref_w * ratio_w / SNAP_PX) * SNAP_PX)
            rect_h = max(SNAP_PX, round(ref_h * ratio_h / SNAP_PX) * SNAP_PX)

        border = Border()
        border.Width   = rect_w
        border.Height  = rect_h
        border.Background    = SolidColorBrush(bg)
        border.BorderBrush   = brush
        border.BorderThickness = Thickness(2)
        border.CornerRadius  = System.Windows.CornerRadius(3)
        border.Tag    = "viewport"
        border.Cursor = System.Windows.Input.Cursors.SizeAll

        lbl = TextBlock()
        lbl.Text       = vname
        lbl.FontSize   = max(18, min(30, ref_w / 70))
        lbl.Foreground = brush
        lbl.FontWeight = System.Windows.FontWeights.SemiBold
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment   = VerticalAlignment.Center
        lbl.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
        lbl.TextWrapping = System.Windows.TextWrapping.Wrap
        lbl.TextAlignment = TextAlignment.Center
        lbl.Margin = Thickness(4)
        lbl.IsHitTestVisible = False
        border.Child = lbl

        # Default position: schedules snap to bottom-right of usable area;
        # views fill a 2-column grid from top-left.
        gap_x = max(SNAP_PX, round(ref_w * 0.03 / SNAP_PX) * SNAP_PX)
        gap_y = max(SNAP_PX, round(ref_h * 0.04 / SNAP_PX) * SNAP_PX)

        if row_data["is_schedule"]:
            # Bottom-right corner of usable area, with gap inset
            x = origin_x + ref_w - rect_w - gap_x
            y = origin_y + ref_h - rect_h - gap_y
        else:
            # Count only non-schedule views already placed
            view_idx = sum(
                1 for k, b in self._canvas_rects.items()
                if getattr(b, "Tag", None) == "viewport"
                and not any(
                    r["view_id"].Value == k and r["is_schedule"]
                    for r in self._view_rows
                )
            )
            col   = view_idx % 2
            row_i = view_idx // 2
            x = origin_x + gap_x + col * (rect_w + gap_x)
            y = origin_y + gap_y + row_i * (rect_h + gap_y)

        cw = self.SheetCanvas.Width
        ch = self.SheetCanvas.Height
        x = max(0, min(x, cw - rect_w))
        y = max(0, min(y, ch - rect_h))

        Canvas.SetLeft(border, x)
        Canvas.SetTop(border, y)

        border.MouseLeftButtonDown += self._on_vp_mouse_down
        border.MouseMove           += self._on_vp_mouse_move
        border.MouseLeftButtonUp   += self._on_vp_mouse_up

        self.SheetCanvas.Children.Add(border)
        self._canvas_rects[key] = border

    def _reposition_all_rects(self):
        """Clamp all viewport rects after canvas resize."""
        cw = self.SheetCanvas.Width
        ch = self.SheetCanvas.Height
        for key, border in self._canvas_rects.items():
            x = max(0, min(Canvas.GetLeft(border), cw - border.Width - 2))
            y = max(0, min(Canvas.GetTop(border),  ch - border.Height - 2))
            Canvas.SetLeft(border, x)
            Canvas.SetTop(border, y)

    def _on_vp_mouse_down(self, sender, e):
        self._dragging    = sender
        self._drag_offset = e.GetPosition(sender)
        sender.CaptureMouse()
        e.Handled = True

    def _on_vp_mouse_move(self, sender, e):
        if self._dragging is None:
            return
        pos = e.GetPosition(self.SheetCanvas)
        cw  = self.SheetCanvas.Width
        ch  = self.SheetCanvas.Height
        rw  = self._dragging.Width
        rh  = self._dragging.Height

        # Snap to 1/4" grid
        raw_x = pos.X - self._drag_offset.X
        raw_y = pos.Y - self._drag_offset.Y
        x = round(raw_x / SNAP_PX) * SNAP_PX
        y = round(raw_y / SNAP_PX) * SNAP_PX

        # Clamp to usable area if defined, else full canvas
        if self._usable_area:
            ua = self._usable_area
            x = max(ua["x"], min(x, ua["x"] + ua["w"] - rw))
            y = max(ua["y"], min(y, ua["y"] + ua["h"] - rh))
        else:
            x = max(0, min(x, cw - rw))
            y = max(0, min(y, ch - rh))

        Canvas.SetLeft(self._dragging, x)
        Canvas.SetTop(self._dragging, y)

        cx = (x + rw / 2.0) / PX_PER_INCH
        cy = self._sheet_h - (y + rh / 2.0) / PX_PER_INCH
        self.CoordLabel.Text = 'Center: {:.2f}", {:.2f}" from bottom-left'.format(cx, cy)
        e.Handled = True

    def _on_vp_mouse_up(self, sender, e):
        if self._dragging:
            self._dragging.ReleaseMouseCapture()
            self._dragging = None
        e.Handled = True

    # ── Titleblock parameters ────────────────────────────────────────

    def ParamHeader_Toggle(self, sender, e):
        """Collapse / expand the TITLEBLOCK PARAMETERS body."""
        if self.ParamBody.Visibility == Visibility.Visible:
            self.ParamBody.Visibility = Visibility.Collapsed
            self.ParamChevron.Text    = "\u25ba"   # ▶ collapsed
            self.ParamHeaderBorder.CornerRadius = System.Windows.CornerRadius(5)
            self._update_param_badge(show=True)
        else:
            self.ParamBody.Visibility = Visibility.Visible
            self.ParamChevron.Text    = "\u25bc"   # ▼ expanded
            self.ParamHeaderBorder.CornerRadius = System.Windows.CornerRadius(5, 5, 0, 0)
            self.ParamCountBadge.Visibility = Visibility.Collapsed

    def _update_param_badge(self, show=False):
        """Update the filled-count badge shown in the collapsed header."""
        if not show:
            self.ParamCountBadge.Visibility = Visibility.Collapsed
            return
        filled = sum(1 for tb in self._param_boxes.values() if tb.Text.strip())
        total  = len(self._param_boxes)
        if total == 0:
            self.ParamCountBadge.Text = ""
            self.ParamCountBadge.Visibility = Visibility.Collapsed
        else:
            self.ParamCountBadge.Text = "{}/{} filled".format(filled, total)
            self.ParamCountBadge.Visibility = Visibility.Visible

    def _populate_params(self, tb_type_id):
        self.ParamRowsPanel.Children.Clear()
        self._param_boxes = {}
        params = get_titleblock_params(tb_type_id)
        if not params:
            self.NoParamsLabel.Visibility = Visibility.Visible
            return
        self.NoParamsLabel.Visibility = Visibility.Collapsed
        for pname in params:
            row = Grid()
            row.Margin = Thickness(0, 0, 0, 6)
            c0 = ColumnDefinition()
            c0.Width = System.Windows.GridLength(100)
            c1 = ColumnDefinition()
            c1.Width = System.Windows.GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(c0)
            row.ColumnDefinitions.Add(c1)

            lbl = TextBlock()
            lbl.Text = pname
            lbl.Foreground = SolidColorBrush(Color.FromRgb(107, 114, 128))
            lbl.FontSize = 11
            lbl.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(lbl, 0)
            row.Children.Add(lbl)

            tb = TextBox()
            tb.FontSize = 11
            tb.Padding = Thickness(4, 3, 4, 3)
            tb.BorderBrush = SolidColorBrush(Color.FromRgb(209, 213, 219))
            tb.BorderThickness = Thickness(1)
            Grid.SetColumn(tb, 1)
            row.Children.Add(tb)

            self.ParamRowsPanel.Children.Add(row)
            self._param_boxes[pname] = tb

        # Section starts collapsed -- show badge immediately so user knows
        # parameters are available without having to expand.
        self._update_param_badge(show=True)

    # ── Presets ──────────────────────────────────────────────────────

    def _refresh_presets(self):
        self.PresetCombo.Items.Clear()
        for name in list_presets():
            item = ComboBoxItem()
            item.Content = name
            self.PresetCombo.Items.Add(item)
        if self.PresetCombo.Items.Count > 0:
            self.PresetCombo.SelectedIndex = 0

    def _build_preset_data(self):
        slots = []
        cw = self.SheetCanvas.Width
        ch = self.SheetCanvas.Height
        for row in self._view_rows:
            if row["checkbox"].IsChecked != True:
                continue
            vid_int = row["view_id"].Value
            border  = self._canvas_rects.get(vid_int)
            if border is None:
                continue
            cx = (Canvas.GetLeft(border) + border.Width  / 2.0) / cw
            cy = (Canvas.GetTop(border)  + border.Height / 2.0) / ch
            scale   = None
            vp_type = None
            if row["scale_combo"]:
                si = row["scale_combo"].SelectedItem
                if si and hasattr(si, "Tag"):
                    scale = si.Tag
            if row["vp_type_combo"]:
                vi = row["vp_type_combo"].SelectedItem
                if vi and hasattr(vi, "Tag"):
                    tid = vi.Tag
                    if tid != ElementId.InvalidElementId:
                        vt = doc.GetElement(tid)
                        if vt:
                            vp_type = vt.Name
            slots.append({
                "view_key":   row["view_key"],
                "x_fraction": round(cx, 4),
                "y_fraction": round(cy, 4),
                "scale":      scale,
                "viewport_type": vp_type,
                "is_schedule":   row["is_schedule"],
            })
        params = {pname: tb.Text.strip()
                  for pname, tb in self._param_boxes.items()
                  if tb.Text.strip()}
        tb_name = ""
        if self.TitleBlockCombo.SelectedItem:
            tb_name = self.TitleBlockCombo.SelectedItem.Content
        return {
            "name": "Preset", "version": 1,
            "titleblock": tb_name,
            "slots": slots, "parameters": params,
        }

    def _apply_preset(self, preset):
        if not preset:
            return
        tb_name = preset.get("titleblock", "")
        if tb_name:
            for i in range(self.TitleBlockCombo.Items.Count):
                if self.TitleBlockCombo.Items[i].Content == tb_name:
                    self.TitleBlockCombo.SelectedIndex = i
                    break
        for pname, val in preset.get("parameters", {}).items():
            if pname in self._param_boxes:
                self._param_boxes[pname].Text = val
        cw = self.SheetCanvas.Width
        ch = self.SheetCanvas.Height
        for pslot in preset.get("slots", []):
            pkey = pslot.get("view_key", "")
            for row in self._view_rows:
                if row["view_key"] != pkey:
                    continue
                row["checkbox"].IsChecked = True
                if pslot.get("scale") and row["scale_combo"]:
                    for i in range(row["scale_combo"].Items.Count):
                        si = row["scale_combo"].Items[i]
                        if hasattr(si, "Tag") and si.Tag == pslot["scale"]:
                            row["scale_combo"].SelectedIndex = i
                            break
                vp_type_name = pslot.get("viewport_type")
                if vp_type_name and row["vp_type_combo"]:
                    for i in range(row["vp_type_combo"].Items.Count):
                        vi = row["vp_type_combo"].Items[i]
                        if vi.Content == vp_type_name:
                            row["vp_type_combo"].SelectedIndex = i
                            break
                vid_int = row["view_id"].Value
                border  = self._canvas_rects.get(vid_int)
                if border:
                    x = pslot["x_fraction"] * cw - border.Width  / 2.0
                    y = pslot["y_fraction"] * ch - border.Height / 2.0
                    Canvas.SetLeft(border, max(0, min(x, cw - border.Width)))
                    Canvas.SetTop(border,  max(0, min(y, ch - border.Height)))
                break

    def LoadPreset_Click(self, sender, e):
        sel = self.PresetCombo.SelectedItem
        if not sel:
            forms.alert("Select a preset first.", exitscript=False)
            return
        preset = load_preset(sel.Content)
        if preset:
            self._apply_preset(preset)
        else:
            forms.alert("Could not load preset.", exitscript=False)

    def SavePreset_Click(self, sender, e):
        name = forms.ask_for_string(
            prompt="Preset name:", title="Save Preset", default="My Layout")
        if not name:
            return
        data = self._build_preset_data()
        data["name"] = name
        save_preset(name, data)
        self._refresh_presets()
        for i in range(self.PresetCombo.Items.Count):
            if self.PresetCombo.Items[i].Content == name:
                self.PresetCombo.SelectedIndex = i
                break

    def DeletePreset_Click(self, sender, e):
        sel = self.PresetCombo.SelectedItem
        if not sel:
            return
        if forms.alert("Delete preset '{}'?".format(sel.Content), yes=True, no=True):
            delete_preset(sel.Content)
            self._refresh_presets()

    # ── Profiles ─────────────────────────────────────────────────────

    def _refresh_profiles(self):
        self.ProfileCombo.Items.Clear()
        for name in list_profiles():
            item = ComboBoxItem()
            item.Content = name
            self.ProfileCombo.Items.Add(item)
        if self.ProfileCombo.Items.Count > 0:
            self.ProfileCombo.SelectedIndex = 0

    def _build_profile_data(self):
        """Capture full tool state into a portable profile dict."""
        # View slots: positions stored as fraction of usable area (or canvas)
        ref_w = self._usable_area["w"] if self._usable_area else self.SheetCanvas.Width
        ref_h = self._usable_area["h"] if self._usable_area else self.SheetCanvas.Height
        ox    = self._usable_area["x"] if self._usable_area else 0.0
        oy    = self._usable_area["y"] if self._usable_area else 0.0

        slots = []
        for row in self._view_rows:
            vid_int = row["view_id"].Value
            border  = self._canvas_rects.get(vid_int)

            scale    = None
            vp_type  = None
            if row.get("scale_combo") and row["scale_combo"].SelectedItem:
                si = row["scale_combo"].SelectedItem
                if hasattr(si, "Tag"):
                    scale = si.Tag
            if row.get("vp_type_combo") and row["vp_type_combo"].SelectedItem:
                vi = row["vp_type_combo"].SelectedItem
                if hasattr(vi, "Tag") and vi.Tag != ElementId.InvalidElementId:
                    vt = doc.GetElement(vi.Tag)
                    if vt:
                        vp_type = vt.Name

            x_frac = y_frac = None
            if border:
                cx = Canvas.GetLeft(border) + border.Width  / 2.0
                cy = Canvas.GetTop(border)  + border.Height / 2.0
                x_frac = round((cx - ox) / ref_w, 4) if ref_w else 0.5
                y_frac = round((cy - oy) / ref_h, 4) if ref_h else 0.5

            slots.append({
                "view_key":      row["view_key"],
                "view_name":     row["view_name"],
                "is_schedule":   row["is_schedule"],
                "enabled":       row["checkbox"].IsChecked == True,
                "scale":         scale,
                "viewport_type": vp_type,
                "x_frac":        x_frac,
                "y_frac":        y_frac,
            })

        # Titleblock
        tb_name = ""
        tb_id_int = None
        if self.TitleBlockCombo.SelectedItem:
            tb_name = self.TitleBlockCombo.SelectedItem.Content
            tb_id = self._combo_tag(self.TitleBlockCombo)
            if tb_id != ElementId.InvalidElementId:
                tb_id_int = tb_id.Value

        # Usable area in sheet-relative inches
        ua_in = None
        if self._usable_area:
            ch = self.SheetCanvas.Height
            ua = self._usable_area
            ua_in = {
                "x_in": ua["x"] / PX_PER_INCH,
                "y_in": (ch - ua["y"] - ua["h"]) / PX_PER_INCH,
                "w_in": ua["w"] / PX_PER_INCH,
                "h_in": ua["h"] / PX_PER_INCH,
            }

        tb_params = {pname: tb.Text.strip()
                     for pname, tb in self._param_boxes.items()
                     if tb.Text.strip()}

        return {
            "version":       2,
            "titleblock":    tb_name,
            "titleblock_id": tb_id_int,
            "sheet_number":  self.SheetNumberBox.Text.strip(),
            "usable_area":   ua_in,
            "tb_margins":    self._tb_margins,
            "tb_params":     tb_params,
            "slots":         slots,
        }

    def _apply_profile_slots(self, slots_with_rows):
        """Apply a list of (slot_dict, row_dict) pairs to the canvas.
        slot_dict comes from the profile; row_dict is the matched current row.
        """
        ref_w = self._usable_area["w"] if self._usable_area else self.SheetCanvas.Width
        ref_h = self._usable_area["h"] if self._usable_area else self.SheetCanvas.Height
        ox    = self._usable_area["x"] if self._usable_area else 0.0
        oy    = self._usable_area["y"] if self._usable_area else 0.0
        cw    = self.SheetCanvas.Width
        ch    = self.SheetCanvas.Height

        for slot, row in slots_with_rows:
            if row is None:
                continue

            # Enable/disable
            if row.get("checkbox"):
                row["checkbox"].IsChecked = slot.get("enabled", True)

            # Scale
            if slot.get("scale") and row.get("scale_combo"):
                for i in range(row["scale_combo"].Items.Count):
                    si = row["scale_combo"].Items[i]
                    if hasattr(si, "Tag") and si.Tag == slot["scale"]:
                        row["scale_combo"].SelectedIndex = i
                        break

            # Viewport type
            vp_name = slot.get("viewport_type")
            if vp_name and row.get("vp_type_combo"):
                for i in range(row["vp_type_combo"].Items.Count):
                    vi = row["vp_type_combo"].Items[i]
                    if vi.Content == vp_name:
                        row["vp_type_combo"].SelectedIndex = i
                        break

            # Canvas position (from usable-area fractions)
            if slot.get("x_frac") is not None and slot.get("y_frac") is not None:
                vid_int = row["view_id"].Value
                border  = self._canvas_rects.get(vid_int)
                if border:
                    cx = ox + slot["x_frac"] * ref_w
                    cy = oy + slot["y_frac"] * ref_h
                    x  = max(0, min(cx - border.Width  / 2.0, cw - border.Width))
                    y  = max(0, min(cy - border.Height / 2.0, ch - border.Height))
                    Canvas.SetLeft(border, x)
                    Canvas.SetTop(border,  y)

    def _apply_profile(self, profile):
        """Load a profile onto the current window state.

        Matches slots to current view rows by (view_key, view_name).
        Unresolved slots are passed to ProfileMapperWindow.
        """
        if not profile:
            return

        # 1. Titleblock
        tb_name = profile.get("titleblock", "")
        if tb_name:
            for i in range(self.TitleBlockCombo.Items.Count):
                if self.TitleBlockCombo.Items[i].Content == tb_name:
                    self.TitleBlockCombo.SelectedIndex = i
                    break

        # 2. Usable area
        ua_in = profile.get("usable_area")
        if ua_in:
            ch  = self.SheetCanvas.Height
            h_px = ua_in["h_in"] * PX_PER_INCH
            self._usable_area = {
                "x": ua_in["x_in"] * PX_PER_INCH,
                "y": ch - ua_in["y_in"] * PX_PER_INCH - h_px,
                "w": ua_in["w_in"] * PX_PER_INCH,
                "h": h_px,
            }
            self._update_usable_area_display()

        # 3. Sheet number / params
        if profile.get("sheet_number"):
            self.SheetNumberBox.Text = profile["sheet_number"]
        for pname, val in profile.get("tb_params", {}).items():
            if pname in self._param_boxes:
                self._param_boxes[pname].Text = val

        # 4. Match slots → current rows
        resolved   = []   # (slot, row)
        unresolved = []   # slot (index in unresolved list)

        for slot in profile.get("slots", []):
            matched = None
            # Exact match: same key AND same name
            for row in self._view_rows:
                if (row["view_key"]  == slot.get("view_key") and
                        row["view_name"] == slot.get("view_name")):
                    matched = row
                    break
            # Fallback: same key only
            if matched is None:
                for row in self._view_rows:
                    if row["view_key"] == slot.get("view_key"):
                        matched = row
                        break
            if matched is not None:
                resolved.append((slot, matched))
            else:
                unresolved.append(slot)

        # 5. If there are unresolved slots, show the mapper
        if unresolved:
            mapper = ProfileMapperWindow(unresolved, self._view_rows)
            mapper.ShowDialog()
            if mapper.result is None:
                # User cancelled the mapper -- still apply resolved slots
                pass
            else:
                for i, row in mapper.result.items():
                    resolved.append((unresolved[i], row))

        # 6. Apply all resolved pairs
        self._apply_profile_slots(resolved)

    def ImportProfile_Click(self, sender, e):
        sel = self.ProfileCombo.SelectedItem
        if not sel:
            forms.alert("Select a profile first.", exitscript=False)
            return
        profile = load_profile(sel.Content)
        if profile:
            self._apply_profile(profile)
        else:
            forms.alert("Could not load profile.", exitscript=False)

    def ExportProfile_Click(self, sender, e):
        name = forms.ask_for_string(
            prompt="Profile name:", title="Save Profile",
            default="My Assembly Profile")
        if not name:
            return
        data = self._build_profile_data()
        data["name"] = name
        save_profile(name, data)
        self._refresh_profiles()
        for i in range(self.ProfileCombo.Items.Count):
            if self.ProfileCombo.Items[i].Content == name:
                self.ProfileCombo.SelectedIndex = i
                break
        forms.alert("Profile '{}' saved.".format(name), exitscript=False)

    def DeleteProfile_Click(self, sender, e):
        sel = self.ProfileCombo.SelectedItem
        if not sel:
            return
        if forms.alert("Delete profile '{}'?".format(sel.Content), yes=True, no=True):
            delete_profile(sel.Content)
            self._refresh_profiles()



    def _combo_tag(self, combo):
        sel = combo.SelectedItem
        return sel.Tag if (sel and hasattr(sel, "Tag")) else ElementId.InvalidElementId

    # ── Footer actions ───────────────────────────────────────────────

    def BatchCreate_Click(self, sender, e):
        """Open batch create window and create one sheet per selected assembly."""
        tb_id = self._combo_tag(self.TitleBlockCombo)
        if tb_id == ElementId.InvalidElementId:
            forms.alert("Select a titleblock first.", exitscript=False)
            return

        all_asm = get_assemblies()
        if not all_asm:
            forms.alert("No assemblies found in this project.", exitscript=False)
            return

        current_params = {pname: tb.Text.strip()
                          for pname, tb in self._param_boxes.items()
                          if tb.Text.strip()}

        batch_win = BatchCreateWindow(
            all_asm,
            self.SheetNumberBox.Text.strip(),
            self.SheetNameBox.Text.strip(),
            current_params,
        )
        batch_win.ShowDialog()
        rows = batch_win.result
        if not rows:
            return

        # Build the slot config from current canvas state (same as single sheet)
        tb_origin_x_ft, tb_origin_y_ft = get_titleblock_sheet_origin(tb_id)
        sheet_h_px = self.SheetCanvas.Height

        slots = []
        for row in self._view_rows:
            if row["checkbox"].IsChecked != True:
                continue
            vid     = row["view_id"]
            vid_int = vid.Value
            border  = self._canvas_rects.get(vid_int)
            if border is None:
                continue
            left_px = Canvas.GetLeft(border)
            top_px  = Canvas.GetTop(border)
            if row["is_schedule"]:
                paper_x_in = left_px / PX_PER_INCH
                paper_y_in = (sheet_h_px - top_px) / PX_PER_INCH
            else:
                paper_x_in = (left_px + border.Width  / 2.0) / PX_PER_INCH
                paper_y_in = (sheet_h_px - (top_px + border.Height / 2.0)) / PX_PER_INCH
            x_ft = tb_origin_x_ft + paper_x_in / 12.0
            y_ft = tb_origin_y_ft + paper_y_in / 12.0
            scale = None
            vp_type_id = ElementId.InvalidElementId
            if row["scale_combo"]:
                si = row["scale_combo"].SelectedItem
                if si and hasattr(si, "Tag"): scale = si.Tag
            if row["vp_type_combo"]:
                vi = row["vp_type_combo"].SelectedItem
                if vi and hasattr(vi, "Tag"): vp_type_id = vi.Tag
            slots.append({
                "view_id":          vid,
                "x":                x_ft,
                "y":                y_ft,
                "scale":            scale,
                "viewport_type_id": vp_type_id,
                "is_schedule":      row["is_schedule"],
            })

        dest = "assembly" if self.RbLocAssembly.IsChecked == True else "sheets"
        folder_val = ""
        if dest == "sheets":
            sel = self.SheetFolderCombo.SelectedItem
            folder_val = (sel.Content if (sel and hasattr(sel, "Content"))
                          else self.SheetFolderCombo.Text.strip())

        # Create one sheet per assembly row
        created = []; errors_all = []
        for row_data in rows:
            # Each assembly gets its own views -- load them fresh
            aid  = row_data["assembly_id"]
            views = get_assembly_views(aid)
            if not views:
                errors_all.append(
                    "{}: no views found, skipped".format(row_data["sheet_name"]))
                continue

            # Remap slots to this assembly's views by view_key
            remapped_slots = []
            for slot in slots:
                # Find the matching view in this assembly
                orig_vid = slot["view_id"]
                orig_row = next(
                    (r for r in self._view_rows
                     if r["view_id"].Value == orig_vid.Value), None)
                if orig_row is None:
                    continue
                vkey = orig_row["view_key"]
                match = next(
                    (v for v in views if v[0] == vkey), None)
                if match:
                    new_slot = dict(slot)
                    new_slot["view_id"]     = match[2]
                    new_slot["is_schedule"] = match[3]
                    remapped_slots.append(new_slot)

            if not remapped_slots:
                errors_all.append(
                    "{}: no matching views, skipped".format(row_data["sheet_name"]))
                continue

            config = {
                "title_block_id":      tb_id,
                "sheet_number":        row_data["sheet_number"],
                "sheet_name":          row_data["sheet_name"],
                "slots":               remapped_slots,
                "parameters":          current_params,
                "destination":         dest,
                "assembly_id":         aid,
                "browser_folder_param": getattr(self, "_browser_folder_param", None),
                "browser_folder_value": folder_val,
            }
            sheet, errs = create_sheet(config)
            if sheet:
                created.append("{} - {}".format(
                    row_data["sheet_number"], row_data["sheet_name"]))
            else:
                errors_all.extend(errs)

        # Report
        msg = "Created {} sheet(s):\n{}".format(
            len(created), "\n".join(created))
        if errors_all:
            msg += "\n\nWarnings:\n" + "\n".join(errors_all)
        forms.alert(msg, exitscript=False)
        self._batch_completed = True
        self.Close()

    def Cancel_Click(self, sender, e):
        self._result = None
        self.Close()

    def CreateSheet_Click(self, sender, e):
        if self._assembly_id is None:
            if self.RbCreate.IsChecked == True:
                forms.alert(
                    "Click 'Create Assembly' first to build the assembly before creating the sheet.",
                    exitscript=False)
            else:
                forms.alert("Select an assembly first.", exitscript=False)
            return

        tb_id = self._combo_tag(self.TitleBlockCombo)
        if tb_id == ElementId.InvalidElementId:
            forms.alert("Select a titleblock.", exitscript=False)
            return

        sheet_num = self.SheetNumberBox.Text.strip()
        if not sheet_num:
            forms.alert("Enter a sheet number.", exitscript=False)
            return

        sheet_name = self.SheetNameBox.Text.strip()
        if not sheet_name:
            forms.alert("Enter a sheet name.", exitscript=False)
            return

        # Duplicate sheet number check
        for s in (FilteredElementCollector(doc)
                  .OfClass(ViewSheet)
                  .WhereElementIsNotElementType()):
            try:
                if s.SheetNumber.lower() == sheet_num.lower():
                    forms.alert(
                        "Sheet number '{}' already exists.".format(sheet_num),
                        exitscript=False)
                    return
            except Exception:
                pass

        # Collect placed view slots from canvas
        # Canvas (0,0) = top-left of paper.
        # Revit sheet coords origin = wherever the titleblock family puts it.
        # We query that offset so placements land exactly where the preview shows.
        tb_origin_x_ft, tb_origin_y_ft = get_titleblock_sheet_origin(tb_id)
        sheet_h_px = self.SheetCanvas.Height  # sheet_h * PX_PER_INCH

        slots = []
        for row in self._view_rows:
            if row["checkbox"].IsChecked != True:
                continue
            vid     = row["view_id"]
            vid_int = vid.Value
            border  = self._canvas_rects.get(vid_int)
            if border is None:
                continue

            # Viewports use center point; schedules use top-left corner.
            # Compute both from the border position in canvas pixels.
            left_px = Canvas.GetLeft(border)
            top_px  = Canvas.GetTop(border)

            if row["is_schedule"]:
                # Top-left corner → paper inches → Revit sheet feet
                paper_x_in = left_px / PX_PER_INCH
                # Canvas Y is top-down; Revit Y is bottom-up.
                # Top of the border in paper inches from bottom = (sheet_h - top_px) / PPI
                paper_y_in = (sheet_h_px - top_px) / PX_PER_INCH
            else:
                # Center point
                paper_x_in = (left_px + border.Width  / 2.0) / PX_PER_INCH
                paper_y_in = (sheet_h_px - (top_px + border.Height / 2.0)) / PX_PER_INCH

            # Paper inches → Revit sheet feet (apply titleblock origin offset)
            x_ft = tb_origin_x_ft + paper_x_in / 12.0
            y_ft = tb_origin_y_ft + paper_y_in / 12.0
            scale = None
            vp_type_id = ElementId.InvalidElementId
            if row["scale_combo"]:
                si = row["scale_combo"].SelectedItem
                if si and hasattr(si, "Tag"):
                    scale = si.Tag
            if row["vp_type_combo"]:
                vi = row["vp_type_combo"].SelectedItem
                if vi and hasattr(vi, "Tag"):
                    vp_type_id = vi.Tag
            slots.append({
                "view_id":          vid,
                "x":                x_ft,
                "y":                y_ft,
                "scale":            scale,
                "viewport_type_id": vp_type_id,
                "is_schedule":      row["is_schedule"],
            })

        if not slots:
            forms.alert("Check at least one view to place on the sheet.", exitscript=False)
            return

        params = {pname: tb.Text.strip()
                  for pname, tb in self._param_boxes.items()
                  if tb.Text.strip()}

        # Sheet destination
        dest = "assembly" if self.RbLocAssembly.IsChecked == True else "sheets"
        folder_val = ""
        if dest == "sheets":
            sel = self.SheetFolderCombo.SelectedItem
            if sel and hasattr(sel, "Content"):
                folder_val = sel.Content
            else:
                folder_val = self.SheetFolderCombo.Text.strip()

        self._result = {
            "title_block_id":      tb_id,
            "sheet_number":        sheet_num,
            "sheet_name":          sheet_name,
            "slots":               slots,
            "parameters":          params,
            "destination":         dest,
            "assembly_id":         self._assembly_id,
            "browser_folder_param": getattr(self, "_browser_folder_param", None),
            "browser_folder_value": folder_val,
        }
        self.Close()

    @property
    def result(self):
        return self._result


# ─────────────────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────────────────

selected = get_selected_elements()
assembly_preselected = None
build_elements = []

for elem in selected:
    if isinstance(elem, AssemblyInstance):
        assembly_preselected = elem
    else:
        build_elements.append(elem)

# Route to the correct mode based on selection
if not build_elements and not assembly_preselected:
    assemblies = get_assemblies()
    if not assemblies:
        forms.alert(
            "No assemblies found in this project and no elements are selected.\n\n"
            "To use Assembly Studio:\n"
            "  - Select fabrication elements to build a new assembly, OR\n"
            "  - Create an assembly first, then run the tool to layout its sheet.",
            exitscript=False,
        )
        # End script naturally -- no SystemExit
        import sys as _sys
        _sys.exit = lambda *a, **k: None  # no-op to be safe

win = AssemblyStudioWindow(
    XAML_PATH,
    build_elements=build_elements if build_elements else None,
    preselected_assembly=assembly_preselected,
)
win.ShowDialog()

# If batch create was used, sheets were already committed inside the window.
# Let the script end naturally -- do NOT call script.exit() which raises
# SystemExit and causes pyRevit to roll back all transactions.
if win._batch_completed:
    pass  # done -- fall off the end of the script naturally

elif win.result is None:
    pass  # user cancelled -- nothing to do

else:
    config = win.result
    sheet, errors = create_sheet(config)

    if sheet is None:
        out.print_md("## Assembly Studio - Sheet Creation Failed")
        for e in errors:
            out.print_md("- **Error:** {}".format(e))
    else:
        out.print_md("## Assembly Studio - Complete")
        out.print_md("**Sheet:** {} - {} (ID {})".format(
            sheet.SheetNumber, sheet.Name, out.linkify(sheet.Id)))
        out.print_md("**Viewports placed:** {}".format(len(config["slots"])))
        if errors:
            out.print_md("---\n**Warnings:**")
            for e in errors:
                out.print_md("- {}".format(e))
