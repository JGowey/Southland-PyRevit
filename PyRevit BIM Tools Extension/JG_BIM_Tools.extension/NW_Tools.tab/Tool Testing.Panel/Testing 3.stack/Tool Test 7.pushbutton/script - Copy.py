# -*- coding: utf-8 -*-
"""
pyRevit button:
- Rebuild Space solids (DirectShape)
- Mark = '#SPACE_SOLID'
- Comments = Space Name (no number)
- Apply ACTIVE VIEW overrides:
    Surface FG/BG = Solid fill + unique color
    Cut FG/BG     = Solid fill + unique color
    Transparency  = set value
"""

from pyrevit import revit, DB, script
from System.Collections.Generic import List
import hashlib

doc = revit.doc
view = doc.ActiveView
output = script.get_output()

# ---------------- SETTINGS ----------------
DS_NAME_PREFIX = "SpaceSolid_"
GEN_TAG = "#SPACE_SOLID"
TARGET_CATEGORY = DB.BuiltInCategory.OST_GenericModel
ONLY_PLACED = True

APPLY_OVERRIDES_IN_ACTIVE_VIEW = True
TRANSPARENCY = 70  # 0..100
# ----------------------------------------


def get_spaces(document):
    return list(
        DB.FilteredElementCollector(document)
          .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
          .WhereElementIsNotElementType()
    )


def get_param_as_string(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p:
            return p.AsString() or ""
    except:
        pass
    return ""


def get_param_by_name(elem, pname):
    try:
        p = elem.LookupParameter(pname)
        if p:
            return p.AsString() or ""
    except:
        pass
    return ""


def get_space_name(space):
    # reliable "Name" for Spaces
    name = get_param_as_string(space, DB.BuiltInParameter.ROOM_NAME)
    if not name:
        name = get_param_by_name(space, "Name")
    if not name:
        try:
            name = space.Name or ""
        except:
            name = ""
    return (name or "").strip()


def set_text_param(elem, bip, value):
    p = elem.get_Parameter(bip)
    if p and not p.IsReadOnly:
        p.Set(value)


def delete_old_generated(document):
    ids = List[DB.ElementId]()
    for ds in DB.FilteredElementCollector(document).OfClass(DB.DirectShape):
        try:
            name_match = (ds.Name and ds.Name.startswith(DS_NAME_PREFIX))
            mark = get_param_as_string(ds, DB.BuiltInParameter.ALL_MODEL_MARK)
            mark_match = (mark == GEN_TAG)
            if name_match or mark_match:
                ids.Add(ds.Id)
        except:
            pass

    if ids.Count > 0:
        document.Delete(ids)
    return ids.Count


def can_override(v):
    try:
        return hasattr(v, "SetElementOverrides") and (not v.IsTemplate)
    except:
        return False


def find_solid_fill_pattern_id(document):
    """
    Find the FillPatternElement whose pattern is Solid fill.
    Works without FillPatternElement.GetSolidFillPatternId.
    """
    for fpe in DB.FilteredElementCollector(document).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            # In the API this is the reliable check
            if fp and fp.IsSolidFill:
                return fpe.Id
        except:
            pass
    return DB.ElementId.InvalidElementId


def color_from_text(text):
    """Deterministic unique-ish color per text."""
    if not text:
        return DB.Color(180, 180, 180)
    h = hashlib.md5(text.encode("utf-8")).hexdigest()
    r = 60 + (int(h[0:2], 16) % 160)
    g = 60 + (int(h[2:4], 16) % 160)
    b = 60 + (int(h[4:6], 16) % 160)
    return DB.Color(r, g, b)


def make_override(label_text, solid_fill_id):
    ogs = DB.OverrideGraphicSettings()
    col = color_from_text(label_text)

    # Transparency
    try:
        ogs.SetSurfaceTransparency(TRANSPARENCY)
    except:
        pass

    # Projection lines (helps edges pop)
    try:
        ogs.SetProjectionLineColor(col)
    except:
        pass

    # The important part: SOLID FILL patterns + color
    if solid_fill_id and solid_fill_id != DB.ElementId.InvalidElementId:
        # Surface
        try:
            ogs.SetSurfaceForegroundPatternId(solid_fill_id)
            ogs.SetSurfaceForegroundPatternColor(col)
        except:
            pass
        try:
            ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
            ogs.SetSurfaceBackgroundPatternColor(col)
        except:
            pass

        # Cut (optional but matches your screenshot “needs to be”)
        try:
            ogs.SetCutForegroundPatternId(solid_fill_id)
            ogs.SetCutForegroundPatternColor(col)
        except:
            pass
        try:
            ogs.SetCutBackgroundPatternId(solid_fill_id)
            ogs.SetCutBackgroundPatternColor(col)
        except:
            pass

    return ogs


# --------- main ----------
spaces = get_spaces(doc)
if not spaces:
    output.print_md("❌ No MEP Spaces found.")
    script.exit()

if APPLY_OVERRIDES_IN_ACTIVE_VIEW and not can_override(view):
    output.print_md("❌ Active view can't accept element overrides (or is a template). Open a 3D view and rerun.")
    script.exit()

solid_fill_id = find_solid_fill_pattern_id(doc)
if solid_fill_id == DB.ElementId.InvalidElementId:
    output.print_md("❌ Could not find 'Solid fill' pattern in this project.")
    output.print_md("   (Check that the project has a Solid fill pattern available.)")
    script.exit()

calc = DB.SpatialElementGeometryCalculator(doc)
cat_id = DB.ElementId(TARGET_CATEGORY)

t = DB.Transaction(doc, "Rebuild Space Solids + SolidFill Overrides")
t.Start()

deleted = created = skipped = failed = 0

try:
    deleted = delete_old_generated(doc)

    for sp in spaces:
        try:
            if ONLY_PLACED and sp.Location is None:
                skipped += 1
                continue

            results = calc.CalculateSpatialElementGeometry(sp)
            solid = results.GetGeometry()

            if solid is None or solid.Volume <= 1e-6:
                skipped += 1
                continue

            name_lbl = get_space_name(sp)  # Name only (no number)

            ds = DB.DirectShape.CreateElement(doc, cat_id)
            ds.Name = DS_NAME_PREFIX + (name_lbl or str(sp.Id.IntegerValue))
            ds.SetShape([solid])

            # Tag for cleanup
            set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)

            # What you want to see in properties/schedules
            set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

            # Apply the same overrides you manually set (solid fill + transparency)
            if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                ogs = make_override(name_lbl, solid_fill_id)
                view.SetElementOverrides(ds.Id, ogs)

            created += 1

        except Exception as e:
            failed += 1
            try:
                msg = e.Message
            except:
                msg = str(e)
            output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, msg))
            continue

    t.Commit()

except Exception as fatal:
    t.RollBack()
    output.print_md("❌ Transaction rolled back")
    output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
    raise

output.print_md("## Space Solids + Solid Fill Overrides")
output.print_md("- Deleted old solids: **{}**".format(deleted))
output.print_md("- Created: **{}**".format(created))
output.print_md("- Skipped: **{}**".format(skipped))
output.print_md("- Failed: **{}**".format(failed))
output.print_md("- Overrides applied in active view: **{}**".format("Yes" if APPLY_OVERRIDES_IN_ACTIVE_VIEW else "No"))
