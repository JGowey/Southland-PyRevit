# -*- coding: utf-8 -*-
"""
pyRevit button:
- Rebuild Space solids (DirectShape)
- Mark      = '#SPACE_SOLID' (ALL_MODEL_MARK)
- Comments  = Space Name (no number) (ALL_MODEL_INSTANCE_COMMENTS)
- NEW: Merge solids that have identical Space Names into a single "zone" DirectShape
       BUT ONLY when they physically touch/overlap (safeguard).

Optional:
- Apply ACTIVE VIEW overrides:
    Surface FG/BG = Solid fill + unique color
    Cut FG/BG     = Solid fill + unique color
    Transparency  = set value
"""

from pyrevit import revit, DB, script
from System.Collections.Generic import List
from collections import defaultdict
import hashlib

# ----------------- USER SETTINGS -----------------

GEN_TAG = "#SPACE_SOLID"
DS_NAME_PREFIX = "ZONE_"

# Target category for DirectShape (Generic Models is common)
TARGET_CATEGORY = int(DB.BuiltInCategory.OST_GenericModel)

# Only process placed spaces (Location exists)
ONLY_PLACED = True

# Apply overrides in current active view
APPLY_OVERRIDES_IN_ACTIVE_VIEW = True

# Transparency (0 = opaque, 100 = invisible). Typical: 40-70
TRANSPARENCY = 50

# Touch tolerance in Revit internal units (feet). ~1/16" ≈ 0.0052 ft
TOUCH_TOL = 0.005

# Minimum solid volume to keep
MIN_VOL = 1e-6

# ----------------- SETUP -----------------

doc = revit.doc
uidoc = revit.uidoc
view = doc.ActiveView
output = script.get_output()


# ----------------- HELPERS -----------------

def get_space_name(space):
    """Return space name only (no number)."""
    try:
        # Many MEP Spaces have Name and Number; "space.Name" is usually the name.
        nm = space.Name
        if nm:
            return nm.strip()
    except:
        pass

    # Fallback: parameter lookup
    p = space.LookupParameter("Name")
    if p and p.HasValue:
        try:
            return (p.AsString() or "").strip()
        except:
            pass
    return "Unnamed"


def set_text_param(elem, bip, textval):
    try:
        p = elem.get_Parameter(bip)
        if p and (not p.IsReadOnly):
            p.Set(textval or "")
            return True
    except:
        pass
    return False


def get_solid_fill_pattern_id(doc_):
    """Return the ElementId of the built-in Solid fill pattern (FillPatternElement)."""
    try:
        fpe = DB.FillPatternElement.GetFillPatternElementByName(doc_, DB.FillPatternTarget.Drafting, "Solid fill")
        if fpe:
            return fpe.Id
    except:
        pass

    # Fallback: search all fill patterns
    col = DB.FilteredElementCollector(doc_).OfClass(DB.FillPatternElement).ToElements()
    for e in col:
        try:
            fp = e.GetFillPattern()
            if fp and fp.IsSolidFill:
                return e.Id
        except:
            continue
    return DB.ElementId.InvalidElementId


def hash_to_color(name_lbl):
    """
    Deterministic unique-ish color per name.
    Revit Color expects 0-255 bytes.
    """
    s = (name_lbl or "Unnamed").encode("utf-8")
    h = hashlib.md5(s).hexdigest()
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)

    # Optional: keep colors from being too dark by biasing upward a bit
    # (still deterministic)
    r = int((r + 64) % 256)
    g = int((g + 64) % 256)
    b = int((b + 64) % 256)

    return DB.Color(r, g, b)


def make_override(name_lbl, solid_fill_id):
    """Create OverrideGraphicSettings with solid fill + unique color and transparency."""
    c = hash_to_color(name_lbl)

    ogs = DB.OverrideGraphicSettings()

    # Projection (surface)
    ogs.SetSurfaceForegroundPatternId(solid_fill_id)
    ogs.SetSurfaceForegroundPatternColor(c)
    ogs.SetSurfaceBackgroundPatternId(solid_fill_id)
    ogs.SetSurfaceBackgroundPatternColor(c)

    # Cut
    ogs.SetCutForegroundPatternId(solid_fill_id)
    ogs.SetCutForegroundPatternColor(c)
    ogs.SetCutBackgroundPatternId(solid_fill_id)
    ogs.SetCutBackgroundPatternColor(c)

    try:
        ogs.SetSurfaceTransparency(int(TRANSPARENCY))
    except:
        pass

    return ogs


def delete_old_generated(doc_):
    """Delete previously generated DirectShapes that match GEN_TAG in Mark param."""
    ids = List[DB.ElementId]()

    col = (DB.FilteredElementCollector(doc_)
           .OfClass(DB.DirectShape)
           .ToElements())

    for e in col:
        try:
            p = e.get_Parameter(DB.BuiltInParameter.ALL_MODEL_MARK)
            if p and p.HasValue and (p.AsString() == GEN_TAG):
                ids.Add(e.Id)
        except:
            continue

    if ids.Count > 0:
        doc_.Delete(ids)
    return ids.Count


def get_spaces(doc_):
    """Collect MEP Spaces."""
    # Spaces live in OST_MEPSpaces (not rooms)
    return (DB.FilteredElementCollector(doc_)
            .OfCategory(DB.BuiltInCategory.OST_MEPSpaces)
            .WhereElementIsNotElementType()
            .ToElements())


# --------- Merge logic helpers (touch+union safeguard) ---------

def bbox_of_solid(s):
    try:
        return s.GetBoundingBox()
    except:
        return None


def bbox_intersects_or_touches(bb1, bb2, tol=TOUCH_TOL):
    """Cheap reject test before attempting boolean union."""
    if bb1 is None or bb2 is None:
        return False

    return (
        (bb1.Min.X - tol) <= (bb2.Max.X + tol) and (bb2.Min.X - tol) <= (bb1.Max.X + tol) and
        (bb1.Min.Y - tol) <= (bb2.Max.Y + tol) and (bb2.Min.Y - tol) <= (bb1.Max.Y + tol) and
        (bb1.Min.Z - tol) <= (bb2.Max.Z + tol) and (bb2.Min.Z - tol) <= (bb1.Max.Z + tol)
    )


def try_union(a, b):
    """Attempt boolean union; return (True, unionSolid) or (False, None)."""
    try:
        u = DB.BooleanOperationsUtils.ExecuteBooleanOperation(a, b, DB.BooleanOperationsType.Union)
        if u and u.Volume > MIN_VOL:
            return True, u
    except:
        pass
    return False, None


def cluster_and_union_solids(solids):
    """
    Given a list of solids, return list of merged solids,
    where each output is a union of a connected cluster.

    Connectivity is defined as:
    - bounding boxes touch/overlap within TOUCH_TOL
    - AND a successful boolean union
    """
    remaining = [s for s in solids if s and s.Volume > MIN_VOL]
    merged_outputs = []

    while remaining:
        cluster = remaining.pop(0)
        cluster_bb = bbox_of_solid(cluster)

        changed = True
        while changed:
            changed = False
            i = len(remaining) - 1
            while i >= 0:
                s = remaining[i]
                bb = bbox_of_solid(s)

                if bbox_intersects_or_touches(cluster_bb, bb):
                    ok, u = try_union(cluster, s)
                    if ok:
                        cluster = u
                        cluster_bb = bbox_of_solid(cluster)
                        remaining.pop(i)
                        changed = True
                i -= 1

        merged_outputs.append(cluster)

    return merged_outputs


# ----------------- MAIN -----------------

spaces = get_spaces(doc)
if not spaces or len(spaces) == 0:
    output.print_md("❌ No MEP Spaces found in this model.")
    script.exit()

solid_fill_id = get_solid_fill_pattern_id(doc)
if solid_fill_id == DB.ElementId.InvalidElementId and APPLY_OVERRIDES_IN_ACTIVE_VIEW:
    output.print_md("❌ Could not find 'Solid fill' pattern in this project.")
    output.print_md("   (Check that the project has a Solid fill pattern available.)")
    script.exit()

calc = DB.SpatialElementGeometryCalculator(doc)
cat_id = DB.ElementId(TARGET_CATEGORY)

t = DB.Transaction(doc, "Rebuild Space Solids (Merged by Name)")
t.Start()

deleted = created = skipped = failed = 0

try:
    # 1) delete old
    deleted = delete_old_generated(doc)

    # 2) collect solids by Space Name
    solids_by_name = defaultdict(list)

    for sp in spaces:
        try:
            if ONLY_PLACED and sp.Location is None:
                skipped += 1
                continue

            results = calc.CalculateSpatialElementGeometry(sp)
            solid = results.GetGeometry()

            if solid is None or solid.Volume <= MIN_VOL:
                skipped += 1
                continue

            name_lbl = get_space_name(sp)
            solids_by_name[name_lbl].append(solid)

        except Exception as e:
            failed += 1
            try:
                msg = e.Message
            except:
                msg = str(e)
            output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, msg))
            continue

    # 3) for each name, merge only touching clusters and create DirectShapes
    for name_lbl, solids in solids_by_name.items():
        merged_solids = cluster_and_union_solids(solids)

        for idx, msolid in enumerate(merged_solids):
            ds = DB.DirectShape.CreateElement(doc, cat_id)

            # If multiple disconnected clusters share same name, make DS names unique
            suffix = "" if idx == 0 else "_{}".format(idx + 1)
            ds.Name = DS_NAME_PREFIX + (name_lbl or "Unnamed") + suffix

            ds.SetShape(List[DB.GeometryObject]([msolid]))

            set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
            set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

            if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                ogs = make_override(name_lbl, solid_fill_id)
                view.SetElementOverrides(ds.Id, ogs)

            created += 1

    t.Commit()

except Exception as fatal:
    t.RollBack()
    output.print_md("❌ Transaction rolled back")
    output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
    raise


# ----------------- REPORT -----------------

output.print_md("## Space Solids (Merged by Name + Touch Safeguard)")
output.print_md("- Deleted old solids: **{}**".format(deleted))
output.print_md("- Created: **{}**".format(created))
output.print_md("- Skipped: **{}**".format(skipped))
output.print_md("- Failed: **{}**".format(failed))
output.print_md("- Overrides applied in active view: **{}**".format("Yes" if APPLY_OVERRIDES_IN_ACTIVE_VIEW else "No"))
output.print_md("- Touch tolerance (ft): **{}**".format(TOUCH_TOL))
