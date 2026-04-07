# -*- coding: utf-8 -*-
"""
Single pyRevit button launcher (ALL-IN-ONE):
1) Create 3D space solids per Space (DirectShape + optional view overrides)
2) Create 3D space solids and merge into zones by Space Name (touch/union safeguard)
3) Create merged zones for Navisworks export (materialized tessellation + optional view overrides)
4) Import merged zones into ACTIVE model by reading Spaces from ANOTHER OPEN model (no overrides)
5) Import per-Space 3D solids into ACTIVE model by reading Spaces from ANOTHER OPEN model (no merge; optional view overrides)

Added:
- Optional pre-clear of existing generated DirectShapes in the TARGET (active) model
  so you don't overlap yourself when re-importing.
"""

from pyrevit import revit, DB, script, forms
from System.Collections.Generic import List
from collections import defaultdict
import hashlib

doc = revit.doc
view = doc.ActiveView
output = script.get_output()

# ---------------- GLOBAL SETTINGS ----------------
GEN_TAG = "#SPACE_SOLID"
TARGET_CATEGORY = DB.BuiltInCategory.OST_GenericModel
ONLY_PLACED = True

# Overrides (Revit-only visuals; export may ignore these)
APPLY_OVERRIDES_IN_ACTIVE_VIEW = True
TRANSPARENCY = 70  # 0..100

# Merge safeguards (Modes 2, 3, 4)
TOUCH_TOL = 0.005     # feet (~1/16")
MIN_VOL   = 1e-6

# Materials for export-safe coloring (Mode 3)
MAT_PREFIX = "OPAL_SPACE__"

# Name prefixes (so cleanup can remove prior runs)
PREFIX_MODE1 = "SpaceSolid_"
PREFIX_MODE2 = "ZONE_"
PREFIX_MODE3 = "ZONE_"
PREFIX_MODE4 = "ZONE_"
PREFIX_MODE5 = "SpaceSolid_"
# -----------------------------------------------


# ---------------- SPACE HELPERS ----------------
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
    name = get_param_as_string(space, DB.BuiltInParameter.ROOM_NAME)
    if not name:
        name = get_param_by_name(space, "Name")
    if not name:
        try:
            name = space.Name or ""
        except:
            name = ""
    return (name or "").strip() or "Unnamed"


def set_text_param(elem, bip, value):
    try:
        p = elem.get_Parameter(bip)
        if p and not p.IsReadOnly:
            p.Set(value or "")
            return True
    except:
        pass
    return False


# ---------------- CLEANUP ----------------
def delete_old_generated(document, prefixes):
    """
    Deletes DirectShapes created by these tools:
    - Mark == GEN_TAG
    - OR Name starts with any supplied prefix
    """
    ids = List[DB.ElementId]()
    col = DB.FilteredElementCollector(document).OfClass(DB.DirectShape)

    for ds in col:
        try:
            nm = ds.Name or ""
            mark = get_param_as_string(ds, DB.BuiltInParameter.ALL_MODEL_MARK)
            if (mark == GEN_TAG) or any(nm.startswith(p) for p in prefixes):
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
    """Find Solid fill FillPatternElement id."""
    for fpe in DB.FilteredElementCollector(document).OfClass(DB.FillPatternElement):
        try:
            fp = fpe.GetFillPattern()
            if fp and fp.IsSolidFill:
                return fpe.Id
        except:
            pass
    return DB.ElementId.InvalidElementId


# ---------------- COLORS + OVERRIDES ----------------
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

    try:
        ogs.SetSurfaceTransparency(int(TRANSPARENCY))
    except:
        pass

    try:
        ogs.SetProjectionLineColor(col)
    except:
        pass

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

        # Cut (optional)
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


# ---------------- MERGE HELPERS (Modes 2, 3, 4) ----------------
def bbox_of_solid(s):
    try:
        return s.GetBoundingBox()
    except:
        return None


def bbox_intersects_or_touches(bb1, bb2, tol=TOUCH_TOL):
    if bb1 is None or bb2 is None:
        return False

    return (
        (bb1.Min.X - tol) <= (bb2.Max.X + tol) and (bb2.Min.X - tol) <= (bb1.Max.X + tol) and
        (bb1.Min.Y - tol) <= (bb2.Max.Y + tol) and (bb2.Min.Y - tol) <= (bb1.Max.Y + tol) and
        (bb1.Min.Z - tol) <= (bb2.Max.Z + tol) and (bb2.Min.Z - tol) <= (bb1.Max.Z + tol)
    )


def try_union(a, b):
    try:
        u = DB.BooleanOperationsUtils.ExecuteBooleanOperation(a, b, DB.BooleanOperationsType.Union)
        if u and u.Volume > MIN_VOL:
            return True, u
    except:
        pass
    return False, None


def cluster_and_union_solids(solids):
    """
    Return list of merged solids, each is the union of a connected cluster.
    Connectivity:
    - bbox touch/overlap within TOUCH_TOL
    - AND successful boolean union
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


# ---------------- MATERIALS + TESSELLATION (Mode 3) ----------------
_ILLEGAL = r'\/:*?"<>|'

def safe_material_name(space_name):
    s = (space_name or "Unknown").strip()
    for ch in _ILLEGAL:
        s = s.replace(ch, "_")
    s = s[:120]
    return MAT_PREFIX + s


def build_material_index(document):
    idx = {}
    for m in DB.FilteredElementCollector(document).OfClass(DB.Material):
        try:
            idx[m.Name] = m
        except:
            pass
    return idx


def get_or_create_space_material(document, mat_index, space_name):
    """Creates/reuses one Revit Material per unique space name."""
    mname = safe_material_name(space_name)

    if mname in mat_index:
        return mat_index[mname].Id

    mid = DB.Material.Create(document, mname)
    mat = document.GetElement(mid)

    col = color_from_text(space_name)
    try:
        mat.Color = col
    except:
        pass

    try:
        mat.Transparency = int(TRANSPARENCY)
    except:
        pass

    mat_index[mname] = mat
    return mat.Id


def solid_to_tessellated_geometry(solid, material_id):
    """
    Convert Solid into tessellated geometry where each triangle carries material_id.
    Reliable for NWC/Navisworks color on DirectShape.
    """
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(True)

    for face in solid.Faces:
        try:
            mesh = face.Triangulate()
            for i in range(mesh.NumTriangles):
                tri = mesh.get_Triangle(i)
                v0 = tri.get_Vertex(0)
                v1 = tri.get_Vertex(1)
                v2 = tri.get_Vertex(2)

                pts = List[DB.XYZ]([v0, v1, v2])
                tface = DB.TessellatedFace(pts, material_id)
                if tface and tface.IsValidObject:
                    builder.AddFace(tface)
        except:
            pass

    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()

    result = builder.GetBuildResult()
    return list(result.GetGeometricalObjects())


# ---------------- MODE IMPLEMENTATIONS ----------------
def validate_overrides_if_needed():
    if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
        if not can_override(view):
            output.print_md("❌ Active view can't accept element overrides (or is a template). Open a 3D view and rerun.")
            script.exit()
        solid_fill_id = find_solid_fill_pattern_id(doc)
        if solid_fill_id == DB.ElementId.InvalidElementId:
            output.print_md("❌ Could not find 'Solid fill' pattern in this project.")
            output.print_md("   (If you only care about export, set APPLY_OVERRIDES_IN_ACTIVE_VIEW = False and rerun.)")
            script.exit()
        return solid_fill_id
    return DB.ElementId.InvalidElementId


def mode1_create_per_space():
    DS_NAME_PREFIX = PREFIX_MODE1
    solid_fill_id = validate_overrides_if_needed()

    spaces = get_spaces(doc)
    if not spaces:
        output.print_md("❌ No MEP Spaces found.")
        return

    calc = DB.SpatialElementGeometryCalculator(doc)
    cat_id = DB.ElementId(TARGET_CATEGORY)

    t = DB.Transaction(doc, "Space Solids (Per Space)")
    t.Start()

    deleted = created = skipped = failed = 0
    try:
        deleted = delete_old_generated(doc, [PREFIX_MODE1, PREFIX_MODE2, PREFIX_MODE3, PREFIX_MODE4, PREFIX_MODE5])

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

                ds = DB.DirectShape.CreateElement(doc, cat_id)
                ds.Name = DS_NAME_PREFIX + (name_lbl or str(sp.Id.IntegerValue))
                ds.SetShape([solid])

                set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
                set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

                if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                    ogs = make_override(name_lbl, solid_fill_id)
                    view.SetElementOverrides(ds.Id, ogs)

                created += 1

            except Exception as e:
                failed += 1
                output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, getattr(e, "Message", str(e))))

        t.Commit()
    except Exception as fatal:
        t.RollBack()
        output.print_md("❌ Transaction rolled back")
        output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
        raise

    output.print_md("## Mode 1 — Space Solids (Per Space)")
    output.print_md("- Deleted old solids: **{}**".format(deleted))
    output.print_md("- Created: **{}**".format(created))
    output.print_md("- Skipped: **{}**".format(skipped))
    output.print_md("- Failed: **{}**".format(failed))


def mode2_merge_by_name_touching():
    DS_NAME_PREFIX = PREFIX_MODE2
    solid_fill_id = validate_overrides_if_needed()

    spaces = get_spaces(doc)
    if not spaces:
        output.print_md("❌ No MEP Spaces found.")
        return

    calc = DB.SpatialElementGeometryCalculator(doc)
    cat_id = DB.ElementId(TARGET_CATEGORY)

    t = DB.Transaction(doc, "Space Solids (Merged Zones by Name)")
    t.Start()

    deleted = created = skipped = failed = 0
    try:
        deleted = delete_old_generated(doc, [PREFIX_MODE1, PREFIX_MODE2, PREFIX_MODE3, PREFIX_MODE4, PREFIX_MODE5])

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
                output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, getattr(e, "Message", str(e))))

        for name_lbl, solids in solids_by_name.items():
            merged_solids = cluster_and_union_solids(solids)

            for idx, msolid in enumerate(merged_solids):
                try:
                    if (msolid is None) or (msolid.Volume <= MIN_VOL):
                        skipped += 1
                        continue

                    ds = DB.DirectShape.CreateElement(doc, cat_id)
                    suffix = "" if idx == 0 else "_{}".format(idx + 1)
                    ds.Name = DS_NAME_PREFIX + (name_lbl or "Unnamed") + suffix

                    ds.SetShape(List[DB.GeometryObject]([msolid]))

                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

                    if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                        ogs = make_override(name_lbl, solid_fill_id)
                        view.SetElementOverrides(ds.Id, ogs)

                    created += 1

                except Exception as e:
                    failed += 1
                    output.print_md(u"**Failed Zone '{}'** — {}".format(name_lbl, getattr(e, "Message", str(e))))

        t.Commit()
    except Exception as fatal:
        t.RollBack()
        output.print_md("❌ Transaction rolled back")
        output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
        raise

    output.print_md("## Mode 2 — Space Solids (Merged by Name + Touch Safeguard)")
    output.print_md("- Deleted old solids: **{}**".format(deleted))
    output.print_md("- Created: **{}**".format(created))
    output.print_md("- Skipped: **{}**".format(skipped))
    output.print_md("- Failed: **{}**".format(failed))
    output.print_md("- Touch tolerance (ft): **{}**".format(TOUCH_TOL))


def mode3_merged_for_navisworks_export():
    DS_NAME_PREFIX = PREFIX_MODE3
    solid_fill_id = validate_overrides_if_needed()

    spaces = get_spaces(doc)
    if not spaces:
        output.print_md("❌ No MEP Spaces found.")
        return

    calc = DB.SpatialElementGeometryCalculator(doc)
    cat_id = DB.ElementId(TARGET_CATEGORY)
    mat_index = build_material_index(doc)

    t = DB.Transaction(doc, "Space Solids (Merged + Materialized for Navisworks Export)")
    t.Start()

    deleted = created = skipped = failed = 0
    try:
        deleted = delete_old_generated(doc, [PREFIX_MODE1, PREFIX_MODE2, PREFIX_MODE3, PREFIX_MODE4, PREFIX_MODE5])

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
                output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, getattr(e, "Message", str(e))))

        for name_lbl, solids in solids_by_name.items():
            merged_solids = cluster_and_union_solids(solids)
            mat_id = get_or_create_space_material(doc, mat_index, name_lbl)

            for idx, msolid in enumerate(merged_solids):
                try:
                    if (msolid is None) or (msolid.Volume <= MIN_VOL):
                        skipped += 1
                        continue

                    geom = solid_to_tessellated_geometry(msolid, mat_id)
                    if not geom:
                        skipped += 1
                        continue

                    ds = DB.DirectShape.CreateElement(doc, cat_id)
                    suffix = "" if idx == 0 else "_{}".format(idx + 1)
                    ds.Name = DS_NAME_PREFIX + (name_lbl or "Unnamed") + suffix
                    ds.SetShape(geom)

                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

                    if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                        ogs = make_override(name_lbl, solid_fill_id)
                        view.SetElementOverrides(ds.Id, ogs)

                    created += 1

                except Exception as e:
                    failed += 1
                    output.print_md(u"**Failed Export Zone '{}'** — {}".format(name_lbl, getattr(e, "Message", str(e))))

        t.Commit()
    except Exception as fatal:
        t.RollBack()
        output.print_md("❌ Transaction rolled back")
        output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
        raise

    output.print_md("## Mode 3 — Merged Zones (Materialized + Tessellated for Navisworks/NWC)")
    output.print_md("- Deleted old solids: **{}**".format(deleted))
    output.print_md("- Created: **{}**".format(created))
    output.print_md("- Skipped: **{}**".format(skipped))
    output.print_md("- Failed: **{}**".format(failed))
    output.print_md("- Touch tolerance (ft): **{}**".format(TOUCH_TOL))
    output.print_md("- Export coloring method: **Revit Materials per unique Space Name**")


# ---------------- MODE 4/5 (READ OTHER DOC, WRITE ACTIVE DOC) ----------------
def get_open_project_docs(app, exclude_doc=None):
    docs = []
    for d in app.Documents:
        try:
            if exclude_doc and d.Equals(exclude_doc):
                continue
            if d.IsFamilyDocument:
                continue
            try:
                if hasattr(d, "IsLinked") and d.IsLinked:
                    continue
            except:
                pass
            docs.append(d)
        except:
            pass
    return docs


class DocItem(object):
    def __init__(self, d):
        self.doc = d
        try:
            self.name = d.Title
        except:
            self.name = "Untitled"
    def __repr__(self):
        return self.name


def mode4_import_merged_from_other_doc_into_active():
    target_doc = revit.doc  # ACTIVE doc: only doc we modify
    app = target_doc.Application

    sources = get_open_project_docs(app, exclude_doc=target_doc)
    if not sources:
        output.print_md("❌ No other open project documents found. Open the source model and rerun.")
        return

    src_item = forms.SelectFromList.show(
        [DocItem(d) for d in sources],
        title="Select SOURCE model (Spaces read from here)",
        multiselect=False,
        button_name="Use This Source"
    )
    if not src_item:
        return
    source_doc = src_item.doc

    # Ask user whether to clear old zones in TARGET first
    do_clear = forms.alert(
        "Clear existing generated DirectShapes (zones) in the ACTIVE model before importing?",
        yes=True, no=True
    )

    spaces = get_spaces(source_doc)
    if not spaces:
        output.print_md("❌ No MEP Spaces found in SOURCE model: **{}**".format(source_doc.Title))
        return

    # Build merged solids from SOURCE
    calc = DB.SpatialElementGeometryCalculator(source_doc)
    solids_by_name = defaultdict(list)

    skipped = failed = 0
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
            output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, getattr(e, "Message", str(e))))

    cat_id = DB.ElementId(TARGET_CATEGORY)

    t = DB.Transaction(target_doc, "Import Merged Zones from Other Model Spaces")
    t.Start()

    deleted = created = 0
    try:
        if do_clear:
            deleted = delete_old_generated(target_doc, [PREFIX_MODE1, PREFIX_MODE2, PREFIX_MODE3, PREFIX_MODE4, PREFIX_MODE5])

        for name_lbl, solids in solids_by_name.items():
            merged_solids = cluster_and_union_solids(solids)

            for idx, msolid in enumerate(merged_solids):
                try:
                    if (msolid is None) or (msolid.Volume <= MIN_VOL):
                        skipped += 1
                        continue

                    ds = DB.DirectShape.CreateElement(target_doc, cat_id)
                    suffix = "" if idx == 0 else "_{}".format(idx + 1)
                    ds.Name = PREFIX_MODE4 + (name_lbl or "Unnamed") + suffix
                    ds.SetShape(List[DB.GeometryObject]([msolid]))

                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
                    set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

                    created += 1

                except Exception as e:
                    failed += 1
                    output.print_md(u"**Failed Zone '{}'** — {}".format(name_lbl, getattr(e, "Message", str(e))))

        t.Commit()
    except:
        t.RollBack()
        raise

    output.print_md("## Mode 4 — Imported Merged Zones into ACTIVE Model (No Overlap)")
    output.print_md("- Source: **{}**".format(source_doc.Title))
    output.print_md("- Target (active): **{}**".format(target_doc.Title))
    output.print_md("- Deleted old generated shapes in target: **{}**".format(deleted))
    output.print_md("- Created zones in target: **{}**".format(created))
    output.print_md("- Skipped: **{}**".format(skipped))
    output.print_md("- Failed: **{}**".format(failed))
    output.print_md("- Touch tolerance (ft): **{}**".format(TOUCH_TOL))


def mode5_import_per_space_from_other_doc_into_active():
    """
    Same as Mode 1 (per-space solids), but reads MEP Spaces from another OPEN project document
    and writes DirectShapes into the ACTIVE model. No merging.
    """
    target_doc = revit.doc  # ACTIVE doc: only doc we modify
    app = target_doc.Application

    sources = get_open_project_docs(app, exclude_doc=target_doc)
    if not sources:
        output.print_md("❌ No other open project documents found. Open the source model and rerun.")
        return

    src_item = forms.SelectFromList.show(
        [DocItem(d) for d in sources],
        title="Select SOURCE model (Spaces read from here)",
        multiselect=False,
        button_name="Use This Source"
    )
    if not src_item:
        return
    source_doc = src_item.doc

    # Ask user whether to clear old shapes in TARGET first
    do_clear = forms.alert(
        "Clear existing generated DirectShapes (space solids) in the ACTIVE model before importing?",
        yes=True, no=True
    )

    solid_fill_id = validate_overrides_if_needed()

    spaces = get_spaces(source_doc)
    if not spaces:
        output.print_md("❌ No MEP Spaces found in SOURCE model: **{}**".format(source_doc.Title))
        return

    calc = DB.SpatialElementGeometryCalculator(source_doc)
    cat_id = DB.ElementId(TARGET_CATEGORY)

    t = DB.Transaction(target_doc, "Import Per-Space Solids from Other Model Spaces")
    t.Start()

    deleted = created = skipped = failed = 0
    try:
        if do_clear:
            deleted = delete_old_generated(target_doc, [PREFIX_MODE1, PREFIX_MODE2, PREFIX_MODE3, PREFIX_MODE4, PREFIX_MODE5])

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

                ds = DB.DirectShape.CreateElement(target_doc, cat_id)

                # Add source space Id to avoid name collisions when multiple Spaces share the same name
                ds.Name = PREFIX_MODE5 + (name_lbl or "Unnamed") + "__" + str(sp.Id.IntegerValue)

                ds.SetShape([solid])

                set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_MARK, GEN_TAG)
                set_text_param(ds, DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS, name_lbl)

                if APPLY_OVERRIDES_IN_ACTIVE_VIEW:
                    ogs = make_override(name_lbl, solid_fill_id)
                    view.SetElementOverrides(ds.Id, ogs)

                created += 1

            except Exception as e:
                failed += 1
                output.print_md(u"**Failed Space ID {}** — {}".format(sp.Id.IntegerValue, getattr(e, "Message", str(e))))

        t.Commit()
    except Exception as fatal:
        t.RollBack()
        output.print_md("❌ Transaction rolled back")
        output.print_md(fatal.ToString() if hasattr(fatal, "ToString") else str(fatal))
        raise

    output.print_md("## Mode 5 — Imported Per-Space Solids into ACTIVE Model (No Merge)")
    output.print_md("- Source: **{}**".format(source_doc.Title))
    output.print_md("- Target (active): **{}**".format(target_doc.Title))
    output.print_md("- Deleted old generated shapes in target: **{}**".format(deleted))
    output.print_md("- Created space solids in target: **{}**".format(created))
    output.print_md("- Skipped: **{}**".format(skipped))
    output.print_md("- Failed: **{}**".format(failed))


# ---------------- UI LAUNCHER ----------------
choice = forms.CommandSwitchWindow.show(
    [
        "1) Create 3D spaces for each area (per Space)",
        "2) Create 3D spaces + merge into common zones (by Space Name)",
        "3) Create merged 3D spaces for Navisworks export (materials/tessellation)",
        "4) Import merged zones into ACTIVE model from OTHER open model (Spaces)",
        "5) Import 3D spaces for each area into ACTIVE model from OTHER open model"
    ],
    message="Select which version to run:"
)

if not choice:
    script.exit()

if choice.startswith("1)"):
    mode1_create_per_space()
elif choice.startswith("2)"):
    mode2_merge_by_name_touching()
elif choice.startswith("3)"):
    mode3_merged_for_navisworks_export()
elif choice.startswith("4)"):
    mode4_import_merged_from_other_doc_into_active()
else:
    mode5_import_per_space_from_other_doc_into_active()
