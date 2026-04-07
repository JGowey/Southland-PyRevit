# -*- coding: utf-8 -*-
import os
import json
import datetime as _dt

import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import CategorySet, InstanceBinding, TypeBinding
import Autodesk.Revit.DB as DB


# ----------------------------- paths (unchanged) -----------------------------
_SHARED_DIR_OVERRIDE = None
_USER_DIR_OVERRIDE = None

def set_paths(shared_root=None, user_root=None, **kwargs):
    global _SHARED_DIR_OVERRIDE, _USER_DIR_OVERRIDE
    sr = shared_root or kwargs.get("shared_cfg_dir") or kwargs.get("shared_dir")
    ur = user_root    or kwargs.get("user_cfg_dir")    or kwargs.get("user_dir")
    _SHARED_DIR_OVERRIDE = sr
    _USER_DIR_OVERRIDE    = ur

def _script_dir():
    return os.path.dirname(__file__)

def shared_dir():
    if _SHARED_DIR_OVERRIDE and os.path.isdir(_SHARED_DIR_OVERRIDE):
        return _SHARED_DIR_OVERRIDE
    p = os.path.join(_script_dir(), "Configs")
    return p if os.path.isdir(p) else None

def user_dir():
    root = _USER_DIR_OVERRIDE
    if not root:
        root = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "ParamBinder", "Configs")
    if not os.path.isdir(root):
        try:
            os.makedirs(root)
        except:
            pass
    return root

def can_edit_shared():
    sd = shared_dir()
    try:
        return bool(sd and os.path.exists(sd) and os.access(sd, os.W_OK))
    except:
        return False

def list_all_configs():
    rows = []
    for scope, base in (("shared", shared_dir()), ("user", user_dir())):
        if not base:
            continue
        try:
            for fn in os.listdir(base):
                if not fn.lower().endswith(".json"):
                    continue
                fp = os.path.join(base, fn)
                disp = None
                try:
                    with open(fp, "r") as f:
                        disp = json.load(f).get("display_name")
                except:
                    pass
                try:
                    mtime = _dt.datetime.fromtimestamp(os.path.getmtime(fp))
                except:
                    mtime = None
                rows.append((scope, fn, fp, disp, mtime))
        except:
            pass
    rows.sort(key=lambda r: (0 if r[0] == "shared" else 1, r[3] or r[1]))
    return rows

def load_json_file(path):
    with open(path, "r") as f:
        return json.load(f)


# ----------------------------- shared parameter files -----------------------
def open_shared_parameter_file(path):
    if not path or not os.path.exists(path):
        return None
    app = __revit__.Application
    original = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = path
        return app.OpenSharedParameterFile()
    finally:
        app.SharedParametersFilename = original


# ----------------------------- group mapping tables -------------------------
# (UI label -> canonical token)
_ALIAS_TO_TOKEN = {
    "GREEN_BUILDING_PROPERTIES": "GREEN_BUILDING",
    "SEGMENTS_FITTINGS": "SEGMENTS_AND_FITTINGS",
    "TITLE": "TITLE_TEXT",
    "ELECTRICAL_-_CIRCUITING": "ELECTRICAL_CIRCUITING",
    "ELECTRICAL_-_LIGHTING": "ELECTRICAL_LIGHTING",
    "ELECTRICAL_-_LOADS": "ELECTRICAL_LOADS",
    "MECHANICAL_-_FLOW": "MECHANICAL_FLOW",
    "MECHANICAL_-_LOADS": "MECHANICAL_LOADS",
    "RELEASES___MEMBER_FORCES": "RELEASES_MEMBER_FORCES",
    "SEGMENTS_AND_FITTINGS": "SEGMENTS_AND_FITTINGS",
    "TITLE_TEXT": "TITLE_TEXT",
    "OVERALL_LEGEND": "OVERALL_LEGEND",
    "MODEL_PROPERTIES": "MODEL_PROPERTIES",
    "MATERIALS_AND_FINISHES": "MATERIALS_AND_FINISHES",
    "DIVISION_GEOMETRY": "DIVISION_GEOMETRY",
    "LIFE_SAFETY": "LIFE_SAFETY",
    "STRUCTURAL_ANALYSIS": "STRUCTURAL_ANALYSIS",
    "IFC_PARAMETERS": "IFC_PARAMETERS",
    "SLAB_SHAPE_EDIT": "SLAB_SHAPE_EDIT",
    "ANALYTICAL_ALIGNMENT": "ANALYTICAL_ALIGNMENT",
    "IDENTITY_DATA": "IDENTITY_DATA",
    "FORCES": "FORCES",
    "LAYERS": "LAYERS",
    "DIMENSIONS": "DIMENSIONS",
}

def _normalize_token(tok):
    s = (tok or "CONSTRUCTION").upper().strip()
    for ch in (" - ", "-", "/", "&"):
        s = s.replace(ch, "_")
    s = s.replace("  ", " ").replace(" ", "_")
    while "__" in s:
        s = s.replace("__", "_")
    return _ALIAS_TO_TOKEN.get(s, s)

def _gt(attr):
    try:
        return getattr(DB.GroupTypeId, attr)
    except:
        return None

# 2024+ / ForgeTypeId table (note: "Other" handled in resolver)
_G23 = {
    "ANALYSIS_RESULTS": _gt("AnalysisResults"),
    "ANALYTICAL_ALIGNMENT": _gt("AnalyticalAlignment"),
    "ANALYTICAL_MODEL": _gt("AnalyticalModel"),
    "CONSTRAINTS": _gt("Constraints"),
    "CONSTRUCTION": _gt("Construction"),
    "DATA": _gt("Data"),
    "DIMENSIONS": _gt("Geometry"),
    "DIVISION_GEOMETRY": _gt("DivisionGeometry"),
    "ELECTRICAL": _gt("Electrical"),
    "ELECTRICAL_CIRCUITING": _gt("ElectricalCircuiting"),
    "ELECTRICAL_LIGHTING": _gt("ElectricalLighting"),
    "ELECTRICAL_LOADS": _gt("ElectricalLoads"),
    "ELECTRICAL_ANALYSIS": _gt("ElectricalAnalysis"),
    "ELECTRICAL_ENGINEERING": _gt("ElectricalEngineering"),
    "ENERGY_ANALYSIS": _gt("EnergyAnalysis"),
    "FIRE_PROTECTION": _gt("FireProtection"),
    "FORCES": _gt("Forces"),
    "GENERAL": _gt("General"),
    "GRAPHICS": _gt("Graphics"),
    "GREEN_BUILDING": _gt("GreenBuilding"),
    "IDENTITY_DATA": _gt("IdentityData"),
    "IFC_PARAMETERS": _gt("Ifc"),
    "LAYERS": _gt("RebarSystemLayers"),
    "LIFE_SAFETY": _gt("LifeSafety"),
    "MATERIALS_AND_FINISHES": _gt("Materials"),
    "MECHANICAL": _gt("Mechanical"),
    "MECHANICAL_FLOW": _gt("MechanicalAirflow"),
    "MECHANICAL_LOADS": _gt("MechanicalLoads"),
    "MODEL_PROPERTIES": _gt("AdskModelProperties"),
    "MOMENTS": _gt("Moments"),
    "OVERALL_LEGEND": _gt("OverallLegend"),
    "PHASING": _gt("Phasing"),
    "PHOTOMETRICS": _gt("LightPhotometrics"),
    "PLUMBING": _gt("Plumbing"),
    "PRIMARY_END": _gt("PrimaryEnd"),
    "REBAR_SET": _gt("RebarArray"),
    "REBAR_ARRAY": _gt("RebarArray"),
    "RELEASES_MEMBER_FORCES": _gt("ReleasesMemberForces"),
    "SECONDARY_END": _gt("SecondaryEnd"),
    "SEGMENTS_AND_FITTINGS": _gt("SegmentsFittings"),
    "SET": _gt("CouplerArray"),
    "SLAB_SHAPE_EDIT": _gt("SlabShapeEdit"),
    "STRUCTURAL": _gt("Structural"),
    "STRUCTURAL_ANALYSIS": _gt("StructuralAnalysis"),
    "STRUCTURAL_SECTION_DIMENSIONS": _gt("StructuralSectionDimensions"),
    "SUB_DIVISION": _gt("ToposolidSubdivision"),
    "TEXT": _gt("Text"),
    "TITLE_TEXT": _gt("Title"),
    "VISIBILITY": _gt("Visibility"),
    "VISUALIZATION": _gt("Visualization"),
}


# 2022/2023 use BuiltInParameterGroup (enum)
_HAS_PG = hasattr(DB, "BuiltInParameterGroup")
_G22 = {}
if _HAS_PG:
    BuiltInParameterGroup = DB.BuiltInParameterGroup
    def _pg(attr, default_attr="PG_CONSTRUCTION"):
        for name in (attr, default_attr, "PG_CONSTRUCTION"):
            try:
                return getattr(BuiltInParameterGroup, name)
            except Exception:
                pass
        return BuiltInParameterGroup.PG_CONSTRUCTION
    _G22 = {
        "ANALYSIS_RESULTS": _pg("PG_ANALYSIS_RESULTS"),
        "ANALYTICAL_ALIGNMENT": _pg("PG_ANALYTICAL_ALIGNMENT"),
        "ANALYTICAL_MODEL": _pg("PG_ANALYTICAL_MODEL"),
        "CONSTRAINTS": _pg("PG_CONSTRAINTS"),
        "CONSTRUCTION": _pg("PG_CONSTRUCTION"),
        "DATA": _pg("PG_IDENTITY_DATA", "PG_DATA"),
        "IDENTITY_DATA": _pg("PG_IDENTITY_DATA", "PG_DATA"),
        "DIMENSIONS": _pg("PG_LENGTH"),
        "DIVISION_GEOMETRY": _pg("PG_DIVISION_GEOMETRY", "PG_GEOMETRY"),
        "ELECTRICAL": _pg("PG_ELECTRICAL"),
        "ELECTRICAL_CIRCUITING": _pg("PG_ELECTRICAL_CIRCUITING"),
        "LIGHTING": _pg("PG_ELECTRICAL_LIGHTING"),
        "ELECTRICAL_LIGHTING": _pg("PG_ELECTRICAL_LIGHTING"),
        "ELECTRICAL_LOADS": _pg("PG_ELECTRICAL_LOADS"),
        "ELECTRICAL_ANALYSIS": _pg("PG_ELECTRICAL_ANALYSIS", "PG_ELECTRICAL"),
        "ELECTRICAL_ENGINEERING": _pg("PG_ELECTRICAL_ENGINEERING", "PG_ELECTRICAL"),
        "ENERGY_ANALYSIS": _pg("PG_ENERGY_ANALYSIS"),
        "FIRE_PROTECTION": _pg("PG_FIRE_PROTECTION"),
        "GENERAL": _pg("PG_GENERAL", "PG_CONSTRUCTION"),
        "GRAPHICS": _pg("PG_GRAPHICS"),
        "GREEN_BUILDING": _pg("PG_GREEN_BUILDING", "PG_DATA"),
        "IFC_PARAMETERS": _pg("PG_IFC"),
        "LAYERS": _pg("PG_GRAPHICS"),
        "LIFE_SAFETY": _pg("PG_LIFE_SAFETY", "PG_DATA"),
        "MATERIALS_AND_FINISHES": _pg("PG_MATERIALS", "PG_MATERIALS_AND_FINISHES"),
        "MECHANICAL": _pg("PG_MECHANICAL"),
        "MECHANICAL_FLOW": _pg("PG_MECHANICAL_AIRFLOW", "PG_MECHANICAL"),
        "MECHANICAL_LOADS": _pg("PG_MECHANICAL_LOADS", "PG_MECHANICAL"),
        "MODEL_PROPERTIES": _pg("PG_MODEL_PROPERTIES", "PG_GENERAL"),
        "MOMENTS": _pg("PG_MOMENTS", "PG_STRUCTURAL"),
        "OTHER": BuiltInParameterGroup.INVALID,  # "Other" == INVALID in enum API
        "OVERALL_LEGEND": _pg("PG_OVERALL_LEGENDS", "PG_OVERALL_LEGEND"),
        "PHASING": _pg("PG_PHASING"),
        "PHOTOMETRICS": _pg("PG_PHOTOMETRICS", "PG_LIGHTING"),
        "PLUMBING": _pg("PG_PLUMBING"),
        "PRIMARY_END": _pg("PG_PRIMARY_END", "PG_STRUCTURAL"),
        "RELEASES_MEMBER_FORCES": _pg("PG_RELEASES_MEMBER_FORCES", "PG_STRUCTURAL"),
        "SECONDARY_END": _pg("PG_SECONDARY_END", "PG_STRUCTURAL"),
        "SEGMENTS_AND_FITTINGS": _pg("PG_SEGMENTS_FITTINGS"),
        "SET": _pg("PG_SET", "PG_CONSTRUCTION"),
        "SLAB_SHAPE_EDIT": _pg("PG_SLAB_SHAPE_EDIT"),
        "STRUCTURAL": _pg("PG_STRUCTURAL"),
        "STRUCTURAL_ANALYSIS": _pg("PG_STRUCTURAL_ANALYSIS"),
        "TEXT": _pg("PG_TEXT"),
        "TITLE_TEXT": _pg("PG_TITLE"),
        "VISIBILITY": _pg("PG_VISIBILITY"),
    }

# ----------------------------- resolver -------------------------
def _rev_year():
    try:
        return int(__revit__.Application.VersionNumber)
    except:
        return 2024


def _try_bipg_from_token(token):
    """Revit 2022/2023: resolve BuiltInParameterGroup from an enum token.
    Accepts:
      - "PG_ELECTRICAL_LOADS"
      - "DB.BuiltInParameterGroup.PG_ELECTRICAL_LOADS"
    """
    if not hasattr(DB, "BuiltInParameterGroup"):
        return None
    t = (token or "").strip()
    if not t:
        return None
    if "." in t:
        t = t.split(".")[-1]
    if "PG_" in t and not t.startswith("PG_"):
        t = "PG_" + t.split("PG_", 1)[1]
    try:
        return getattr(DB.BuiltInParameterGroup, t)
    except:
        return None


def _try_grouptypeid_from_token(token):
    """Revit 2024+: resolve GroupTypeId ForgeTypeId from a property token.
    Accepts:
      - "ElectricalLoads"
      - "DB.GroupTypeId.ElectricalLoads"
    """
    t = (token or "").strip()
    if not t:
        return None
    if "." in t:
        t = t.split(".")[-1]
    try:
        return getattr(DB.GroupTypeId, t)
    except:
        return None

def resolve_group_for_json(token, rev_year):
    """
    Returns:
      - Revit 2024+: DB.ForgeTypeId (use empty ForgeTypeId() for "Other")
      - Revit 2022–2023: DB.BuiltInParameterGroup (use INVALID for "Other")
    """
    t_raw = (token or "").strip()

    if rev_year >= 2024:
        # Prefer direct property token resolution (e.g. "ElectricalLoads")
        direct = _try_grouptypeid_from_token(t_raw)
        if direct is not None:
            return direct

        # Legacy: labels / loose tokens from old configs
        t = _normalize_token(t_raw)
        if t == "OTHER":
            try:
                return DB.ForgeTypeId()  # "Other" in 2024+
            except:
                pass
        g = _G23.get(t)
        return g if g is not None else _G23.get("CONSTRUCTION")

    else:
        # Prefer direct enum token resolution (e.g. "PG_ELECTRICAL_LOADS")
        direct = _try_bipg_from_token(t_raw)
        if direct is not None:
            return direct

        # Legacy: labels / loose tokens from old configs
        t = _normalize_token(t_raw)
        g = _G22.get(t) if _G22 else None
        return g if g is not None else (_G22.get("CONSTRUCTION") if _G22 else None)



# ----------------------------- installer (unchanged logic) -------------------------
def _find_definition_in_spf(spf, name):
    if not spf:
        return None
    for g in spf.Groups:
        for d in g.Definitions:
            try:
                if d.Name == name:
                    return d
            except:
                pass
    return None

def _category_set_from_names(doc, names, skipped):
    cs = CategorySet()
    ac = doc.Settings.Categories
    for nm in names or []:
        try:
            cat = ac.get_Item(nm)
            if cat:
                cs.Insert(cat)
        except:
            skipped.append("bad category: " + nm)
    return cs

def install_from_config(cfg, cfg_path):
    from Autodesk.Revit.UI import TaskDialog

    uidoc = __revit__.ActiveUIDocument
    if uidoc is None:
        TaskDialog.Show("Param Binder - Error", "No active document.")
        return
    doc = uidoc.Document
    app = __revit__.Application
    rev_year = _rev_year()

    top_spf = (cfg.get("shared_parameter_file") or "").strip()
    persist_top = bool(cfg.get("persist_shared_parameter_file", False))
    original_spf = app.SharedParametersFilename

    added, updated, skipped = [], [], []

    try:
        if top_spf:
            app.SharedParametersFilename = top_spf

        spf_cache = {}
        def _open(path):
            key = os.path.normcase(path or "")
            if key not in spf_cache:
                spf_cache[key] = open_shared_parameter_file(path) if path else open_shared_parameter_file(app.SharedParametersFilename)
            return spf_cache[key]

        if doc.IsFamilyDocument:
            fam = doc.FamilyManager
            existing = set([p.Definition.Name for p in fam.Parameters])
            tx = DB.Transaction(doc, "Add Shared Parameters (Family)")
            tx.Start()
            try:
                for p in cfg.get("parameters", []):
                    name = (p.get("name") or "").strip()
                    if not name:
                        skipped.append("missing name")
                        continue
                    if name in existing:
                        skipped.append(name + " (already present)")
                        continue
                    spf = _open(p.get("shared_parameter_file"))
                    d = _find_definition_in_spf(spf, name)
                    if not d:
                        skipped.append(name + " (not found in SP file)")
                        continue
                    grp = resolve_group_for_json(p.get("group_key"), rev_year)
                    is_inst = bool(p.get("rfa_is_instance", True))
                    try:
                        fam.AddParameter(d, grp, is_inst)
                        added.append(name)
                    except Exception as e:
                        skipped.append(name + " (add failed: " + str(e) + ")")
                tx.Commit()
            except:
                tx.RollBack()
                raise
        else:
            bindings = doc.ParameterBindings
            tx = DB.Transaction(doc, "Bind Shared Parameters (Project)")
            tx.Start()
            try:
                for p in cfg.get("parameters", []):
                    name = (p.get("name") or "").strip()
                    if not name:
                        skipped.append("missing name")
                        continue
                    spf = _open(p.get("shared_parameter_file"))
                    d = _find_definition_in_spf(spf, name)
                    if not d:
                        skipped.append(name + " (not found in SP file)")
                        continue
                    cs = _category_set_from_names(doc, p.get("rvt_categories", []), skipped)
                    if cs.Size == 0:
                        skipped.append(name + " (no valid categories)")
                        continue
                    is_inst = bool(p.get("rvt_is_instance", True))
                    vary = bool(p.get("rvt_allow_vary_between_groups", True))
                    grp = resolve_group_for_json(p.get("group_key"), rev_year)
                    b = InstanceBinding(cs) if is_inst else TypeBinding(cs)
                    try:
                        if hasattr(b, "AllowVaryBetweenGroups"):
                            b.AllowVaryBetweenGroups = vary
                    except:
                        pass
                    if not bindings.Contains(d):
                        bindings.Insert(d, b, grp)
                        added.append(name)
                    else:
                        bindings.ReInsert(d, b, grp)
                        updated.append(name)
                tx.Commit()
            except:
                tx.RollBack()
                raise

        if persist_top and top_spf:
            app.SharedParametersFilename = top_spf
        else:
            app.SharedParametersFilename = original_spf

        msg = []
        msg.append("Done.")
        msg.append("")
        msg.append("Added:    " + str(len(added)))
        msg.append("Updated: " + str(len(updated)))
        msg.append("Skipped: " + str(len(skipped)))
        if added:
            msg.append("\nAdded:\n" + "\n".join(sorted(added)))
        if updated:
            msg.append("\nUpdated:\n" + "\n".join(sorted(updated)))
        if skipped:
            msg.append("\nSkipped:\n" + "\n".join(skipped))
        TaskDialog.Show("Param Binder", "\n".join(msg))
    except Exception as e:
        app.SharedParametersFilename = original_spf
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Param Binder - Error", str(e))

def install_from_json(path, doc=None):
    try:
        cfg = load_json_file(path)
        install_from_config(cfg, path)
    except Exception as e:
        from Autodesk.Revit.UI import TaskDialog
        TaskDialog.Show("Param Binder - Install Error", str(e))
