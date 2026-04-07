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




def _rev_year():
    """Return the running Revit major version as int (e.g., 2022, 2024, 2026)."""
    try:
        return int(__revit__.Application.VersionNumber)
    except:
        try:
            return int(__revit__.Application.VersionName)
        except:
            return 2024

# ----------------------------- group mapping & helpers -------------------------
# Goal:
#   - Revit 2022/2023 uses BuiltInParameterGroup (enum)
#   - Revit 2024+ uses GroupTypeId (ForgeTypeId)
# We prefer storing stable API tokens in JSON:
#   - 2022/2023: "PG_ELECTRICAL_LOADS" (enum name)
#   - 2024+:    "ElectricalLoads"      (DB.GroupTypeId property name)
# UI labels (e.g. "Electrical - Loads") are accepted for backward compatibility.

# Legacy label -> canonical token (only used as fallback)
_ALIAS_TO_TOKEN = {
    "GREEN_BUILDING_PROPERTIES": "GREEN_BUILDING",
    "SEGMENTS_FITTINGS": "SEGMENTS_AND_FITTINGS",
    "TITLE": "TITLE_TEXT",
    "ELECTRICAL_-_CIRCUITING": "ELECTRICAL_CIRCUITING",
    "ELECTRICAL_-_LIGHTING": "ELECTRICAL_LIGHTING",
    "ELECTRICAL_-_LOADS": "ELECTRICAL_LOADS",
    "MECHANICAL_-_FLOW": "MECHANICAL_AIRFLOW",
    "MECHANICAL_-_LOADS": "MECHANICAL_LOADS",
    "RELEASES___MEMBER_FORCES": "RELEASES_MEMBER_FORCES",
    "MATERIALS_AND_FINISHES": "MATERIALS",
    "IFC_PARAMETERS": "IFC",
    "MODEL_PROPERTIES": "ADSK_MODEL_PROPERTIES",
    "PHOTOMETRICS": "LIGHT_PHOTOMETRICS",
    "SET": "COUPLER_ARRAY",
    "REBAR_SET": "REBAR_ARRAY",
    "LAYERS": "REBAR_SYSTEM_LAYERS",
    "DIMENSIONS": "GEOMETRY",
}

def _normalize_token(tok):
    s = (tok or "CONSTRUCTION").upper().strip()
    for ch in (" - ", "-", "/", "&"):
        s = s.replace(ch, "_")
    s = s.replace("  ", " ").replace(" ", "_")
    while "__" in s:
        s = s.replace("__", "_")
    return _ALIAS_TO_TOKEN.get(s, s)

def _try_bipg_from_token(token):
    """
    Revit 2022/2023:
      token examples:
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
    """
    Revit 2024+:
      token examples:
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

def _gt(attr):
    try:
        return getattr(DB.GroupTypeId, attr)
    except:
        return None

# 2024+/ForgeTypeId fallback table (must match the actual user-assignable groups)
# Source of truth: your dump script output. (Revit 2024/2025 family user-assignable list)
_G23 = {
    "ANALYSIS_RESULTS": _gt("AnalysisResults"),
    "ANALYTICAL_ALIGNMENT": _gt("AnalyticalAlignment"),
    "ANALYTICAL_MODEL": _gt("AnalyticalModel"),
    "CONSTRAINTS": _gt("Constraints"),
    "CONSTRUCTION": _gt("Construction"),
    "DATA": _gt("Data"),
    "GEOMETRY": _gt("Geometry"),  # UI label "Dimensions"
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
    "IFC": _gt("Ifc"),
    "REBAR_SYSTEM_LAYERS": _gt("RebarSystemLayers"),
    "LIFE_SAFETY": _gt("LifeSafety"),
    "MATERIALS": _gt("Materials"),
    "MECHANICAL": _gt("Mechanical"),
    "MECHANICAL_AIRFLOW": _gt("MechanicalAirflow"),
    "MECHANICAL_LOADS": _gt("MechanicalLoads"),
    "ADSK_MODEL_PROPERTIES": _gt("AdskModelProperties"),
    "MOMENTS": _gt("Moments"),
    "OVERALL_LEGEND": _gt("OverallLegend"),
    "PHASING": _gt("Phasing"),
    "LIGHT_PHOTOMETRICS": _gt("LightPhotometrics"),
    "PLUMBING": _gt("Plumbing"),
    "PRIMARY_END": _gt("PrimaryEnd"),
    "REBAR_ARRAY": _gt("RebarArray"),
    "RELEASES_MEMBER_FORCES": _gt("ReleasesMemberForces"),
    "SECONDARY_END": _gt("SecondaryEnd"),
    "SEGMENTS_AND_FITTINGS": _gt("SegmentsFittings"),
    "COUPLER_ARRAY": _gt("CouplerArray"),
    "SLAB_SHAPE_EDIT": _gt("SlabShapeEdit"),
    "STRUCTURAL": _gt("Structural"),
    "STRUCTURAL_ANALYSIS": _gt("StructuralAnalysis"),
    "STRUCTURAL_SECTION_DIMENSIONS": _gt("StructuralSectionDimensions"),
    "TOPOSOLID_SUBDIVISION": _gt("ToposolidSubdivision"),
    "TEXT": _gt("Text"),
    "TITLE_TEXT": _gt("Title"),
    "VISIBILITY": _gt("Visibility"),
    "VISUALIZATION": _gt("Visualization"),  # present in 2025 family list
}

# 2022/2023: we keep a tiny fallback map, but prefer direct enum token resolution
_HAS_PG = hasattr(DB, "BuiltInParameterGroup")
_G22 = {}
if _HAS_PG:
    BuiltInParameterGroup = DB.BuiltInParameterGroup
    def _pg(attr, default_attr="PG_CONSTRUCTION"):
        for name in (attr, default_attr, "PG_CONSTRUCTION"):
            try:
                return getattr(BuiltInParameterGroup, name)
            except:
                pass
        return BuiltInParameterGroup.PG_CONSTRUCTION
    _G22 = {
        "CONSTRUCTION": _pg("PG_CONSTRUCTION"),
        "DATA": _pg("PG_DATA"),
        "IDENTITY_DATA": _pg("PG_IDENTITY_DATA", "PG_DATA"),
        "GEOMETRY": _pg("PG_GEOMETRY", "PG_LENGTH"),  # UI label "Dimensions"
        "ELECTRICAL": _pg("PG_ELECTRICAL"),
        "MECHANICAL": _pg("PG_MECHANICAL"),
        "PLUMBING": _pg("PG_PLUMBING"),
        "STRUCTURAL": _pg("PG_STRUCTURAL"),
        "GRAPHICS": _pg("PG_GRAPHICS"),
        "GENERAL": _pg("PG_GENERAL"),
        "OTHER": getattr(BuiltInParameterGroup, "INVALID", BuiltInParameterGroup.PG_CONSTRUCTION),
    }

def resolve_group_for_json_ex(token, rev_year):
    """
    Returns: (group_obj, used_fallback_bool, resolved_desc_str)

    - group_obj type:
        * 2022/2023: BuiltInParameterGroup
        * 2024+: ForgeTypeId (GroupTypeId)
    """
    raw = (token or "").strip()

    # Preferred: direct token resolution
    if rev_year >= 2024:
        direct = _try_grouptypeid_from_token(raw)
        if direct is not None:
            return direct, False, "GroupTypeId." + (raw.split(".")[-1] if raw else "")
    else:
        direct = _try_bipg_from_token(raw)
        if direct is not None:
            nm = raw.split(".")[-1] if raw else ""
            return direct, False, "BuiltInParameterGroup." + nm

    # Backward compatibility: label/alias normalization
    t = _normalize_token(raw)
    if rev_year >= 2024:
        if t == "OTHER":
            try:
                return DB.ForgeTypeId(), False, "ForgeTypeId()"
            except:
                pass
        g = _G23.get(t)
        if g is not None:
            return g, True, t
        return _G23.get("CONSTRUCTION"), True, "CONSTRUCTION (default)"
    else:
        g = _G22.get(t) if _G22 else None
        if g is not None:
            return g, True, t
        return (_G22.get("CONSTRUCTION") if _G22 else None), True, "CONSTRUCTION (default)"

def resolve_group_for_json(token, rev_year):
    return resolve_group_for_json_ex(token, rev_year)[0]

def list_user_assignable_group_choices(doc):
    """
    Returns list of (label, token) suitable for UI.
    token is what should be stored in JSON.
    """
    rev_year = _rev_year()
    out = []

    if rev_year >= 2024:
        import clr
        from System.Reflection import BindingFlags

        fm = None
        try:
            if doc and doc.IsFamilyDocument:
                fm = doc.FamilyManager
        except:
            fm = None

        gt_type = clr.GetClrType(DB.GroupTypeId)
        props = gt_type.GetProperties(BindingFlags.Public | BindingFlags.Static)

        for p in props:
            try:
                gid = p.GetValue(None, None)
            except:
                continue

            if fm:
                try:
                    if not fm.IsUserAssignableParameterGroup(gid):
                        continue
                except:
                    pass

            try:
                lab = DB.LabelUtils.GetLabelForGroup(gid)
            except:
                lab = None
            if lab:
                out.append((lab, p.Name))

        # Allow "Other" in config even if UI might not show it in some contexts
        out.append(("Other", "OTHER"))
        out = sorted(set(out), key=lambda r: r[0].lower())
        return out

    # 2022/2023: enumerate BuiltInParameterGroup names
    if hasattr(DB, "BuiltInParameterGroup"):
        Bipg = DB.BuiltInParameterGroup
        for nm in sorted([n for n in dir(Bipg) if n.startswith("PG_")]):
            try:
                val = getattr(Bipg, nm)
            except:
                continue
            try:
                lab = DB.LabelUtils.GetLabelFor(val)
            except:
                lab = None
            out.append((lab or nm, nm))
        # "Other" typically maps to INVALID
        out.append(("Other", "OTHER"))
        out = sorted(set(out), key=lambda r: r[0].lower())
    return out


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
    group_warnings = []  # group_key that had to fall back/default

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
                    grp, used_fallback, grp_desc = resolve_group_for_json_ex(p.get('group_key'), rev_year)
                    if used_fallback:
                        group_warnings.append("%s -> %s" % (p.get('group_key'), grp_desc))
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
                    grp, used_fallback, grp_desc = resolve_group_for_json_ex(p.get('group_key'), rev_year)
                    if used_fallback:
                        group_warnings.append("%s -> %s" % (p.get('group_key'), grp_desc))
                    grp = resolve_group_for_json(p.get("group_key"), rev_year)
                    b = InstanceBinding(cs) if is_inst else TypeBinding(cs)
                    vary = bool(p.get("rvt_allow_vary_between_groups", True))
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

        # Friendly group notes for UI (avoid confusing internal ForgeTypeId names)
        friendly_group_notes = []
        for w in sorted(set(group_warnings)):
            try:
                lhs, rhs = [x.strip() for x in w.split('->', 1)]
            except:
                friendly_group_notes.append(w)
                continue
            # Revit 2024+ maps the UI group label 'Set' to internal GroupTypeId.CouplerArray.
            if lhs.lower() == 'set' and rhs.upper() in ('COUPLER_ARRAY', 'COUPLERARRAY'):
                friendly_group_notes.append("Set (expected in Revit 2024+: internal group is CouplerArray)")
            else:
                friendly_group_notes.append("{} -> {}".format(lhs, rhs))

        msg = []
        msg.append("Status: SUCCESS")
        msg.append("Done.")
        msg.append("")
        msg.append("Added:    " + str(len(added)))
        msg.append("Updated: " + str(len(updated)))
        msg.append("Skipped: " + str(len(skipped)))
        if friendly_group_notes:
            msg.append("\nGroup notes:\n" + "\n".join(friendly_group_notes))
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