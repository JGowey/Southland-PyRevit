# -*- coding: utf-8 -*-
"""
NW Weight Tool - Model Setup Actions (ExternalEvent)

- Revit 2022 compatible: uses System.Collections.Generic.List
- Uses hard-coded shell path by year (no override support)
- Dumps shell schedule names when names don't match
- No popups; all reporting goes to Status

Schedule name mapping updated based on your shell preview:
  - Weight: "Weight Schedule"
  - Spool:  "NW Spool Sheet Weight Schedule"
  - QC:     "NW Weight Schedule QC"
"""

from __future__ import division

import os
import time

import Autodesk.Revit.DB as DB
from System.Collections.Generic import List


NW_SHARED_PARAMS_TXT = r"\\si.net\si\TCM\BIM\Shared Parameters\NW_Shared Parameters.txt"
SCHEDULE_ASSET_DIR = r"\\si.net\si\TCM\BIM\PyRevit\PyRevit_Content_Library\Schedule_Assets"

def shell_rvt_for_year(year):
    yy = str(int(year))[-2:]
    return os.path.join(SCHEDULE_ASSET_DIR, "PyRevit_Schedule_Asset_R_{}.rvt".format(yy))


SCHEDULE_NAMES = {
    "weight": "NW_Weight Schedule",
    "qc":     "NW_Weight Schedule QC report",
    "cat":    "NW_Weight Schedule with Category",
}
CATEGORY_NAMES = [
    "Air Terminals","Cable Tray Fittings","Cable Trays","Columns","Communication Devices",
    "Conduit Fittings","Conduits","Data Devices","Doors","Duct Accessories","Duct Fittings","Ducts",
    "Electrical Equipment","Electrical Fixtures","Fire Alarm Devices","Fire Protection","Flex Ducts","Flex Pipes",
    "Food Service Equipment","Generic Models","Lighting Devices","Lighting Fixtures","Mechanical Control Devices",
    "Mechanical Equipment","Medical Equipment","MEP Fabrication Ductwork","MEP Fabrication Hangers","MEP Fabrication Pipework",
    "Parts","Pipe Accessories","Pipe Fittings","Pipes","Plumbing Equipment","Plumbing Fixtures","Railings","Ramps",
    "Security Devices","Specialty Equipment","Sprinklers","Structural Columns","Structural Connections",
    "Structural Foundations","Structural Framing","Structural Rebar","Structural Stiffeners","Structural Trusses",
    "Telephone Devices","Walls","Windows","Wires",
]


def _strip_quotes(s):
    s = (s or "").strip()
    if len(s) >= 2 and ((s[0] == '"' and s[-1] == '"') or (s[0] == "'" and s[-1] == "'")):
        s = s[1:-1].strip()
    return s

def _now_tag():
    return time.strftime("%Y-%m-%d %H%M%S")

def _rev_year(uiapp):
    try:
        return int(uiapp.Application.VersionNumber)
    except:
        return 2024

def _construction_group(uiapp):
    yr = _rev_year(uiapp)
    if yr >= 2024:
        try:
            return DB.GroupTypeId.Construction
        except:
            return None
    else:
        try:
            return _resolve_param_group("CONSTRUCTION")
        except:
            return None

def _make_logger(ui_status):
    lines = []
    def log(msg=""):
        lines.append(msg)
        try:
            ui_status("\n".join(lines))
        except:
            pass
    return log

def with_temp_shared_param_file(uiapp, sp_txt_path, func):
    app = uiapp.Application
    original = None
    try:
        original = app.SharedParametersFilename
    except:
        original = None
    try:
        try:
            app.SharedParametersFilename = sp_txt_path
        except:
            pass
        return func()
    finally:
        try:
            app.SharedParametersFilename = original
        except:
            pass

def _open_shared_param_file(uiapp):
    try:
        return uiapp.Application.OpenSharedParameterFile()
    except:
        return None

def _find_def(spf, name):
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

def _category_set(doc, names):
    cs = DB.CategorySet()
    missing = []
    cats = doc.Settings.Categories
    for nm in names:
        try:
            cat = cats.get_Item(nm)
            if cat:
                cs.Insert(cat)
        except:
            missing.append(nm)
    return cs, missing

def _collect_schedules_by_name(doc):
    d = {}
    for vs in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule):
        try:
            nm = vs.Name
            if nm:
                d.setdefault(nm, []).append(vs)
        except:
            pass
    return d

def _list_shell_schedule_names(shell_doc, limit=200):
    names = []
    for vs in DB.FilteredElementCollector(shell_doc).OfClass(DB.ViewSchedule):
        try:
            nm = vs.Name
            if nm:
                names.append(nm)
        except:
            pass
    names = sorted(set(names))
    return names[:limit], len(names)

def _find_schedule_in_shell(shell_doc, name):
    for vs in DB.FilteredElementCollector(shell_doc).OfClass(DB.ViewSchedule):
        try:
            if vs.Name == name:
                return vs
        except:
            pass
    return None

def _open_shell_doc(app, path):
    if not path or not os.path.exists(path):
        return None, "Shell RVT not found: {}".format(path)
    try:
        return app.OpenDocumentFile(path), None
    except Exception as ex:
        return None, "Failed opening shell RVT: {}".format(ex)

def _close_doc_nosave(doc):
    try:
        doc.Close(False)
    except:
        pass




# -----------------------------------------------------------------------------
# Families (RFA editor mode) - add shared parameters as Family Parameters
# -----------------------------------------------------------------------------
def _rev_year():
    try:
        return int(uiapp.Application.VersionNumber)
    except:
        try:
            return int(__revit__.Application.VersionNumber)
        except:
            return 2024


def _group_type_id_by_ui_label(label_text):
    """Revit 2024+: scan DB.GroupTypeId static properties and return the one whose UI label matches."""
    try:
        if not hasattr(DB, "GroupTypeId"):
            return None
        import clr
        from System.Reflection import BindingFlags
        gt_type = clr.GetClrType(DB.GroupTypeId)
        props = gt_type.GetProperties(BindingFlags.Public | BindingFlags.Static)
        want = (label_text or "").strip().lower()
        for p in props:
            try:
                v = p.GetValue(None, None)  # ForgeTypeId
            except:
                continue
            try:
                lab = DB.LabelUtils.GetLabelFor(v)
            except:
                lab = None
            if lab and lab.strip().lower() == want:
                return v
    except:
        pass
    return None


def _gtid(name):
    """Get DB.GroupTypeId.<name> via .NET reflection (works even if IronPython can't see the property)."""
    try:
        if not hasattr(DB, "GroupTypeId"):
            return None
        import clr
        gt_type = clr.GetClrType(DB.GroupTypeId)
        prop = gt_type.GetProperty(name)
        if prop:
            return prop.GetValue(None, None)  # ForgeTypeId
    except:
        pass
    return None


def _resolve_param_group(token, uiapp=None):
    """
    Return a group identifier usable with FamilyManager.AddParameter.

    Revit 2024+:
      - "CONSTRUCTION" -> DB.GroupTypeId.Construction
      - "SET"          -> DB.GroupTypeId.CouplerArray  (your environment labels this as "Set")
    Revit 2022-2023:
      - BuiltInParameterGroup enums when available (PG_CONSTRUCTION / PG_SET)
    """
    t = (token or "CONSTRUCTION").strip().upper().replace(" ", "_").replace("-", "_")
    if t == "PG_CONSTRUCTION":
        t = "CONSTRUCTION"
    if t == "PG_SET":
        t = "SET"

    # 2024+ ForgeTypeId groups
    if hasattr(DB, "GroupTypeId"):
        if t == "CONSTRUCTION":
            return _gtid("Construction")
        if t == "SET":
            # IMPORTANT: In your Revit build, the UI label "Set" maps to GroupTypeId.CouplerArray
            return _gtid("CouplerArray")
        if t == "DATA":
            return _gtid("Data")
        # fallback
        return _gtid("Construction")

    # 2022-2023 BuiltInParameterGroup
    try:
        if hasattr(DB, "BuiltInParameterGroup"):
            pg = DB.BuiltInParameterGroup
            if t == "CONSTRUCTION":
                return pg.PG_CONSTRUCTION
            if t == "SET":
                return getattr(pg, "PG_SET", pg.PG_CONSTRUCTION)
            if t == "DATA":
                return pg.PG_DATA
            return pg.PG_CONSTRUCTION
    except:
        pass

    return None


class _FamilyLoadOptions(DB.IFamilyLoadOptions):
    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try:
            overwriteParameterValues.Value = True
        except:
            pass
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        try:
            overwriteParameterValues.Value = True
        except:
            pass
        return True


def _fam_has_param(fm, name):
    try:
        for p in fm.Parameters:
            try:
                if p.Definition and p.Definition.Name == name:
                    return True
            except:
                pass
    except:
        pass
    return False


def _load_rfa_family_parameters(uiapp, sp_txt, param_specs, log):
    """
    If running inside an RFA (family editor) document, adds parameters to THAT open family document.
    If running inside an RVT project document, iterates editable loaded families via EditFamily.
    param_specs: list of (param_name, BuiltInParameterGroup)
    Adds the shared parameter as an INSTANCE family parameter.
    """
    log("[Families]")
    if not param_specs:
        log("  None selected (skipped).")
        log("")
        return True

    if not sp_txt or not os.path.exists(sp_txt):
        log("  Shared parameter TXT not found:")
        log("  {}".format(sp_txt))
        log("")
        return False

    doc = uiapp.ActiveUIDocument.Document
    opts = _FamilyLoadOptions()

    def _do():
        spf = _open_shared_param_file(uiapp)
        if not spf:
            log("  ERROR: Could not open shared parameter file inside Revit.")
            log("")
            return False

        # Resolve definitions
        defs = {}
        for nm, _grp in param_specs:
            d = _find_def(spf, nm)
            if not d:
                log("  ERROR: Not found in shared param file: {}".format(nm))
                log("")
                return False
            defs[nm] = d

        # ------------------------------------------------------------
        # CASE A: We are already in the RFA editor (Family Document)
        # ------------------------------------------------------------
        try:
            is_family_doc = doc.IsFamilyDocument
        except:
            is_family_doc = False

        if is_family_doc:
            log("  Active document is a FAMILY (RFA editor).")
            log("  Adding family parameters to the OPEN family...")
            for nm, grp in param_specs:
                try:
                    lab = DB.LabelUtils.GetLabelFor(grp)
                except:
                    lab = str(grp)
                log("    - {} -> {}".format(nm, lab))
            log("")

            added = 0
            skipped = 0
            failed = 0

            try:
                fm = doc.FamilyManager
            except Exception as ex:
                log("  ERROR: Could not access FamilyManager ({})".format(ex))
                log("")
                return False

            t = DB.Transaction(doc, "NW Model Setup - Add Family Params")
            t.Start()
            try:
                for nm, grp in param_specs:
                    if _fam_has_param(fm, nm):
                        skipped += 1
                        continue
                    try:
                        fm.AddParameter(defs[nm], grp, True)  # instance
                        added += 1
                    except Exception as ex_add:
                        failed += 1
                        log("  FAIL add: {} ({})".format(nm, ex_add))
                t.Commit()
            except:
                t.RollBack()
                raise

            log("")
            log("  Family param loading complete (OPEN family).")
            log("  Added:   {}".format(added))
            log("  Skipped: {}".format(skipped))
            log("  Failed:  {}".format(failed))
            log("")
            return failed == 0

        # ------------------------------------------------------------
        # CASE B: We are in an RVT project document
        # ------------------------------------------------------------
        fams = list(DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements())
        editable = [f for f in fams if getattr(f, "IsEditable", False)]

        log("  Active document is a PROJECT (RVT).")
        log("  Editable families found: {}".format(len(editable)))
        log("  Adding family parameters (in RFA editor mode via EditFamily):")
        for nm, grp in param_specs:
            try:
                lab = DB.LabelUtils.GetLabelFor(grp)
            except:
                lab = str(grp)
            log("    - {} -> {}".format(nm, lab))
        log("")

        added = 0
        skipped = 0
        failed = 0

        for fam in editable:
            try:
                fam_name = fam.Name
            except:
                fam_name = "<family>"

            try:
                famdoc = doc.EditFamily(fam)
            except Exception as ex:
                failed += 1
                log("  FAIL open: {} ({})".format(fam_name, ex))
                continue

            try:
                fm = famdoc.FamilyManager
                t = DB.Transaction(famdoc, "NW Model Setup - Add Family Params")
                t.Start()
                try:
                    for nm, grp in param_specs:
                        if _fam_has_param(fm, nm):
                            skipped += 1
                            continue
                        try:
                            fm.AddParameter(defs[nm], grp, True)  # instance
                            added += 1
                        except Exception as ex_add:
                            failed += 1
                            log("  FAIL add: {} -> {} ({})".format(fam_name, nm, ex_add))
                    t.Commit()
                except:
                    t.RollBack()
                    raise

                try:
                    famdoc.LoadFamily(doc, opts)
                except Exception as ex_load:
                    failed += 1
                    log("  FAIL load: {} ({})".format(fam_name, ex_load))

            except Exception as ex:
                failed += 1
                log("  FAIL: {} ({})".format(fam_name, ex))
            finally:
                try:
                    famdoc.Close(False)
                except:
                    pass

        log("")
        log("  Family param loading complete (PROJECT pass).")
        log("  Added:   {}".format(added))
        log("  Skipped: {}".format(skipped))
        log("  Failed:  {}".format(failed))
        log("")
        return failed == 0

    return bool(with_temp_shared_param_file(uiapp, sp_txt, _do))
def run(uiapp, req, ui_status):
    log = _make_logger(ui_status)
    ok = True

    sp_txt = _strip_quotes(getattr(req, "shared_param_file", None) or NW_SHARED_PARAMS_TXT)

    log("Starting Model Setup...")
    log("")

    if getattr(req, "install_params", False):
        ok = _install_required_params(uiapp, sp_txt, log) and ok
    else:
        log("[Parameters]")
        log("  Not selected (skipped).")
        log("")

    schedule_keys = []
    if getattr(req, "import_weight", False): schedule_keys.append("weight")
    if getattr(req, "import_spool", False):  schedule_keys.append("spool")
    if getattr(req, "import_qc", False):     schedule_keys.append("qc")

    if schedule_keys:
        # req.shell_year may exist but be None ("Auto" selection)
        year = getattr(req, "shell_year", None) or _rev_year(uiapp)
        collision_mode = getattr(req, "collision_mode", "rename")

        shell_path = shell_rvt_for_year(year)
        names = [SCHEDULE_NAMES[k] for k in schedule_keys if k in SCHEDULE_NAMES]

        ok = _import_schedules(uiapp, shell_path, names, collision_mode, log) and ok
    else:
        log("[Schedules]")
        log("  None selected (skipped).")
        log("")

    # Families (RFA editor) - optional
    specs = []
    if bool(getattr(req, "rfa_add_cp_weight", False)):
        specs.append(("CP_Weight", _resolve_param_group("CONSTRUCTION")))
    if bool(getattr(req, "rfa_add_cp_volume", False)):
        specs.append(("CP_Volume", _resolve_param_group("CONSTRUCTION")))
    if bool(getattr(req, "rfa_add_cp_wpf", False)):
        specs.append(("CP_Weight_Per_Foot", _resolve_param_group("SET")))

    ok = _load_rfa_family_parameters(uiapp, sp_txt, specs, log) and ok

    log("Done." if ok else "Finished with warnings/errors.")
    return ok


def _install_required_params(uiapp, sp_txt, log):
    log("[Parameters]")
    if not sp_txt or not os.path.exists(sp_txt):
        log("  Shared parameter TXT not found:")
        log("  {}".format(sp_txt))
        log("")
        return False

    doc = uiapp.ActiveUIDocument.Document
    grp = _construction_group(uiapp)

    log("  Installing/binding required parameters...")
    log("  Using TXT: {}".format(sp_txt))
    log("  Will restore user's SharedParametersFilename afterward.")
    log("")

    def _do():
        spf = _open_shared_param_file(uiapp)
        if not spf:
            log("  ERROR: Could not open shared parameter file inside Revit.")
            log("")
            return False

        cs, missing = _category_set(doc, CATEGORY_NAMES)
        if missing:
            log("  Note: Some categories not found in this Revit/template:")
            log("    " + ", ".join(sorted(set(missing))))

        names = ["CP_BOM_Category", "CP_Weight", "CP_Volume"]
        defs = []
        for nm in names:
            d = _find_def(spf, nm)
            if not d:
                log("  ERROR: Not found in shared param file: {}".format(nm))
                log("")
                return False
            defs.append(d)

        binding = DB.InstanceBinding(cs)
        bindings = doc.ParameterBindings

        added = []
        updated = []

        t = DB.Transaction(doc, "NW Model Setup - Bind CP Parameters")
        t.Start()
        try:
            for d in defs:
                if not bindings.Contains(d):
                    bindings.Insert(d, binding, grp)
                    added.append(d.Name)
                else:
                    bindings.ReInsert(d, binding, grp)
                    updated.append(d.Name)
            t.Commit()
        except:
            t.RollBack()
            raise

        log("  Parameter binding complete.")
        log("  Added:   {}".format(", ".join(added) if added else "none"))
        log("  Updated: {}".format(", ".join(updated) if updated else "none"))
        log("")
        return True

    return bool(with_temp_shared_param_file(uiapp, sp_txt, _do))


def _import_schedules(uiapp, shell_path, schedule_names, collision_mode, log):
    log("[Schedules]")
    log("  Requested: {}".format(", ".join(schedule_names)))
    log("  Shell: {}".format(shell_path))
    log("  Collision: {}".format(collision_mode))
    log("")

    doc = uiapp.ActiveUIDocument.Document
    app = uiapp.Application

    shell_doc, err = _open_shell_doc(app, shell_path)
    if err:
        log("  " + err)
        log("")
        return False

    try:
        existing = _collect_schedules_by_name(doc)

        shell_views = []
        not_found = []
        for nm in schedule_names:
            vs = _find_schedule_in_shell(shell_doc, nm)
            if not vs:
                not_found.append(nm)
            else:
                shell_views.append(vs)

        if not_found:
            log("  Not found in shell:")
            for nm in not_found:
                log("    - {}".format(nm))
            log("")

            names_preview, total = _list_shell_schedule_names(shell_doc, limit=200)
            log("  Shell contains {} schedules. Preview (up to 200):".format(total))
            for nm in names_preview:
                log("    * {}".format(nm))
            log("")

        if not shell_views:
            log("  No schedules found to import (check names above).")
            log("")
            return False

        log("  Copying schedules...")

        t = DB.Transaction(doc, "NW Model Setup - Import Schedules")
        t.Start()
        try:
            for shell_vs in shell_views:
                target_name = shell_vs.Name
                collided = target_name in existing and existing[target_name]

                if collided and collision_mode == "rename":
                    for ex_vs in existing[target_name]:
                        try:
                            ex_vs.Name = "{} (backup {})".format(target_name, _now_tag())
                        except:
                            pass
                    log("    Renamed existing: {}".format(target_name))

                elif collided and collision_mode == "overwrite":
                    for ex_vs in existing[target_name]:
                        try:
                            doc.Delete(ex_vs.Id)
                        except:
                            pass
                    log("    Deleted existing: {}".format(target_name))

                elif collided and collision_mode == "skip":
                    log("    Skip existing: {}".format(target_name))
                    continue

                ids = List[DB.ElementId]()
                ids.Add(shell_vs.Id)

                try:
                    DB.ElementTransformUtils.CopyElements(
                        shell_doc, ids, doc, None, DB.CopyPasteOptions()
                    )
                    log("    Imported: {}".format(target_name))
                except Exception as ex:
                    log("    FAILED importing {}: {}".format(target_name, ex))

            t.Commit()
        except:
            t.RollBack()
            raise

        log("")
        log("  Schedule import complete.")
        log("")
        return True

    finally:
        log("  Closing shell (no save).")
        _close_doc_nosave(shell_doc)
        log("")
