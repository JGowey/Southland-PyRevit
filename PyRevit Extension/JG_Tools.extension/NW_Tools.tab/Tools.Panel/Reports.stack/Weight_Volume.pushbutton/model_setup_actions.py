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
            return DB.BuiltInParameterGroup.PG_CONSTRUCTION
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


def run(uiapp, req, ui_status):
    log = _make_logger(ui_status)
    ok = True

    log("Starting Model Setup...")
    log("")

    if getattr(req, "install_params", False):
        sp_txt = _strip_quotes(getattr(req, "shared_param_file", None) or NW_SHARED_PARAMS_TXT)
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

    log("Done." if ok else "Finished with warnings/errors.")
    return ok



def _install_family_params(uiapp, sp_txt, log):
    """In an RFA (Family Editor), 'Install/Bind' means add shared params as FAMILY parameters."""
    log("  Active document is a FAMILY (RFA). Installing as Family Parameters...")

    if not sp_txt or not os.path.exists(sp_txt):
        log("  Shared parameter TXT not found:")
        log("  {}".format(sp_txt))
        log("")
        return False

    doc = uiapp.ActiveUIDocument.Document

    # Group identifier for FamilyManager.AddParameter
    grp = _construction_group(uiapp)

    def _do():
        spf = _open_shared_param_file(uiapp)
        if not spf:
            log("  ERROR: Could not open shared parameter file inside Revit.")
            log("")
            return False

        names = ["CP_BOM_Category", "CP_Weight", "CP_Volume"]
        defs = []
        for nm in names:
            d = _find_def(spf, nm)
            if not d:
                log("  ERROR: Not found in shared param file: {}".format(nm))
                log("")
                return False
            defs.append(d)

        try:
            fm = doc.FamilyManager
        except Exception as ex:
            log("  ERROR: Could not access FamilyManager: {}".format(ex))
            log("")
            return False

        def _fam_has(nm):
            try:
                for p in fm.Parameters:
                    try:
                        if p.Definition and p.Definition.Name == nm:
                            return True
                    except:
                        pass
            except:
                pass
            return False

        added = []
        skipped = []
        failed = []

        t = DB.Transaction(doc, "NW Model Setup - Add Family Params")
        t.Start()
        try:
            for d in defs:
                nm = d.Name
                if _fam_has(nm):
                    skipped.append(nm)
                    continue
                try:
                    fm.AddParameter(d, grp, True)  # instance
                    added.append(nm)
                except Exception as ex_add:
                    failed.append("{} ({})".format(nm, ex_add))
            t.Commit()
        except:
            t.RollBack()
            raise

        log("  Family parameter install complete.")
        log("  Added:   {}".format(", ".join(added) if added else "none"))
        log("  Skipped: {}".format(", ".join(skipped) if skipped else "none"))
        if failed:
            log("  Failed:  {}".format("; ".join(failed)))
            log("")
            return False
        log("")
        return True

    return bool(with_temp_shared_param_file(uiapp, sp_txt, _do))

def _install_required_params(uiapp, sp_txt, log):
    log("[Parameters]")
    # If running in the Family Editor (RFA), install as FAMILY parameters.
    try:
        if uiapp.ActiveUIDocument.Document.IsFamilyDocument:
            return _install_family_params(uiapp, sp_txt, log)
    except:
        pass
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
