# -*- coding: utf-8 -*-
import clr
import sys
import traceback

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    Transaction, TransactionGroup, BuiltInCategory, ExternalDefinitionCreationOptions,
    SpecTypeId, InstanceBinding
)
from Autodesk.Revit.UI import TaskDialog


# -------------------------
# SETTINGS
# -------------------------
# Set this to your shared parameter file. Must be accessible.
SHARED_PARAM_FILE = r"\\si.net\si\TCM\BIM\Shared Parameters\NW_Shared Parameters.txt"

# Shared parameter group to use/create inside the SPF
SPF_GROUP_NAME = "_PB_SELFTEST"

# Name prefix for test definitions (unique per group)
DEF_PREFIX_RFA = "_PB_TEST_FAM_"
DEF_PREFIX_RVT = "_PB_TEST_PRJ_"

# RVT categories to test binding against (safe/common)
RVT_TEST_CATEGORIES = [BuiltInCategory.OST_Walls, BuiltInCategory.OST_GenericModel]


# -------------------------
# REVIT CONTEXT
# -------------------------
uidoc = __revit__.ActiveUIDocument
if uidoc is None:
    TaskDialog.Show("Param Binder Self-Test", "No active Revit document.")
    sys.exit()

doc = uidoc.Document
app = doc.Application
REV_YEAR = int(app.VersionNumber)


# -------------------------
# HELPERS
# -------------------------
def _log(lines, msg):
    lines.append(msg)

def _open_shared_parameter_file_with_temp_path(app, sp_path):
    """Open SPF by temporarily setting SharedParametersFilename, then restore."""
    original = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = sp_path
        spf = app.OpenSharedParameterFile()
        return spf
    finally:
        app.SharedParametersFilename = original

def _get_or_create_spf_group(spf, name):
    grp = None
    for g in spf.Groups:
        if g.Name == name:
            grp = g
            break
    if grp is None:
        grp = spf.Groups.Create(name)
    return grp

def _get_or_create_definition(spf_group, def_name):
    # Reuse if already in file
    for d in spf_group.Definitions:
        if d.Name == def_name:
            return d
    opt = ExternalDefinitionCreationOptions(def_name, SpecTypeId.String.Text)
    opt.Visible = True
    return spf_group.Definitions.Create(opt)

def _iter_groups_current_year():
    """
    Returns list of (token_name, group_id) pairs for this year:
    - 2022/2023: BuiltInParameterGroup values
    - 2024+: GroupTypeId ForgeTypeId values
    """
    groups = []
    if REV_YEAR >= 2024:
        from Autodesk.Revit.DB import GroupTypeId, ForgeTypeId
        for name in dir(GroupTypeId):
            if name.startswith("_"):
                continue
            try:
                gid = getattr(GroupTypeId, name)
                # Keep only ForgeTypeId instances
                if isinstance(gid, ForgeTypeId):
                    groups.append((name, gid))
            except:
                pass
    else:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        for name in dir(BuiltInParameterGroup):
            if name.startswith("_"):
                continue
            # Most useful ones start with PG_
            if not name.startswith("PG_"):
                continue
            try:
                gid = getattr(BuiltInParameterGroup, name)
                groups.append((name, gid))
            except:
                pass
    return groups

def _safe_categoryset_for_project(doc, app, wanted_bics):
    """Create a CategorySet with categories that exist and allow bound params."""
    catset = app.Create.NewCategorySet()
    cats = doc.Settings.Categories
    for bic in wanted_bics:
        try:
            cat = cats.get_Item(bic)
            if cat and cat.AllowsBoundParameters:
                catset.Insert(cat)
        except:
            pass
    return catset

def _format_report(lines):
    # Keep it readable in TaskDialog: chunk if huge
    return "\n".join(lines)


# -------------------------
# SELF TESTS
# -------------------------
def self_test_family(doc, app, spf, groups, lines):
    fam_mgr = doc.FamilyManager

    _log(lines, "Doc Type: RFA (Family)")
    _log(lines, "Revit Year: {}".format(REV_YEAR))
    _log(lines, "Total Groups Found: {}".format(len(groups)))
    _log(lines, "")

    spf_group = _get_or_create_spf_group(spf, SPF_GROUP_NAME)

    # For 2024+ we only test user-assignable groups (same as “Group parameter under” list)
    # For 2022/2023 we just try them (PG_ list)
    tested = 0
    ok = 0
    skipped = 0
    failed = 0

    t = Transaction(doc, "PB SelfTest (RFA)")
    t.Start()
    try:
        for token, gid in groups:

            # Skip non user-assignable groups in 2024+ (this is expected)
            if REV_YEAR >= 2024:
                try:
                    if not fam_mgr.IsUserAssignableParameterGroup(gid):
                        _log(lines, "⚠️ SKIP (not user assignable): {}".format(token))
                        skipped += 1
                        continue
                except Exception as ex:
                    _log(lines, "❌ ERROR (assignable check): {} -> {}".format(token, ex))
                    failed += 1
                    continue

            # UNIQUE parameter definition per group (fixes “already in use”)
            def_name = "{}{}".format(DEF_PREFIX_RFA, token)

            try:
                extdef = _get_or_create_definition(spf_group, def_name)
            except Exception as ex:
                _log(lines, "❌ ERROR (create def): {} -> {}".format(token, ex))
                failed += 1
                continue

            # Add shared parameter to family using this group
            tested += 1
            try:
                fam_mgr.AddParameter(extdef, gid, False)  # False = Type param
                _log(lines, "✅ RFA OK: {}".format(token))
                ok += 1
            except Exception as ex:
                # This is the meaningful failure (real incompatibility)
                _log(lines, "❌ RFA ERROR: {} -> {}".format(token, ex))
                failed += 1

        _log(lines, "")
        _log(lines, "RFA Summary: tested={} ok={} skipped={} failed={}".format(tested, ok, skipped, failed))

    finally:
        # Rollback so nothing stays
        t.RollBack()


def self_test_project(doc, app, spf, groups, lines):
    _log(lines, "Doc Type: RVT (Project)")
    _log(lines, "Revit Year: {}".format(REV_YEAR))
    _log(lines, "Total Groups Found: {}".format(len(groups)))
    _log(lines, "")

    spf_group = _get_or_create_spf_group(spf, SPF_GROUP_NAME)

    catset = _safe_categoryset_for_project(doc, app, RVT_TEST_CATEGORIES)
    if catset.IsEmpty:
        _log(lines, "❌ RVT: No valid test categories allow bound parameters.")
        return

    binding = app.Create.NewInstanceBinding(catset)
    pb = doc.ParameterBindings

    tested = 0
    ok = 0
    failed = 0

    t = Transaction(doc, "PB SelfTest (RVT)")
    t.Start()
    try:
        for token, gid in groups:
            # Unique definition per group so we don't collide
            def_name = "{}{}".format(DEF_PREFIX_RVT, token)

            try:
                extdef = _get_or_create_definition(spf_group, def_name)
            except Exception as ex:
                _log(lines, "❌ RVT ERROR (create def): {} -> {}".format(token, ex))
                failed += 1
                continue

            tested += 1
            try:
                inserted = pb.Insert(extdef, binding, gid)
                if not inserted:
                    inserted = pb.ReInsert(extdef, binding, gid)

                if inserted:
                    _log(lines, "✅ RVT OK: {}".format(token))
                    ok += 1
                else:
                    _log(lines, "⚠️ RVT FAILED (Insert/ReInsert returned False): {}".format(token))
                    failed += 1

            except Exception as ex:
                _log(lines, "❌ RVT ERROR: {} -> {}".format(token, ex))
                failed += 1

        _log(lines, "")
        _log(lines, "RVT Summary: tested={} ok={} failed={}".format(tested, ok, failed))

    finally:
        # Rollback so nothing stays
        t.RollBack()


# -------------------------
# MAIN
# -------------------------
def run():
    lines = []
    _log(lines, "=== PARAM BINDER FULL SELF TEST ===")

    # Load SPF
    spf = _open_shared_parameter_file_with_temp_path(app, SHARED_PARAM_FILE)
    if spf is None:
        TaskDialog.Show("Param Binder Self-Test",
                        "Shared Parameter file could not be opened.\n\nPath:\n{}".format(SHARED_PARAM_FILE))
        return

    groups = _iter_groups_current_year()

    tg = TransactionGroup(doc, "PB FULL SELF TEST (Rollback)")
    tg.Start()
    try:
        if doc.IsFamilyDocument:
            self_test_family(doc, app, spf, groups, lines)
        else:
            self_test_project(doc, app, spf, groups, lines)

        _log(lines, "")
        _log(lines, "✅ TEST COMPLETE (ALL ROLLED BACK)")
    except Exception:
        _log(lines, "")
        _log(lines, "❌ UNHANDLED ERROR:")
        _log(lines, traceback.format_exc())
    finally:
        tg.RollBack()

    TaskDialog.Show("Param Binder Full Self Test", _format_report(lines))


run()
