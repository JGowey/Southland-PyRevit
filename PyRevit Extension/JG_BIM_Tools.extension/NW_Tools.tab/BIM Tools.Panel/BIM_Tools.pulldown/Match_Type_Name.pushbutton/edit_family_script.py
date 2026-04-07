# -*- coding: utf-8 -*-
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import DB
from Autodesk.Revit.UI import TaskDialog

# Revit context
doc = __revit__.ActiveUIDocument.Document
app = doc.Application

if not doc.IsFamilyDocument:
    raise Exception("Run this in the Family Editor (RFA), not a project (RVT).")

REV_YEAR = int(app.VersionNumber)

# -----------------------
# SETTINGS
# -----------------------
param_name = "CP_Type Name"  # shared parameter name in your SPF
is_instance = False          # False = Type param, True = Instance param
shared_param_file = r"\\si.net\si\TCM\BIM\Shared Parameters\NW_Shared Parameters.txt"

# Choose group properly by Revit year
if REV_YEAR >= 2024:
    group_id = DB.GroupTypeId.Construction
else:
    group_id = DB.BuiltInParameterGroup.PG_CONSTRUCTION


def open_spf(application, sp_path):
    """Open shared parameter file by temporarily setting SharedParametersFilename."""
    original = application.SharedParametersFilename
    try:
        application.SharedParametersFilename = sp_path
        return application.OpenSharedParameterFile()
    finally:
        application.SharedParametersFilename = original


def find_definition(spf, name):
    for g in spf.Groups:
        for d in g.Definitions:
            if d.Name == name:
                return d
    return None


fam_mgr = doc.FamilyManager

# Open SPF + locate definition
spf = open_spf(app, shared_param_file)
if spf is None:
    raise Exception("Could not open Shared Parameter file:\n{}".format(shared_param_file))

param_def = find_definition(spf, param_name)
if param_def is None:
    raise Exception("Shared Parameter '{}' not found in:\n{}".format(param_name, shared_param_file))

# Remove existing parameter if present (so AddParameter doesn't fail)
existing = None
for p in fam_mgr.Parameters:
    try:
        if p.Definition and p.Definition.Name == param_name:
            existing = p
            break
    except:
        pass

# Main transaction (ALWAYS closes properly)
t = DB.Transaction(doc, "Match Type Name - Add/Update Shared Parameter")
try:
    t.Start()

    if existing is not None:
        fam_mgr.RemoveParameter(existing)

    # Add the shared parameter to the family
    new_param = fam_mgr.AddParameter(param_def, group_id, is_instance)

    # Write the type name into the parameter for all types
    written = []
    for fam_type in fam_mgr.Types:
        fam_mgr.CurrentType = fam_type
        try:
            fam_mgr.Set(new_param, fam_type.Name)
            written.append(fam_type.Name)
        except:
            pass

    t.Commit()

except Exception as ex:
    if t.HasStarted():
        t.RollBack()
    raise

# Report
msg = []
msg.append("Revit Year: {}".format(REV_YEAR))
msg.append("Added '{}' as {} parameter.".format(param_name, "INSTANCE" if is_instance else "TYPE"))
msg.append("Group: {}".format("Construction"))
msg.append("")
msg.append("Wrote values to types ({}):".format(len(written)))
msg.append("\n".join(written) if written else "(none)")

TaskDialog.Show("Match Type Name", "\n".join(msg))
