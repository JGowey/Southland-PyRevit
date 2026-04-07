# -*- coding: utf-8 -*-
import clr
import sys

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

# Access current document
uidoc = __revit__.ActiveUIDocument
if uidoc is None:
    TaskDialog.Show("PyRevitLoader - Error", "No active Revit document.")
    sys.exit()

doc = uidoc.Document
if not doc.IsFamilyDocument:
    TaskDialog.Show("PyRevitLoader - Error", "You must run this inside an open RFA (Family Editor) document.")
    sys.exit()

# Get Revit version
app_version = int(doc.Application.VersionNumber)

# Load shared parameter file
app = doc.Application
shared_param_file = app.OpenSharedParameterFile()
if shared_param_file is None:
    TaskDialog.Show("PyRevitLoader - Error", "Shared parameter file not found.")
    sys.exit()

# Find "CP_Type Name" in the shared parameter file
target_def = None
for group in shared_param_file.Groups:
    for definition in group.Definitions:
        if definition.Name == "CP_Type Name":
            target_def = definition
            break
    if target_def:
        break

if not target_def:
    TaskDialog.Show("PyRevitLoader - Error", "Could not find shared parameter: CP_Type Name.")
    sys.exit()

# Check if parameter already exists
fam_mgr = doc.FamilyManager
param_exists = any(p.Definition.Name == "CP_Type Name" for p in fam_mgr.Parameters)

if param_exists:
    TaskDialog.Show("PyRevitLoader - Info", "Parameter already exists.")
    sys.exit()

# Start transaction and add parameter
t = Transaction(doc, "Add Shared Parameter: CP_Type Name")
t.Start()
try:
    if app_version >= 2023:
        group_id = GroupTypeId.Construction  # ForgeTypeId for Revit 2023+
    else:
        group_id = BuiltInParameterGroup.PG_CONSTRUCTION  # Legacy enum for Revit 2022

    fam_mgr.AddParameter(target_def, group_id, False)  # False = Type Parameter
    t.Commit()
    TaskDialog.Show("PyRevitLoader - Success", "Parameter added successfully.")
except Exception as e:
    t.RollBack()
    TaskDialog.Show("PyRevitLoader - Error", "Failed to add parameter:\n{}".format(str(e)))
