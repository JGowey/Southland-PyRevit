# -*- coding: ascii -*-
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import ExternalDefinitionCreationOptions, ElementId
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

import os

# Revit context
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

if not doc.IsFamilyDocument:
    raise Exception("This script must be run in the Family Editor.")

# Settings
param_name = "CP_Type Name"
param_group_id = ElementId(-2000011)  # PG_CONSTRUCTION
is_instance = False
shared_param_file = r"\\si.net\si\TCM\BIM\Shared Parameters\NW_Shared Parameters.txt"

# Temporarily set shared parameter file
original_file = app.SharedParametersFilename
app.SharedParametersFilename = shared_param_file
shared_file = app.OpenSharedParameterFile()
app.SharedParametersFilename = original_file

if shared_file is None:
    raise Exception("Could not open shared parameter file.")

# Locate parameter definition
param_def = None
for group in shared_file.Groups:
    for definition in group.Definitions:
        if definition.Name == param_name:
            param_def = definition
            break
    if param_def:
        break

if param_def is None:
    raise Exception("Parameter '{}' not found.".format(param_name))

fam_mgr = doc.FamilyManager

# Remove existing parameter if present
for param in fam_mgr.Parameters:
    if param.Definition.Name == param_name:
        TransactionManager.Instance.EnsureInTransaction(doc)
        fam_mgr.RemoveParameter(param)
        TransactionManager.Instance.TransactionTaskDone()
        break

# Add the parameter
TransactionManager.Instance.EnsureInTransaction(doc)
try:
    new_param = fam_mgr.AddParameter(param_def, param_group_id, is_instance)
except Exception as ex:
    TransactionManager.Instance.TransactionTaskDone()
    raise Exception("Failed to add parameter: {}".format(str(ex)))
TransactionManager.Instance.TransactionTaskDone()

# Write values to types
TransactionManager.Instance.EnsureInTransaction(doc)
written = []
for fam_type in fam_mgr.Types:
    fam_mgr.CurrentType = fam_type
    try:
        fam_mgr.Set(new_param, fam_type.Name)
        written.append(fam_type.Name)
    except:
        pass
TransactionManager.Instance.TransactionTaskDone()

# Show result
from Autodesk.Revit.UI import TaskDialog
result_msg = "Bound '{}' to Construction group.\nWrote values to types:\n\n{}".format(param_name, "\n".join(written))
TaskDialog.Show("pyRevit", result_msg)
