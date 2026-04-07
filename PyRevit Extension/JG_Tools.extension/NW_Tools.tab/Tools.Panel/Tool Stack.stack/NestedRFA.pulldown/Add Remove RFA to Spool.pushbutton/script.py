# ToggleNestedRFAs_OutsideEditMode.py
# Date: 2025-07-08

import clr
import sys
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from RevitServices.Persistence import DocumentManager
from System.Collections.Generic import List
from pyrevit import forms

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Select a parent FamilyInstance
try:
    ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick parent RFA to toggle in its assembly.")
    parent_elem = doc.GetElement(ref.ElementId)
except:
    sys.exit()

if not isinstance(parent_elem, FamilyInstance):
    forms.alert("Selected element is not a FamilyInstance.", exitscript=True)

parent_id = parent_elem.Id

# Get all assemblies in the model
all_assemblies = FilteredElementCollector(doc).OfClass(AssemblyInstance).ToElements()

# Find which assembly this element belongs to
target_assembly = None
for assembly in all_assemblies:
    if parent_id in assembly.GetMemberIds():
        target_assembly = assembly
        break

# If not found, let user pick the assembly manually
if not target_assembly:
    assy_ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick Assembly to add/remove this element.")
    target_assembly = doc.GetElement(assy_ref.ElementId)

if not isinstance(target_assembly, AssemblyInstance):
    forms.alert("Selected element is not an Assembly.", exitscript=True)

# Collect nested RFAs from the parent
nested_ids = []
for sub_id in parent_elem.GetSubComponentIds():
    sub_elem = doc.GetElement(sub_id)
    if isinstance(sub_elem, FamilyInstance):
        nested_ids.append(sub_elem.Id)

# Combine parent and nested RFAs
all_ids = [parent_id] + nested_ids
member_ids = set([eid.IntegerValue for eid in target_assembly.GetMemberIds()])

# Determine if parent is already in assembly
t = Transaction(doc, "Toggle Parent and Nested RFAs")
t.Start()

if parent_id.IntegerValue in member_ids:
    # Remove parent only; nested will auto-remove
    target_assembly.RemoveMemberIds(List[ElementId]([parent_id]))
    msg = "Removed parent and nested RFAs from assembly."
else:
    to_add = [eid for eid in all_ids if eid.IntegerValue not in member_ids]
    if to_add:
        target_assembly.AddMemberIds(List[ElementId](to_add))
        msg = "Added parent and nested RFAs to assembly."
    else:
        msg = "All elements already in assembly."

t.Commit()
forms.alert(msg)
