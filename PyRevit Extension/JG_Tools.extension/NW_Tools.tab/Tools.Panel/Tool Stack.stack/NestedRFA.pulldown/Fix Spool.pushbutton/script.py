# PullNestedRFAsIntoAssembly.py
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

# Pick an AssemblyInstance
try:
    assy_ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick an Assembly to pull nested RFAs into.")
    assembly = doc.GetElement(assy_ref.ElementId)
except:
    sys.exit()

if not isinstance(assembly, AssemblyInstance):
    forms.alert("Selected element is not an Assembly.", exitscript=True)

existing_ids = set([eid.IntegerValue for eid in assembly.GetMemberIds()])
to_add = []

# Loop through all current assembly members
for eid in assembly.GetMemberIds():
    elem = doc.GetElement(eid)
    if isinstance(elem, FamilyInstance):
        for sub_id in elem.GetSubComponentIds():
            sub_elem = doc.GetElement(sub_id)
            if isinstance(sub_elem, FamilyInstance):
                if sub_elem.Id.IntegerValue not in existing_ids:
                    to_add.append(sub_elem.Id)

# Add new nested RFAs
if not to_add:
    forms.alert("No new nested RFAs to add. Assembly already includes all nested elements.")
    sys.exit()

t = Transaction(doc, "Pull Nested RFAs into Assembly")
t.Start()
assembly.AddMemberIds(List[ElementId](to_add))
t.Commit()

forms.alert("Added {} nested RFA(s) to the assembly.".format(len(to_add)))

