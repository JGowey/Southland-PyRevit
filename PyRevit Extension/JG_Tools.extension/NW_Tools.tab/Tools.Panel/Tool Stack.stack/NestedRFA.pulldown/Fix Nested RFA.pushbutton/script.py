# AddNestedRFAsPersistent.py
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

forms.alert("Tool active: Click parent RFAs one-by-one.\nPress ESC or Right-Click to exit.")

while True:
    try:
        ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick a parent RFA already in an assembly.")
        parent_elem = doc.GetElement(ref.ElementId)

        if not isinstance(parent_elem, FamilyInstance):
            forms.alert("Selected element is not a FamilyInstance.")
            continue

        parent_id = parent_elem.Id

        # Find the assembly it belongs to
        all_assemblies = FilteredElementCollector(doc).OfClass(AssemblyInstance).ToElements()
        target_assembly = None
        for assembly in all_assemblies:
            if parent_id in assembly.GetMemberIds():
                target_assembly = assembly
                break

        if not target_assembly:
            forms.alert("That element is not currently in an assembly.")
            continue

        # Get nested FamilyInstances
        nested_ids = []
        for sub_id in parent_elem.GetSubComponentIds():
            sub_elem = doc.GetElement(sub_id)
            if isinstance(sub_elem, FamilyInstance):
                nested_ids.append(sub_elem.Id)

        existing_ids = set([eid.IntegerValue for eid in target_assembly.GetMemberIds()])
        to_add = [eid for eid in nested_ids if eid.IntegerValue not in existing_ids]

        if not to_add:
            forms.alert("All nested RFAs already in assembly.")
            continue

        t = Transaction(doc, "Add Nested RFAs to Assembly")
        t.Start()
        target_assembly.AddMemberIds(List[ElementId](to_add))
        t.Commit()

        forms.alert("Added {} nested RFA(s) to assembly.".format(len(to_add)))

    except:
        break  # user cancelled or ESC pressed
