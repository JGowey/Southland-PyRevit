# -*- coding: utf-8 -*-
import clr
import csv
import os
import codecs
import System

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import SaveFileDialog, MessageBox, MessageBoxButtons

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from RevitServices.Persistence import DocumentManager

from pyrevit import script

# Setup context
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
output = script.get_output()
output.close_others()

# Prompt user
result = MessageBox.Show("Select a nested FamilyInstance (RFA) to export its type parameters.", "Export Type Parameters", MessageBoxButtons.OKCancel)
if result != System.Windows.Forms.DialogResult.OK:
    script.exit()

# Select element
try:
    ref = uidoc.Selection.PickObject(ObjectType.Element, "Select a nested family instance")
    inst = doc.GetElement(ref.ElementId)
    symbol = inst.Symbol
    family = symbol.Family
except:
    MessageBox.Show("No valid nested family instance selected.", "Error")
    script.exit()

# Collect all types in that family
collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
family_types = [t for t in collector if t.Family.Id == family.Id]

# Collect unique parameter definitions
all_param_defs = {}
for ftype in family_types:
    for param in ftype.Parameters:
        if param.Definition:
            pname = param.Definition.Name
            if pname not in all_param_defs:
                all_param_defs[pname] = param.Definition

# Sort parameters and build CSV header
sorted_param_names = sorted(all_param_defs.keys())
header = ["Type Name", "Family Name"]
for pname in sorted_param_names:
    pdef = all_param_defs[pname]
    if hasattr(pdef, "ParameterType") and pdef.ParameterType == ParameterType.Length:
        header.append("{0} (Inches)".format(pname))
    else:
        header.append(pname)

# Build rows
rows = []
for ftype in family_types:
    try:
        type_name = ftype.Name.strip()
    except:
        type_name = "Unnamed"

    try:
        family_name = family.Name.strip()
    except:
        family_name = "UnknownFamily"

    row = [type_name, family_name]
    param_map = {p.Definition.Name: p for p in ftype.Parameters}

    for pname in sorted_param_names:
        val = ""
        p = param_map.get(pname)
        if p and p.HasValue:
            try:
                if p.Definition.ParameterType == ParameterType.Length:
                    val = str(round(p.AsDouble() * 12.0, 6))  # Convert to decimal inches
                else:
                    val = p.AsValueString() or ""
            except:
                # fallback per storage type
                if p.StorageType == StorageType.Integer:
                    val = str(p.AsInteger())
                elif p.StorageType == StorageType.Double:
                    val = str(round(p.AsDouble() * 12.0, 6)) if p.Definition.ParameterType == ParameterType.Length else str(round(p.AsDouble(), 6))
                elif p.StorageType == StorageType.String:
                    val = p.AsString() or ""
                elif p.StorageType == StorageType.ElementId:
                    val = str(p.AsElementId().IntegerValue)
        row.append(val.strip())

    rows.append(row)

# Save CSV
save_dialog = SaveFileDialog()
save_dialog.Title = "Save Type Parameter Export"
save_dialog.Filter = "CSV files (*.csv)|*.csv"
save_dialog.FileName = "FamilyType_Export_Inches.csv"

if save_dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK:
    script.exit()

csv_path = save_dialog.FileName

# Write CSV
with codecs.open(csv_path, 'w', encoding='utf-8') as f:
    writer = csv.writer(f)
    writer.writerow(header)
    for row in rows:
        writer.writerow(row)

# Open file
os.startfile(csv_path)
