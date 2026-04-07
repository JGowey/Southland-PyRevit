import clr
import sys
import csv
import time

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import OpenFileDialog, DialogResult

doc = __revit__.ActiveUIDocument.Document
if doc is None or not doc.IsFamilyDocument:
    TaskDialog.Show("Error", "Open a family document before running this script.")
    sys.exit()

def clean_header(header):
    """Removes units and extra characters from CSV headers to match Revit parameters."""
    return header.split("##")[0].strip()

def convert_value(storage_type, value):
    """Converts CSV values to Revit data types efficiently."""
    if not value.strip():
        return None  # Skip empty values

    try:
        if storage_type == StorageType.String:
            return value
        elif storage_type == StorageType.Integer:
            return int(value) if value.isdigit() else None  # Ensure integer-only values
        elif storage_type == StorageType.Double:
            return float(value.replace(',', '').strip()) if value.replace(',', '').replace('.', '').isdigit() else None
        elif storage_type == StorageType.ElementId:
            return ElementId(int(value)) if value.isdigit() else ElementId.InvalidElementId
    except Exception as e:
        print("Error converting value '{}': {}".format(value, e))
    return None

def get_valid_parameters(family_manager, headers):
    """Pre-checks CSV vs. Revit parameters, skipping formula-driven ones."""
    param_map = {clean_header(header): idx for idx, header in enumerate(headers)}
    
    revit_params = {p.Definition.Name.strip(): p for p in family_manager.Parameters}
    param_storage = {name: param.StorageType for name, param in revit_params.items()}
    
    # Identify common parameters (skip formula-driven ones)
    valid_params = [
        name for name in revit_params 
        if name in param_map and not revit_params[name].IsDeterminedByFormula
    ]

    return param_map, revit_params, param_storage, valid_params

def import_family_types(csv_path):
    """Optimized family type import with correct type-value mapping."""
    family_manager = doc.FamilyManager

    with open(csv_path, 'r') as file:
        reader = csv.reader(file)
        headers = next(reader)
        rows = list(reader)

        # Pre-check parameters once
        param_map, revit_params, param_storage, valid_params = get_valid_parameters(family_manager, headers)

    # Create a dictionary mapping type names to their CSV row data
    type_data_map = {}
    for row in rows:
        type_name = row[0].strip()
        if type_name:
            type_data_map[type_name] = row  # Store the full row for later lookup

    # Store references to created types
    type_lookup = {}

    t = Transaction(doc, "Create Family Types")
    t.Start()
    try:
        for type_name in type_data_map.keys():
            if type_name:
                type_lookup[type_name] = family_manager.NewType(type_name)
        t.Commit()
    except:
        t.RollBack()
        raise

    # Apply parameters only to the corresponding type
    t = Transaction(doc, "Set Family Parameters")
    t.Start()
    try:
        for type_name, row in type_data_map.items():
            if type_name not in type_lookup:
                continue  # Skip if type creation failed

            family_manager.CurrentType = type_lookup[type_name]  # Set correct type

            param_values = {}
            for param_name in valid_params:
                if param_name in param_map:  # Ensure param exists in CSV
                    csv_index = param_map[param_name]
                    if csv_index < len(row):  # Ensure index is in range
                        param_value = row[csv_index].strip()
                        converted_value = convert_value(param_storage[param_name], param_value)
                        if converted_value is not None:
                            param_values[revit_params[param_name]] = converted_value
                        else:
                            print("Warning: Skipping '{}' for '{}'. Invalid data.".format(param_name, type_name))

            # Assign parameter values
            for param, val in param_values.items():
                try:
                    family_manager.Set(param, val)
                except Exception as e:
                    print("Failed to set '{}' for '{}': {}".format(param.Definition.Name, type_name, e))

        t.Commit()
        return len(rows)
    except:
        t.RollBack()
        raise

def delete_all_family_types(family_manager):
    """Deletes all family types except one placeholder."""
    types = list(family_manager.Types)
    if len(types) <= 1:
        return 0

    original_type_count = len(types)
    t = Transaction(doc, "Delete Existing Types")
    t.Start()
    try:
        temp_type = family_manager.NewType("TEMP_DELETE")
        types_to_delete = [ft for ft in types if ft.Name != "TEMP_DELETE"]
        for ft in types_to_delete:
            family_manager.CurrentType = ft
            family_manager.DeleteCurrentType()
        family_manager.CurrentType = temp_type
        family_manager.RenameCurrentType("Default Type")
        t.Commit()
        return original_type_count
    except:
        t.RollBack()
        raise

def delete_default_type(family_manager):
    """Deletes 'Default Type' if it exists."""
    t = Transaction(doc, "Delete Default Type")
    t.Start()
    try:
        default_type = next((ft for ft in family_manager.Types if ft.Name == "Default Type"), None)
        if default_type:
            family_manager.CurrentType = default_type
            family_manager.DeleteCurrentType()
        t.Commit()
    except:
        t.RollBack()
    finally:
        if t.GetStatus() == TransactionStatus.Started:
            t.RollBack()

if __name__ == "__main__":
    dialog = OpenFileDialog()
    dialog.Filter = "CSV Files (*.csv)|*.csv"
    if dialog.ShowDialog() == DialogResult.OK:
        total_start = time.time()
        try:
            family_manager = doc.FamilyManager

            # Delete old types
            del_start = time.time()
            deleted_count = delete_all_family_types(family_manager)
            del_time = time.time() - del_start

            # Import new types
            imp_start = time.time()
            imported_count = import_family_types(dialog.FileName)
            imp_time = time.time() - imp_start

            # Delete "Default Type" after import
            delete_default_type(family_manager)

            # Success message
            success_msg = (
                "Deleted {} old types in {:.1f}s\n"
                "Imported {} new types in {:.1f}s\n"
                "Total time: {:.1f}s"
            ).format(deleted_count, del_time, imported_count, imp_time, time.time() - total_start)

            TaskDialog.Show("Success", success_msg)
        except Exception as e:
            TaskDialog.Show("Error", str(e))
