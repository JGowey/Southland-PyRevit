import clr
import sys
import csv
import System  # Explicitly import System

# Add Revit API references
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import TaskDialog

clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import SaveFileDialog, DialogResult, Button, Form, Label

import System.Diagnostics  # Import for opening files

# Debug: Script started
print("Script started")

# Get active document
uidoc = __revit__.ActiveUIDocument
if uidoc is None or uidoc.Document is None:
    TaskDialog.Show("Error", "No document is open. Please open a family document.")
    sys.exit()

doc = uidoc.Document

# Ensure the document is a family document
if not doc.IsFamilyDocument:
    TaskDialog.Show("Error", "This script must be run in a Family Document.")
    sys.exit()

# Get the family manager
family_manager = doc.FamilyManager
parameters = family_manager.Parameters

# Prepare data for CSV
csv_data = []

# Create header row for parameter names
header_row = [""]  # Leave the first cell (A1) blank
for param in parameters:
    param_name = param.Definition.Name
    storage_type = "OTHER"  # Default if we don't recognize the parameter type

    # Determine unit type based on StorageType
    if param.StorageType == StorageType.Double:
        if "MASS" in param_name.upper():
            storage_type = "MASS##POUNDS_MASS"
        elif "AREA" in param_name.upper():
            storage_type = "AREA##SQUARE_FEET"
        else:
            storage_type = "LENGTH##FEET"  # Default for double values is length

    header_row.append(param_name + "##" + storage_type)

csv_data.append(header_row)  # Add the header row for type catalog

# Iterate over all family types
for ftype in family_manager.Types:
    data_row = [ftype.Name]  # Add the type name as the first entry in the row

    for param in parameters:
        param_value = ""

        if param.StorageType == StorageType.Double:
            param_value = str(ftype.AsDouble(param))  # Keep in Revit's native internal units

        elif param.StorageType == StorageType.Integer:
            param_value = str(ftype.AsInteger(param))

        elif param.StorageType == StorageType.String:
            param_value = ftype.AsString(param) or ""

        elif param.StorageType == StorageType.ElementId:
            param_value = str(ftype.AsElementId(param).IntegerValue)

        data_row.append(param_value)

    csv_data.append(data_row)  # Append the row with type name and parameter values

# Show save file dialog
save_file_dialog = SaveFileDialog()
save_file_dialog.Filter = "CSV files (*.csv)|*.csv"
save_file_dialog.Title = "Save Type Catalog"
save_file_dialog.InitialDirectory = System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments)
save_file_dialog.DefaultExt = "csv"  # Set default extension to .csv

if save_file_dialog.ShowDialog() == DialogResult.OK:
    file_path = save_file_dialog.FileName
    if not file_path.endswith('.csv'):  # Check if the file doesn't have .csv extension
        file_path += '.csv'  # Add .csv extension if missing

    with open(file_path, 'wb') as file:  # 'wb' for IronPython compatibility
        writer = csv.writer(file, delimiter=',', lineterminator='\n')
        for row in csv_data:
            file.write(','.join(row).encode('utf-8') + b'\n')  # Write manually to ensure correct encoding
    
    # Create a unified form with success message and buttons
    class CSVForm(Form):
        def __init__(self):
            self.Text = "CSV File Saved"
            self.Width = 300
            self.Height = 200

            # Display success message
            self.success_label = Label()
            self.success_label.Text = "Thanks For Using Jacob's Export Tool"
            self.success_label.Top = 30
            self.success_label.Left = 50
            self.success_label.Width = 200
            self.Controls.Add(self.success_label)

            # Open CSV Button
            self.open_button = Button()
            self.open_button.Text = "Open CSV"
            self.open_button.Top = 100
            self.open_button.Left = 50
            self.open_button.Click += self.open_csv
            self.Controls.Add(self.open_button)

            # Close Button
            self.close_button = Button()
            self.close_button.Text = "Close"
            self.close_button.Top = 100
            self.close_button.Left = 150
            self.close_button.Click += self.close_form
            self.Controls.Add(self.close_button)

        def open_csv(self, sender, args):
            # Open the saved CSV file
            System.Diagnostics.Process.Start(file_path)

        def close_form(self, sender, args):
            self.Close()

    # Show the form with the success message and buttons
    form = CSVForm()
    form.ShowDialog()

else:
    TaskDialog.Show("Cancelled", "Operation cancelled by user.")
