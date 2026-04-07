# -*- coding: utf-8 -*-
import csv
import os
import subprocess

from Autodesk.Revit.DB import TextNote
from Autodesk.Revit.UI import (
    TaskDialog,
    TaskDialogCommonButtons,
    TaskDialogResult,
    TaskDialogCommandLinkId
)

from pyrevit import revit, script
from System.Windows.Forms import SaveFileDialog, DialogResult
from System import Environment
from System.Environment import SpecialFolder

# Optional output for pyRevit terminal
output = script.get_output()

def get_save_file_path():
    dialog = SaveFileDialog()
    dialog.Title = "Save Text Notes CSV"
    dialog.Filter = "CSV files (*.csv)|*.csv"

    # ✅ Use Windows "Documents" folder explicitly
    documents_path = Environment.GetFolderPath(SpecialFolder.MyDocuments)
    dialog.InitialDirectory = documents_path
    dialog.FileName = "TextNotesExport.csv"

    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None

def show_success_dialog(filepath):
    dialog = TaskDialog("Export Successful")
    dialog.MainInstruction = "Text Notes exported successfully!"
    dialog.MainContent = "File saved to:\n\n{}".format(filepath)
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open CSV File")
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    dialog.DefaultButton = TaskDialogResult.Close

    result = dialog.Show()

    if result == TaskDialogResult.CommandLink1:
        try:
            subprocess.Popen(filepath, shell=True)
        except:
            TaskDialog.Show("Error", "Could not open the file.")

def export_selected_textnotes_to_csv(doc, uidoc):
    selection_ids = uidoc.Selection.GetElementIds()

    if not selection_ids or selection_ids.Count == 0:
        TaskDialog.Show("Export Text Notes", "Please select some Text Notes before running the script.")
        return

    textnotes = []
    for elid in selection_ids:
        el = doc.GetElement(elid)
        if isinstance(el, TextNote):
            textnotes.append(el)

    if not textnotes:
        TaskDialog.Show("Export Text Notes", "No Text Notes found in the selection.")
        return

    filepath = get_save_file_path()
    if not filepath:
        TaskDialog.Show("Export Cancelled", "No file selected. Export cancelled.")
        return

    try:
        with open(filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['ElementId', 'Text'])
            for tn in textnotes:
                clean_text = tn.Text.replace('\r', '').replace('\n', '').encode('utf-8')
                writer.writerow([tn.Id.IntegerValue, clean_text])
        show_success_dialog(filepath)
    except Exception as e:
        TaskDialog.Show("Export Failed", str(e))

# Main execution
doc = revit.doc
uidoc = revit.uidoc
export_selected_textnotes_to_csv(doc, uidoc)
