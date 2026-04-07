# -*- coding: utf-8 -*-
import csv
import os

from Autodesk.Revit.DB import TextNote, ElementId, Transaction
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, script

from System.Windows.Forms import OpenFileDialog, DialogResult
from System import Environment
from System.Environment import SpecialFolder

# Optional console output
output = script.get_output()

def get_csv_file_path():
    dialog = OpenFileDialog()
    dialog.Title = "Select CSV File to Import"
    dialog.Filter = "CSV files (*.csv)|*.csv"

    # ✅ Use real Windows Documents path
    documents_path = Environment.GetFolderPath(SpecialFolder.MyDocuments)
    dialog.InitialDirectory = documents_path

    if dialog.ShowDialog() == DialogResult.OK:
        return dialog.FileName
    return None

def import_textnote_updates(doc):
    filepath = get_csv_file_path()
    if not filepath:
        TaskDialog.Show("Import Cancelled", "No file selected.")
        return

    updated_count = 0
    skipped_count = 0

    try:
        with open(filepath, 'rb') as csvfile:
            reader = csv.reader(csvfile)
            header = reader.next()  # Skip header row

            with Transaction(doc, "Update Text Notes from CSV") as t:
                t.Start()
                for row in reader:
                    try:
                        element_id = int(row[0])
                        new_text = row[1].decode('utf-8')
                        element = doc.GetElement(ElementId(element_id))

                        if element and isinstance(element, TextNote):
                            element.Text = new_text
                            updated_count += 1
                        else:
                            skipped_count += 1
                    except:
                        skipped_count += 1
                t.Commit()

        TaskDialog.Show(
            "Import Complete",
            "Updated {} Text Notes.\nSkipped {} rows.".format(updated_count, skipped_count)
        )

    except Exception as e:
        TaskDialog.Show("Import Failed", str(e))

# Main execution
doc = revit.doc
import_textnote_updates(doc)
