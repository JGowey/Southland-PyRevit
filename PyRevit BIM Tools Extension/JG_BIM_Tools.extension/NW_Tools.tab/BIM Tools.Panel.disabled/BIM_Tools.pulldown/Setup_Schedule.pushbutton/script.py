import clr
import os
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from System.Collections.Generic import List

# Revit context from pyRevit
uidoc = __revit__.ActiveUIDocument
if uidoc is None:
    TaskDialog.Show("Error", "No active Revit document is open.")
    raise Exception("ActiveUIDocument is None")

doc = uidoc.Document
app = __revit__.Application

def run(doc, app):
    shell_path = r"\\si.net\si\TCM\BIM\PyRevit\PyRevit_Content_Library\Schedule_Assets\PyRevit_Schedule_Asset_R_22.rvt"
    schedule_name_to_find = "Weight"

    if not os.path.exists(shell_path):
        TaskDialog.Show("Error", "Shell project not found:\n" + shell_path)
        return

    shell_doc = None
    try:
        shell_doc = app.OpenDocumentFile(shell_path)

        # Find the schedule
        collector = FilteredElementCollector(shell_doc).OfClass(ViewSchedule)
        schedule_to_copy = None
        for vs in collector:
            if not vs.IsTemplate and vs.Name == schedule_name_to_find:
                schedule_to_copy = vs
                break

        if not schedule_to_copy:
            TaskDialog.Show("Not Found", "Schedule '{}' was not found.".format(schedule_name_to_find))
            return

        # Begin transaction in the current document
        t = Transaction(doc, "Import Schedule: " + schedule_name_to_find)
        t.Start()

        ids_to_copy = List[ElementId]()
        ids_to_copy.Add(schedule_to_copy.Id)
        copied_ids = ElementTransformUtils.CopyElements(shell_doc, ids_to_copy, doc, None, None)

        t.Commit()

        TaskDialog.Show("Success", "Schedule '{}' copied successfully.".format(schedule_name_to_find))

    except Exception as e:
        if shell_doc:
            try:
                shell_doc.Close(False)
            except:
                pass
        TaskDialog.Show("Exception", "Error:\n" + str(e))
        return

    finally:
        if shell_doc:
            try:
                shell_doc.Close(False)
            except:
                pass

run(doc, app)
