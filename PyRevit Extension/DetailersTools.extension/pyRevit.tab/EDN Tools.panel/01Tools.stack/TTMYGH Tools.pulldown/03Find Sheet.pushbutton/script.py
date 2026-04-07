# -*- coding: utf-8 -*-
"""Find Sheet - Find which sheet a view is placed on."""

from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc


def get_selected_view():
    """Get single view from selection."""
    selection_ids = uidoc.Selection.GetElementIds()
    
    for eid in selection_ids:
        elem = doc.GetElement(eid)
        if elem and isinstance(elem, DB.View):
            if isinstance(elem, DB.ViewSheet):
                continue
            if isinstance(elem, DB.ViewSchedule):
                continue
            if elem.IsTemplate:
                continue
            return elem
    
    return None


def find_sheets_for_view(view):
    """Find all sheets containing this view."""
    sheets = []
    
    all_sheets = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSheet)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    for sheet in all_sheets:
        if view.Id in sheet.GetAllPlacedViews():
            sheets.append(sheet)
    
    return sheets


def main():
    # Check selection first
    view = get_selected_view()
    
    if not view:
        forms.alert(
            "Select a view in the Project Browser first, then run this tool.",
            title="Find Sheet"
        )
        return
    
    # Find sheets
    sheets = find_sheets_for_view(view)
    
    # Show result
    if not sheets:
        forms.alert(
            "'{}' is not placed on any sheet.".format(view.Name),
            title="Find Sheet"
        )
    elif len(sheets) == 1:
        sheet = sheets[0]
        if forms.alert(
            "'{}' is on:\n\n{} - {}\n\nGo to sheet?".format(
                view.Name, sheet.SheetNumber, sheet.Name),
            title="Find Sheet",
            yes=True, no=True
        ):
            uidoc.ActiveView = sheet
    else:
        # Multiple sheets
        sheet_dict = {}
        for s in sheets:
            key = "{} - {}".format(s.SheetNumber, s.Name)
            sheet_dict[key] = s
        
        selected = forms.SelectFromList.show(
            sorted(sheet_dict.keys()),
            title="'{}' on {} sheets".format(view.Name, len(sheets)),
            button_name="Go to Sheet",
            multiselect=False
        )
        
        if selected:
            uidoc.ActiveView = sheet_dict[selected]


if __name__ == "__main__":
    main()
