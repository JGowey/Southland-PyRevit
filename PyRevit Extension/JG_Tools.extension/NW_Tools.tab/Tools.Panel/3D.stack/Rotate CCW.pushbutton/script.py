# -*- coding: utf-8 -*-
import clr
import sys
import math
import json
import os

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from System.Collections.Generic import List  # <-- added for RotateElements

# Get Revit context
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Define a folder and file for user-specific settings in AppData
appdata_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Roaming')
TOOL_FOLDER = os.path.join(appdata_path, "PyRevit JG_Tools", "3D Roate Settings")
SETTINGS_FILE = os.path.join(TOOL_FOLDER, "settings.json")
DEFAULT_ANGLE = 90.0

def get_rotation_angle():
    """Reads the rotation angle from the settings file. Returns default if not found."""
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                return float(settings.get("RotationAngle", DEFAULT_ANGLE))
        except Exception:
            return DEFAULT_ANGLE
    return DEFAULT_ANGLE

def get_rotation_axis():
    """
    Determines the rotation axis based on the active view type and its orientation.
    Returns the rotation axis (XYZ) or None if the view is not supported.
    """
    active_view = doc.ActiveView
    
    # Check for Plan or Section Views
    if isinstance(active_view, ViewPlan):
        return XYZ.BasisZ # Plan views always rotate around the vertical Z-axis.
    elif isinstance(active_view, ViewSection):
        return active_view.ViewDirection # Section views rotate around the axis perpendicular to the plane.
    
    # Check for Orthogonal 3D Views (Top, Bottom, Front, Back, Left, Right)
    elif isinstance(active_view, View3D) and not active_view.IsTemplate:
        forward = active_view.GetOrientation().ForwardDirection
        
        # Check if the forward direction is aligned with a primary axis
        # and ignore small floating point errors
        if abs(forward.X) > 0.9999 or abs(forward.Y) > 0.9999 or abs(forward.Z) > 0.9999:
            return forward
        else:
            TaskDialog.Show("Error", "Rotation is only supported in orthogonal 3D views (like Top, Front, etc.).")
            return None
    
    # Fallback for all other unsupported view types
    TaskDialog.Show("Error", "Please switch to a supported view (Plan, Section, or an orthogonal 3D view) before running this script.")
    return None

def rotate_selected_element():
    """Rotate the selected Revit elements counter-clockwise as a unit using the user's settings."""
    selection_ids = uidoc.Selection.GetElementIds()
    if not selection_ids:
        TaskDialog.Show("Error", "No element selected. Please select one or more RFA or ITM instances in Revit.")
        return

    # Get the rotation axis based on the active view
    view_forward = get_rotation_axis()
    if not view_forward:
        return

    # Determine the angle from the settings file
    angle_degrees = get_rotation_angle()
    
    # Correct the angle for counter-clockwise rotation based on the view type.
    active_view = doc.ActiveView
    
    # For Plan and Section views, the rotation direction is inverted.
    # We apply a positive angle for counter-clockwise.
    if isinstance(active_view, ViewPlan) or isinstance(active_view, ViewSection):
        angle_degrees = angle_degrees
    # For 3D views, a positive angle is counter-clockwise when looking towards the origin.
    # Since the axis is based on the forward direction, we need to negate it.
    elif isinstance(active_view, View3D):
        angle_degrees = -angle_degrees

    angle_radians = math.radians(angle_degrees)

    # --- combined bounding box for the whole selection ---
    has_bbox = False
    min_x = min_y = min_z = None
    max_x = max_y = max_z = None

    for elem_id in selection_ids:
        elem = doc.GetElement(elem_id)
        if elem is None:
            continue
        bbox = elem.get_BoundingBox(None)
        if bbox is None:
            continue

        if not has_bbox:
            min_x, min_y, min_z = bbox.Min.X, bbox.Min.Y, bbox.Min.Z
            max_x, max_y, max_z = bbox.Max.X, bbox.Max.Y, bbox.Max.Z
            has_bbox = True
        else:
            min_x = min(min_x, bbox.Min.X)
            min_y = min(min_y, bbox.Min.Y)
            min_z = min(min_z, bbox.Min.Z)
            max_x = max(max_x, bbox.Max.X)
            max_y = max(max_y, bbox.Max.Y)
            max_z = max(max_z, bbox.Max.Z)

    if not has_bbox:
        TaskDialog.Show("Error", "Unable to determine bounding box for selection. Rotation failed.")
        return

    # Center of the combined bounding box = rotation point for the whole set
    rotation_point = XYZ(
        (min_x + max_x) / 2.0,
        (min_y + max_y) / 2.0,
        (min_z + max_z) / 2.0
    )

    axis = Line.CreateBound(rotation_point, rotation_point + view_forward)

    # Convert selection_ids to a .NET List[ElementId] for RotateElements
    id_list = List[ElementId]()
    for eid in selection_ids:
        id_list.Add(eid)

    t = Transaction(doc, "Rotate Selection Counterclockwise")
    t.Start()
    try:
        ElementTransformUtils.RotateElements(doc, id_list, axis, angle_radians)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Rotation failed: {}".format(str(e)))

def rotate_family_instance(element, angle_radians, view_forward):
    """Rotate a FamilyInstance (RFA)."""
    location = element.Location
    if not isinstance(location, LocationPoint):
        TaskDialog.Show("Error", "Invalid RFA location. Rotation failed.")
        return

    rotation_point = location.Point
    axis = Line.CreateBound(rotation_point, rotation_point + view_forward)

    t = Transaction(doc, "Rotate RFA Counterclockwise")
    t.Start()
    try:
        ElementTransformUtils.RotateElement(doc, element.Id, axis, angle_radians)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Rotation failed: {}".format(str(e)))

def rotate_itm(element, angle_radians, view_forward):
    """Rotate an ITM (MEPFabricationPart) based on the user's view direction."""
    bbox = element.get_BoundingBox(None)
    if not bbox:
        TaskDialog.Show("Error", "Invalid ITM bounding box. Rotation failed.")
        return

    rotation_point = XYZ(
        (bbox.Min.X + bbox.Max.X) / 2,
        (bbox.Min.Y + bbox.Max.Y) / 2,
        (bbox.Min.Z + bbox.Max.Z) / 2
    )

    axis = Line.CreateBound(rotation_point, rotation_point + view_forward)

    t = Transaction(doc, "Rotate ITM Counterclockwise")
    t.Start()
    try:
        ElementTransformUtils.RotateElement(doc, element.Id, axis, angle_radians)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "ITM Rotation failed: {}".format(str(e)))

# Rotate the selected element(s)
rotate_selected_element()
