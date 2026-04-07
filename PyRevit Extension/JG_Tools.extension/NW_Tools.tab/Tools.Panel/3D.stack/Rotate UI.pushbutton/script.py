# -*- coding: utf-8 -*-
import clr
import sys
import math
import json
import os

clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from System.Windows.Forms import Application, Form, Label, TextBox, Button
from System.Drawing import Point, Size
from System.Collections.Generic import List  # needed for RotateElements

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

def save_rotation_angle(angle):
    """Saves the rotation angle to the settings file."""
    if not os.path.exists(TOOL_FOLDER):
        os.makedirs(TOOL_FOLDER)
    
    settings = {"RotationAngle": angle}
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        TaskDialog.Show("Error", "Failed to save settings file: {}".format(str(e)))

def get_rotation_axis():
    """Determines the rotation axis based on the active view type."""
    active_view = doc.ActiveView
    
    if isinstance(active_view, ViewPlan):
        return XYZ.BasisZ
    elif isinstance(active_view, ViewSection):
        return active_view.ViewDirection
    elif isinstance(active_view, View3D) and not active_view.IsTemplate:
        forward = active_view.GetOrientation().ForwardDirection
        if abs(forward.X) > 0.9999 or abs(forward.Y) > 0.9999 or abs(forward.Z) > 0.9999:
            return forward
        else:
            TaskDialog.Show("Error", "Rotation only supported in orthogonal 3D views.")
            return None

    TaskDialog.Show("Error", "Unsupported view for rotation.")
    return None

class RotateParameterHandler(IExternalEventHandler):
    """Handles rotation inside the Revit API context."""
    def __init__(self):
        # Existing view-based rotation flag
        self.clockwise = True
        # NEW: flags & data for model-axis rotation
        self.use_model_axis = False
        self.model_axis_vector = None
        self.model_axis_angle = 0.0
        self.model_axis_label = ""

    def Execute(self, uiapp):
        # Called in Revit API context
        try:
            if self.use_model_axis and self.model_axis_vector is not None:
                # Model-axis rotation
                rotate_selection_about_model_axis(
                    self.model_axis_vector,
                    self.model_axis_angle,
                    self.model_axis_label
                )
            else:
                # Original view-based rotation
                rotate_selected_element(self.clockwise)
        finally:
            # Reset mode after each execution
            self.use_model_axis = False
            self.model_axis_vector = None
            self.model_axis_angle = 0.0
            self.model_axis_label = ""

    def GetName(self):
        return "Rotate Parameter Handler"

def compute_selection_center(selection_ids):
    """Compute center of combined bounding box for selection. Returns (center, has_bbox)."""
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
        return None, False

    center = XYZ(
        (min_x + max_x) / 2.0,
        (min_y + max_y) / 2.0,
        (min_z + max_z) / 2.0
    )
    return center, True

def rotate_selected_element(clockwise=True):
    """Rotate selected elements as one unit, using view-based axis (CW/CCW)."""
    selection_ids = uidoc.Selection.GetElementIds()
    if not selection_ids:
        TaskDialog.Show("Error", "No elements selected.")
        return

    view_forward = get_rotation_axis()
    if not view_forward:
        return

    angle_degrees = get_rotation_angle()
    active_view = doc.ActiveView

    # rotation direction logic
    if isinstance(active_view, ViewPlan) or isinstance(active_view, ViewSection):
        if clockwise:
            angle_degrees = -angle_degrees
    elif isinstance(active_view, View3D):
        if not clockwise:
            angle_degrees = -angle_degrees

    angle_radians = math.radians(angle_degrees)

    rotation_point, has_bbox = compute_selection_center(selection_ids)
    if not has_bbox:
        TaskDialog.Show("Error", "Bounding box error.")
        return

    axis = Line.CreateBound(rotation_point, rotation_point + view_forward)

    id_list = List[ElementId]()
    for eid in selection_ids:
        id_list.Add(eid)

    t = Transaction(doc, "Rotate Selection")
    t.Start()
    try:
        ElementTransformUtils.RotateElements(doc, id_list, axis, angle_radians)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", str(e))

def rotate_selection_about_model_axis(axis_vector, angle_degrees, label):
    """
    Rotate selected elements as one unit around a model axis (XYZ.BasisX/Y/Z),
    independent of the view. CALLED FROM EXTERNAL EVENT HANDLER ONLY.
    """
    selection_ids = uidoc.Selection.GetElementIds()
    if not selection_ids:
        TaskDialog.Show("Error", "No elements selected.")
        return

    rotation_point, has_bbox = compute_selection_center(selection_ids)
    if not has_bbox:
        TaskDialog.Show("Error", "Bounding box error.")
        return

    angle_radians = math.radians(angle_degrees)
    axis = Line.CreateBound(rotation_point, rotation_point + axis_vector)

    id_list = List[ElementId]()
    for eid in selection_ids:
        id_list.Add(eid)

    t = Transaction(doc, "Rotate Selection about {}".format(label))
    t.Start()
    try:
        ElementTransformUtils.RotateElements(doc, id_list, axis, angle_radians)
        t.Commit()
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "Rotation about {} failed: {}".format(label, str(e)))

class RotationParameterForm(Form):
    """Floating UI for entering a rotation angle and auto-saving."""
    def __init__(self):
        self.Text = "Set Rotation Angle"
        self.Size = Size(420, 350)
        self.TopMost = True
        
        # View-based angle (existing behavior)
        self.label1 = Label(Text="View-based rotation (degrees):", Location=Point(20, 20),AutoSize=False,Size=Size(220, 20))
        self.Controls.Add(self.label1)

        self.textbox = TextBox(Location=Point(20, 45), Size=Size(100, 20), Text=str(get_rotation_angle()))
        self.Controls.Add(self.textbox)

        # AUTO SAVE WHEN LEAVING
        self.textbox.Leave += self.on_textbox_leave

        self.rotate_cw_button = Button(Text="Rotate CW", Location=Point(20, 80), Size=Size(140, 30))
        self.rotate_cw_button.Click += self.on_rotate_cw_clicked
        self.Controls.Add(self.rotate_cw_button)

        self.rotate_ccw_button = Button(Text="Rotate CCW", Location=Point(180, 80), Size=Size(140, 30))
        self.rotate_ccw_button.Click += self.on_rotate_ccw_clicked
        self.Controls.Add(self.rotate_ccw_button)

        # --- Model-axis rotation fields (X/Y/Z) ---

        self.label_axes = Label(Text="Model axis rotation (degrees):", Location=Point(20, 125),AutoSize=False,Size=Size(220, 20))
        self.Controls.Add(self.label_axes)

        # X
        self.label_x = Label(Text="Rotate X:", Location=Point(20, 155), Size=Size(80, 20))
        self.Controls.Add(self.label_x)

        self.textbox_x = TextBox(Location=Point(100, 155), Size=Size(60, 20), Text="0")
        self.Controls.Add(self.textbox_x)

        self.button_x = Button(Text="Apply X", Location=Point(180, 150), Size=Size(80, 30))
        self.button_x.Click += self.on_rotate_x_clicked
        self.Controls.Add(self.button_x)

        # Y
        self.label_y = Label(Text="Rotate Y:", Location=Point(20, 190), Size=Size(80, 20))
        self.Controls.Add(self.label_y)

        self.textbox_y = TextBox(Location=Point(100, 190), Size=Size(60, 20), Text="0")
        self.Controls.Add(self.textbox_y)

        self.button_y = Button(Text="Apply Y", Location=Point(180, 185), Size=Size(80, 30))
        self.button_y.Click += self.on_rotate_y_clicked
        self.Controls.Add(self.button_y)

        # Z
        self.label_z = Label(Text="Rotate Z:", Location=Point(20, 225), Size=Size(80, 20))
        self.Controls.Add(self.label_z)

        self.textbox_z = TextBox(Location=Point(100, 225), Size=Size(60, 20), Text="0")
        self.Controls.Add(self.textbox_z)

        self.button_z = Button(Text="Apply Z", Location=Point(180, 220), Size=Size(80, 30))
        self.button_z.Click += self.on_rotate_z_clicked
        self.Controls.Add(self.button_z)

        # Close button
        self.cancel_button = Button(Text="Close", Location=Point(280, 260), Size=Size(100, 30))
        self.cancel_button.Click += self.on_cancel_clicked
        self.Controls.Add(self.cancel_button)

    # Save helper for main angle
    def save_angle_from_textbox(self, show_error=True):
        txt = self.textbox.Text.strip()
        if not txt:
            return
        try:
            angle = float(txt)
            save_rotation_angle(angle)
        except:
            if show_error:
                TaskDialog.Show("Error", "Invalid number for view-based angle.")

    def on_textbox_leave(self, sender, e):
        self.save_angle_from_textbox(show_error=False)

    def on_rotate_cw_clicked(self, sender, e):
        self.save_angle_from_textbox(show_error=False)
        rotate_handler.use_model_axis = False  # ensure view-based mode
        rotate_handler.clockwise = True
        rotate_event.Raise()

    def on_rotate_ccw_clicked(self, sender, e):
        self.save_angle_from_textbox(show_error=False)
        rotate_handler.use_model_axis = False  # ensure view-based mode
        rotate_handler.clockwise = False
        rotate_event.Raise()

    # X/Y/Z buttons now just set handler fields and raise ExternalEvent
    def on_rotate_x_clicked(self, sender, e):
        txt = self.textbox_x.Text.strip()
        if not txt:
            return
        try:
            angle = float(txt)
        except:
            TaskDialog.Show("Error", "Invalid number for Rotate X.")
            return

        rotate_handler.use_model_axis = True
        rotate_handler.model_axis_vector = XYZ.BasisX
        rotate_handler.model_axis_angle = angle
        rotate_handler.model_axis_label = "X"
        rotate_event.Raise()

    def on_rotate_y_clicked(self, sender, e):
        txt = self.textbox_y.Text.strip()
        if not txt:
            return
        try:
            angle = float(txt)
        except:
            TaskDialog.Show("Error", "Invalid number for Rotate Y.")
            return

        rotate_handler.use_model_axis = True
        rotate_handler.model_axis_vector = XYZ.BasisY
        rotate_handler.model_axis_angle = angle
        rotate_handler.model_axis_label = "Y"
        rotate_event.Raise()

    def on_rotate_z_clicked(self, sender, e):
        txt = self.textbox_z.Text.strip()
        if not txt:
            return
        try:
            angle = float(txt)
        except:
            TaskDialog.Show("Error", "Invalid number for Rotate Z.")
            return

        rotate_handler.use_model_axis = True
        rotate_handler.model_axis_vector = XYZ.BasisZ
        rotate_handler.model_axis_angle = angle
        rotate_handler.model_axis_label = "Z"
        rotate_event.Raise()

    def on_cancel_clicked(self, sender, e):
        self.Close()

# Create External Events and Handlers
rotate_handler = RotateParameterHandler()
rotate_event = ExternalEvent.Create(rotate_handler)

Application.EnableVisualStyles()
rotation_form = RotationParameterForm()
rotation_form.Show()
