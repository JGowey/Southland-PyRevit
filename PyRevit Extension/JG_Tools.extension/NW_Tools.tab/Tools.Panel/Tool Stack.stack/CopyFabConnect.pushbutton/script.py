# -*- coding: utf-8 -*-
# TOOL: Smart Paste & Align Chain
# DESCRIPTION: Live placement with Spacebar rotation, non-blocking UI
# AUTHOR: JAcob gowey, Chat GPT, Gemini
# DATE: 2025-07-31
# VERSION: 12.9

import clr
import sys
import math
import System
import System.Reflection

clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

import System.Windows.Forms as WinForms
from System.Windows.Forms import Keys
from System.Collections.Generic import List

from pyrevit import revit, forms

# Globals
uidoc = revit.uidoc
doc = revit.doc
TOLERANCE = 1e-6

# Keyboard listener - Initialized as None, set up later in main()
GetAsyncKeyState = None

# UI Form
class InteractivePlacementForm(WinForms.Form):
    def __init__(self):
        self.Text = "Live Placement Control"
        self.TopMost = True
        self.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
        self.Width = 320
        self.Height = 160

        self.StartPosition = WinForms.FormStartPosition.Manual
        screen_area = WinForms.Screen.PrimaryScreen.WorkingArea
        self.Top = 10
        self.Left = screen_area.Left + (screen_area.Width - self.Width) // 2

        label = WinForms.Label()
        label.Text = ("Placement Mode Active:\n\n"
                      "Spacebar = Rotate 45°\n"
                      "Enter = Place Chain\n"
                      "Esc = Cancel Placement")
        label.Dock = WinForms.DockStyle.Fill
        label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        label.Font = System.Drawing.Font("Arial", 10, System.Drawing.FontStyle.Bold)
        self.Controls.Add(label)

# Helpers
def get_connector_manager(elem):
    if hasattr(elem, 'MEPModel') and elem.MEPModel:
        return elem.MEPModel.ConnectorManager
    elif hasattr(elem, 'ConnectorManager') and elem.ConnectorManager:
        return elem.ConnectorManager
    return None

def find_closest_connector(connectors, point):
    if not connectors:
        return None
    return min(connectors, key=lambda c: c.Origin.DistanceTo(point))

class PreselectionFilter(ISelectionFilter):
    def __init__(self, valid_ids):
        self.valid_ids = valid_ids
    def AllowElement(self, elem):
        return elem.Id.IntegerValue in self.valid_ids
    def AllowReference(self, ref, pt):
        return True

class OpenConnectorFilter(ISelectionFilter):
    def AllowElement(self, elem):
        mgr = get_connector_manager(elem)
        if mgr:
            return any(not c.IsConnected for c in mgr.Connectors)
        return False

    def AllowReference(self, ref, pt):
        elem = doc.GetElement(ref.ElementId)
        if elem:
            mgr = get_connector_manager(elem)
            if mgr:
                closest_conn = find_closest_connector(list(mgr.Connectors), pt)
                if closest_conn and not closest_conn.IsConnected:
                    return True
        return False

def rotate_elements_around_point(element_ids, center, axis, angle_rad):
    if not element_ids or abs(angle_rad) < TOLERANCE:
        return
    valid_ids = [eid for eid in element_ids if doc.GetElement(eid) is not None]
    if not valid_ids:
        raise Exception("All copied elements were removed by Revit. Likely due to welds/flanges.")
    line = Line.CreateBound(center, center + axis)
    ElementTransformUtils.RotateElements(doc, List[ElementId](valid_ids), line, angle_rad)

def get_perpendicular_axis(vec):
    if abs(vec.X) < 0.1 and abs(vec.Y) < 0.1:
        return XYZ.BasisX
    else:
        return vec.CrossProduct(XYZ.BasisZ).Normalize()

def show_user_warning(message, title="Warning"):
    WinForms.MessageBox.Show(message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning)

def get_all_connectors_in_selection(elements):
    """
    Collects all unique connectors from all elements in the provided set.
    Uses a dictionary for efficient de-duplication based on connector properties.
    """
    unique_connectors_dict = {}
    for el in elements:
        mgr = get_connector_manager(el)
        if mgr:
            for conn in mgr.Connectors:
                # Create a robust unique key for each connector using integer values
                # for coordinates and direction to handle floating point imprecision.
                origin_tuple = (int(conn.Origin.X * 1000), int(conn.Origin.Y * 1000), int(conn.Origin.Z * 1000))
                direction_tuple = (int(conn.CoordinateSystem.BasisZ.X * 1000),
                                   int(conn.CoordinateSystem.BasisZ.Y * 1000),
                                   int(conn.CoordinateSystem.BasisZ.Z * 1000))
                owner_id_int = conn.Owner.Id.IntegerValue

                connector_key = (origin_tuple, direction_tuple, owner_id_int)

                # Add to dictionary if not already present
                if connector_key not in unique_connectors_dict:
                    unique_connectors_dict[connector_key] = conn

    return list(unique_connectors_dict.values())

# Main Script
def main():
    # Initialize keyboard listener for interactive features
    global GetAsyncKeyState
    if GetAsyncKeyState is None:
        try:
            import ctypes
            user32 = ctypes.WinDLL('user32', use_last_error=True)
            GetAsyncKeyState = user32.GetAsyncKeyState
            GetAsyncKeyState.argtypes = [ctypes.c_int]
            GetAsyncKeyState.restype = ctypes.c_short
        except Exception as e:
            forms.alert("Error loading user32.dll for keyboard polling. Interactive features will not work.\nDetails: {}".format(e), title="Initialization Error", exitscript=True)
            return

    if not GetAsyncKeyState:
        forms.alert("Failed to initialize keyboard listener. Script cannot run.", title="Initialization Error", exitscript=True)
        return

    # Get initial selection or prompt user to select elements
    selection_ids = uidoc.Selection.GetElementIds()
    if not selection_ids:
        try:
            refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements to copy, then click Finish")
            selection_ids = List[ElementId]([r.ElementId for r in refs])
        except Exception as e:
            if "cancel" in str(e).lower():
                sys.exit()
            forms.alert("Selection canceled or failed.\n" + str(e), title="Selection Error", exitscript=True)

    if not selection_ids or len(selection_ids) == 0:
        forms.alert("No elements selected. Exiting.", title="No Selection", exitscript=True)
        return

    # Convert ElementIds to Elements
    source_elements = [doc.GetElement(id) for id in selection_ids]
    source_id_set = set(e.Id.IntegerValue for e in source_elements)

    # Determine if any selected element is a Fabrication Part
    FabPartType = None
    for e in source_elements:
        if "FabricationPart" in e.GetType().FullName:
            FabPartType = e.GetType()
            break

    # Collect all connectors from the selected elements for the "transfer" pick
    all_source_connectors = get_all_connectors_in_selection(source_elements)
    if not all_source_connectors:
        forms.alert("No connectors found on selected elements. Please select elements with MEP connectors.", title="Error", exitscript=True)
        return
    elif len(all_source_connectors) == 1:
        forms.alert("Only one connector found on selected elements. Please select a chain with at least two connectors.", title="Single Connector", exitscript=True)
        return

    class TransferConnectorFilter(ISelectionFilter):
        def __init__(self, valid_ids, valid_connectors):
            self.valid_ids = valid_ids
            self.valid_connectors = valid_connectors

        def AllowElement(self, elem):
            return elem.Id.IntegerValue in self.valid_ids

        def AllowReference(self, ref, pt):
            elem = doc.GetElement(ref.ElementId)
            if elem and elem.Id.IntegerValue in self.valid_ids:
                mgr = get_connector_manager(elem)
                if mgr:
                    closest_conn = find_closest_connector(list(mgr.Connectors), pt)
                    # Check if the closest connector is one of the valid source connectors
                    for c in self.valid_connectors:
                        if c.Origin.IsAlmostEqualTo(closest_conn.Origin, TOLERANCE) and \
                           c.CoordinateSystem.BasisZ.IsAlmostEqualTo(closest_conn.CoordinateSystem.BasisZ, TOLERANCE) and \
                           c.Owner.Id == closest_conn.Owner.Id: # Use closest_conn from find_closest_connector if it was found
                            return True
            return False

    try:
        # User picks the "transfer" connector on the selected chain
        transfer_ref = uidoc.Selection.PickObject(ObjectType.Face, TransferConnectorFilter(source_id_set, all_source_connectors), "Pick the connector face on the selected chain to use as the transfer point")
        transfer_elem = doc.GetElement(transfer_ref.ElementId)
        transfer_mgr = get_connector_manager(transfer_elem)
        transfer_conns = list(transfer_mgr.Connectors)
        transfer_conn = find_closest_connector(transfer_conns, transfer_ref.GlobalPoint)

        # Validate that the picked connector is indeed from the source elements
        is_valid_transfer_conn = False
        for c in all_source_connectors:
            if c.Origin.IsAlmostEqualTo(transfer_conn.Origin, TOLERANCE) and \
               c.CoordinateSystem.BasisZ.IsAlmostEqualTo(transfer_conn.CoordinateSystem.BasisZ, TOLERANCE) and \
               c.Owner.Id == c.Owner.Id: # Corrected from closest_conn.Owner.Id to c.Owner.Id for validity check against original source connectors
                is_valid_transfer_conn = True
                break

        if not transfer_conn or not is_valid_transfer_conn:
            forms.alert("The picked face does not belong to a valid connector on the selected chain.", title="Invalid Selection", exitscript=True)

    except Exception as e:
        if "cancel" in str(e).lower():
            sys.exit()
        forms.alert("Error selecting transfer connector.\n" + str(e), title="Error", exitscript=True)

    # Main placement loop
    while True:
        try:
            # User picks the target open connector face
            target_ref = uidoc.Selection.PickObject(ObjectType.Face, OpenConnectorFilter(), "Pick TARGET open connector face (ESC to exit)")
            target_elem = doc.GetElement(target_ref.ElementId)
            target_mgr = get_connector_manager(target_elem)
            all_target_conns = list(target_mgr.Connectors)
            open_target_conns = [c for c in all_target_conns if not c.IsConnected]
            target_conn = find_closest_connector(open_target_conns, target_ref.GlobalPoint)

            if not target_conn:
                forms.alert("No open connectors found at the picked location. Try another target.", title="No Open Connector")
                continue # Allow user to pick again
        except Exception: # This catches the case where the user presses ESC during PickObject
            # print("Tool exited.") # REMOVED THIS LINE
            break # Exit the placement loop if picking is cancelled

        t_group = TransactionGroup(doc, "Paste & Align Group")
        t_group.Start()
        copied_ids = []
        form = None

        try:
            with Transaction(doc, "Copy & Initial Align") as t:
                t.Start()
                # Calculate move vector and copy elements
                move_vector = target_conn.Origin - transfer_conn.Origin
                copied_ids = ElementTransformUtils.CopyElements(doc, List[ElementId]([e.Id for e in source_elements]), move_vector)

                # Check if any copied elements were deleted by Revit (e.g., due to welding requirements)
                copied_ids = [eid for eid in copied_ids if doc.GetElement(eid) is not None]
                if len(copied_ids) != len(source_elements):
                    t.RollBack()
                    t_group.RollBack()
                    show_user_warning(
                        "Some elements were automatically deleted by Revit during duplication.\n"
                        "This happens with welds, flanges, or fittings that require a host connection.\n\n"
                        "Try selecting only pipe or fittings that can exist independently.",
                        title="Copy Failed"
                    )
                    break # Break out of placement loop if initial copy fails

                # Find the corresponding transfer element in the newly copied chain
                copied_transfer_elem = None
                original_element_index = -1
                for i, original_elem in enumerate(source_elements):
                    if original_elem.Id == transfer_elem.Id:
                        original_element_index = i
                        break

                if original_element_index != -1 and original_element_index < len(copied_ids):
                    copied_transfer_elem = doc.GetElement(copied_ids[original_element_index])

                if not copied_transfer_elem:
                    raise Exception("Could not find the copied version of the transfer element.")

                # Find the specific transfer connector on the copied element
                copied_transfer_mgr = get_connector_manager(copied_transfer_elem)
                copied_transfer_conns = list(copied_transfer_mgr.Connectors)

                # Find the connector on the copied element that matches the original transfer connector's properties
                # (origin relative to element and direction) after the initial move.
                expected_copied_transfer_conn_origin = transfer_conn.Origin + move_vector
                copied_transfer_conn = None
                min_dist = float('inf')

                for c_temp in copied_transfer_conns:
                    dist = c_temp.Origin.DistanceTo(expected_copied_transfer_conn_origin)
                    direction_match = c_temp.CoordinateSystem.BasisZ.IsAlmostEqualTo(transfer_conn.CoordinateSystem.BasisZ, TOLERANCE)

                    if dist < min_dist:
                        min_dist = dist
                        copied_transfer_conn = c_temp
                    elif dist == min_dist and direction_match:
                        copied_transfer_conn = c_temp # Prefer direction match as a tie-breaker

                if not copied_transfer_conn:
                    raise Exception("Could not find the transfer connector on the copied chain after initial move.")

                # Calculate rotation to align copied chain with target
                v1 = copied_transfer_conn.CoordinateSystem.BasisZ
                v2 = -target_conn.CoordinateSystem.BasisZ
                axis = v1.CrossProduct(v2)
                angle = v1.AngleTo(v2)

                if axis.IsAlmostEqualTo(XYZ.Zero, TOLERANCE):
                    axis = get_perpendicular_axis(v1) # Handle collinear vectors (180 deg)

                if angle > TOLERANCE:
                    try:
                        rotate_elements_around_point(copied_ids, target_conn.Origin, axis, angle)
                    except Exception as rot_err:
                        t.RollBack()
                        t_group.RollBack()
                        show_user_warning(str(rot_err), title="Rotation Failed")
                        break # Break out of placement loop if initial rotation fails
                t.Commit()

            uidoc.RefreshActiveView()

            placement_status = {"result": "pending"}

            # Define interactive actions
            def rotate_action():
                with Transaction(doc, "Rotate 45°") as t:
                    t.Start()
                    try:
                        rotate_elements_around_point(copied_ids, target_conn.Origin, target_conn.CoordinateSystem.BasisZ, math.pi / 4)
                    except Exception as rot_err:
                        t.RollBack()
                        show_user_warning(str(rot_err), title="Rotation Failed")
                        return
                    t.Commit()
                uidoc.RefreshActiveView()

            def place_action():
                placement_status["result"] = "place"
            def cancel_action():
                placement_status["result"] = "cancel"

            form = InteractivePlacementForm()
            form.Show()

            # Keyboard polling loop for interactive placement
            key_states = {
                Keys.Space: False,
                Keys.Enter: False,
                Keys.Escape: False
            }

            while placement_status["result"] == "pending":
                WinForms.Application.DoEvents() # Process UI events

                is_space_down = (GetAsyncKeyState(int(Keys.Space)) & 0x8000) != 0
                if is_space_down and not key_states[Keys.Space]:
                    rotate_action()
                key_states[Keys.Space] = is_space_down

                is_enter_down = (GetAsyncKeyState(int(Keys.Enter)) & 0x8000) != 0
                if is_enter_down and not key_states[Keys.Enter]:
                    place_action()
                    break
                key_states[Keys.Enter] = is_enter_down

                is_escape_down = (GetAsyncKeyState(int(Keys.Escape)) & 0x8000) != 0
                if is_escape_down and not key_states[Keys.Escape]:
                    cancel_action()
                    break
                key_states[Keys.Escape] = is_escape_down

                System.Threading.Thread.Sleep(50) # Small delay to reduce CPU usage

            # Handle outcome of interactive placement
            if placement_status["result"] != "place":
                t_group.RollBack()
                # print("Placement cancelled.") # REMOVED THIS LINE
                break # Exit loop, user cancelled

            # Final connection transaction
            with Transaction(doc, "Connect Parts") as t:
                t.Start()
                # Re-find the copied transfer element and its connector after all transformations
                copied_transfer_elem_final = None
                original_element_index = -1
                for i, original_elem in enumerate(source_elements):
                    if original_elem.Id == transfer_elem.Id:
                        original_element_index = i
                        break

                if original_element_index != -1 and original_element_index < len(copied_ids):
                    copied_transfer_elem_final = doc.GetElement(copied_ids[original_element_index])

                if not copied_transfer_elem_final:
                    raise Exception("Could not find the copied version of the transfer element for final connection.")

                new_mgr = get_connector_manager(copied_transfer_elem_final)

                # Find the connector on the copied element now closest to the target connector's origin
                final_conn = find_closest_connector(list(new_mgr.Connectors), target_conn.Origin)

                if final_conn and not final_conn.IsConnected and not target_conn.IsConnected:
                    # Attempt fabrication part specific connection if applicable, otherwise generic connect
                    if FabPartType and isinstance(final_conn.Owner, FabPartType):
                        flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static
                        connect_method = FabPartType.GetMethod("ConnectAndCouple", flags)
                        if connect_method:
                            connect_method.Invoke(None, System.Array[System.Object]([doc, final_conn, target_conn]))
                        else:
                            final_conn.ConnectTo(target_conn)
                    else:
                        final_conn.ConnectTo(target_conn)

                    t.Commit()
                    t_group.Assimilate() # Commit the entire group of transactions
                    uidoc.Selection.SetElementIds(List[ElementId](copied_ids))
                    # Removed: print("✔ Connected. Ready for next placement.")
                else:
                    t.RollBack() # Rollback connection if conditions are not met
                    raise Exception("Connectors not properly aligned, already used, or suitable connectors not found for final connection.")

        except Exception as ex:
            print("An error occurred during placement: {}".format(ex)) # This print is helpful for debugging actual errors, but can be removed too if desired
            if t_group and t_group.HasStarted():
                t_group.RollBack() # Ensure transaction group is rolled back on error
            break # Exit the placement loop on error

        finally:
            if form and not form.IsDisposed:
                form.Close() # Close the UI form

if __name__ == "__main__":
    if not doc or not uidoc:
        forms.alert("This script requires an open Revit document.", title="Error", exitscript=True)
    main()