# -*- coding: utf-8 -*-
"""
Check Duct Orientation - MULTIPLE SELECTION OPTIONS

Provides various selection methods:
- Pick individual elements
- Box Select
- All In View
- Visible in View
- Use Current Selection
"""

from __future__ import division
from pyrevit import revit, DB, forms, script
import clr

doc = revit.doc
uidoc = revit.uidoc
out = script.get_output()

# .NET List for selection sets
clr.AddReference("System")
from System.Collections.Generic import List

# ========================== DEBUG SETTING ==========================
# Set to True to see detailed debug output in a popup window
DEBUG_MODE = False
# ===================================================================

# Global debug storage
DEBUG_OUTPUT = []

# ---------------------------- HELPER FUNCTIONS ---------------------------

def feet_to_inches(ft):
    """Convert feet to inches"""
    try:
        return round(float(ft) * 12.0, 2)
    except:
        return None

def get_parameter(el, name):
    """Get parameter if it exists and has value"""
    try:
        p = el.LookupParameter(name)
        return p if (p and p.HasValue) else None
    except:
        return None

def get_param_val(el, name):
    """Get parameter value"""
    p = get_parameter(el, name)
    if not p:
        return None
    try:
        st = p.StorageType
        if st == DB.StorageType.Double:
            return p.AsDouble()
        elif st == DB.StorageType.Integer:
            return p.AsInteger()
        elif st == DB.StorageType.String:
            return p.AsString()
    except:
        pass
    return None

def get_fabrication_dimensions(e):
    """Get width and depth from fabrication part parameters (in inches)"""
    try:
        width_ft = get_param_val(e, "Main Primary Width")
        depth_ft = get_param_val(e, "Main Primary Depth")
        
        if width_ft is not None and depth_ft is not None:
            width_in = feet_to_inches(width_ft)
            depth_in = feet_to_inches(depth_ft)
            return width_in, depth_in
    except:
        pass
    
    return None, None

def get_bounding_box_dimensions(e, view):
    """Get the bounding box dimensions (in inches)"""
    try:
        bb = e.get_BoundingBox(view)
        if bb is None:
            bb = e.get_BoundingBox(doc.ActiveView)
        
        if bb is None:
            return None, None, None
        
        x_extent = abs(bb.Max.X - bb.Min.X)
        y_extent = abs(bb.Max.Y - bb.Min.Y)
        z_extent = abs(bb.Max.Z - bb.Min.Z)
        
        x_in = feet_to_inches(x_extent)
        y_in = feet_to_inches(y_extent)
        z_in = feet_to_inches(z_extent)
        
        return x_in, y_in, z_in
    except:
        return None, None, None

def get_duct_axis_direction(e):
    """Get the primary axis direction of the duct"""
    try:
        loc = e.Location
        if hasattr(loc, 'Curve'):
            curve = loc.Curve
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            
            dx = abs(p1.X - p0.X)
            dy = abs(p1.Y - p0.Y)
            dz = abs(p1.Z - p0.Z)
            
            max_delta = max(dx, dy, dz)
            
            if dx == max_delta:
                return 'X'
            elif dy == max_delta:
                return 'Y'
            else:
                return 'Z'
    except:
        pass
    
    return None

def debug_element(e):
    """Collect debug information about an element"""
    if not DEBUG_MODE:
        return
    
    elem_id = e.Id.IntegerValue
    
    debug_lines = []
    debug_lines.append("\n" + "="*60)
    debug_lines.append("DEBUG: Element ID {}".format(elem_id))
    debug_lines.append("="*60)
    
    # Get dimensions
    width_ft = get_param_val(e, "Main Primary Width")
    depth_ft = get_param_val(e, "Main Primary Depth")
    size_str = get_param_val(e, "Size")
    
    width_in = feet_to_inches(width_ft) if width_ft else None
    depth_in = feet_to_inches(depth_ft) if depth_ft else None
    
    debug_lines.append("\nLabeled Dimensions:")
    debug_lines.append("  Width: {} ft = {} in".format(width_ft, width_in if width_in else "?"))
    debug_lines.append("  Depth: {} ft = {} in".format(depth_ft, depth_in if depth_in else "?"))
    debug_lines.append("  Size: {}".format(size_str if size_str else "?"))
    
    if width_in and depth_in:
        if width_in > depth_in:
            debug_lines.append("  >> Width is LARGER ({} > {})".format(width_in, depth_in))
        elif depth_in > width_in:
            debug_lines.append("  >> Depth is LARGER ({} > {})".format(depth_in, width_in))
        else:
            debug_lines.append("  >> Width equals Depth ({} = {})".format(width_in, depth_in))
    
    # Bounding box
    bb_x, bb_y, bb_z = get_bounding_box_dimensions(e, doc.ActiveView)
    axis = get_duct_axis_direction(e)
    
    debug_lines.append("\nBounding Box:")
    debug_lines.append("  X extent: {:.2f} in".format(bb_x if bb_x else 0))
    debug_lines.append("  Y extent: {:.2f} in".format(bb_y if bb_y else 0))
    debug_lines.append("  Z extent: {:.2f} in".format(bb_z if bb_z else 0))
    debug_lines.append("  Axis direction: {}".format(axis if axis else "?"))
    
    # Cross-section
    if axis == 'X':
        horizontal = bb_y
        vertical = bb_z
        debug_lines.append("\nCross-section:")
        debug_lines.append("  Horizontal (Y) = {:.2f} in".format(horizontal))
        debug_lines.append("  Vertical (Z) = {:.2f} in".format(vertical))
        if horizontal > vertical:
            debug_lines.append("  >> Horizontal is LARGER ({:.2f} > {:.2f})".format(horizontal, vertical))
        elif vertical > horizontal:
            debug_lines.append("  >> Vertical is LARGER ({:.2f} > {:.2f})".format(vertical, horizontal))
    elif axis == 'Y':
        horizontal = bb_x
        vertical = bb_z
        debug_lines.append("\nCross-section:")
        debug_lines.append("  Horizontal (X) = {:.2f} in".format(horizontal))
        debug_lines.append("  Vertical (Z) = {:.2f} in".format(vertical))
        if horizontal > vertical:
            debug_lines.append("  >> Horizontal is LARGER ({:.2f} > {:.2f})".format(horizontal, vertical))
        elif vertical > horizontal:
            debug_lines.append("  >> Vertical is LARGER ({:.2f} > {:.2f})".format(vertical, horizontal))
    else:
        debug_lines.append("\nCross-section: Cannot determine (vertical duct?)")
    
    # Store in global list
    DEBUG_OUTPUT.extend(debug_lines)


def check_orientation_mismatch(e):
    """Check if duct has mismatched Width x Depth using ratio-based comparison"""
    width_labeled, depth_labeled = get_fabrication_dimensions(e)
    
    if width_labeled is None or depth_labeled is None:
        return False, {"reason": "Could not read labeled dimensions"}
    
    bb_x, bb_y, bb_z = get_bounding_box_dimensions(e, doc.ActiveView)
    
    if bb_x is None or bb_y is None or bb_z is None:
        return False, {"reason": "Could not read bounding box"}
    
    axis_dir = get_duct_axis_direction(e)
    
    if axis_dir is None:
        return False, {"reason": "Could not determine duct axis"}
    
    if axis_dir in ['X', 'Y']:
        if axis_dir == 'X':
            horizontal_extent = bb_y
        else:
            horizontal_extent = bb_x
        
        vertical_extent = bb_z
        
        width_is_larger = width_labeled > depth_labeled
        horizontal_is_larger = horizontal_extent > vertical_extent
        
        details = {
            "width_labeled": width_labeled,
            "depth_labeled": depth_labeled,
            "bb_horizontal": horizontal_extent,
            "bb_vertical": vertical_extent,
            "axis_direction": axis_dir
        }
        
        if width_is_larger == horizontal_is_larger:
            details["status"] = "Correct orientation"
            return False, details
        else:
            details["issue"] = "Width and Depth are swapped (rotated 90 degrees)"
            return True, details
    else:
        return False, {"reason": "Vertical duct - not yet supported"}

def get_fabrication_parts_from_elements(elements):
    """Filter elements to only fabrication parts"""
    FabricationPart = None
    try:
        from Autodesk.Revit.DB.Mechanical import FabricationPart as _Fab
        FabricationPart = _Fab
    except:
        try:
            from Autodesk.Revit.DB.Fabrication import FabricationPart as _Fab
            FabricationPart = _Fab
        except:
            FabricationPart = None
    
    if FabricationPart is not None:
        return [e for e in elements if isinstance(e, FabricationPart)]
    else:
        return [e for e in elements if e.Category and 
                e.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_FabricationDuctwork)]

def prompt_for_selection_method():
    """Show selection method options"""
    from Autodesk.Revit.UI.Selection import ObjectType
    
    # Define selection options
    options = {
        'Pick': 'Click individual elements',
        'Box Select': 'Drag a selection box',
        'All In View': 'All fabrication duct in current view',
        'Visible in View': 'Only visible fabrication duct (respects filters)',
        'Use Current Selection': 'Use already-selected elements'
    }
    
    # Show selection method dialog
    selected_method = forms.CommandSwitchWindow.show(
        options.keys(),
        message='Select fabrication duct to check:',
        recognize_access_key=False
    )
    
    if not selected_method:
        return None
    
    parts = []
    
    try:
        if selected_method == 'Pick':
            # Pick individual elements
            selected_refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select fabrication duct (ESC when done)"
            )
            elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]
            parts = get_fabrication_parts_from_elements(elements)
        
        elif selected_method == 'Box Select':
            # Box selection
            selected_refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Drag box to select fabrication duct (ESC when done)"
            )
            elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]
            parts = get_fabrication_parts_from_elements(elements)
        
        elif selected_method == 'All In View':
            # All in current view
            collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)\
                .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)\
                .WhereElementIsNotElementType()
            parts = list(collector)
        
        elif selected_method == 'Visible in View':
            # Only visible (respects view filters)
            collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)\
                .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork)\
                .WhereElementIsNotElementType()
            
            # Filter out hidden elements
            parts = []
            for elem in collector:
                if not elem.IsHidden(doc.ActiveView):
                    parts.append(elem)
        
        elif selected_method == 'Use Current Selection':
            # Use current selection
            selection_ids = uidoc.Selection.GetElementIds()
            if not selection_ids:
                forms.alert(
                    "No elements currently selected.\n\nPlease select elements first.",
                    title="No Selection"
                )
                return None
            
            elements = [doc.GetElement(eid) for eid in selection_ids]
            parts = get_fabrication_parts_from_elements(elements)
    
    except Exception as e:
        # User cancelled
        return None
    
    return parts

def analyze_parts(parts):
    """Analyze parts and return results"""
    results = {
        "total": len(parts),
        "rectangular": 0,
        "mismatched": 0,
        "correct": 0,
        "skipped": 0,
        "elements": []
    }
    
    for e in parts:
        elem_id = e.Id.IntegerValue
        
        # Collect debug info
        debug_element(e)
        
        # Check orientation
        is_mismatched, details = check_orientation_mismatch(e)
        
        if "reason" in details:
            results["skipped"] += 1
            results["elements"].append({
                "id": elem_id,
                "status": "SKIPPED",
                "details": details
            })
            if DEBUG_MODE:
                DEBUG_OUTPUT.append("\nResult: SKIPPED - {}".format(details.get("reason")))
        else:
            results["rectangular"] += 1
            
            if is_mismatched:
                results["mismatched"] += 1
                results["elements"].append({
                    "id": elem_id,
                    "status": "MISMATCHED",
                    "details": details
                })
                if DEBUG_MODE:
                    DEBUG_OUTPUT.append("\nResult: *** MISMATCHED ***")
                    DEBUG_OUTPUT.append("  {}".format(details.get("issue")))
                    DEBUG_OUTPUT.append("  Should be: {}x{} in (currently labeled {}x{} in)".format(
                        details.get("depth_labeled"),
                        details.get("width_labeled"),
                        details.get("width_labeled"),
                        details.get("depth_labeled")
                    ))
            else:
                results["correct"] += 1
                results["elements"].append({
                    "id": elem_id,
                    "status": "CORRECT",
                    "details": details
                })
                if DEBUG_MODE:
                    DEBUG_OUTPUT.append("\nResult: CORRECT - Orientation matches")
    
    return results

def select_mismatched_elements(results):
    """Select elements with mismatched orientations"""
    ids = List[DB.ElementId]()
    for elem in results["elements"]:
        if elem["status"] == "MISMATCHED":
            ids.Add(DB.ElementId(int(elem["id"])))
    
    try:
        uidoc.Selection.SetElementIds(ids)
        if ids.Count > 0:
            uidoc.ShowElements(ids)
    except:
        pass

def print_results(results):
    """Print formatted results to output window (only if debugging)"""
    if not DEBUG_MODE:
        return  # Don't show output window unless debugging
    
    out.print_md("## Check Duct Orientation Results")
    out.print_md("**Method:** Ratio-based (works with any connector size)")
    out.print_md("---")
    out.print_md("**Total analyzed:** {}".format(results["total"]))
    out.print_md("**Rectangular ducts:** {}".format(results["rectangular"]))
    out.print_md("**✅ Correct orientation:** {}".format(results["correct"]))
    out.print_md("**❌ Mismatched orientation:** {}".format(results["mismatched"]))
    out.print_md("**⚠️ Skipped:** {}".format(results["skipped"]))
    out.print_md("---")
    
    if results["mismatched"] > 0:
        out.print_md("### ❌ Mismatched Elements (Rotated 90°):")
        for elem in results["elements"]:
            if elem["status"] == "MISMATCHED":
                d = elem["details"]
                out.print_md(
                    "- **ID {}**: Labeled {}×{} in → Should be {}×{} in".format(
                        elem["id"],
                        d.get("width_labeled", "?"),
                        d.get("depth_labeled", "?"),
                        d.get("depth_labeled", "?"),
                        d.get("width_labeled", "?")
                    )
                )
    
    if results["correct"] > 0:
        out.print_md("### ✅ Correct Orientation:")
        count = 0
        for elem in results["elements"]:
            if elem["status"] == "CORRECT":
                d = elem["details"]
                out.print_md(
                    "- **ID {}**: {}×{} in - Correct".format(
                        elem["id"],
                        d.get("width_labeled", "?"),
                        d.get("depth_labeled", "?")
                    )
                )
                count += 1
                if count >= 10:
                    remaining = results["correct"] - 10
                    if remaining > 0:
                        out.print_md("- ... and {} more".format(remaining))
                    break

# ---------------------------- MAIN -----------------------------------

def main():
    """Main execution"""
    # Clear debug output at start
    global DEBUG_OUTPUT
    DEBUG_OUTPUT = []
    
    if DEBUG_MODE:
        out.print_md("# Check Duct Orientation")
        out.print_md("---")
    
    # Get selection
    parts = prompt_for_selection_method()
    
    if parts is None:
        return  # User cancelled - just exit silently
    
    if not parts:
        forms.alert(
            "No fabrication duct found.\n\n"
            "Please select fabrication duct elements.",
            title="No Fabrication Duct"
        )
        return
    
    if DEBUG_MODE:
        out.print_md("Analyzing {} elements...".format(len(parts)))
        out.print_md("")
    
    # Analyze
    results = analyze_parts(parts)
    print_results(results)
    
    # Select mismatched
    if results["mismatched"] > 0:
        select_mismatched_elements(results)
    
    # Show debug window if enabled
    if DEBUG_MODE and DEBUG_OUTPUT:
        debug_text = "\n".join(DEBUG_OUTPUT)
        forms.alert(
            debug_text,
            title="Debug Information ({} elements)".format(results["total"]),
            warn_icon=False
        )
    
    # Show summary dialog
    msg_lines = [
        "Check Duct Orientation Complete",
        "",
        "Analyzed: {}".format(results["total"]),
        "Rectangular: {}".format(results["rectangular"]),
        "Correct: {}".format(results["correct"]),
        "MISMATCHED: {}".format(results["mismatched"]),
        "Skipped: {}".format(results["skipped"]),
        ""
    ]
    
    if results["mismatched"] > 0:
        msg_lines.append("❌ {} duct(s) rotated 90° - NOW SELECTED".format(results["mismatched"]))
    elif results["rectangular"] > 0:
        msg_lines.append("✅ All rectangular ducts have correct orientation!")
    
    forms.alert("\n".join(msg_lines), title="Check Duct Orientation")

if __name__ == "__main__":
    main()