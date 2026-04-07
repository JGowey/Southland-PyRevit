# -*- coding: utf-8 -*-
"""
Renumber Tool for pyRevit
Author: Jeremiah Griffith
Version: 2.2.5

Intelligently renumber elements with multiple ordering modes and comparison options.
Includes grouping by Category/Family/Type with custom ordering.
Includes post-pick filtering.
Now supports Assembly Name and Assembly Mark.
"""

from pyrevit import revit, DB, forms, script
import math
import os
import System  # Import the System namespace
from System.Collections.ObjectModel import ObservableCollection
from System.Collections.Generic import List

import clr
clr.AddReference("System.Xml")
from System.Xml import XmlDocument

# =============================================================================
# GLOBALS
# =============================================================================

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# Grouping modes
GROUPING_MODES = [
    "All at Once",
    "By Category",
    "By Family",
    "By Type"
]

ORDER_MODES = [
    "Largest -> Smallest",
    "Smallest -> Largest",
    "Clockwise",
    "Counterclockwise",
    "Follow spline"
]

DEFAULT_COMPARE_FIELDS = [
    "Alias", "Angle", "Category", "CutType", "Cid", "Connectors",
    "Dimensions", "FamilyInfo", "InsulationSpecification", "InsulationType",
    "Length", "LiningThickness", "Material", "Service", "Size",
    "Specification", "Status", "FabricationNotes"
]

# Special parameters not directly on elements
# "Assembly Name" is on the element, so it is NO LONGER needed here.
ASSEMBLY_MARK_PARAM = "Assembly Mark" # This one IS on the instance.


CONFIG_SECTION = "RenumberTool"
WINDOW_INSTANCE = None


# =============================================================================
# DATA CLASSES
# =============================================================================

class CompareField(object):
    """Field for comparing identical parts"""
    def __init__(self, name, is_included=False):
        self.Name = name
        self.Include = is_included

# NEW Data class for the Group Order ListBox
class GroupItem(object):
    """Item for the group order list"""
    def __init__(self, name):
        self.Name = name

# Data class for the Deselect Elements ListBox
class ElementItem(object):
    """Item for the deselect elements list"""
    def __init__(self, elem_id, category, name):
        self.Id = elem_id
        self.Category = category
        self.Name = name

# Data class for grouped deselect items (Category/Family/Type)
class DeselectItem(object):
    """Item for deselect list showing name and count"""
    def __init__(self, name, count, element_ids=None):
        self.Name = name
        self.Count = "({} elements)".format(count) if count > 1 else "(1 element)"
        self.ElementIds = element_ids or []


class Settings(object):
    """Application settings"""
    def __init__(self):
        self.prefix = ""
        self.suffix = ""
        self.start_number = "1"
        self.order_mode = ORDER_MODES[0]
        self.grouping_mode = GROUPING_MODES[0]
        self.specify_order = False
        self.place_tags = False
        self.prompt_for_positions = False
        self.same_for_identical = False
        self.target_a = ""
        self.target_b = ""
        self.target_c = ""
        self.compare_param_1 = ""
        self.compare_param_2 = ""
        self.compare_param_3 = ""
        self.compare_param_4 = ""
        self.picked_curve_id = None
        self.saved_group_order = "" # NEW
        
        self.compare_fields = ObservableCollection[CompareField]()
        for name in DEFAULT_COMPARE_FIELDS:
            self.compare_fields.Add(CompareField(name))
        
        self.order_modes = ObservableCollection[str]()
        for mode in ORDER_MODES:
            self.order_modes.Add(mode)

        self.grouping_modes = ObservableCollection[str]()
        for mode in GROUPING_MODES:
            self.grouping_modes.Add(mode)
        
        self.group_order_items = ObservableCollection[GroupItem]() # NEW
        
        self.available_params = ObservableCollection[str]()


# =============================================================================
# CONFIGURATION MANAGEMENT
# =============================================================================

class ConfigManager(object):
    """Handles loading and saving settings"""
    
    @staticmethod
    def load():
        """Load settings from pyRevit config"""
        settings = Settings()
        cfg = script.get_config(CONFIG_SECTION)
        
        settings.prefix = ConfigManager._safe_get(cfg, "Prefix", "")
        settings.suffix = ConfigManager._safe_get(cfg, "Suffix", "")
        settings.start_number = ConfigManager._safe_get(cfg, "StartNumber", "1")
        settings.order_mode = ConfigManager._safe_get(cfg, "OrderMode", ORDER_MODES[0])
        settings.grouping_mode = ConfigManager._safe_get(cfg, "GroupingMode", GROUPING_MODES[0])
        settings.place_tags = bool(ConfigManager._safe_get(cfg, "PlaceTags", False))
        settings.prompt_for_positions = bool(ConfigManager._safe_get(cfg, "PromptPos", False))
        settings.same_for_identical = bool(ConfigManager._safe_get(cfg, "SameIdent", False))
        settings.specify_order = bool(ConfigManager._safe_get(cfg, "SpecifyOrder", False))
        settings.target_a = ConfigManager._safe_get(cfg, "TargetA", "")
        settings.target_b = ConfigManager._safe_get(cfg, "TargetB", "")
        settings.target_c = ConfigManager._safe_get(cfg, "TargetC", "")
        settings.compare_param_1 = ConfigManager._safe_get(cfg, "CompareParam1", "")
        settings.compare_param_2 = ConfigManager._safe_get(cfg, "CompareParam2", "")
        settings.compare_param_3 = ConfigManager._safe_get(cfg, "CompareParam3", "")
        settings.compare_param_4 = ConfigManager._safe_get(cfg, "CompareParam4", "")
        settings.saved_group_order = ConfigManager._safe_get(cfg, "SavedGroupOrder", "") # NEW
        
        compare_str = ConfigManager._safe_get(cfg, "CompareFields", "")
        selected_fields = set(f.strip() for f in compare_str.split(",") if f.strip())
        
        all_field_names = set(DEFAULT_COMPARE_FIELDS)
        all_field_names.update(selected_fields) 

        settings.compare_fields.Clear()
        
        for name in DEFAULT_COMPARE_FIELDS:
            settings.compare_fields.Add(CompareField(name, name in selected_fields))
        
        for name in sorted(list(all_field_names)):
            if name not in DEFAULT_COMPARE_FIELDS:
                 settings.compare_fields.Add(CompareField(name, name in selected_fields))

        return settings
    
    @staticmethod
    def save(settings):
        """Save settings to pyRevit config"""
        try:
            cfg = script.get_config(CONFIG_SECTION)
            cfg.Prefix = settings.prefix
            cfg.Suffix = settings.suffix
            cfg.StartNumber = settings.start_number
            cfg.OrderMode = settings.order_mode
            cfg.GroupingMode = settings.grouping_mode
            cfg.PlaceTags = settings.place_tags
            cfg.PromptPos = settings.prompt_for_positions
            cfg.SameIdent = settings.same_for_identical
            cfg.SpecifyOrder = settings.specify_order
            cfg.TargetA = settings.target_a
            cfg.TargetB = settings.target_b
            cfg.TargetC = settings.target_c
            cfg.CompareParam1 = settings.compare_param_1
            cfg.CompareParam2 = settings.compare_param_2
            cfg.CompareParam3 = settings.compare_param_3
            cfg.CompareParam4 = settings.compare_param_4
            cfg.SavedGroupOrder = settings.saved_group_order # NEW
            cfg.CompareFields = ",".join([f.Name for f in settings.compare_fields if f.Include])
            script.save_config()
        except:
            pass
    
    @staticmethod
    def _safe_get(cfg, key, default=None):
        """Safely get config value"""
        try:
            return getattr(cfg, key, default) if hasattr(cfg, key) else default
        except:
            return default


# =============================================================================
# PROFILE IMPORT/EXPORT
# =============================================================================

class ProfileManager(object):
    """Handles XML profile import/export"""
    
    @staticmethod
    def export(settings):
        """Export settings to XML"""
        xml = XmlDocument()
        root = xml.CreateElement("RenumberProfile")
        xml.AppendChild(root)
        
        elements = {
            "Prefix": settings.prefix,
            "Suffix": settings.suffix,
            "StartNumber": settings.start_number,
            "OrderMode": settings.order_mode,
            "GroupingMode": settings.grouping_mode,
            "PlaceTags": str(settings.place_tags),
            "PromptForTextPos": str(settings.prompt_for_positions),
            "SameForIdentical": str(settings.same_for_identical),
            "SpecifyOrder": str(settings.specify_order),
            "TargetA": settings.target_a,
            "TargetB": settings.target_b,
            "TargetC": settings.target_c,
            "CompareParam1": settings.compare_param_1,
            "CompareParam2": settings.compare_param_2,
            "CompareParam3": settings.compare_param_3,
            "CompareParam4": settings.compare_param_4,
            "SavedGroupOrder": settings.saved_group_order # NEW
        }
        
        for name, value in elements.items():
            el = xml.CreateElement(name)
            el.InnerText = str(value) if value is not None else ""
            root.AppendChild(el)
        
        fields_node = xml.CreateElement("CompareFields")
        root.AppendChild(fields_node)
        for field in settings.compare_fields:
            if field.Include:
                field_node = xml.CreateElement("Field")
                field_node.SetAttribute("name", field.Name)
                fields_node.AppendChild(field_node)
        
        return xml
    
    @staticmethod
    def import_to(xml_doc, settings):
        """Import settings from XML"""
        root = xml_doc.DocumentElement
        
        def get_text(node_name):
            node = root.SelectSingleNode(node_name)
            return node.InnerText if node is not None else ""
        
        settings.prefix = get_text("Prefix")
        settings.suffix = get_text("Suffix")
        settings.start_number = get_text("StartNumber") or "1"
        
        order = get_text("OrderMode")
        settings.order_mode = order if order in ORDER_MODES else ORDER_MODES[0]

        group_mode = get_text("GroupingMode")
        settings.grouping_mode = group_mode if group_mode in GROUPING_MODES else GROUPING_MODES[0]
        
        settings.place_tags = get_text("PlaceTags") == "True"
        settings.prompt_for_positions = get_text("PromptForTextPos") == "True"
        settings.same_for_identical = get_text("SameForIdentical") == "True"
        settings.specify_order = get_text("SpecifyOrder") == "True"
        settings.target_a = get_text("TargetA")
        settings.target_b = get_text("TargetB")
        settings.target_c = get_text("TargetC")
        settings.compare_param_1 = get_text("CompareParam1")
        settings.compare_param_2 = get_text("CompareParam2")
        settings.compare_param_3 = get_text("CompareParam3")
        settings.compare_param_4 = get_text("CompareParam4")
        settings.saved_group_order = get_text("SavedGroupOrder") # NEW
        
        selected = set()
        fields_node = root.SelectSingleNode("CompareFields")
        if fields_node is not None:
            for node in fields_node.ChildNodes:
                try:
                    name = node.GetAttribute("name")
                    if name:
                        selected.add(name)
                except:
                    pass
        
        all_field_names = set(DEFAULT_COMPARE_FIELDS)
        all_field_names.update(selected) 

        settings.compare_fields.Clear()
        
        for name in DEFAULT_COMPARE_FIELDS:
            settings.compare_fields.Add(CompareField(name, name in selected))
        
        for name in sorted(list(all_field_names)):
            if name not in DEFAULT_COMPARE_FIELDS:
                 settings.compare_fields.Add(CompareField(name, name in selected))


# =============================================================================
# PARAMETER UTILITIES
# =============================================================================

class ParameterHelper(object):
    """Helper for parameter operations"""
    
    @staticmethod
    def get_writable_parameters(element):
        """Get writable parameter names"""
        params = set()
        try:
            for p in element.Parameters:
                try:
                    if p and not p.IsReadOnly and p.Definition:
                        name = p.Definition.Name
                        if name and name.strip():
                            params.add(name)
                except:
                    pass
        except:
            pass
        return params
    
    @staticmethod
    def get_group_key(elem, group_mode):
        """Get the group key (Category, Family, or Type name) for an element"""
        key = "Other"
        try:
            if group_mode == "By Category":
                key = elem.Category.Name if elem.Category else "Other"
            elif group_mode == "By Family":
                type_id = elem.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type and hasattr(elem_type, 'FamilyName'):
                            key = elem_type.FamilyName
                    elif elem.Category: # Fallback for system families
                            key = elem.Category.Name 
                    else:
                            key = "Other"
                elif elem.Category: # Fallback for elements with no type
                    key = elem.Category.Name
                else:
                    key = "Other"
            elif group_mode == "By Type":
                type_id = elem.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    key = doc.GetElement(type_id).Name
                else:
                    key = "Other"
        except:
            key = "Other"
        return key

    @staticmethod
    def get_common_parameters(elements):
        """Get parameters common to all elements"""
        if not elements:
            return []
        
        common = None
        all_in_assembly = True
        
        for elem in elements:
            # Use Parameters property for *all* parameters (read-only included)
            # for filtering purposes.
            params = set()
            try:
                for p in elem.Parameters:
                    if p and p.Definition:
                        name = p.Definition.Name
                        if name and name.strip():
                            params.add(name)
            except:
                pass

            # NEW: Check for assembly membership
            try:
                if elem.AssemblyInstanceId == DB.ElementId.InvalidElementId:
                    all_in_assembly = False
            except:
                all_in_assembly = False
                
            common = params if common is None else common.intersection(params)
            if not common and not all_in_assembly:
                break
        
        common_list = list(common or [])
        
        # NEW: If all elements are in an assembly, add Assembly parameters
        if all_in_assembly:
            # "Assembly Name" will be found by the standard lookup above.
            # "Assembly Mark" is on the instance, so we add it manually.
            common_list.append(ASSEMBLY_MARK_PARAM)
            
        return sorted(list(set(common_list)), key=lambda s: s.lower())
    
    @staticmethod
    def set_value(element, param_name, value_text):
        """Set parameter value from text"""
        if not param_name:
            return True
        
        # *** BUG FIX HERE ***
        # Handle setting assembly parameters
        # ONLY "Assembly Mark" is settable this way. "Assembly Name" is read-only
        # on the element and settable on the Type, which we won't do here.
        if param_name == ASSEMBLY_MARK_PARAM:
            try:
                if element.AssemblyInstanceId != DB.ElementId.InvalidElementId:
                    assembly_instance = doc.GetElement(element.AssemblyInstanceId)
                    if not assembly_instance:
                        return False
                    
                    # Set the Assembly INSTANCE parameter
                    param = assembly_instance.get_Parameter(DB.BuiltInParameter.ASSEMBLY_MARK)
                    if param and not param.IsReadOnly:
                        param.Set(value_text)
                        return True
                            
                return False # Failed to set assembly param
            except:
                return False # Error setting assembly param
        # *** END BUG FIX ***

        try:
            param = element.LookupParameter(param_name)
            # This will now correctly find "Assembly Name", see it's ReadOnly,
            # and return False, which is the correct behavior.
            if not param or param.IsReadOnly:
                return False
            
            storage = param.StorageType
            
            if storage == DB.StorageType.String:
                param.Set(value_text)
                return True
            elif storage == DB.StorageType.Integer:
                digits = "".join(c for c in value_text if c.isdigit())
                param.Set(int(digits) if digits else 0)
                return True
            else:
                try:
                    param.Set(value_text)
                    return True
                except:
                    try:
                        param.SetValueString(value_text)
                        return True
                    except:
                        return False
        except:
            return False
    
    @staticmethod
    def get_value(element, param_name):
        """Get parameter value as string"""
        if not param_name:
            return ""
        
        if param_name.lower() == "category":
            try:
                return element.Category.Name if element.Category else ""
            except:
                return ""
        
        # *** BUG FIX HERE ***
        # Special handling ONLY for params NOT on the element itself
        if param_name == ASSEMBLY_MARK_PARAM:
            try:
                if element.AssemblyInstanceId != DB.ElementId.InvalidElementId:
                    assembly_instance = doc.GetElement(element.AssemblyInstanceId)
                    if assembly_instance:
                        # Get the Assembly MARK from the INSTANCE
                        param = assembly_instance.get_Parameter(DB.BuiltInParameter.ASSEMBLY_MARK)
                        if param:
                            return param.AsString() or ""
                return "" # Not in assembly or param not found
            except:
                return "" # Error
        # *** END BUG FIX ***

        # Standard logic block (will now correctly find "Assembly Name")
        try:
            param = element.LookupParameter(param_name)
            if not param:
                return ""
            
            storage = param.StorageType
            if storage == DB.StorageType.String:
                return param.AsString() or ""
            elif storage == DB.StorageType.Integer:
                return str(param.AsInteger())
            elif storage == DB.StorageType.Double:
                return param.AsValueString() or str(param.AsDouble())
            elif storage == DB.StorageType.ElementId:
                eid = param.AsElementId()
                e_val = doc.GetElement(eid)
                return e_val.Name if e_val else ""
        except:
            pass
        
        return ""
    
    @staticmethod
    def get_element_signature(element, field_names):
        """Create signature for comparing elements"""
        sig_values = []
        has_non_empty = False
        
        for name in field_names:
            value = ParameterHelper.get_value(element, name)
            sig_values.append((name, value))
            if value and str(value).strip():
                has_non_empty = True
        
        # If all values are empty, include ElementId to prevent false matches
        if not has_non_empty:
            sig_values.append(("_ElementId_", str(element.Id.IntegerValue)))
        
        return tuple(sig_values)


# =============================================================================
# ELEMENT SELECTION
# =============================================================================

class ElementPicker(object):
    """Handles element picking operations"""
    
    @staticmethod
    def pick_multiple():
        """Pick multiple elements"""
        from Autodesk.Revit.UI.Selection import ObjectType
        
        picked = []
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select elements to renumber (Finish or ESC when done)"
            )
            for ref in refs:
                elem = doc.GetElement(ref.ElementId)
                if elem:
                    picked.append(elem)
        except:
            pass
        
        return picked
    
    @staticmethod
    def pick_in_order():
        """Pick elements one by one and keep them highlighted"""
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.Exceptions import OperationCanceledException
        from System.Collections.Generic import List
        
        sequence = []
        picked_ids = []
        
        # Clear selection
        try:
            empty_list = List[DB.ElementId]()
            uidoc.Selection.SetElementIds(empty_list)
        except:
            pass
        
        pick_number = 1
        while True:
            try:
                prompt = "Pick element #{} (Press ESC when finished)".format(pick_number)
                
                # Critical pick - only OperationCanceledException should break the loop
                ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
                
                # Get element
                try:
                    elem = doc.GetElement(ref.ElementId)
                    if not elem:
                        continue
                except:
                    continue
                
                # Check if already picked
                elem_id_int = elem.Id.IntegerValue
                already_picked = False
                for e in sequence:
                    if e.Id.IntegerValue == elem_id_int:
                        already_picked = True
                        break
                
                if already_picked:
                    continue
                
                # Add to sequence
                sequence.append(elem)
                picked_ids.append(elem.Id)
                pick_number += 1
                
                # Try to highlight - but don't break if this fails
                try:
                    id_collection = List[DB.ElementId]()
                    for eid in picked_ids:
                        id_collection.Add(eid)
                    uidoc.Selection.SetElementIds(id_collection)
                except:
                    pass
                
            except OperationCanceledException:
                break
            except:
                continue
        
        return sequence
    
    @staticmethod
    def pick_curve():
        """Pick a curve element"""
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.Exceptions import OperationCanceledException
        
        try:
            ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                "Pick a curve/spline for ordering"
            )
            elem = doc.GetElement(ref.ElementId)
            
            if isinstance(elem, DB.CurveElement):
                return elem
            else:
                forms.alert(
                    "Please select a CurveElement (model line, arc, spline, etc.)",
                    title="Renumber"
                )
                return None
        except OperationCanceledException:
            return None
        except:
            return None


# =============================================================================
# FABRICATION PART UTILITIES
# =============================================================================

class FabricationHelper(object):
    """Helper for fabrication parts"""
    
    @staticmethod
    def is_fabrication_part(element):
        """Check if element is fabrication part"""
        try:
            from Autodesk.Revit.DB.Mechanical import FabricationPart as MechFabPart
            if isinstance(element, MechFabPart):
                return True
        except:
            pass
        
        try:
            from Autodesk.Revit.DB.Fabrication import FabricationPart
            if isinstance(element, FabricationPart):
                return True
        except:
            pass
        
        try:
            type_name = element.GetType().FullName
            if type_name and "FabricVlPart" in type_name:
                return True
        except:
            pass
        
        return False
    
    @staticmethod
    def get_size_metric(element):
        """Get size metric for fabrication parts"""
        if not FabricationHelper.is_fabrication_part(element):
            return None, ""
        
        def safe_float(value):
            try:
                return float(value)
            except:
                return None
        
        # Try direct area properties
        for attr in ("Area", "CrossSectionalArea", "SectionArea", "CrossSectionArea"):
            try:
                val = getattr(element, attr, None)
                area = safe_float(val)
                if area is not None and area > 0:
                    return area, "Profile Area"
            except:
                pass
        
        # Try area parameter
        try:
            area_param = element.LookupParameter("Area")
            if area_param:
                area = area_param.AsDouble()
                if area and area > 0:
                    return area, "Profile Area"
        except:
            pass
        
        # Calculate from diameter
        diameter = None
        for attr in ("Diameter", "NominalDiameter"):
            try:
                val = getattr(element, attr, None)
                diameter = safe_float(val)
                if diameter is not None and diameter > 0:
                    break
            except:
                pass
        
        if diameter is None:
            try:
                size_param = element.LookupParameter("Size")
                if size_param:
                    size_str = size_param.AsString()
                    if size_str:
                        import re
                        numbers = re.findall(r'\d+\.?\d*', size_str)
                        if numbers:
                            diameter = float(numbers[0])
            except:
                pass
        
        if diameter and diameter > 0:
            area = math.pi * (diameter / 2.0) ** 2
            return area, "Profile Area (πd²/4)"
        
        # Calculate from width and height
        width = None
        height = None
        
        for attr in ("Width", "DimA"):
            try:
                val = getattr(element, attr, None)
                width = safe_float(val)
                if width is not None and width > 0:
                    break
            except:
                pass
        
        for attr in ("Height", "DimB"):
            try:
                val = getattr(element, attr, None)
                height = safe_float(val)
                if height is not None and height > 0:
                    break
            except:
                pass
        
        if width and height and width > 0 and height > 0:
            return width * height, "Profile Area (w×h)"
        
        if width and width > 0:
            return width * width, "Profile Area (square)"
        
        return None, ""


# =============================================================================
# GEOMETRY & ORDERING
# =============================================================================

class GeometryHelper(object):
    """Helper for geometry operations"""
    
    @staticmethod
    def get_bounding_box_info(element, view=None):
        """Get bounding box volume, diagonal, and center"""
        bbox = None
        try:
            bbox = element.get_BoundingBox(view)
        except:
            pass
        
        if not bbox:
            try:
                bbox = element.get_BoundingBox(None)
            except:
                pass
        
        if not bbox:
            return 0.0, 0.0, DB.XYZ(0, 0, 0)
        
        dx = max(0, bbox.Max.X - bbox.Min.X)
        dy = max(0, bbox.Max.Y - bbox.Min.Y)
        dz = max(0, bbox.Max.Z - bbox.Min.Z)
        
        volume = dx * dy * dz
        diagonal = math.sqrt(dx*dx + dy*dy + dz*dz)
        center = DB.XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0
        )
        
        return volume, diagonal, center
    
    @staticmethod
    def get_anchor_point(element, view):
        """Get representative point for element"""
        try:
            location = element.Location
            if location:
                if hasattr(location, "Point"):
                    pt = location.Point
                    if isinstance(pt, DB.XYZ):
                        return pt
                
                if hasattr(location, "Curve"):
                    curve = location.Curve
                    if curve:
                        try:
                            return curve.Evaluate(0.5, True)
                        except:
                            pass
        except:
            pass
        
        return GeometryHelper.get_bounding_box_info(element, view)[2]
    
    @staticmethod
    def project_to_view_plane(point, view):
        """Project 3D point to view 2D"""
        if not isinstance(view, DB.View):
            return point.X, point.Y
        
        try:
            origin = view.Origin
            right = view.RightDirection
            up = view.UpDirection
            
            vector = point - origin
            u = vector.DotProduct(right)
            v = vector.DotProduct(up)
            
            return u, v
        except:
            return point.X, point.Y


class ElementSorter(object):
    """Sorts elements by various criteria"""
    
    @staticmethod
    def by_size(elements, view, reverse=False):
        """Order by size"""
        sorted_items = []
        
        for elem in elements:
            metric, desc = FabricationHelper.get_size_metric(elem)
            
            if metric is None:
                volume, diagonal, center = GeometryHelper.get_bounding_box_info(elem, view)
                metric = volume if volume > 0 else diagonal
                desc = "Volume" if volume > 0 else "Diagonal"
            
            sorted_items.append((metric, desc, elem))
        
        sorted_items.sort(key=lambda x: x[0], reverse=reverse)
        return [elem for _, _, elem in sorted_items]
    
    @staticmethod
    def by_angle(elements, view, counterclockwise=False):
        """Order by angle around center"""
        if not isinstance(view, DB.View):
            view = revit.active_view
        
        points = []
        for elem in elements:
            pt3d = GeometryHelper.get_anchor_point(elem, view)
            u, v = GeometryHelper.project_to_view_plane(pt3d, view)
            points.append((u, v, elem))
        
        if not points:
            return elements
        
        cu = sum(p[0] for p in points) / len(points)
        cv = sum(p[1] for p in points) / len(points)
        
        with_angles = []
        for u, v, elem in points:
            angle = math.atan2(v - cv, u - cu)
            with_angles.append((angle, elem))
        
        with_angles.sort(key=lambda x: x[0], reverse=not counterclockwise)
        return [elem for _, elem in with_angles]
    
    @staticmethod
    def by_curve(elements, curve_element, view):
        """Order by proximity along curve"""
        try:
            curve = curve_element.GeometryCurve
        except:
            curve = None
        
        if curve is None:
            return elements
        
        sorted_items = []
        for elem in elements:
            point = GeometryHelper.get_anchor_point(elem, view)
            
            try:
                result = curve.Project(point)
                parameter = result.Parameter if result else 0.0
            except:
                parameter = 0.0
            
            sorted_items.append((parameter, elem))
        
        sorted_items.sort(key=lambda x: x[0])
        return [elem for _, elem in sorted_items]


# =============================================================================
# TEXT NOTE CREATION
# =============================================================================

class TextNoteCreator(object):
    """Creates text notes"""
    
    @staticmethod
    def create(view, text, location):
        """Create text note at location"""
        try:
            text_type = DB.FilteredElementCollector(doc)\
                .OfClass(DB.TextNoteType)\
                .FirstElement()
            
            if not text_type:
                return None
            
            return DB.TextNote.Create(doc, view.Id, location, text, text_type.Id)
        except:
            return None


# =============================================================================
# RENUMBERING ENGINE
# =============================================================================

class RenumberEngine(object):
    """Main renumbering logic"""
    
    @staticmethod
    def execute(elements, settings, picked_curve=None, manual_order=None, custom_group_order=None): # MODIFIED
        """Execute renumbering"""
        targets = [t for t in [settings.target_a, settings.target_b, settings.target_c] if t]
        if not targets:
            forms.alert("Please select at least one target parameter", title="Renumber")
            return 0, 0, []
        
        # Use the filtered 'elements' list passed from the window
        if not elements:
            forms.alert("No elements selected (or filter resulted in 0 elements)", title="Renumber")
            return 0, 0, []
        
        try:
            start_num = int(settings.start_number)
        except:
            start_num = 1
        
        
        if settings.specify_order:
            # If manual order, grouping is ignored. Use the manual_order list.
            if not manual_order or len(manual_order) == 0:
                forms.alert(
                    "Manual pick order is enabled but no order was picked.\n\n"
                    "Please use the 'Pick Order' button to select elements in order.",
                    title="Renumber"
                )
                return 0, 0, []
            # Make sure the manual order only contains elements from the filtered list
            filtered_manual_order_ids = set([e.Id for e in elements])
            final_ordered_list = [e for e in manual_order if e.Id in filtered_manual_order_ids]

        else:
            # NEW: Group elements first
            grouped_lists = RenumberEngine._get_grouped_elements(
                elements, # Use the filtered list
                settings.grouping_mode,
                custom_group_order or [] # MODIFIED
            )
            
            final_ordered_list = []
            for group in grouped_lists:
                # Then apply spatial ordering within each group
                ordered_group = RenumberEngine._get_ordered_elements(
                    group, settings, picked_curve
                )
                final_ordered_list.extend(ordered_group)

        if not final_ordered_list:
            return 0, 0, []
        
        compare_fields = []
        if settings.same_for_identical:
            # Get static compare fields
            compare_fields = [f.Name for f in settings.compare_fields if f.Include]
            
            # Add dynamic compare parameters
            for param in [settings.compare_param_1,
                          settings.compare_param_2,
                          settings.compare_param_3,
                          settings.compare_param_4]:
                if param and param.strip():
                    compare_fields.append(param)

            if not compare_fields:
                forms.toast("No compare fields selected", title="Renumber", appid="renumber")
                settings.same_for_identical = False
        
        return RenumberEngine._apply_numbering(
            final_ordered_list, settings, targets, start_num, compare_fields
        )

    @staticmethod
    def _get_grouped_elements(elements, group_mode, custom_order): # MODIFIED
        """Group elements by the selected mode"""
        if group_mode == "All at Once":
            return [list(elements)]
        
        groups = {}
        for elem in elements:
            key = ParameterHelper.get_group_key(elem, group_mode) # MODIFIED
            
            if key not in groups:
                groups[key] = []
            groups[key].append(elem)
            
        # MODIFIED: Sort groups based on custom_order, then alphabetically
        final_list_of_lists = []
        processed_keys = set()

        # Add groups based on the custom order first
        for key in custom_order:
            if key in groups:
                final_list_of_lists.append(groups[key])
                processed_keys.add(key)
        
        # Add any remaining groups (newly selected, etc.) alphabetically
        for key in sorted(groups.keys()):
            if key not in processed_keys:
                final_list_of_lists.append(groups[key])
        
        return final_list_of_lists

    @staticmethod
    def _get_ordered_elements(elements, settings, picked_curve):
        """Get elements in desired order"""
        # This function is now only called for non-manual-order
        
        ordered = list(elements)
        view = revit.active_view
        mode = settings.order_mode
        
        if mode == "Largest -> Smallest":
            return ElementSorter.by_size(ordered, view, reverse=True)
        elif mode == "Smallest -> Largest":
            return ElementSorter.by_size(ordered, view, reverse=False)
        elif mode == "Clockwise":
            return ElementSorter.by_angle(ordered, view, counterclockwise=False)
        elif mode == "Counterclockwise":
            return ElementSorter.by_angle(ordered, view, counterclockwise=True)
        elif mode == "Follow spline":
            if not picked_curve:
                forms.alert("Please pick a curve first", title="Renumber")
                return []
            return ElementSorter.by_curve(ordered, picked_curve, view)
        
        return ordered
    
    @staticmethod
    def _apply_numbering(ordered_elements, settings, targets, start_num, compare_fields):
        """Apply numbering to elements"""
        t = DB.Transaction(doc, "Renumber Elements")
        t.Start()
        
        try:
            current_num = start_num
            signature_map = {}
            success_count = 0
            skip_count = 0
            errors = []
            
            for elem in ordered_elements:
                try:
                    if settings.same_for_identical and compare_fields:
                        sig = ParameterHelper.get_element_signature(elem, compare_fields)
                        if sig in signature_map:
                            number_text = signature_map[sig]
                        else:
                            number_text = "{}{}{}".format(
                                settings.prefix, current_num, settings.suffix
                            )
                            signature_map[sig] = number_text
                            current_num += 1
                    else:
                        number_text = "{}{}{}".format(
                            settings.prefix, current_num, settings.suffix
                        )
                        current_num += 1
                    
                    wrote_any = False
                    for param_name in targets:
                        if ParameterHelper.set_value(elem, param_name, number_text):
                            wrote_any = True
                    
                    if wrote_any:
                        success_count += 1
                        
                        if settings.place_tags and isinstance(revit.active_view, DB.View):
                            try:
                                if settings.prompt_for_positions:
                                    pt = uidoc.Selection.PickPoint(
                                        "Pick location for '{}'".format(number_text)
                                    )
                                else:
                                    pt = GeometryHelper.get_anchor_point(elem, revit.active_view)
                                
                                TextNoteCreator.create(revit.active_view, number_text, pt)
                            except:
                                pass
                    else:
                        skip_count += 1
                        errors.append("Element {} - No writable parameters".format(elem.Id))
                    
                except Exception as ex:
                    skip_count += 1
                    errors.append("Element {} - {}".format(elem.Id, str(ex)))
            
            t.Commit()
            return success_count, skip_count, errors
            
        except Exception as ex:
            t.RollBack()
            raise ex


# =============================================================================
# UI WINDOW
# =============================================================================

class RenumberWindow(forms.WPFWindow):
    """Main application window"""
    
    def __init__(self, xaml_path, preloaded_elements=None):
        forms.WPFWindow.__init__(self, xaml_path)
        
        self.settings = ConfigManager.load()
        self.elements = preloaded_elements or []
        self.all_picked_elements = list(self.elements) # Store original selection
        self.picked_curve = None
        self.manual_order = []
        self._executing_save = False
        self._executing_load = False
        self._results = None
        self._next_action = None
        self._initializing = True
        
        self._get_controls()
        self._wire_events()
        
        self._initialize_ui()
        
        self._initializing = False
        
        if self.elements:
            self.ok_button.IsEnabled = True
        else:
            self.ok_button.IsEnabled = False

        # This must be called after _initializing = False
        # and after elements are loaded
        if self.elements:
            self._update_element_count()
            self._refresh_parameters() # This will call _update_group_order_list
        
        self._update_filter_box_visibility() # Set initial filter box state
    
    def _get_controls(self):
        """Get UI controls"""
        self.picked_count_label = self.FindName("PickedCountLbl")
        self.pick_button = self.FindName("PickElemsBtn")
        self.save_button = self.FindName("SaveProfileBtn")
        self.load_button = self.FindName("LoadProfileBtn")
        self.target_a_combo = self.FindName("TargetParamBoxA")
        self.target_b_combo = self.FindName("TargetParamBoxB")
        self.target_c_combo = self.FindName("TargetParamBoxC")

        # NEW Filter Controls
        self.filter_box = self.FindName("FilterBox")
        self.filter_param_combo = self.FindName("FilterParamBox")
        self.filter_value_box = self.FindName("FilterValueBox")
        self.filter_apply_btn = self.FindName("FilterApplyBtn")
        self.filter_status_label = self.FindName("FilterStatusLbl")

        self.prefix_box = self.FindName("PrefixBox")
        self.suffix_box = self.FindName("SuffixBox")
        self.start_box = self.FindName("StartNumberBox")
        self.order_combo = self.FindName("OrderModeBox")
        self.grouping_mode_combo = self.FindName("GroupingModeBox")
        self.group_order_box = self.FindName("GroupOrderBox") # NEW
        self.group_order_list = self.FindName("GroupOrderList") # NEW
        self.move_up_btn = self.FindName("MoveUpBtn") # NEW
        self.move_down_btn = self.FindName("MoveDownBtn") # NEW
        self.curve_label = self.FindName("CurvePickedLabel")
        self.curve_button = self.FindName("PickCurveBtn")
        self.place_tags_check = self.FindName("PlaceTagsChk")
        self.prompt_pos_check = self.FindName("PromptPosChk")
        self.same_ident_check = self.FindName("SameIdentChk")
        self.specify_order_check = self.FindName("SpecifyOrderChk")
        self.order_status_label = self.FindName("OrderStatusLabel")
        self.pick_order_button = self.FindName("PickOrderBtn")
        self.options_list = self.FindName("OptionsList")
        self.check_all_button = self.FindName("CheckAllBtn")
        self.check_none_button = self.FindName("CheckNoneBtn")
        self.compare_param_1_combo = self.FindName("CompareParamBox1")
        self.compare_param_2_combo = self.FindName("CompareParamBox2")
        self.compare_param_3_combo = self.FindName("CompareParamBox3")
        self.compare_param_4_combo = self.FindName("CompareParamBox4")
        
        # Deselect controls
        self.deselect_individual_radio = self.FindName("DeselectIndividualRadio")
        self.deselect_category_radio = self.FindName("DeselectCategoryRadio")
        self.deselect_family_radio = self.FindName("DeselectFamilyRadio")
        self.deselect_type_radio = self.FindName("DeselectTypeRadio")
        self.deselect_items_list = self.FindName("DeselectItemsList")
        
        self.ok_button = self.FindName("OKBtn")
        self.cancel_button = self.FindName("CancelBtn")
    
    def _wire_events(self):
        """Wire event handlers"""
        if self.order_combo:
            self.order_combo.SelectionChanged += self._on_order_changed
        if self.grouping_mode_combo:
            self.grouping_mode_combo.SelectionChanged += self._on_grouping_mode_changed
        if self.specify_order_check:
            self.specify_order_check.Checked += self._on_specify_order_changed
            self.specify_order_check.Unchecked += self._on_specify_order_changed
        if self.move_up_btn: # NEW
            self.move_up_btn.Click += self.OnMoveUp
        if self.move_down_btn: # NEW
            self.move_down_btn.Click += self.OnMoveDown
        if self.filter_apply_btn: # NEW
            self.filter_apply_btn.Click += self.OnFilterApply
    
    def _initialize_ui(self):
        """Initialize UI state"""
        if self.order_combo:
            self.order_combo.ItemsSource = self.settings.order_modes
            self.order_combo.SelectedItem = self.settings.order_mode
        
        if self.grouping_mode_combo:
            self.grouping_mode_combo.ItemsSource = self.settings.grouping_modes
            self.grouping_mode_combo.SelectedItem = self.settings.grouping_mode

        if self.group_order_list: # NEW
            self.group_order_list.ItemsSource = self.settings.group_order_items

        if self.options_list:
            self.options_list.ItemsSource = self.settings.compare_fields
        
        if self.prefix_box:
            self.prefix_box.Text = self.settings.prefix
        if self.suffix_box:
            self.suffix_box.Text = self.settings.suffix
        if self.start_box:
            self.start_box.Text = self.settings.start_number
        if self.place_tags_check:
            self.place_tags_check.IsChecked = self.settings.place_tags
        if self.prompt_pos_check:
            self.prompt_pos_check.IsChecked = self.settings.prompt_for_positions
        if self.same_ident_check:
            self.same_ident_check.IsChecked = self.settings.same_for_identical
        if self.specify_order_check:
            self.specify_order_check.IsChecked = self.settings.specify_order
        
        for combo in [
            self.target_a_combo, self.target_b_combo, self.target_c_combo,
            self.compare_param_1_combo, self.compare_param_2_combo,
            self.compare_param_3_combo, self.compare_param_4_combo,
            self.filter_param_combo # NEW
        ]:
            if combo:
                combo.ItemsSource = self.settings.available_params
        
        if self.target_a_combo:
            self.target_a_combo.Text = self.settings.target_a
        if self.target_b_combo:
            self.target_b_combo.Text = self.settings.target_b
        if self.target_c_combo:
            self.target_c_combo.Text = self.settings.target_c
        
        if self.compare_param_1_combo:
            self.compare_param_1_combo.Text = self.settings.compare_param_1
        if self.compare_param_2_combo:
            self.compare_param_2_combo.Text = self.settings.compare_param_2
        if self.compare_param_3_combo:
            self.compare_param_3_combo.Text = self.settings.compare_param_3
        if self.compare_param_4_combo:
            self.compare_param_4_combo.Text = self.settings.compare_param_4

        if self.pick_order_button:
            is_manual = self.specify_order_check.IsChecked if self.specify_order_check else False
            self.pick_order_button.IsEnabled = is_manual
        
        if self.order_status_label:
            if len(self.manual_order) > 0:
                self.order_status_label.Text = "Order: {} elements".format(len(self.manual_order))
            else:
                self.order_status_label.Text = "(not set)"
        
        if self.order_combo and self.specify_order_check:
            self.order_combo.IsEnabled = not self.specify_order_check.IsChecked
        
        # Defer calling update_group_order_list until elements are loaded
        self._on_order_changed(None, None)
        self._on_specify_order_changed(None, None)
        self._update_filter_box_visibility() # NEW
        self._refresh_deselect_list()  # Initialize deselect list
    
    def _update_element_count(self):
        """Update element count display to show filtered/total"""
        if self.picked_count_label:
            total_count = len(self.all_picked_elements)
            filtered_count = len(self.elements)
            
            if total_count == 0:
                self.picked_count_label.Content = "(none)"
            elif total_count == filtered_count:
                self.picked_count_label.Content = "{} elements".format(total_count)
            else:
                self.picked_count_label.Content = "{}/{} elements".format(filtered_count, total_count)
    
    def _refresh_parameters(self):
        """Refresh available parameters based on self.elements"""
        # Get params from the *original* selection so filter list is complete
        params = ParameterHelper.get_common_parameters(self.all_picked_elements) 
        
        # Store current filter selection before clearing
        current_filter_param = self.filter_param_combo.SelectedItem
        
        self.settings.available_params.Clear()
        for param in params:
            self.settings.available_params.Add(param)
        
        # Restore filter selection
        if current_filter_param in self.settings.available_params:
            self.filter_param_combo.SelectedItem = current_filter_param
        
        # Update group order list based on the *filtered* elements
        self._update_group_order_list() 

    def _update_filter_box_visibility(self):
        """Show/Hide the filter box based on element selection"""
        if self.filter_box:
            if self.all_picked_elements and len(self.all_picked_elements) > 0:
                self.filter_box.Visibility = System.Windows.Visibility.Visible
            else:
                self.filter_box.Visibility = System.Windows.Visibility.Collapsed

    def _read_ui_to_settings(self):
        """Read UI values to settings"""
        if self.prefix_box:
            self.settings.prefix = self.prefix_box.Text or ""
        if self.suffix_box:
            self.settings.suffix = self.suffix_box.Text or ""
        if self.start_box:
            self.settings.start_number = self.start_box.Text or "1"
        if self.order_combo:
            self.settings.order_mode = self.order_combo.SelectedItem or ORDER_MODES[0]
        if self.grouping_mode_combo:
            self.settings.grouping_mode = self.grouping_mode_combo.SelectedItem or GROUPING_MODES[0]
        if self.place_tags_check:
            self.settings.place_tags = bool(self.place_tags_check.IsChecked)
        if self.prompt_pos_check:
            self.settings.prompt_for_positions = bool(self.prompt_pos_check.IsChecked)
        if self.same_ident_check:
            self.settings.same_for_identical = bool(self.same_ident_check.IsChecked)
        if self.specify_order_check:
            self.settings.specify_order = bool(self.specify_order_check.IsChecked)
        if self.target_a_combo:
            self.settings.target_a = self.target_a_combo.Text or ""
        if self.target_b_combo:
            self.settings.target_b = self.target_b_combo.Text or ""
        if self.target_c_combo:
            self.settings.target_c = self.target_c_combo.Text or ""
            
        if self.compare_param_1_combo:
            self.settings.compare_param_1 = self.compare_param_1_combo.Text or ""
        if self.compare_param_2_combo:
            self.settings.compare_param_2 = self.compare_param_2_combo.Text or ""
        if self.compare_param_3_combo:
            self.settings.compare_param_3 = self.compare_param_3_combo.Text or ""
        if self.compare_param_4_combo:
            self.settings.compare_param_4 = self.compare_param_4_combo.Text or ""
        
        # NEW: Save the custom group order
        if self.group_order_list and self.settings.group_order_items:
            self.settings.saved_group_order = ",".join(
                [item.Name for item in self.settings.group_order_items]
            )
    
    def _refresh_options_list(self):
        """Refresh options list"""
        if self.options_list:
            self.options_list.ItemsSource = None
            self.options_list.ItemsSource = self.settings.compare_fields
    
    @staticmethod
    def transfer_settings_static(new_window, saved_settings, saved_curve):
        """Static method to transfer settings"""
        new_window.settings.prefix = saved_settings.prefix
        new_window.settings.suffix = saved_settings.suffix
        new_window.settings.start_number = saved_settings.start_number
        new_window.settings.order_mode = saved_settings.order_mode
        new_window.settings.grouping_mode = saved_settings.grouping_mode
        new_window.settings.place_tags = saved_settings.place_tags
        new_window.settings.prompt_for_positions = saved_settings.prompt_for_positions
        new_window.settings.same_for_identical = saved_settings.same_for_identical
        new_window.settings.specify_order = saved_settings.specify_order
        new_window.settings.target_a = saved_settings.target_a
        new_window.settings.target_b = saved_settings.target_b
        new_window.settings.target_c = saved_settings.target_c
        new_window.settings.compare_param_1 = saved_settings.compare_param_1
        new_window.settings.compare_param_2 = saved_settings.compare_param_2
        new_window.settings.compare_param_3 = saved_settings.compare_param_3
        new_window.settings.compare_param_4 = saved_settings.compare_param_4
        new_window.settings.saved_group_order = saved_settings.saved_group_order # NEW
        
        saved_fields_lookup = {f.Name: f.Include for f in saved_settings.compare_fields}

        for field in new_window.settings.compare_fields:
            if field.Name in saved_fields_lookup:
                field.Include = saved_fields_lookup[field.Name]
        
        if saved_curve and not saved_settings.specify_order:
            new_window.picked_curve = saved_curve
            if hasattr(new_window, 'curve_label') and new_window.curve_label:
                new_window.curve_label.Text = "Curve Id: {}".format(saved_curve.Id.IntegerValue)
    
    # XAML Click handlers
    def OnPickElements(self, sender, args):
        self._handle_pick_elements()
    
    def OnSaveProfile(self, sender, args):
        self._handle_save_profile()
    
    def OnLoadProfile(self, sender, args):
        self._handle_load_profile()
    
    def OnPickCurve(self, sender, args):
        self._handle_pick_curve()
    
    def OnPickOrder(self, sender, args):
        self._handle_pick_order()
    
    def OnCheckAll(self, sender, args):
        for field in self.settings.compare_fields:
            field.Include = True
        self._refresh_options_list()
    
    def OnCheckNone(self, sender, args):
        for field in self.settings.compare_fields:
            field.Include = False
        self._refresh_options_list()

    # NEW: Move Up Button
    def OnMoveUp(self, sender, args):
        if self.group_order_list and self.group_order_list.SelectedItem:
            selected_index = self.group_order_list.SelectedIndex
            if selected_index > 0:
                item_to_move = self.settings.group_order_items[selected_index]
                self.settings.group_order_items.RemoveAt(selected_index)
                self.settings.group_order_items.Insert(selected_index - 1, item_to_move)
                self.group_order_list.SelectedIndex = selected_index - 1

    # NEW: Move Down Button
    def OnMoveDown(self, sender, args):
        if self.group_order_list and self.group_order_list.SelectedItem:
            selected_index = self.group_order_list.SelectedIndex
            if selected_index < len(self.settings.group_order_items) - 1:
                item_to_move = self.settings.group_order_items[selected_index]
                self.settings.group_order_items.RemoveAt(selected_index)
                self.settings.group_order_items.Insert(selected_index + 1, item_to_move)
                self.group_order_list.SelectedIndex = selected_index + 1
    
    # NEW: Filter Apply Button
    def OnFilterApply(self, sender, args):
        """Handle the Apply Filter button click"""
        param_name = self.filter_param_combo.SelectedItem
        value_to_match = self.filter_value_box.Text
        
        if not param_name:
            # If param field is empty, clear the filter
            self.elements = list(self.all_picked_elements)
            self.filter_status_label.Text = "(filter cleared)"
            self.filter_value_box.Text = ""
        else:
            # Apply the filter
            new_filtered_list = []
            value_lower = value_to_match.lower()
            
            for elem in self.all_picked_elements:
                elem_val = ParameterHelper.get_value(elem, param_name)
                # Check for 'is not None' to handle empty strings ""
                if elem_val is not None and elem_val.lower() == value_lower:
                    new_filtered_list.append(elem)
            
            self.elements = new_filtered_list
            self.filter_status_label.Text = "Filtered {} items".format(len(self.elements))
        
        # Refresh UI elements based on the new self.elements list
        self._update_element_count()
        self._refresh_parameters() # This re-populates group lists, etc.
        
        # Re-check if OK button should be enabled
        self.ok_button.IsEnabled = len(self.elements) > 0

    def OnDeselectModeChanged(self, sender, args):
        """Handle deselect mode radio button change"""
        self._refresh_deselect_list()

    def OnDeselectSelected(self, sender, args):
        """Handle deselect button - removes selected items based on current mode"""
        if not self.deselect_items_list:
            return
        
        selected_items = list(self.deselect_items_list.SelectedItems)
        if not selected_items:
            forms.alert("Please select items to deselect from the list.", title="No Selection")
            return
        
        # Determine which mode we're in
        if self.deselect_individual_radio and self.deselect_individual_radio.IsChecked:
            self._deselect_by_individual(selected_items)
        elif self.deselect_category_radio and self.deselect_category_radio.IsChecked:
            self._deselect_by_category(selected_items)
        elif self.deselect_family_radio and self.deselect_family_radio.IsChecked:
            self._deselect_by_family(selected_items)
        elif self.deselect_type_radio and self.deselect_type_radio.IsChecked:
            self._deselect_by_type(selected_items)
    
    def OnDeselectAll(self, sender, args):
        """Handle deselect all button - clears all selected elements"""
        if not self.elements:
            forms.alert("No elements to deselect.", title="Already Empty")
            return
        
        count = len(self.elements)
        
        # Clear all elements
        self.elements = []
        self.all_picked_elements = []
        self.manual_order = []
        
        # Refresh the UI
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        
        # Disable OK button
        if self.ok_button:
            self.ok_button.IsEnabled = False
        
        forms.toast("Deselected all {} element(s)".format(count), title="Deselect All", appid="renumber")

    def _deselect_by_individual(self, selected_items):
        """Deselect specific individual elements"""
        # For individual mode, selected_items are DeselectItem objects with ElementIds
        ids_to_remove = set()
        for item in selected_items:
            for elem_id in item.ElementIds:
                ids_to_remove.add(str(elem_id))
        
        initial_count = len(self.elements)
        self.elements = [e for e in self.elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self.all_picked_elements = [e for e in self.all_picked_elements if str(e.Id.IntegerValue) not in ids_to_remove]
        
        removed_count = initial_count - len(self.elements)
        
        # Refresh the UI
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        
        if self.ok_button:
            self.ok_button.IsEnabled = len(self.elements) > 0
        
        forms.toast("Deselected {} element(s)".format(removed_count), title="Deselect", appid="renumber")
    
    def _deselect_by_category(self, selected_items):
        """Deselect all elements matching selected categories"""
        # For category mode, selected_items are DeselectItem objects
        categories_to_remove = set([item.Name for item in selected_items])
        
        initial_count = len(self.elements)
        self.elements = [e for e in self.elements if not (e.Category and e.Category.Name in categories_to_remove)]
        self.all_picked_elements = [e for e in self.all_picked_elements if not (e.Category and e.Category.Name in categories_to_remove)]
        
        removed_count = initial_count - len(self.elements)
        
        # Refresh the UI
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        
        if self.ok_button:
            self.ok_button.IsEnabled = len(self.elements) > 0
        
        forms.toast("Deselected {} element(s) from {} category(ies)".format(removed_count, len(categories_to_remove)), 
                   title="Deselect", appid="renumber")
    
    def _deselect_by_family(self, selected_items):
        """Deselect all elements matching selected families"""
        families_to_remove = set([item.Name for item in selected_items])
        
        initial_count = len(self.elements)
        new_elements = []
        new_all_picked = []
        
        for e in self.elements:
            family_name = self._get_element_family(e)
            if family_name not in families_to_remove:
                new_elements.append(e)
        
        for e in self.all_picked_elements:
            family_name = self._get_element_family(e)
            if family_name not in families_to_remove:
                new_all_picked.append(e)
        
        self.elements = new_elements
        self.all_picked_elements = new_all_picked
        
        removed_count = initial_count - len(self.elements)
        
        # Refresh the UI
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        
        if self.ok_button:
            self.ok_button.IsEnabled = len(self.elements) > 0
        
        forms.toast("Deselected {} element(s) from {} family(ies)".format(removed_count, len(families_to_remove)), 
                   title="Deselect", appid="renumber")
    
    def _deselect_by_type(self, selected_items):
        """Deselect all elements matching selected types"""
        types_to_remove = set([item.Name for item in selected_items])
        
        initial_count = len(self.elements)
        new_elements = []
        new_all_picked = []
        
        for e in self.elements:
            type_name = self._get_element_type_name(e)
            if type_name not in types_to_remove:
                new_elements.append(e)
        
        for e in self.all_picked_elements:
            type_name = self._get_element_type_name(e)
            if type_name not in types_to_remove:
                new_all_picked.append(e)
        
        self.elements = new_elements
        self.all_picked_elements = new_all_picked
        
        removed_count = initial_count - len(self.elements)
        
        # Refresh the UI
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        
        if self.ok_button:
            self.ok_button.IsEnabled = len(self.elements) > 0
        
        forms.toast("Deselected {} element(s) from {} type(s)".format(removed_count, len(types_to_remove)), 
                   title="Deselect", appid="renumber")

    def _refresh_deselect_list(self):
        """Refresh the deselect list based on current mode"""
        if not self.deselect_items_list:
            return
        
        # Determine which mode we're in
        if self.deselect_individual_radio and self.deselect_individual_radio.IsChecked:
            self._populate_individual_list()
        elif self.deselect_category_radio and self.deselect_category_radio.IsChecked:
            self._populate_category_list()
        elif self.deselect_family_radio and self.deselect_family_radio.IsChecked:
            self._populate_family_list()
        elif self.deselect_type_radio and self.deselect_type_radio.IsChecked:
            self._populate_type_list()
        else:
            # Default to individual
            self._populate_individual_list()
    
    def _populate_individual_list(self):
        """Populate list with individual elements"""
        items = ObservableCollection[DeselectItem]()
        
        for elem in self.elements:
            elem_id = str(elem.Id.IntegerValue)
            category = elem.Category.Name if elem.Category else "Unknown"
            
            # Try to get element name
            name = "Element {}".format(elem_id)
            try:
                name_param = elem.get_Parameter(DB.BuiltInParameter.ELEM_NAME_PARAM)
                if name_param and name_param.HasValue:
                    name = name_param.AsString()
                else:
                    # Try to get Type name
                    elem_type = doc.GetElement(elem.GetTypeId())
                    if elem_type:
                        type_name = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                        if type_name and type_name.HasValue:
                            name = type_name.AsString()
            except:
                pass
            
            # Create a display name that combines ID, category, and name
            display_name = "ID {} - {} - {}".format(elem_id, category, name)
            items.Add(DeselectItem(display_name, 1, [elem_id]))
        
        self.deselect_items_list.ItemsSource = items
    
    def _populate_category_list(self):
        """Populate list with categories and counts"""
        category_dict = {}
        
        for elem in self.elements:
            cat_name = elem.Category.Name if elem.Category else "No Category"
            if cat_name not in category_dict:
                category_dict[cat_name] = []
            category_dict[cat_name].append(elem.Id.IntegerValue)
        
        items = ObservableCollection[DeselectItem]()
        for cat_name in sorted(category_dict.keys()):
            item_count = len(category_dict[cat_name])
            items.Add(DeselectItem(cat_name, item_count, category_dict[cat_name]))
        
        self.deselect_items_list.ItemsSource = items
    
    def _populate_family_list(self):
        """Populate list with families and counts"""
        family_dict = {}
        
        for elem in self.elements:
            family_name = self._get_element_family(elem)
            if family_name not in family_dict:
                family_dict[family_name] = []
            family_dict[family_name].append(elem.Id.IntegerValue)
        
        items = ObservableCollection[DeselectItem]()
        for family_name in sorted(family_dict.keys()):
            item_count = len(family_dict[family_name])
            items.Add(DeselectItem(family_name, item_count, family_dict[family_name]))
        
        self.deselect_items_list.ItemsSource = items
    
    def _populate_type_list(self):
        """Populate list with types and counts"""
        type_dict = {}
        
        for elem in self.elements:
            type_name = self._get_element_type_name(elem)
            if type_name not in type_dict:
                type_dict[type_name] = []
            type_dict[type_name].append(elem.Id.IntegerValue)
        
        items = ObservableCollection[DeselectItem]()
        for type_name in sorted(type_dict.keys()):
            item_count = len(type_dict[type_name])
            items.Add(DeselectItem(type_name, item_count, type_dict[type_name]))
        
        self.deselect_items_list.ItemsSource = items
    
    def _get_element_family(self, elem):
        """Get family name for an element"""
        try:
            elem_type = doc.GetElement(elem.GetTypeId())
            if elem_type:
                family_param = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if family_param and family_param.HasValue:
                    return family_param.AsString()
        except:
            pass
        return "Unknown Family"
    
    def _get_element_type_name(self, elem):
        """Get type name for an element"""
        try:
            elem_type = doc.GetElement(elem.GetTypeId())
            if elem_type:
                type_name = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if type_name and type_name.HasValue:
                    return type_name.AsString()
                # Fallback to element type name
                return elem_type.Name if hasattr(elem_type, 'Name') else "Unknown Type"
        except:
            pass
        return "Unknown Type"


    def OnOK(self, sender, args):
        """Execute renumbering"""
        global WINDOW_INSTANCE
        
        try:
            # self.elements is now the filtered list
            if not self.elements:
                forms.alert("No elements selected (or filter resulted in 0 elements)", title="Renumber")
                return
            
            self._read_ui_to_settings()
            
            if not self.settings.specify_order and self.settings.order_mode == "Follow spline" and not self.picked_curve:
                forms.alert("Please pick a curve for 'Follow spline' mode", title="Renumber")
                return
            
            ConfigManager.save(self.settings)
            
            # NEW: Get the custom group order from the UI list
            custom_group_order = None
            if self.settings.grouping_mode != GROUPING_MODES[0]:
                custom_group_order = [item.Name for item in self.settings.group_order_items]

            # Pass the filtered list (self.elements) to the engine
            success, skipped, errors = RenumberEngine.execute(
                self.elements,
                self.settings,
                self.picked_curve,
                self.manual_order,
                custom_group_order # NEW
            )
            
            self._results = {
                'success': success,
                'skipped': skipped,
                'errors': errors,
                'total': len(self.elements) # Report based on filtered list
            }
            
            WINDOW_INSTANCE = None
            self.DialogResult = True
            self.Close()
            
        except Exception as ex:
            WINDOW_INSTANCE = None
            self._results = None
            try:
                self.DialogResult = False
                self.Close()
            except:
                pass
            forms.alert("Renumbering failed:\n\n{}".format(ex), title="Error")
    
    def OnCancel(self, sender, args):
        global WINDOW_INSTANCE
        WINDOW_INSTANCE = None
        self.DialogResult = False
        self.Close()
    
    def _on_order_changed(self, sender, args):
        """Handle order mode change"""
        try:
            if hasattr(self, '_initializing') and self._initializing:
                return
            
            if self.order_combo and self.curve_button:
                mode = self.order_combo.SelectedItem
                is_spline_mode = (mode == "Follow spline")
                
                is_manual_order = self.specify_order_check.IsChecked if self.specify_order_check else False
                
                if is_manual_order:
                    self.curve_button.IsEnabled = False
                else:
                    self.curve_button.IsEnabled = is_spline_mode
                
                if self.curve_label:
                    if is_spline_mode and self.picked_curve:
                        self.curve_label.Text = "Curve Id: {}".format(self.picked_curve.Id.IntegerValue)
                    elif is_spline_mode:
                        self.curve_label.Text = "(pick curve needed)"
                    else:
                        self.curve_label.Text = "(not applicable)"
        except:
            pass

    def _on_grouping_mode_changed(self, sender, args):
        """Handle grouping mode change"""
        try:
            if hasattr(self, '_initializing') and self._initializing:
                return
            
            self._update_group_order_list() # NEW
            
            if self.grouping_mode_combo and self.specify_order_check:
                mode = self.grouping_mode_combo.SelectedItem
                is_default_mode = (mode == GROUPING_MODES[0])  # "All at Once"
                
                self.specify_order_check.IsEnabled = is_default_mode
                if not is_default_mode:
                    self.specify_order_check.IsChecked = False
        except:
            pass

    def _on_specify_order_changed(self, sender, args):
        """Handle manual order checkbox change"""
        try:
            is_checked = self.specify_order_check.IsChecked if self.specify_order_check else False
            
            if self.pick_order_button:
                self.pick_order_button.IsEnabled = is_checked
            
            if self.order_status_label:
                # Update status label based on the *filtered* manual order
                filtered_manual_order_ids = set([e.Id for e in self.elements])
                filtered_manual_count = len([e for e in self.manual_order if e.Id in filtered_manual_order_ids])

                if is_checked and filtered_manual_count > 0:
                    self.order_status_label.Text = "Order: {} elements".format(filtered_manual_count)
                else:
                    self.order_status_label.Text = "(not set)"
            
            if self.order_combo:
                self.order_combo.IsEnabled = not is_checked
            
            if self.grouping_mode_combo:
                self.grouping_mode_combo.IsEnabled = not is_checked
                if is_checked:
                    self.grouping_mode_combo.SelectedItem = GROUPING_MODES[0] 

            if self.group_order_box:
                if is_checked:
                    self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
                else:
                    self._update_group_order_list() # Will set visibility

            if self.curve_button:
                if is_checked:
                    self.curve_button.IsEnabled = False
                else:
                    mode = self.order_combo.SelectedItem if self.order_combo else None
                    self.curve_button.IsEnabled = (mode == "Follow spline")
            
            if not is_checked:
                self.manual_order = []
                
        except:
            pass
    
    # NEW: Helper to populate and sort the Group Order ListBox
    def _update_group_order_list(self):
        """Populates the Group Order list based on selection and saved order."""
        if not hasattr(self, 'settings') or not self.group_order_box or self._initializing:
            return

        try:
            mode = self.grouping_mode_combo.SelectedItem
            
            # Use self.elements (the filtered list)
            if mode == GROUPING_MODES[0] or not self.elements:
                self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
                self.settings.group_order_items.Clear()
                return

            self.group_order_box.Visibility = System.Windows.Visibility.Visible
            
            # Get all unique group keys from the current *filtered* selection
            unique_keys = set()
            for elem in self.elements: # Use self.elements
                key = ParameterHelper.get_group_key(elem, mode)
                unique_keys.add(key)
            
            # Get the saved order
            saved_order = []
            if self.settings.saved_group_order:
                saved_order = self.settings.saved_group_order.split(',')
            
            # Create a list of GroupItem objects
            items = [GroupItem(key) for key in unique_keys]
            
            # Sort the items:
            # 1. By their index in the saved_order list
            # 2. Alphabetically for any new items not in saved_order
            def sort_key(item):
                try:
                    idx = saved_order.index(item.Name)
                    return (idx, item.Name)
                except:
                    # Item not in saved list, put it at the end alphabetically
                    return (len(saved_order), item.Name)

            items.sort(key=sort_key)
            
            # Repopulate the observable collection
            self.settings.group_order_items.Clear()
            for item in items:
                self.settings.group_order_items.Add(item)

        except Exception as ex:
            # Failsafe
            if self.group_order_box:
                self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
            if self.settings and self.settings.group_order_items:
                self.settings.group_order_items.Clear()

    def _handle_pick_elements(self):
        """Handle pick elements"""
        global WINDOW_INSTANCE
        
        if self.elements:
            result = forms.alert(
                "You already have {} elements selected.\n\nDo you want to select different elements?".format(len(self.all_picked_elements)),
                title="Re-pick Elements?",
                ok=False,
                yes=True,
                no=True
            )
            if not result:
                return
        
        try:
            self._read_ui_to_settings()
            saved_settings = self.settings
            saved_curve = self.picked_curve
            
            self._next_action = {
                'action': 'pick_elements',
                'saved_settings': saved_settings,
                'saved_curve': saved_curve
            }
            
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
                
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")
    
    def _handle_pick_order(self):
        """Handle pick order button - picks elements AND defines order in one step"""
        global WINDOW_INSTANCE
        
        try:
            self._read_ui_to_settings()
            saved_settings = self.settings
            saved_curve = self.picked_curve
            
            self._next_action = {
                'action': 'pick_order',
                'saved_settings': saved_settings,
                'saved_curve': saved_curve
            }
            
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
            
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")
    
    def _handle_pick_curve(self):
        """Handle pick curve"""
        global WINDOW_INSTANCE
        
        try:
            self._read_ui_to_settings()
            saved_settings = self.settings
            saved_elements = self.elements # Pass the filtered list
            saved_all_elements = self.all_picked_elements # Pass the original list
            saved_manual_order = self.manual_order
            
            self._next_action = {
                'action': 'pick_curve',
                'saved_settings': saved_settings,
                'saved_elements': saved_elements,
                'saved_all_elements': saved_all_elements, # NEW
                'saved_manual_order': saved_manual_order
            }
            
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
            
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")
    
    def _handle_save_profile(self):
        """Save profile to XML"""
        if self._executing_save:
            return
        
        self._executing_save = True
        try:
            self._read_ui_to_settings()
            
            path = forms.save_file(
                file_ext="xml",
                default_name="RenumberProfile.xml",
                title="Save Profile As"
            )
            if not path:
                return
            
            xml_doc = ProfileManager.export(self.settings)
            xml_doc.Save(path)
            
            forms.toast("Profile saved", title="Renumber", appid="renumber")
        except Exception as ex:
            forms.alert("Could not save profile:\n\n{}".format(ex), title="Error")
        finally:
            self._executing_save = False
    
    def _handle_load_profile(self):
        """Load profile from XML"""
        if self._executing_load:
            return
        
        self._executing_load = True
        try:
            path = forms.pick_file(file_ext="xml", title="Select Profile to Load")
            if not path:
                return
            
            xml_doc = XmlDocument()
            xml_doc.Load(path)
            
            ProfileManager.import_to(xml_doc, self.settings)
            
            self._initialize_ui()
            self._refresh_options_list()
            self._update_group_order_list() # NEW: Refresh group list after loading profile
            
            forms.toast("Profile loaded", title="Renumber", appid="renumber")
        except Exception as ex:
            forms.alert("Could not load profile:\n\n{}".format(ex), title="Error")
        finally:
            self._executing_load = False
    
# =============================================================================
# ENTRY POINT
# =============================================================================

def run():
    """Main entry point"""
    global WINDOW_INSTANCE
    
    WINDOW_INSTANCE = None
    
    preselection = list(revit.get_selection() or [])
    
    try:
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "window.xaml")
        
        if not os.path.exists(xaml_path):
            error_msg = "XAML file not found: {}".format(xaml_path)
            forms.alert(error_msg, title="Renumber Error")
            return
        
        window = RenumberWindow(xaml_path, preselection if preselection else [])
        
        while True:
            WINDOW_INSTANCE = window
            result = window.ShowDialog()
            
            next_action = window._next_action if hasattr(window, '_next_action') else None
            results = window._results if hasattr(window, '_results') else None
            
            WINDOW_INSTANCE = None
            
            if results:
                success = results['success']
                skipped = results['skipped']
                errors = results['errors']
                total = results['total']
                
                if errors and len(errors) > 0:
                    output.print_md("---")
                    output.print_md("## Renumber Results")
                    output.print_md("- **Total Renumbered:** {}".format(total))
                    output.print_md("- **Success:** {}".format(success))
                    output.print_md("- **Skipped:** {}".format(skipped))
                    if skipped > 0:
                        output.print_md("\n**Issues:**")
                        for error in errors[:20]:
                            output.print_md("- {}".format(error))
                        if len(errors) > 20:
                            output.print_md("- *(and {} more)*".format(len(errors) - 20))
                
                if success > 0:
                    if skipped > 0:
                        forms.toast(
                            "Renumbered {} elements ({} skipped)".format(success, skipped),
                            title="Renumber Complete",
                            appid="renumber"
                        )
                    else:
                        forms.toast(
                            "Successfully renumbered {} elements".format(success),
                            title="Renumber Complete",
                            appid="renumber"
                        )
                else:
                    forms.alert(
                        "No elements were updated.\n\nCheck the output window for details.",
                        title="Renumber Failed"
                    )
                break
            
            if not next_action:
                break
            
            action_type = next_action.get('action')
            
            if action_type == 'pick_elements':
                elements = ElementPicker.pick_multiple()
                if not elements:
                    break
                
                saved_settings = next_action['saved_settings']
                saved_curve = next_action['saved_curve']
                
                window = RenumberWindow(xaml_path, elements)
                window.manual_order = []
                RenumberWindow.transfer_settings_static(window, saved_settings, saved_curve)
                window._initialize_ui()
                # _refresh_parameters() is called in __init__
                if window.ok_button:
                    window.ok_button.IsEnabled = True
                continue
            
            elif action_type == 'pick_curve':
                curve = ElementPicker.pick_curve()
                
                saved_settings = next_action['saved_settings']
                saved_elements = next_action['saved_elements'] # Filtered list
                saved_all_elements = next_action['saved_all_elements'] # Original list
                saved_manual_order = next_action['saved_manual_order']
                
                window = RenumberWindow(xaml_path, saved_elements)
                window.all_picked_elements = saved_all_elements # Restore original list
                window.manual_order = saved_manual_order
                RenumberWindow.transfer_settings_static(window, saved_settings, curve)
                
                if curve:
                    window.picked_curve = curve
                    window.settings.picked_curve_id = curve.Id.IntegerValue
                
                window._initialize_ui()
                # _refresh_parameters() is called in __init__
                
                if curve and window.curve_label:
                    window.curve_label.Text = "Curve Id: {}".format(curve.Id.IntegerValue)
                
                if window.elements and window.ok_button:
                    window.ok_button.IsEnabled = True
                continue
            
            elif action_type == 'pick_order':
                forms.alert(
                    "Click on elements in the order you want them numbered.\n\n"
                    "This will select the elements AND define their numbering order.\n\n"
                    "Press ESC when finished.",
                    title="Pick Elements in Order",
                    ok=True
                )
                
                manual_order = ElementPicker.pick_in_order()
                
                if not manual_order or len(manual_order) == 0:
                    break
                
                saved_settings = next_action['saved_settings']
                saved_curve = next_action['saved_curve']
                
                saved_settings.specify_order = True
                
                window = RenumberWindow(xaml_path, manual_order)
                window.manual_order = list(manual_order) # Set manual order
                window.all_picked_elements = list(manual_order) # Set all_picked
                
                RenumberWindow.transfer_settings_static(window, saved_settings, saved_curve)
                
                window._initialize_ui()
                # _refresh_parameters() is called in __init__
                
                if window.order_status_label:
                    window.order_status_label.Text = "Order: {} elements".format(len(manual_order))
                
                if window.specify_order_check:
                    window.specify_order_check.IsChecked = True
                
                if window.pick_order_button:
                    window.pick_order_button.IsEnabled = True
                
                if window.ok_button:
                    window.ok_button.IsEnabled = True
                
                continue
            
            else:
                break
        
    except Exception as ex:
        WINDOW_INSTANCE = None
        forms.alert(
            "Failed to open Renumber window.\n\nError: {}".format(ex),
            title="Renumber Error"
        )


if __name__ == "__main__":
    run()