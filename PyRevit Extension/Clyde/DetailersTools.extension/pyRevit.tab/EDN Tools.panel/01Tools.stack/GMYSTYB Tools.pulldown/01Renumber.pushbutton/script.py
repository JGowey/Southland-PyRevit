# -*- coding: utf-8 -*-
"""
Renumber Tool for pyRevit
Author: Jeremiah Griffith
Version: 2.2.11 (Filter Value Dropdown)

Intelligently renumber elements with multiple ordering modes and comparison options.
Includes grouping by Category/Family/Type with custom ordering.
Includes post-pick filtering with dynamic value dropdown.
Now supports Assembly Name and Assembly Mark.
"""

from pyrevit import revit, DB, forms, script
import math
import os
import System
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

# SET THIS TO TRUE TO SEE DIAGNOSTICS IN OUTPUT WINDOW
DEBUG_MODE = False

def log_debug(msg):
    if DEBUG_MODE:
        print("[DEBUG] " + msg)

# Grouping modes
GROUPING_MODES = [
    "All at Once",
    "By Category",
    "By Family",
    "By Type",
    "Order by Parameter"
]

ORDER_MODES = [
    "Largest -> Smallest",
    "Smallest -> Largest",
    "Clockwise",
    "Counterclockwise",
    "Follow spline"
]

# Fabrication-only compare fields (FabricationPart class properties)
FAB_COMPARE_FIELDS = [
    "Alias", "Angle", "Cid", "CutType", "Extension", 
    "FabricationNotes", "Gauge", "Offset", "Service", "SheetMetalArea", "Specification"
]

# Common compare fields (work on Fab + other element types via parameter lookup or ConnectorManager)
COMMON_COMPARE_FIELDS = [
    "Category", "CenterlineLength", "Connectors", "Dimensions", 
    "EndConnector1", "EndConnector2", "FamilyInfo", "InsulationSpecification", 
    "InsulationType", "Length", "LiningThickness", "Material", 
    "Size", "Status", "Weight"
]

# Combined for backward compatibility
DEFAULT_COMPARE_FIELDS = FAB_COMPARE_FIELDS + COMMON_COMPARE_FIELDS

ASSEMBLY_MARK_PARAM = "Assembly Mark"
CONFIG_SECTION = "RenumberTool"
WINDOW_INSTANCE = None


# =============================================================================
# DATA CLASSES
# =============================================================================

class CompareField(object):
    def __init__(self, name, is_included=False):
        self.Name = name
        self.Include = is_included

class GroupItem(object):
    def __init__(self, name):
        self.Name = name

class ElementItem(object):
    def __init__(self, elem_id, category, name):
        self.Id = elem_id
        self.Category = category
        self.Name = name

class DeselectItem(object):
    def __init__(self, name, count, element_ids=None):
        self.Name = name
        self.Count = "({} elements)".format(count) if count > 1 else "(1 element)"
        self.ElementIds = element_ids or []


class Settings(object):
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
        self.saved_group_order = ""
        self.order_parameter = ""
        self.order_parameter_ascending = True
        self.include_nested_families = False
        self.zero_pad_digits = 0  # 0 = no padding, positive = pad to that many digits
        
        # Separate collections for Fab and Common compare fields
        self.fab_compare_fields = ObservableCollection[CompareField]()
        for name in FAB_COMPARE_FIELDS:
            self.fab_compare_fields.Add(CompareField(name))
        
        self.common_compare_fields = ObservableCollection[CompareField]()
        for name in COMMON_COMPARE_FIELDS:
            self.common_compare_fields.Add(CompareField(name))
        
        self.order_modes = ObservableCollection[str]()
        for mode in ORDER_MODES:
            self.order_modes.Add(mode)

        self.grouping_modes = ObservableCollection[str]()
        for mode in GROUPING_MODES:
            self.grouping_modes.Add(mode)
        
        self.group_order_items = ObservableCollection[GroupItem]()
        
        self.available_params = ObservableCollection[str]()
        
        self.filter_values = ObservableCollection[str]()


# =============================================================================
# CONFIGURATION MANAGEMENT
# =============================================================================

class ConfigManager(object):
    @staticmethod
    def load():
        settings = Settings()
        cfg = script.get_config(CONFIG_SECTION)
        
        settings.prefix = ConfigManager._safe_get(cfg, "Prefix", "")
        settings.suffix = ConfigManager._safe_get(cfg, "Suffix", "")
        settings.start_number = ConfigManager._safe_get(cfg, "StartNumber", "1")
        
        loaded_order_mode = ConfigManager._safe_get(cfg, "OrderMode", ORDER_MODES[0])
        settings.order_mode = loaded_order_mode if loaded_order_mode in ORDER_MODES else ORDER_MODES[0]
        
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
        settings.saved_group_order = ConfigManager._safe_get(cfg, "SavedGroupOrder", "")
        settings.order_parameter = ConfigManager._safe_get(cfg, "OrderParameter", "")
        settings.order_parameter_ascending = bool(ConfigManager._safe_get(cfg, "OrderParameterAscending", True))
        settings.include_nested_families = bool(ConfigManager._safe_get(cfg, "IncludeNested", False))
        settings.zero_pad_digits = int(ConfigManager._safe_get(cfg, "ZeroPadDigits", 0))
        
        compare_str = ConfigManager._safe_get(cfg, "CompareFields", "")
        selected_fields = set(f.strip() for f in compare_str.split(",") if f.strip())
        
        # Populate Fab compare fields
        settings.fab_compare_fields.Clear()
        for name in FAB_COMPARE_FIELDS:
            settings.fab_compare_fields.Add(CompareField(name, name in selected_fields))
        
        # Populate Common compare fields
        settings.common_compare_fields.Clear()
        for name in COMMON_COMPARE_FIELDS:
            settings.common_compare_fields.Add(CompareField(name, name in selected_fields))

        return settings
    
    @staticmethod
    def save(settings):
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
            cfg.SavedGroupOrder = settings.saved_group_order
            cfg.OrderParameter = settings.order_parameter
            cfg.OrderParameterAscending = settings.order_parameter_ascending
            cfg.IncludeNested = settings.include_nested_families
            cfg.ZeroPadDigits = settings.zero_pad_digits
            # Combine fab and common compare fields for saving
            selected_fields = []
            for f in settings.fab_compare_fields:
                if f.Include:
                    selected_fields.append(f.Name)
            for f in settings.common_compare_fields:
                if f.Include:
                    selected_fields.append(f.Name)
            cfg.CompareFields = ",".join(selected_fields)
            script.save_config()
        except:
            pass
    
    @staticmethod
    def _safe_get(cfg, key, default=None):
        try:
            return getattr(cfg, key, default) if hasattr(cfg, key) else default
        except:
            return default


# =============================================================================
# PROFILE IMPORT/EXPORT
# =============================================================================

class ProfileManager(object):
    @staticmethod
    def export(settings):
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
            "SavedGroupOrder": settings.saved_group_order,
            "OrderParameter": settings.order_parameter,
            "OrderParameterAscending": str(settings.order_parameter_ascending),
            "IncludeNested": str(settings.include_nested_families)
        }
        
        for name, value in elements.items():
            el = xml.CreateElement(name)
            el.InnerText = str(value) if value is not None else ""
            root.AppendChild(el)
        
        fields_node = xml.CreateElement("CompareFields")
        root.AppendChild(fields_node)
        for field in settings.fab_compare_fields:
            if field.Include:
                field_node = xml.CreateElement("Field")
                field_node.SetAttribute("name", field.Name)
                fields_node.AppendChild(field_node)
        for field in settings.common_compare_fields:
            if field.Include:
                field_node = xml.CreateElement("Field")
                field_node.SetAttribute("name", field.Name)
                fields_node.AppendChild(field_node)
        
        return xml
    
    @staticmethod
    def import_to(xml_doc, settings):
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
        settings.saved_group_order = get_text("SavedGroupOrder")
        settings.order_parameter = get_text("OrderParameter")
        settings.order_parameter_ascending = get_text("OrderParameterAscending") == "True"
        settings.include_nested_families = get_text("IncludeNested") == "True"
        
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
        
        # Populate Fab compare fields
        settings.fab_compare_fields.Clear()
        for name in FAB_COMPARE_FIELDS:
            settings.fab_compare_fields.Add(CompareField(name, name in selected))
        
        # Populate Common compare fields
        settings.common_compare_fields.Clear()
        for name in COMMON_COMPARE_FIELDS:
            settings.common_compare_fields.Add(CompareField(name, name in selected))


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
        
        # For fabrication parts, ensure Item Number is available
        if FabricationHelper.is_fabrication_part(element):
            try:
                param = element.get_Parameter(DB.BuiltInParameter.FABRICATION_PART_ITEM_NUMBER)
                if param and not param.IsReadOnly:
                    params.add("Item Number")
            except:
                pass
        
        return params
    
    @staticmethod
    def get_group_key(elem, group_mode, order_parameter=None):
        """Get the group key (Category, Family, Type name, or Parameter value) for an element"""
        key = "Other"
        try:
            if group_mode == "By Category":
                key = elem.Category.Name if elem.Category else "Other"
            elif group_mode == "By Family":
                type_id = elem.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        if hasattr(elem_type, 'FamilyName'):
                            family_name = elem_type.FamilyName
                            if family_name:
                                key = family_name
                            else:
                                key = elem.Category.Name if elem.Category else "Other"
                        elif elem.Category:
                            key = elem.Category.Name
                        else:
                            key = "Other"
                    else:
                        key = elem.Category.Name if elem.Category else "Other"
                elif elem.Category:
                    key = elem.Category.Name
                else:
                    key = "Other"
            elif group_mode == "By Type":
                type_id = elem.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        family_name = ""
                        if hasattr(elem_type, 'FamilyName'):
                            family_name = elem_type.FamilyName or ""
                        elif elem.Category:
                            family_name = elem.Category.Name or ""
                        
                        type_name = elem_type.Name if hasattr(elem_type, 'Name') else ""
                        key = "{}:{}".format(family_name, type_name) if family_name or type_name else "Other"
                    else:
                        key = "Other"
                else:
                    key = "Other"
            elif group_mode == "Order by Parameter":
                if not order_parameter:
                    key = "Other"
                else:
                    val = ParameterHelper.get_value(elem, order_parameter)
                    if val is not None and str(val).strip():
                        key = str(val).strip()
                    else:
                        key = "(empty)"
        except:
            key = "Other"
        return key

    @staticmethod
    def get_common_parameters(elements):
        """Get all parameters that exist in ANY element, plus special pseudo-parameters"""
        if not elements:
            return []
        
        all_params = set()
        has_assembly = False
        has_fabrication = False
        has_mep = False
        
        for elem in elements:
            try:
                # Get standard parameters from element
                for p in elem.Parameters:
                    if p and p.Definition:
                        name = p.Definition.Name
                        if name and name.strip():
                            all_params.add(name)
            except:
                pass
            
            # Check for fabrication part
            if FabricationHelper.is_fabrication_part(elem):
                has_fabrication = True
                try:
                    param = elem.get_Parameter(DB.BuiltInParameter.FABRICATION_PART_ITEM_NUMBER)
                    if param:
                        all_params.add("Item Number")
                except:
                    pass

            # Check for assembly
            try:
                if elem.AssemblyInstanceId != DB.ElementId.InvalidElementId:
                    has_assembly = True
            except:
                pass
            
            # Check for MEP elements with connectors
            try:
                cm = None
                if hasattr(elem, "ConnectorManager"): 
                    cm = elem.ConnectorManager
                elif hasattr(elem, "MEPModel") and elem.MEPModel: 
                    cm = elem.MEPModel.ConnectorManager
                if cm and cm.Connectors and cm.Connectors.Size > 0:
                    has_mep = True
            except:
                pass
        
        # Add special pseudo-parameters that always work via get_value
        all_params.add("Category")
        all_params.add("Family")
        all_params.add("Type")
        all_params.add("Family and Type")
        
        # Add assembly mark if any element is in assembly
        if has_assembly:
            all_params.add(ASSEMBLY_MARK_PARAM)
        
        # Add connector-based parameters if MEP elements present
        if has_mep:
            all_params.add("Connectors")
            all_params.add("EndConnector1")
            all_params.add("EndConnector2")
        
        # Add fabrication-specific properties if fab parts present
        if has_fabrication:
            fab_props = ["Alias", "Angle", "Cid", "CutType", "Extension", 
                        "FabricationNotes", "Gauge", "Offset", "Service", 
                        "SheetMetalArea", "Specification", "CenterlineLength"]
            for prop in fab_props:
                all_params.add(prop)
            
        return sorted(list(all_params), key=lambda s: s.lower())
    
    @staticmethod
    def set_value(element, param_name, value_text):
        if not param_name:
            return True
        
        if param_name == ASSEMBLY_MARK_PARAM:
            try:
                if element.AssemblyInstanceId != DB.ElementId.InvalidElementId:
                    assembly_instance = doc.GetElement(element.AssemblyInstanceId)
                    if not assembly_instance:
                        return False
                    param = assembly_instance.get_Parameter(DB.BuiltInParameter.ASSEMBLY_MARK)
                    if param and not param.IsReadOnly:
                        param.Set(value_text)
                        return True
                return False
            except:
                return False

        # Special handling for Item Number on Fabrication Parts
        if param_name == "Item Number":
            try:
                # Try built-in FABRICATION_PART_ITEM_NUMBER first
                param = element.get_Parameter(DB.BuiltInParameter.FABRICATION_PART_ITEM_NUMBER)
                if param and not param.IsReadOnly:
                    param.Set(value_text)
                    return True
            except:
                pass
            # Fall through to normal LookupParameter if not a fabrication part

        try:
            param = element.LookupParameter(param_name)
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
    def _get_std_param(element, param_name):
        """Helper to get standard parameter string value"""
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
    def get_value(element, param_name):
        """Get parameter value as string"""
        if not param_name:
            return ""
        
        # --- Standard Special Cases ---
        if param_name.lower() == "category":
            try:
                return element.Category.Name if element.Category else ""
            except:
                return ""
        
        param_lower_clean = " ".join(param_name.lower().split())
        
        if param_lower_clean == "family and type" or param_lower_clean == "familyinfo":
            try:
                type_id = element.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        family_name = ""
                        if hasattr(elem_type, 'FamilyName'):
                            try: family_name = str(elem_type.FamilyName)
                            except: pass
                        if not family_name and element.Category:
                            try: family_name = str(element.Category.Name)
                            except: pass
                        
                        type_name = ""
                        try:
                            if hasattr(elem_type, 'Name'): type_name = str(elem_type.Name)
                        except: pass
                        if not type_name:
                            try: type_name = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                            except: pass
                        
                        return "{}:{}".format(family_name or "", type_name or "")
                return ""
            except:
                return ""
        
        if param_lower_clean == "family":
            try:
                type_id = element.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type and hasattr(elem_type, 'FamilyName'):
                        return elem_type.FamilyName
                    elif element.Category:
                        return element.Category.Name
                return ""
            except:
                return ""
        
        if param_lower_clean == "type":
            try:
                type_id = element.GetTypeId()
                if type_id and type_id != DB.ElementId.InvalidElementId:
                    elem_type = doc.GetElement(type_id)
                    if elem_type:
                        try:
                            if hasattr(elem_type, 'Name'): return str(elem_type.Name)
                        except: pass
                        try:
                            return elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                        except: pass
                return ""
            except:
                return ""
        
        if param_name == ASSEMBLY_MARK_PARAM:
            try:
                if element.AssemblyInstanceId != DB.ElementId.InvalidElementId:
                    assembly_instance = doc.GetElement(element.AssemblyInstanceId)
                    if assembly_instance:
                        param = assembly_instance.get_Parameter(DB.BuiltInParameter.ASSEMBLY_MARK)
                        if param: return param.AsString() or ""
                return ""
            except:
                return ""
        
        # --- FABRICATION SPECIFIC LOGIC ---
        if FabricationHelper.is_fabrication_part(element):
            try:
                # Item Number - use built-in parameter
                if param_name == "Item Number":
                    param = element.get_Parameter(DB.BuiltInParameter.FABRICATION_PART_ITEM_NUMBER)
                    if param:
                        val = param.AsString()
                        if val: return val
                        # Try AsValueString as fallback
                        val = param.AsValueString()
                        if val: return val
                    # Also try standard lookup as fallback
                    val = ParameterHelper._get_std_param(element, "Item Number")
                    if val: return val
                
                # 1. Length Smart Fallback
                if param_name == "Length" or param_name == "CenterlineLength":
                    # Try explicit property first (most accurate for Fab parts)
                    if hasattr(element, "CenterlineLength"):
                        val = str(element.CenterlineLength)
                        if val and val != "0.0" and val != "0": return val
                    
                    # Try standard parameter
                    val = ParameterHelper._get_std_param(element, "Length")
                    if val and val != "0.0" and val != "0": return val
                    
                    val = ParameterHelper._get_std_param(element, "Centerline Length")
                    if val: return val

                # 2. Weight Smart Fallback
                if param_name == "Weight":
                    # Try property
                    if hasattr(element, "Weight"):
                        val = str(element.Weight)
                        if val and val != "0.0" and val != "0": return val
                    # Try params
                    val = ParameterHelper._get_std_param(element, "Weight")
                    if val: return val
                    val = ParameterHelper._get_std_param(element, "Item Weight")
                    if val: return val

                # 3. Sheet Metal Area Smart Fallback
                if param_name == "SheetMetalArea":
                    if hasattr(element, "SheetMetalArea"):
                        return str(element.SheetMetalArea)
                    if hasattr(element, "SurfaceArea"):
                        return str(element.SurfaceArea)
                    val = ParameterHelper._get_std_param(element, "Sheet Metal Area")
                    if val: return val

                # 4. Connectors
                if param_name == "Connectors":
                    try:
                        connector_info = []
                        cm = None
                        if hasattr(element, "ConnectorManager"): cm = element.ConnectorManager
                        elif hasattr(element, "MEPModel") and element.MEPModel: cm = element.MEPModel.ConnectorManager
                             
                        if cm and cm.Connectors:
                            for conn in cm.Connectors:
                                if conn:
                                    desc = ""
                                    if hasattr(conn, "Description"): desc = conn.Description
                                    
                                    size = ""
                                    if hasattr(conn, 'Radius') and conn.Radius > 0.001:
                                        size = "{:.2f}".format(conn.Radius * 2)
                                    elif hasattr(conn, 'Width') and hasattr(conn, 'Height'):
                                        size = "{}x{}".format(conn.Width, conn.Height)
                                        
                                    if desc and size: connector_info.append("{} ({})".format(desc, size))
                                    elif desc: connector_info.append(desc)
                                    elif size: connector_info.append(size)
                        
                        if connector_info:
                            connector_info.sort()
                            return ", ".join(connector_info)
                    except:
                        pass
                
                # 5. End Connectors
                if param_name in ["EndConnector1", "EndConnector2"]:
                    try:
                        cm = None
                        if hasattr(element, "ConnectorManager"): cm = element.ConnectorManager
                        elif hasattr(element, "MEPModel") and element.MEPModel: cm = element.MEPModel.ConnectorManager
                        
                        if cm and cm.Connectors:
                            connectors = [c for c in cm.Connectors]
                            idx = 0 if param_name == "EndConnector1" else 1
                            if idx < len(connectors):
                                conn = connectors[idx]
                                desc = getattr(conn, "Description", "")
                                size = ""
                                if hasattr(conn, 'Radius') and conn.Radius > 0.001:
                                    size = "{:.3f}".format(conn.Radius * 2)
                                elif hasattr(conn, 'Width'):
                                     size = "{}x{}".format(conn.Width, conn.Height)
                                
                                val = "{} ({})".format(desc, size).strip()
                                if val.endswith("()"): val = val[:-2].strip()
                                return val
                    except:
                        pass

                # 6. Other Fab Properties
                if param_name == "Extension":
                    if hasattr(element, 'Extension'): return str(element.Extension)
                
                # Cid - uses ItemCustomId property
                if param_name == "Cid":
                    if hasattr(element, 'ItemCustomId'): 
                        return str(element.ItemCustomId)
                
                # FabricationNotes - check for Notes property or parameter
                if param_name == "FabricationNotes":
                    # Try property first
                    if hasattr(element, 'Notes'):
                        val = element.Notes
                        if val: return str(val)
                    # Try parameter lookup
                    val = ParameterHelper._get_std_param(element, "Fabrication Notes")
                    if val: return val
                    val = ParameterHelper._get_std_param(element, "Notes")
                    if val: return val
                
                for p_check in ["Gauge", "Offset", "Alias", "Angle", "CutType", "Service", "Specification"]:
                    if param_name == p_check and hasattr(element, p_check):
                        val = getattr(element, p_check)
                        if val is not None: return str(val)

            except:
                pass

        # --- Default Parameter Lookup ---
        return ParameterHelper._get_std_param(element, param_name)
    
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
        
        if not has_non_empty:
            sig_values.append(("_ElementId_", str(element.Id.IntegerValue)))
        
        if DEBUG_MODE:
            # Shorten for log readability
            debug_sig = []
            for k, v in sig_values:
                if v: debug_sig.append((k, v))
            log_debug("Elem {} Sig: {}".format(element.Id, debug_sig))

        return tuple(sig_values)


# =============================================================================
# ELEMENT SELECTION
# =============================================================================

class ElementPicker(object):
    @staticmethod
    def pick_multiple():
        from Autodesk.Revit.UI.Selection import ObjectType
        picked = []
        try:
            refs = uidoc.Selection.PickObjects(
                ObjectType.Element,
                "Select elements to renumber (Finish or ESC when done)"
            )
            for ref in refs:
                elem = doc.GetElement(ref.ElementId)
                if elem: picked.append(elem)
        except:
            pass
        return picked
    
    @staticmethod
    def pick_in_order():
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.Exceptions import OperationCanceledException
        
        sequence = []
        picked_ids = []
        
        try:
            empty_list = List[DB.ElementId]()
            uidoc.Selection.SetElementIds(empty_list)
        except:
            pass
        
        pick_number = 1
        while True:
            try:
                prompt = "Pick element #{} (Press ESC when finished)".format(pick_number)
                ref = uidoc.Selection.PickObject(ObjectType.Element, prompt)
                
                try:
                    elem = doc.GetElement(ref.ElementId)
                    if not elem: continue
                except: continue
                
                elem_id_int = elem.Id.IntegerValue
                already_picked = False
                for e in sequence:
                    if e.Id.IntegerValue == elem_id_int:
                        already_picked = True
                        break
                
                if already_picked: continue
                
                sequence.append(elem)
                picked_ids.append(elem.Id)
                pick_number += 1
                
                try:
                    id_collection = List[DB.ElementId]()
                    for eid in picked_ids: id_collection.Add(eid)
                    uidoc.Selection.SetElementIds(id_collection)
                except: pass
                
            except OperationCanceledException:
                break
            except:
                continue
        return sequence
    
    @staticmethod
    def pick_curve():
        from Autodesk.Revit.UI.Selection import ObjectType
        from Autodesk.Revit.Exceptions import OperationCanceledException
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, "Pick a curve/spline for ordering")
            elem = doc.GetElement(ref.ElementId)
            if isinstance(elem, DB.CurveElement): return elem
            else:
                forms.alert("Please select a CurveElement (model line, arc, spline, etc.)", title="Renumber")
                return None
        except OperationCanceledException: return None
        except: return None
    
    @staticmethod
    def collect_nested_families(picked_elements, include_nested):
        """
        Expands selection to include nested family instances if enabled.
        Returns expanded list of elements including nested components.
        """
        if not include_nested:
            return picked_elements
        
        # Keep track of original element IDs to preserve them
        original_ids = set(elem.Id.IntegerValue for elem in picked_elements)
        expanded_ids = set(original_ids)
        host_ids = set(original_ids)
        
        # Collect nested components recursively (only for FamilyInstances)
        for elem in picked_elements:
            if isinstance(elem, DB.FamilyInstance):
                ElementPicker._collect_nested_subcomponents(elem, expanded_ids)
        
        # Collect children via SuperComponent
        ElementPicker._collect_children_by_supercomponent(host_ids, expanded_ids)
        
        # Convert back to element list - KEEP ALL ORIGINAL ELEMENTS plus nested FamilyInstances
        result = []
        for eid_int in expanded_ids:
            elem = doc.GetElement(DB.ElementId(eid_int))
            if elem:
                # Keep if it was originally picked OR if it's a nested FamilyInstance
                if eid_int in original_ids or isinstance(elem, DB.FamilyInstance):
                    result.append(elem)
        
        return result
    
    @staticmethod
    def _collect_nested_subcomponents(host_fi, out_ids):
        """Recursively collect nested components using GetSubComponentIds"""
        try:
            sub_ids = list(host_fi.GetSubComponentIds())
        except:
            sub_ids = []
        
        for sid in sub_ids:
            sub_el = doc.GetElement(sid)
            if isinstance(sub_el, DB.FamilyInstance):
                eid = sub_el.Id.IntegerValue
                if eid not in out_ids:
                    out_ids.add(eid)
                    ElementPicker._collect_nested_subcomponents(sub_el, out_ids)
    
    @staticmethod
    def _collect_children_by_supercomponent(host_ids_int, out_ids):
        """Collect child instances via SuperComponent property"""
        try:
            from Autodesk.Revit.DB import FilteredElementCollector, FamilyInstance
            for fi in FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements():
                try:
                    sc = fi.SuperComponent
                except:
                    sc = None
                
                if sc and sc.Id and sc.Id.IntegerValue in host_ids_int:
                    out_ids.add(fi.Id.IntegerValue)
        except:
            pass


# =============================================================================
# FABRICATION PART UTILITIES
# =============================================================================

class FabricationHelper(object):
    @staticmethod
    def is_fabrication_part(element):
        try:
            from Autodesk.Revit.DB.Fabrication import FabricationPart
            if isinstance(element, FabricationPart): return True
        except: pass
        return False
    
    @staticmethod
    def get_size_metric(element):
        if not FabricationHelper.is_fabrication_part(element): return None, ""
        
        def safe_float(value):
            try: return float(value)
            except: return None
        
        for attr in ("Area", "CrossSectionalArea", "SectionArea", "CrossSectionArea"):
            try:
                val = getattr(element, attr, None)
                area = safe_float(val)
                if area is not None and area > 0: return area, "Profile Area"
            except: pass
        
        try:
            area_param = element.LookupParameter("Area")
            if area_param:
                area = area_param.AsDouble()
                if area and area > 0: return area, "Profile Area"
        except: pass
        
        diameter = None
        for attr in ("Diameter", "NominalDiameter"):
            try:
                val = getattr(element, attr, None)
                diameter = safe_float(val)
                if diameter is not None and diameter > 0: break
            except: pass
        
        if diameter is None:
            try:
                size_param = element.LookupParameter("Size")
                if size_param:
                    size_str = size_param.AsString()
                    if size_str:
                        import re
                        numbers = re.findall(r'\d+\.?\d*', size_str)
                        if numbers: diameter = float(numbers[0])
            except: pass
        
        if diameter and diameter > 0:
            area = math.pi * (diameter / 2.0) ** 2
            return area, "Profile Area (πd²/4)"
        
        width = None
        height = None
        
        for attr in ("Width", "DimA"):
            try:
                val = getattr(element, attr, None)
                width = safe_float(val)
                if width is not None and width > 0: break
            except: pass
        
        for attr in ("Height", "DimB"):
            try:
                val = getattr(element, attr, None)
                height = safe_float(val)
                if height is not None and height > 0: break
            except: pass
        
        if width and height and width > 0 and height > 0:
            return width * height, "Profile Area (w×h)"
        if width and width > 0:
            return width * width, "Profile Area (square)"
        return None, ""


# =============================================================================
# GEOMETRY & ORDERING
# =============================================================================

class GeometryHelper(object):
    @staticmethod
    def get_bounding_box_info(element, view=None):
        bbox = None
        try: bbox = element.get_BoundingBox(view)
        except: pass
        
        if not bbox:
            try: bbox = element.get_BoundingBox(None)
            except: pass
        
        if not bbox: return 0.0, 0.0, DB.XYZ(0, 0, 0)
        
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
        try:
            location = element.Location
            if location:
                if hasattr(location, "Point"):
                    pt = location.Point
                    if isinstance(pt, DB.XYZ): return pt
                if hasattr(location, "Curve"):
                    curve = location.Curve
                    if curve:
                        try: return curve.Evaluate(0.5, True)
                        except: pass
        except: pass
        return GeometryHelper.get_bounding_box_info(element, view)[2]
    
    @staticmethod
    def project_to_view_plane(point, view):
        if not isinstance(view, DB.View): return point.X, point.Y
        try:
            origin = view.Origin
            right = view.RightDirection
            up = view.UpDirection
            
            vector = point - origin
            u = vector.DotProduct(right)
            v = vector.DotProduct(up)
            return u, v
        except: return point.X, point.Y


class ElementSorter(object):
    @staticmethod
    def by_size(elements, view, reverse=False):
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
        if not isinstance(view, DB.View): view = revit.active_view
        
        points = []
        for elem in elements:
            pt3d = GeometryHelper.get_anchor_point(elem, view)
            u, v = GeometryHelper.project_to_view_plane(pt3d, view)
            points.append((u, v, elem))
        
        if not points: return elements
        
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
        try: curve = curve_element.GeometryCurve
        except: curve = None
        
        if curve is None: return elements
        
        sorted_items = []
        for elem in elements:
            point = GeometryHelper.get_anchor_point(elem, view)
            try:
                result = curve.Project(point)
                parameter = result.Parameter if result else 0.0
            except: parameter = 0.0
            sorted_items.append((parameter, elem))
        
        sorted_items.sort(key=lambda x: x[0])
        return [elem for _, elem in sorted_items]
    
    @staticmethod
    def by_parameter(elements, parameter_name, ascending=True):
        sorted_items = []
        for elem in elements:
            value = ParameterHelper.get_value(elem, parameter_name)
            try:
                numeric_value = float(value) if value else 0.0
                sorted_items.append((numeric_value, elem))
            except (ValueError, TypeError):
                str_value = str(value) if value else ""
                sorted_items.append((str_value, elem))
        sorted_items.sort(key=lambda x: x[0], reverse=not ascending)
        return [elem for _, elem in sorted_items]


# =============================================================================
# TEXT NOTE CREATION
# =============================================================================

class TextNoteCreator(object):
    @staticmethod
    def create(view, text, location):
        try:
            text_type = DB.FilteredElementCollector(doc)\
                .OfClass(DB.TextNoteType)\
                .FirstElement()
            if not text_type: return None
            return DB.TextNote.Create(doc, view.Id, location, text, text_type.Id)
        except: return None


# =============================================================================
# RENUMBERING ENGINE
# =============================================================================

class RenumberEngine(object):
    @staticmethod
    def execute(elements, settings, picked_curve=None, manual_order=None, custom_group_order=None):
        targets = [t for t in [settings.target_a, settings.target_b, settings.target_c] if t]
        if not targets:
            forms.alert("Please select at least one target parameter", title="Renumber")
            return 0, 0, []
        
        if not elements:
            forms.alert("No elements selected", title="Renumber")
            return 0, 0, []
        
        try: start_num = int(settings.start_number)
        except: start_num = 1
        
        if settings.specify_order:
            if not manual_order or len(manual_order) == 0:
                forms.alert("Manual pick order is enabled but no order was picked.", title="Renumber")
                return 0, 0, []
            filtered_manual_order_ids = set([e.Id for e in elements])
            final_ordered_list = [e for e in manual_order if e.Id in filtered_manual_order_ids]
        else:
            grouped_lists = RenumberEngine._get_grouped_elements(
                elements,
                settings.grouping_mode,
                custom_group_order or [],
                settings.order_parameter
            )
            final_ordered_list = []
            for group in grouped_lists:
                ordered_group = RenumberEngine._get_ordered_elements(group, settings, picked_curve)
                final_ordered_list.extend(ordered_group)

        if not final_ordered_list: return 0, 0, []
        
        compare_fields = []
        if settings.same_for_identical:
            # Combine fab and common compare fields
            compare_fields = [f.Name for f in settings.fab_compare_fields if f.Include]
            compare_fields.extend([f.Name for f in settings.common_compare_fields if f.Include])
            for param in [settings.compare_param_1, settings.compare_param_2, settings.compare_param_3, settings.compare_param_4]:
                if param and param.strip(): compare_fields.append(param)

            if not compare_fields:
                forms.toast("No compare fields selected. Enabling 'Same number for identical' has no effect.", title="Warning", appid="renumber")
        
        return RenumberEngine._apply_numbering(
            final_ordered_list, settings, targets, start_num, compare_fields
        )

    @staticmethod
    def _get_grouped_elements(elements, group_mode, custom_order, order_parameter=None):
        if group_mode == "All at Once": return [list(elements)]
        
        groups = {}
        for elem in elements:
            key = ParameterHelper.get_group_key(elem, group_mode, order_parameter)
            if key not in groups: groups[key] = []
            groups[key].append(elem)
            
        final_list_of_lists = []
        processed_keys = set()

        for key in custom_order:
            if key in groups:
                final_list_of_lists.append(groups[key])
                processed_keys.add(key)
        
        for key in sorted(groups.keys()):
            if key not in processed_keys:
                final_list_of_lists.append(groups[key])
        
        return final_list_of_lists

    @staticmethod
    def _get_ordered_elements(elements, settings, picked_curve):
        ordered = list(elements)
        view = revit.active_view
        mode = settings.order_mode
        
        if mode == "Largest -> Smallest": return ElementSorter.by_size(ordered, view, reverse=True)
        elif mode == "Smallest -> Largest": return ElementSorter.by_size(ordered, view, reverse=False)
        elif mode == "Clockwise": return ElementSorter.by_angle(ordered, view, counterclockwise=False)
        elif mode == "Counterclockwise": return ElementSorter.by_angle(ordered, view, counterclockwise=True)
        elif mode == "Follow spline":
            if not picked_curve:
                forms.alert("Please pick a curve first", title="Renumber")
                return []
            return ElementSorter.by_curve(ordered, picked_curve, view)
        return ordered
    
    @staticmethod
    def _apply_numbering(ordered_elements, settings, targets, start_num, compare_fields):
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
                            num_str = str(current_num).zfill(settings.zero_pad_digits) if settings.zero_pad_digits > 0 else str(current_num)
                            number_text = "{}{}{}".format(settings.prefix, num_str, settings.suffix)
                            signature_map[sig] = number_text
                            current_num += 1
                    else:
                        num_str = str(current_num).zfill(settings.zero_pad_digits) if settings.zero_pad_digits > 0 else str(current_num)
                        number_text = "{}{}{}".format(settings.prefix, num_str, settings.suffix)
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
                                    pt = uidoc.Selection.PickPoint("Pick location for '{}'".format(number_text))
                                else:
                                    pt = GeometryHelper.get_anchor_point(elem, revit.active_view)
                                TextNoteCreator.create(revit.active_view, number_text, pt)
                            except: pass
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
    def __init__(self, xaml_path, preloaded_elements=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self.settings = ConfigManager.load()
        
        # Expand to include nested families if setting is enabled
        initial_elements = preloaded_elements or []
        expanded_elements = ElementPicker.collect_nested_families(
            initial_elements, 
            self.settings.include_nested_families
        )
        
        self.elements = expanded_elements
        self.all_picked_elements = list(self.elements)
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
            self._update_element_count()
            self._refresh_parameters()
        else:
            self.ok_button.IsEnabled = False
            self._update_group_order_list()
        
        self._update_filter_box_visibility()
    
    def _get_controls(self):
        self.picked_count_label = self.FindName("PickedCountLbl")
        self.pick_button = self.FindName("PickElemsBtn")
        self.save_button = self.FindName("SaveProfileBtn")
        self.load_button = self.FindName("LoadProfileBtn")
        self.target_a_combo = self.FindName("TargetParamBoxA")
        self.target_b_combo = self.FindName("TargetParamBoxB")
        self.target_c_combo = self.FindName("TargetParamBoxC")
        self.filter_box = self.FindName("FilterBox")
        self.filter_param_combo = self.FindName("FilterParamBox")
        self.filter_value_box = self.FindName("FilterValueBox")
        self.filter_apply_btn = self.FindName("FilterApplyBtn")
        self.filter_status_label = self.FindName("FilterStatusLbl")
        self.prefix_box = self.FindName("PrefixBox")
        self.suffix_box = self.FindName("SuffixBox")
        self.start_box = self.FindName("StartNumberBox")
        self.zero_pad_box = self.FindName("ZeroPadBox")
        self.order_combo = self.FindName("OrderModeBox")
        self.grouping_mode_combo = self.FindName("GroupingModeBox")
        self.group_order_box = self.FindName("GroupOrderBox")
        self.group_order_list = self.FindName("GroupOrderList")
        self.move_up_btn = self.FindName("MoveUpBtn")
        self.move_down_btn = self.FindName("MoveDownBtn")
        self.curve_label = self.FindName("CurvePickedLabel")
        self.curve_button = self.FindName("PickCurveBtn")
        self.order_param_box = self.FindName("OrderByParameterBox")
        self.order_param_combo = self.FindName("OrderParameterBox")
        self.place_tags_check = self.FindName("PlaceTagsChk")
        self.prompt_pos_check = self.FindName("PromptPosChk")
        self.same_ident_check = self.FindName("SameIdentChk")
        self.include_nested_check = self.FindName("IncludeNestedChk")
        self.specify_order_check = self.FindName("SpecifyOrderChk")
        self.order_status_label = self.FindName("OrderStatusLabel")
        self.pick_order_button = self.FindName("PickOrderBtn")
        self.options_list = self.FindName("OptionsList")  # Keep for backward compatibility
        self.fab_options_list = self.FindName("FabOptionsList")
        self.common_options_list = self.FindName("CommonOptionsList")
        self.check_all_button = self.FindName("CheckAllBtn")
        self.check_none_button = self.FindName("CheckNoneBtn")
        self.compare_param_1_combo = self.FindName("CompareParamBox1")
        self.compare_param_2_combo = self.FindName("CompareParamBox2")
        self.compare_param_3_combo = self.FindName("CompareParamBox3")
        self.compare_param_4_combo = self.FindName("CompareParamBox4")
        self.deselect_individual_radio = self.FindName("DeselectIndividualRadio")
        self.deselect_category_radio = self.FindName("DeselectCategoryRadio")
        self.deselect_family_radio = self.FindName("DeselectFamilyRadio")
        self.deselect_type_radio = self.FindName("DeselectTypeRadio")
        self.deselect_parameter_radio = self.FindName("DeselectParameterRadio")
        self.deselect_parameter_controls = self.FindName("DeselectParameterControls")
        self.deselect_parameter_combo = self.FindName("DeselectParameterCombo")
        self.deselect_items_list = self.FindName("DeselectItemsList")
        self.ok_button = self.FindName("OKBtn")
        self.cancel_button = self.FindName("CancelBtn")
    
    def _wire_events(self):
        if self.order_combo: self.order_combo.SelectionChanged += self._on_order_changed
        if self.grouping_mode_combo: self.grouping_mode_combo.SelectionChanged += self._on_grouping_mode_changed
        if self.specify_order_check:
            self.specify_order_check.Checked += self._on_specify_order_changed
            self.specify_order_check.Unchecked += self._on_specify_order_changed
        # MoveUp/MoveDown buttons are wired in XAML via Click="OnMoveUp"/Click="OnMoveDown"
        if self.filter_apply_btn: self.filter_apply_btn.Click += self.OnFilterApply
        if self.filter_param_combo: self.filter_param_combo.SelectionChanged += self._on_filter_param_changed
        if self.order_param_combo:
            self.order_param_combo.SelectionChanged += self._on_order_parameter_changed
            self.order_param_combo.LostFocus += self._on_order_parameter_changed
        if self.deselect_parameter_combo: self.deselect_parameter_combo.SelectionChanged += self._on_deselect_parameter_changed
    
    def _initialize_ui(self):
        if self.order_combo:
            self.order_combo.ItemsSource = self.settings.order_modes
            self.order_combo.SelectedItem = self.settings.order_mode
        if self.grouping_mode_combo:
            self.grouping_mode_combo.ItemsSource = self.settings.grouping_modes
            self.grouping_mode_combo.SelectedItem = self.settings.grouping_mode
        if self.group_order_list:
            self.group_order_list.ItemsSource = self.settings.group_order_items
        if self.fab_options_list:
            self.fab_options_list.ItemsSource = self.settings.fab_compare_fields
        if self.common_options_list:
            self.common_options_list.ItemsSource = self.settings.common_compare_fields
        if self.prefix_box: self.prefix_box.Text = self.settings.prefix
        if self.suffix_box: self.suffix_box.Text = self.settings.suffix
        if self.start_box: self.start_box.Text = self.settings.start_number
        if self.zero_pad_box: self.zero_pad_box.Text = str(self.settings.zero_pad_digits) if self.settings.zero_pad_digits > 0 else ""
        if self.place_tags_check: self.place_tags_check.IsChecked = self.settings.place_tags
        if self.prompt_pos_check: self.prompt_pos_check.IsChecked = self.settings.prompt_for_positions
        if self.same_ident_check: self.same_ident_check.IsChecked = self.settings.same_for_identical
        if self.include_nested_check: self.include_nested_check.IsChecked = self.settings.include_nested_families
        if self.specify_order_check: self.specify_order_check.IsChecked = self.settings.specify_order
        
        for combo in [self.target_a_combo, self.target_b_combo, self.target_c_combo, 
                      self.compare_param_1_combo, self.compare_param_2_combo, self.compare_param_3_combo, self.compare_param_4_combo,
                      self.filter_param_combo, self.order_param_combo, self.deselect_parameter_combo]:
            if combo: combo.ItemsSource = self.settings.available_params
        
        if self.target_a_combo: self.target_a_combo.Text = self.settings.target_a
        if self.target_b_combo: self.target_b_combo.Text = self.settings.target_b
        if self.target_c_combo: self.target_c_combo.Text = self.settings.target_c
        
        if self.compare_param_1_combo: self.compare_param_1_combo.Text = self.settings.compare_param_1
        if self.compare_param_2_combo: self.compare_param_2_combo.Text = self.settings.compare_param_2
        if self.compare_param_3_combo: self.compare_param_3_combo.Text = self.settings.compare_param_3
        if self.compare_param_4_combo: self.compare_param_4_combo.Text = self.settings.compare_param_4
        
        if self.order_param_combo: self.order_param_combo.Text = self.settings.order_parameter
        
        if self.pick_order_button:
            is_manual = self.specify_order_check.IsChecked if self.specify_order_check else False
            self.pick_order_button.IsEnabled = is_manual
        
        if self.order_status_label:
            if len(self.manual_order) > 0: self.order_status_label.Text = "Order: {} elements".format(len(self.manual_order))
            else: self.order_status_label.Text = "(not set)"
        
        if self.order_combo and self.specify_order_check:
            self.order_combo.IsEnabled = not self.specify_order_check.IsChecked
        
        self._on_order_changed(None, None)
        self._on_specify_order_changed(None, None)
        self._update_filter_box_visibility()
        self._refresh_deselect_list()
    
    def _update_element_count(self):
        if self.picked_count_label:
            total_count = len(self.all_picked_elements)
            filtered_count = len(self.elements)
            if total_count == 0: self.picked_count_label.Content = "(none)"
            elif total_count == filtered_count: self.picked_count_label.Content = "{} elements".format(total_count)
            else: self.picked_count_label.Content = "{}/{} elements".format(filtered_count, total_count)
    
    def _refresh_parameters(self):
        params = ParameterHelper.get_common_parameters(self.all_picked_elements) 
        current_filter_param = self.filter_param_combo.SelectedItem
        current_filter_value = (self.filter_value_box.Text or "").strip() if self.filter_value_box else ""
        current_deselect_param = self.deselect_parameter_combo.SelectedItem if self.deselect_parameter_combo else None
        self.settings.available_params.Clear()
        for param in params: self.settings.available_params.Add(param)
        if current_filter_param in self.settings.available_params:
            self.filter_param_combo.SelectedItem = current_filter_param
            # Repopulate filter values for the current parameter
            self._populate_filter_values()
            # Restore the filter value if one was set
            if current_filter_value and self.filter_value_box:
                self.filter_value_box.Text = current_filter_value
        if current_deselect_param and current_deselect_param in self.settings.available_params:
            self.deselect_parameter_combo.SelectedItem = current_deselect_param
        self._update_group_order_list() 

    def _update_filter_box_visibility(self):
        if self.filter_box:
            if self.all_picked_elements and len(self.all_picked_elements) > 0:
                self.filter_box.Visibility = System.Windows.Visibility.Visible
            else: self.filter_box.Visibility = System.Windows.Visibility.Collapsed

    def _read_ui_to_settings(self):
        if self.prefix_box: self.settings.prefix = self.prefix_box.Text or ""
        if self.suffix_box: self.settings.suffix = self.suffix_box.Text or ""
        if self.start_box: self.settings.start_number = self.start_box.Text or "1"
        if self.zero_pad_box:
            try:
                self.settings.zero_pad_digits = int(self.zero_pad_box.Text) if self.zero_pad_box.Text.strip() else 0
            except:
                self.settings.zero_pad_digits = 0
        if self.order_combo: self.settings.order_mode = self.order_combo.SelectedItem or ORDER_MODES[0]
        if self.grouping_mode_combo: self.settings.grouping_mode = self.grouping_mode_combo.SelectedItem or GROUPING_MODES[0]
        if self.place_tags_check: self.settings.place_tags = bool(self.place_tags_check.IsChecked)
        if self.prompt_pos_check: self.settings.prompt_for_positions = bool(self.prompt_pos_check.IsChecked)
        if self.same_ident_check: self.settings.same_for_identical = bool(self.same_ident_check.IsChecked)
        if self.include_nested_check: self.settings.include_nested_families = bool(self.include_nested_check.IsChecked)
        if self.specify_order_check: self.settings.specify_order = bool(self.specify_order_check.IsChecked)
        if self.target_a_combo: self.settings.target_a = self.target_a_combo.Text or ""
        if self.target_b_combo: self.settings.target_b = self.target_b_combo.Text or ""
        if self.target_c_combo: self.settings.target_c = self.target_c_combo.Text or ""
        if self.compare_param_1_combo: self.settings.compare_param_1 = self.compare_param_1_combo.Text or ""
        if self.compare_param_2_combo: self.settings.compare_param_2 = self.compare_param_2_combo.Text or ""
        if self.compare_param_3_combo: self.settings.compare_param_3 = self.compare_param_3_combo.Text or ""
        if self.compare_param_4_combo: self.settings.compare_param_4 = self.compare_param_4_combo.Text or ""
        if self.order_param_combo: self.settings.order_parameter = self.order_param_combo.Text or ""
        if self.group_order_list and self.settings.group_order_items:
            self.settings.saved_group_order = ",".join([item.Name for item in self.settings.group_order_items])
    
    def _refresh_options_list(self):
        if self.fab_options_list:
            self.fab_options_list.ItemsSource = None
            self.fab_options_list.ItemsSource = self.settings.fab_compare_fields
        if self.common_options_list:
            self.common_options_list.ItemsSource = None
            self.common_options_list.ItemsSource = self.settings.common_compare_fields
    
    @staticmethod
    def transfer_settings_static(new_window, saved_settings, saved_curve):
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
        new_window.settings.saved_group_order = saved_settings.saved_group_order
        
        # Transfer fab compare fields
        saved_fab_lookup = {f.Name: f.Include for f in saved_settings.fab_compare_fields}
        for field in new_window.settings.fab_compare_fields:
            if field.Name in saved_fab_lookup:
                field.Include = saved_fab_lookup[field.Name]
        
        # Transfer common compare fields
        saved_common_lookup = {f.Name: f.Include for f in saved_settings.common_compare_fields}
        for field in new_window.settings.common_compare_fields:
            if field.Name in saved_common_lookup:
                field.Include = saved_common_lookup[field.Name]
        
        if saved_curve and not saved_settings.specify_order:
            new_window.picked_curve = saved_curve
            if hasattr(new_window, 'curve_label') and new_window.curve_label:
                new_window.curve_label.Text = "Curve Id: {}".format(saved_curve.Id.IntegerValue)
    
    # Handlers
    def OnPickElements(self, sender, args): self._handle_pick_elements()
    def OnSaveProfile(self, sender, args): self._handle_save_profile()
    def OnLoadProfile(self, sender, args): self._handle_load_profile()
    def OnPickCurve(self, sender, args): self._handle_pick_curve()
    def OnPickOrder(self, sender, args): self._handle_pick_order()
    def OnCheckAll(self, sender, args):
        for field in self.settings.fab_compare_fields: field.Include = True
        for field in self.settings.common_compare_fields: field.Include = True
        self._refresh_options_list()
    def OnCheckNone(self, sender, args):
        for field in self.settings.fab_compare_fields: field.Include = False
        for field in self.settings.common_compare_fields: field.Include = False
        self._refresh_options_list()
    def OnFabCheckAll(self, sender, args):
        for field in self.settings.fab_compare_fields: field.Include = True
        self._refresh_options_list()
    def OnFabCheckNone(self, sender, args):
        for field in self.settings.fab_compare_fields: field.Include = False
        self._refresh_options_list()
    def OnCommonCheckAll(self, sender, args):
        for field in self.settings.common_compare_fields: field.Include = True
        self._refresh_options_list()
    def OnCommonCheckNone(self, sender, args):
        for field in self.settings.common_compare_fields: field.Include = False
        self._refresh_options_list()
    def OnMoveUp(self, sender, args):
        if self.group_order_list and self.group_order_list.SelectedItem:
            idx = self.group_order_list.SelectedIndex
            if idx > 0:
                item = self.settings.group_order_items[idx]
                self.settings.group_order_items.RemoveAt(idx)
                self.settings.group_order_items.Insert(idx - 1, item)
                self.group_order_list.SelectedIndex = idx - 1
    def OnMoveDown(self, sender, args):
        if self.group_order_list and self.group_order_list.SelectedItem:
            idx = self.group_order_list.SelectedIndex
            if idx < len(self.settings.group_order_items) - 1:
                item = self.settings.group_order_items[idx]
                self.settings.group_order_items.RemoveAt(idx)
                self.settings.group_order_items.Insert(idx + 1, item)
                self.group_order_list.SelectedIndex = idx + 1
    def OnFilterApply(self, sender, args):
        param_name = self.filter_param_combo.SelectedItem
        value_to_match = (self.filter_value_box.Text or "").strip()
        
        if not param_name:
            self.elements = list(self.all_picked_elements)
            self.filter_status_label.Text = "(filter cleared)"
            self.filter_value_box.SelectedItem = None
            self.filter_value_box.Text = ""
            self.settings.filter_values.Clear()
        elif not value_to_match:
            # No value specified - clear filter
            self.elements = list(self.all_picked_elements)
            self.filter_status_label.Text = "(filter cleared)"
        else:
            new_filtered_list = []
            value_lower = value_to_match.lower()
            for elem in self.all_picked_elements:
                elem_val = ParameterHelper.get_value(elem, param_name)
                if elem_val is not None:
                    elem_val_clean = str(elem_val).strip().lower()
                    if elem_val_clean == value_lower:
                        new_filtered_list.append(elem)
            self.elements = new_filtered_list
            self.filter_status_label.Text = "Filtered {} items".format(len(self.elements))
        self._update_element_count()
        self._refresh_parameters()
        self.ok_button.IsEnabled = len(self.elements) > 0

    def OnDeselectModeChanged(self, sender, args): self._refresh_deselect_list()
    def OnDeselectSelected(self, sender, args):
        if not self.deselect_items_list: return
        selected = list(self.deselect_items_list.SelectedItems)
        if not selected: return
        if self.deselect_individual_radio.IsChecked: self._deselect_by_individual(selected)
        elif self.deselect_category_radio.IsChecked: self._deselect_by_category(selected)
        elif self.deselect_family_radio.IsChecked: self._deselect_by_family(selected)
        elif self.deselect_type_radio.IsChecked: self._deselect_by_type(selected)
        elif self.deselect_parameter_radio.IsChecked: self._deselect_by_parameter(selected)
    
    def OnDeselectAll(self, sender, args):
        if not self.elements: return
        count = len(self.elements)
        self.elements = []
        self.all_picked_elements = []
        self.manual_order = []
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        self.ok_button.IsEnabled = False
        forms.toast("Deselected all {} element(s)".format(count), title="Deselect All", appid="renumber")

    def _deselect_by_individual(self, selected_items):
        ids_to_remove = set()
        for item in selected_items:
            for eid in item.ElementIds: ids_to_remove.add(str(eid))
        self.elements = [e for e in self.elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self.all_picked_elements = [e for e in self.all_picked_elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        self.ok_button.IsEnabled = len(self.elements) > 0
    
    def _deselect_by_category(self, selected_items):
        cats_remove = set([item.Name for item in selected_items])
        self.elements = [e for e in self.elements if not (e.Category and e.Category.Name in cats_remove)]
        self.all_picked_elements = [e for e in self.all_picked_elements if not (e.Category and e.Category.Name in cats_remove)]
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        self.ok_button.IsEnabled = len(self.elements) > 0

    def _deselect_by_family(self, selected_items):
        fams_remove = set([item.Name for item in selected_items])
        self.elements = [e for e in self.elements if self._get_element_family(e) not in fams_remove]
        self.all_picked_elements = [e for e in self.all_picked_elements if self._get_element_family(e) not in fams_remove]
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        self.ok_button.IsEnabled = len(self.elements) > 0

    def _deselect_by_type(self, selected_items):
        types_remove = set([item.Name for item in selected_items])
        self.elements = [e for e in self.elements if self._get_element_type_name(e) not in types_remove]
        self.all_picked_elements = [e for e in self.all_picked_elements if self._get_element_type_name(e) not in types_remove]
        self._update_element_count()
        self._refresh_deselect_list()
        self._refresh_parameters()
        self.ok_button.IsEnabled = len(self.elements) > 0

    def _deselect_by_parameter(self, selected_items):
        param_name = self.deselect_parameter_combo.SelectedItem
        if not param_name: return
        ids_to_remove = set()
        for item in selected_items:
            for eid in item.ElementIds: ids_to_remove.add(str(eid))
        self.elements = [e for e in self.elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self.all_picked_elements = [e for e in self.all_picked_elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self._update_element_count()
        self._refresh_parameters()  # Refresh params first to preserve combo selection
        self._refresh_deselect_list()  # Then refresh list with preserved selection
        self.ok_button.IsEnabled = len(self.elements) > 0

    def _refresh_deselect_list(self):
        if not self.deselect_items_list: return
        if self.deselect_parameter_controls:
            self.deselect_parameter_controls.Visibility = System.Windows.Visibility.Visible if self.deselect_parameter_radio.IsChecked else System.Windows.Visibility.Collapsed
        if self.deselect_individual_radio.IsChecked: self._populate_individual_list()
        elif self.deselect_category_radio.IsChecked: self._populate_category_list()
        elif self.deselect_family_radio.IsChecked: self._populate_family_list()
        elif self.deselect_type_radio.IsChecked: self._populate_type_list()
        elif self.deselect_parameter_radio.IsChecked: self._populate_parameter_list()
        else: self._populate_individual_list()

    def _populate_individual_list(self):
        items = ObservableCollection[DeselectItem]()
        for elem in self.elements:
            elem_id = str(elem.Id.IntegerValue)
            category = elem.Category.Name if elem.Category else "Unknown"
            name = "Element {}".format(elem_id)
            try:
                np = elem.get_Parameter(DB.BuiltInParameter.ELEM_NAME_PARAM)
                if np and np.HasValue: name = np.AsString()
            except: pass
            display = "ID {} - {} - {}".format(elem_id, category, name)
            items.Add(DeselectItem(display, 1, [elem_id]))
        self.deselect_items_list.ItemsSource = items

    def _populate_category_list(self):
        cat_dict = {}
        for elem in self.elements:
            cn = elem.Category.Name if elem.Category else "No Category"
            if cn not in cat_dict: cat_dict[cn] = []
            cat_dict[cn].append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for cn in sorted(cat_dict.keys()): items.Add(DeselectItem(cn, len(cat_dict[cn]), cat_dict[cn]))
        self.deselect_items_list.ItemsSource = items

    def _populate_family_list(self):
        fam_dict = {}
        for elem in self.elements:
            fn = self._get_element_family(elem)
            if fn not in fam_dict: fam_dict[fn] = []
            fam_dict[fn].append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for fn in sorted(fam_dict.keys()): items.Add(DeselectItem(fn, len(fam_dict[fn]), fam_dict[fn]))
        self.deselect_items_list.ItemsSource = items

    def _populate_type_list(self):
        type_dict = {}
        for elem in self.elements:
            tn = self._get_element_type_name(elem)
            if tn not in type_dict: type_dict[tn] = []
            type_dict[tn].append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for tn in sorted(type_dict.keys()): items.Add(DeselectItem(tn, len(type_dict[tn]), type_dict[tn]))
        self.deselect_items_list.ItemsSource = items

    def _populate_parameter_list(self):
        param_name = self.deselect_parameter_combo.SelectedItem
        if not param_name:
            self.deselect_items_list.ItemsSource = ObservableCollection[DeselectItem]()
            return
        p_dict = {}
        for elem in self.elements:
            val = ParameterHelper.get_value(elem, param_name) or "<empty>"
            if val not in p_dict: p_dict[val] = []
            p_dict[val].append(elem.Id.IntegerValue)
        items = ObservableCollection[DeselectItem]()
        for val in sorted(p_dict.keys()):
            items.Add(DeselectItem("{}: {}".format(param_name, val), len(p_dict[val]), p_dict[val]))
        self.deselect_items_list.ItemsSource = items

    def _get_element_family(self, elem):
        try:
            et = doc.GetElement(elem.GetTypeId())
            if et:
                fp = et.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if fp and fp.HasValue: return fp.AsString()
        except: pass
        return "Unknown Family"

    def _get_element_type_name(self, elem):
        try:
            et = doc.GetElement(elem.GetTypeId())
            if et:
                tn = et.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if tn and tn.HasValue: return tn.AsString()
                return et.Name if hasattr(et, 'Name') else "Unknown Type"
        except: pass
        return "Unknown Type"

    def OnOK(self, sender, args):
        global WINDOW_INSTANCE
        try:
            if not self.elements:
                forms.alert("No elements selected", title="Renumber")
                return
            self._read_ui_to_settings()
            if not self.settings.specify_order and self.settings.order_mode == "Follow spline" and not self.picked_curve:
                forms.alert("Please pick a curve", title="Renumber")
                return
            if not self.settings.specify_order and self.settings.order_mode == "Order by Parameter" and not self.settings.order_parameter:
                forms.alert("Please select a parameter", title="Renumber")
                return
            
            ConfigManager.save(self.settings)
            
            custom_group_order = None
            if self.settings.grouping_mode != GROUPING_MODES[0]:
                custom_group_order = [item.Name for item in self.settings.group_order_items]

            success, skipped, errors = RenumberEngine.execute(
                self.elements, self.settings, self.picked_curve, self.manual_order, custom_group_order
            )
            
            self._results = {'success': success, 'skipped': skipped, 'errors': errors, 'total': len(self.elements)}
            WINDOW_INSTANCE = None
            self.DialogResult = True
            self.Close()
        except Exception as ex:
            WINDOW_INSTANCE = None
            self._results = None
            try: self.DialogResult = False; self.Close()
            except: pass
            forms.alert("Renumbering failed:\n\n{}".format(ex), title="Error")

    def OnCancel(self, sender, args):
        global WINDOW_INSTANCE
        WINDOW_INSTANCE = None
        self.DialogResult = False
        self.Close()

    def OnRunAndPlaceTags(self, sender, args):
        """Run renumbering then launch Place Tags tool."""
        global WINDOW_INSTANCE
        try:
            if not self.elements:
                forms.alert("No elements selected", title="Renumber")
                return
            self._read_ui_to_settings()
            if not self.settings.specify_order and self.settings.order_mode == "Follow spline" and not self.picked_curve:
                forms.alert("Please pick a curve", title="Renumber")
                return
            if not self.settings.specify_order and self.settings.order_mode == "Order by Parameter" and not self.settings.order_parameter:
                forms.alert("Please select a parameter", title="Renumber")
                return
            
            ConfigManager.save(self.settings)
            
            custom_group_order = None
            if self.settings.grouping_mode != GROUPING_MODES[0]:
                custom_group_order = [item.Name for item in self.settings.group_order_items]

            success, skipped, errors = RenumberEngine.execute(
                self.elements, self.settings, self.picked_curve, self.manual_order, custom_group_order
            )
            
            self._results = {'success': success, 'skipped': skipped, 'errors': errors, 'total': len(self.elements)}
            self._launch_place_tags = True
            self._place_tags_elements = list(self.elements)  # Save elements for Place Tags
            WINDOW_INSTANCE = None
            self.DialogResult = True
            self.Close()
        except Exception as ex:
            WINDOW_INSTANCE = None
            self._results = None
            self._launch_place_tags = False
            try: self.DialogResult = False; self.Close()
            except: pass
            forms.alert("Renumbering failed:\n\n{}".format(ex), title="Error")

    def _on_order_changed(self, sender, args):
        if hasattr(self, '_initializing') and self._initializing: return
        if self.order_combo and self.curve_button:
            is_spline = (self.order_combo.SelectedItem == "Follow spline")
            is_manual = self.specify_order_check.IsChecked
            self.curve_button.IsEnabled = is_spline and not is_manual
            if self.curve_label:
                if is_spline and self.picked_curve: self.curve_label.Text = "Curve Id: {}".format(self.picked_curve.Id.IntegerValue)
                elif is_spline: self.curve_label.Text = "(pick curve needed)"
                else: self.curve_label.Text = "(not applicable)"

    def _on_grouping_mode_changed(self, sender, args):
        if hasattr(self, '_initializing') and self._initializing: return
        self._update_group_order_list()
        if self.grouping_mode_combo and self.specify_order_check:
            is_default = (self.grouping_mode_combo.SelectedItem == GROUPING_MODES[0])
            self.specify_order_check.IsEnabled = is_default
            if not is_default: self.specify_order_check.IsChecked = False

    def _on_order_parameter_changed(self, sender, args):
        if hasattr(self, '_initializing') and self._initializing: return
        if self.order_param_combo:
            # Use SelectedItem first, then fall back to Text for typed values
            val = self.order_param_combo.SelectedItem
            if val is None:
                val = self.order_param_combo.Text
            self.settings.order_parameter = str(val) if val else ""
        self._update_group_order_list()

    def _on_filter_param_changed(self, sender, args):
        """Populate filter value dropdown with unique values from selected parameter"""
        if hasattr(self, '_initializing') and self._initializing: return
        self._populate_filter_values()
    
    def _populate_filter_values(self):
        """Get unique values for the selected filter parameter and populate dropdown"""
        self.settings.filter_values.Clear()
        
        if not self.filter_param_combo or not self.filter_value_box:
            return
        
        param_name = self.filter_param_combo.SelectedItem
        if not param_name or not self.all_picked_elements:
            return
        
        # Collect unique values from all picked elements
        unique_values = set()
        for elem in self.all_picked_elements:
            try:
                val = ParameterHelper.get_value(elem, param_name)
                if val is not None and str(val).strip():
                    unique_values.add(str(val).strip())
            except:
                pass
        
        # Sort and add to collection
        for val in sorted(unique_values, key=lambda s: s.lower()):
            self.settings.filter_values.Add(val)
        
        # Update the ComboBox ItemsSource
        self.filter_value_box.ItemsSource = self.settings.filter_values
        
        # Clear the current selection/text
        self.filter_value_box.SelectedItem = None
        self.filter_value_box.Text = ""

    def _on_deselect_parameter_changed(self, sender, args):
        if hasattr(self, '_initializing') and self._initializing: return
        # Close dropdown explicitly for editable ComboBox
        if self.deselect_parameter_combo and hasattr(self.deselect_parameter_combo, 'IsDropDownOpen'):
            self.deselect_parameter_combo.IsDropDownOpen = False
        if self.deselect_parameter_radio.IsChecked: self._populate_parameter_list()

    def _on_specify_order_changed(self, sender, args):
        is_checked = self.specify_order_check.IsChecked
        if self.pick_order_button: self.pick_order_button.IsEnabled = is_checked
        if self.order_status_label:
            fids = set([e.Id for e in self.elements])
            fc = len([e for e in self.manual_order if e.Id in fids])
            self.order_status_label.Text = "Order: {} elements".format(fc) if (is_checked and fc > 0) else "(not set)"
        if self.order_combo: self.order_combo.IsEnabled = not is_checked
        if self.grouping_mode_combo:
            self.grouping_mode_combo.IsEnabled = not is_checked
            if is_checked: self.grouping_mode_combo.SelectedItem = GROUPING_MODES[0]
        if self.group_order_box:
            if is_checked: self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
            else: self._update_group_order_list()
        if self.curve_button:
            mode = self.order_combo.SelectedItem
            self.curve_button.IsEnabled = (mode == "Follow spline") and not is_checked
        if not is_checked: self.manual_order = []

    def _update_group_order_list(self):
        if not hasattr(self, 'settings') or not self.group_order_box: return
        if hasattr(self, '_initializing') and self._initializing:
            if not self.grouping_mode_combo.SelectedItem: return
        
        mode = self.grouping_mode_combo.SelectedItem
        if not mode: return
        
        if mode == GROUPING_MODES[0]:
            self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
            self.settings.group_order_items.Clear()
            if self.order_param_box: self.order_param_box.Visibility = System.Windows.Visibility.Collapsed
            return

        if mode == "Order by Parameter":
            if self.order_param_box: self.order_param_box.Visibility = System.Windows.Visibility.Visible
            if self.group_order_box: self.group_order_box.Visibility = System.Windows.Visibility.Visible
            if not self.elements or not self.settings.order_parameter or not self.settings.order_parameter.strip():
                self.settings.group_order_items.Clear()
                return
        else:
            if self.order_param_box: self.order_param_box.Visibility = System.Windows.Visibility.Collapsed
            if not self.elements:
                self.group_order_box.Visibility = System.Windows.Visibility.Collapsed
                self.settings.group_order_items.Clear()
                return
            self.group_order_box.Visibility = System.Windows.Visibility.Visible
        
        unique_keys = set()
        for elem in self.elements:
            key = ParameterHelper.get_group_key(elem, mode, self.settings.order_parameter)
            unique_keys.add(key)
        
        saved = []
        if self.settings.saved_group_order: saved = self.settings.saved_group_order.split(',')
        
        items = [GroupItem(key) for key in unique_keys]
        def sort_key(item):
            try: return (0, saved.index(item.Name), item.Name)
            except:
                try: return (1, float(item.Name.replace("(empty)", "0")), item.Name)
                except: return (1, 0, item.Name.lower())
        items.sort(key=sort_key)
        
        self.settings.group_order_items.Clear()
        for item in items: self.settings.group_order_items.Add(item)

    def _handle_pick_elements(self):
        global WINDOW_INSTANCE
        if self.elements:
            if not forms.alert("Select different elements?", title="Re-pick Elements?", ok=False, yes=True, no=True): return
        
        try:
            self._read_ui_to_settings()
            self._next_action = {'action': 'pick_elements', 'saved_settings': self.settings, 'saved_curve': self.picked_curve}
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")

    def _handle_pick_order(self):
        global WINDOW_INSTANCE
        try:
            self._read_ui_to_settings()
            self._next_action = {'action': 'pick_order', 'saved_settings': self.settings, 'saved_curve': self.picked_curve}
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")

    def _handle_pick_curve(self):
        global WINDOW_INSTANCE
        try:
            self._read_ui_to_settings()
            self._next_action = {
                'action': 'pick_curve',
                'saved_settings': self.settings,
                'saved_elements': self.elements,
                'saved_all_elements': self.all_picked_elements,
                'saved_manual_order': self.manual_order
            }
            WINDOW_INSTANCE = None
            self.DialogResult = False
            self.Close()
        except Exception as ex:
            WINDOW_INSTANCE = None
            forms.alert("Error: {}".format(ex), title="Error")

    def _handle_save_profile(self):
        if self._executing_save: return
        self._executing_save = True
        try:
            self._read_ui_to_settings()
            path = forms.save_file(file_ext="xml", default_name="RenumberProfile.xml")
            if not path: return
            ProfileManager.export(self.settings).Save(path)
            forms.toast("Profile saved", title="Renumber", appid="renumber")
        except Exception as ex: forms.alert("Error: {}".format(ex), title="Error")
        finally: self._executing_save = False

    def _handle_load_profile(self):
        if self._executing_load: return
        self._executing_load = True
        try:
            path = forms.pick_file(file_ext="xml")
            if not path: return
            doc = XmlDocument()
            doc.Load(path)
            ProfileManager.import_to(doc, self.settings)
            self._initialize_ui()
            self._refresh_options_list()
            self._update_group_order_list()
            forms.toast("Profile loaded", title="Renumber", appid="renumber")
        except Exception as ex: forms.alert("Error: {}".format(ex), title="Error")
        finally: self._executing_load = False

def run():
    global WINDOW_INSTANCE
    WINDOW_INSTANCE = None
    preselection = list(revit.get_selection() or [])
    
    try:
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, "window.xaml")
        if not os.path.exists(xaml_path):
            forms.alert("XAML file not found", title="Error")
            return
        
        window = RenumberWindow(xaml_path, preselection)
        launch_place_tags = False
        place_tags_elements = []
        
        while True:
            WINDOW_INSTANCE = window
            window.ShowDialog()
            next_action = window._next_action if hasattr(window, '_next_action') else None
            results = window._results if hasattr(window, '_results') else None
            launch_place_tags = getattr(window, '_launch_place_tags', False)
            place_tags_elements = getattr(window, '_place_tags_elements', [])
            WINDOW_INSTANCE = None
            
            if results:
                s, k, e, t = results['success'], results['skipped'], results['errors'], results['total']
                if DEBUG_MODE and e:
                    print("--- Errors ---")
                    for err in e: print(err)
                
                if s > 0: forms.toast("Renumbered {} elements".format(s), title="Complete", appid="renumber")
                else: forms.alert("No elements updated", title="Failed")
                break
            
            if not next_action: break
            
            act = next_action['action']
            if act == 'pick_elements':
                elems = ElementPicker.pick_multiple()
                if not elems: break
                window = RenumberWindow(xaml_path, elems)
                RenumberWindow.transfer_settings_static(window, next_action['saved_settings'], next_action['saved_curve'])
                window._initialize_ui()
                if window.elements: 
                    window.ok_button.IsEnabled = True
                    window._update_element_count()
                    window._refresh_parameters()
                if DEBUG_MODE and elems:
                    e1 = elems[0]
                    print("--- PARAMETER DUMP FOR ELEM {} ---".format(e1.Id))
                    for p in e1.Parameters:
                        if p.HasValue:
                            val = p.AsValueString() or p.AsString()
                            print("{}: {}".format(p.Definition.Name, val))
                continue
                
            elif act == 'pick_curve':
                c = ElementPicker.pick_curve()
                window = RenumberWindow(xaml_path, next_action['saved_elements'])
                window.all_picked_elements = next_action['saved_all_elements']
                window.manual_order = next_action['saved_manual_order']
                RenumberWindow.transfer_settings_static(window, next_action['saved_settings'], c)
                if c:
                    window.picked_curve = c
                    window.settings.picked_curve_id = c.Id.IntegerValue
                window._initialize_ui()
                if window.elements: 
                    window.ok_button.IsEnabled = True
                    window._update_element_count()
                    window._refresh_parameters()
                continue
                
            elif act == 'pick_order':
                forms.alert("Pick elements in order. ESC when done.", title="Pick Order", ok=True)
                mo = ElementPicker.pick_in_order()
                if not mo: break
                window = RenumberWindow(xaml_path, mo)
                window.manual_order = list(mo)
                window.all_picked_elements = list(mo)
                s = next_action['saved_settings']
                s.specify_order = True
                RenumberWindow.transfer_settings_static(window, s, next_action['saved_curve'])
                window._initialize_ui()
                window.order_status_label.Text = "Order: {} elements".format(len(mo))
                window.specify_order_check.IsChecked = True
                window.pick_order_button.IsEnabled = True
                window.ok_button.IsEnabled = True
                continue
            else: break
        
        # Launch Place Tags if requested
        if launch_place_tags and place_tags_elements:
            try:
                # Find Place Tags script
                place_tags_dir = os.path.join(os.path.dirname(script_dir), "Place Tags.pushbutton")
                place_tags_script = os.path.join(place_tags_dir, "script.py")
                place_tags_xaml = os.path.join(place_tags_dir, "window.xaml")
                
                if os.path.exists(place_tags_script) and os.path.exists(place_tags_xaml):
                    # Import and run Place Tags
                    import sys
                    if place_tags_dir not in sys.path:
                        sys.path.insert(0, place_tags_dir)
                    
                    # Execute Place Tags script with preselection
                    exec(compile(open(place_tags_script).read(), place_tags_script, 'exec'), {
                        '__name__': '__main__',
                        '__file__': place_tags_script,
                        '_preselected_elements': place_tags_elements
                    })
                else:
                    forms.alert("Place Tags tool not found at:\n{}".format(place_tags_dir), title="Error")
            except Exception as ex:
                forms.alert("Failed to launch Place Tags:\n{}".format(ex), title="Error")
            
    except Exception as ex:
        WINDOW_INSTANCE = None
        forms.alert("Error: {}".format(ex), title="Renumber Error")

if __name__ == "__main__":
    run()