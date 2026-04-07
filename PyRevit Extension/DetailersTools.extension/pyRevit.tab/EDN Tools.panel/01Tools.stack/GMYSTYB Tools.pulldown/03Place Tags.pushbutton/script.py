# -*- coding: utf-8 -*-
"""
Place Tags - Multi-Category
Author: Jeremiah Griffith

Features:
- Works on ANY model category
- Modern WPF UI
- Robust "Tag All" style tag discovery
"""

from pyrevit import revit, DB, forms, script
import clr
import os
import json
import math
import System
from System.Collections.ObjectModel import ObservableCollection
import System.Windows

doc = revit.doc
uidoc = revit.uidoc

# ---------------------------- UTILITIES ---------------------------

def get_safe_symbol_info(symbol):
    """
    Safely retrieve Family Name and Type Name using BuiltInParameters.
    This prevents 'AttributeError: Name' crashes on phantom elements.
    """
    fam_name = "Unknown Family"
    typ_name = "Unknown Type"
    
    if not symbol: return fam_name, typ_name

    # 1. Try to get Type Name via Parameter (Safest)
    try:
        p_type = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
        if p_type and p_type.HasValue:
            typ_name = p_type.AsString()
        elif hasattr(symbol, "Name"):
            typ_name = symbol.Name
    except: pass

    # 2. Try to get Family Name via Parameter (Safest)
    try:
        p_fam = symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p_fam and p_fam.HasValue:
            fam_name = p_fam.AsString()
        elif hasattr(symbol, "FamilyName"): # Revit 2022+ property
            fam_name = symbol.FamilyName
        elif hasattr(symbol, "Family") and symbol.Family:
            fam_name = symbol.Family.Name
    except: pass

    return fam_name, typ_name

def find_tag_by_family_type(family_name, type_name):
    """
    Find a tag FamilySymbol by family name and type name.
    Returns the FamilySymbol or None if not found.
    """
    try:
        collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
        for symbol in collector:
            # Check if it's a tag (annotation category)
            if symbol.Category and symbol.Category.CategoryType == DB.CategoryType.Annotation:
                fam, typ = get_safe_symbol_info(symbol)
                if fam == family_name and typ == type_name:
                    return symbol
    except:
        pass
    return None

def inches_to_feet(inches):
    return float(inches) / 12.0

def collect_nested_families(picked_elements, include_nested):
    """
    Expands selection to include nested family instances if enabled.
    Returns expanded list of elements including nested components.
    """
    if not include_nested:
        return picked_elements
    
    # Keep track of original element IDs to preserve them
    original_ids = set(elem.Id.IntegerValue for elem in picked_elements if elem)
    expanded_ids = set(original_ids)
    host_ids = set(original_ids)
    
    # Collect nested components recursively (only for FamilyInstances)
    for elem in picked_elements:
        if elem and isinstance(elem, DB.FamilyInstance):
            _collect_nested_subcomponents(elem, expanded_ids)
    
    # Collect children via SuperComponent
    _collect_children_by_supercomponent(host_ids, expanded_ids)
    
    # Convert back to element list - KEEP ALL ORIGINAL ELEMENTS plus nested FamilyInstances
    result = []
    for eid_int in expanded_ids:
        elem = doc.GetElement(DB.ElementId(eid_int))
        if elem:
            # Keep if it was originally picked OR if it's a nested FamilyInstance
            if eid_int in original_ids or isinstance(elem, DB.FamilyInstance):
                result.append(elem)
    
    return result

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
                _collect_nested_subcomponents(sub_el, out_ids)

def _collect_children_by_supercomponent(host_ids_int, out_ids):
    """Collect child instances via SuperComponent property"""
    try:
        for fi in DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).ToElements():
            try:
                sc = fi.SuperComponent
            except:
                sc = None
            
            if sc and sc.Id and sc.Id.IntegerValue in host_ids_int:
                out_ids.add(fi.Id.IntegerValue)
    except:
        pass

def get_profiles_folder():
    script_dir = os.path.dirname(__file__)
    profiles_dir = os.path.join(script_dir, "TagProfiles")
    if not os.path.exists(profiles_dir):
        os.makedirs(profiles_dir)
    return profiles_dir

def save_profile(profile_data):
    profiles_dir = get_profiles_folder()
    safe_name = "".join(c for c in profile_data["name"] if c.isalnum() or c in (' ', '-', '_')).strip()
    filepath = os.path.join(profiles_dir, "{}.json".format(safe_name))
    try:
        with open(filepath, 'w') as f:
            json.dump(profile_data, f, indent=2)
        return True
    except Exception as ex:
        forms.alert("Error saving profile: {}".format(ex))
        return False

def load_profiles():
    profiles_dir = get_profiles_folder()
    profiles = []
    if os.path.exists(profiles_dir):
        for f in os.listdir(profiles_dir):
            if f.endswith('.json'):
                try:
                    with open(os.path.join(profiles_dir, f), 'r') as file:
                        profiles.append(json.load(file))
                except: pass
    return profiles

def find_tags_like_tag_all(category_name):
    """
    Emulates 'Tag All' logic to find relevant tags.
    """
    collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
    
    # Matching Heuristics
    potential_names = []
    cat_lower = category_name.lower()
    
    # 1. Exact "Category Tags"
    potential_names.append(cat_lower + " tags")
    
    # 2. Singularize (Ducts -> Duct Tags)
    if cat_lower.endswith("s"):
        potential_names.append(cat_lower[:-1] + " tags")
        
    # 3. Handle "Accessories" -> "Accessory"
    if cat_lower.endswith("ies"):
        potential_names.append(cat_lower[:-3] + "y tags")
    
    valid_symbols = []
    
    for symbol in collector:
        try:
            # Check Category Name of the Tag Family
            if not symbol.Family: continue
            tag_cat = symbol.Family.FamilyCategory
            if not tag_cat: continue
            
            tag_cat_name = tag_cat.Name.lower()
            
            # Must be a "Tags" category
            if not tag_cat_name.endswith("tags"): continue
            
            # Check Matches
            is_match = False
            
            # A. Direct Match
            if tag_cat_name in potential_names:
                is_match = True
                
            # B. Substring Match (e.g. "Duct" inside "MEP Fabrication Ductwork Tags")
            elif cat_lower in tag_cat_name:
                is_match = True
            
            # C. Singular Substring (e.g. "Duct" from "Ducts" inside "Duct Tags")
            elif cat_lower.endswith("s") and cat_lower[:-1] in tag_cat_name:
                is_match = True

            if is_match:
                valid_symbols.append(symbol)

        except:
            continue
            
    return valid_symbols

def get_element_direction(elem):
    """Get orientation vector for rotation logic."""
    try:
        if hasattr(elem, "Location") and isinstance(elem.Location, DB.LocationCurve):
            curve = elem.Location.Curve
            p0 = curve.GetEndPoint(0)
            p1 = curve.GetEndPoint(1)
            return (p1 - p0).Normalize()

        if isinstance(elem, DB.FamilyInstance):
            return elem.FacingOrientation
    except:
        pass
    return DB.XYZ.BasisX


def get_anchor_point(elem, view):
    """
    Content-aware leader target point, in priority order:
    1. LocationCurve midpoint  -- pipes, ducts, conduit, tray, beams
    2. LocationPoint           -- families, equipment
    3. Single connector origin -- MEP equipment with one obvious connector
    4. Bounding box center     -- fallback for anything else
    """
    try:
        loc = getattr(elem, 'Location', None)
        if isinstance(loc, DB.LocationCurve) and loc.Curve:
            curve = loc.Curve
            return curve.Evaluate(0.5, True)
        if isinstance(loc, DB.LocationPoint):
            return loc.Point
    except:
        pass

    try:
        if hasattr(elem, 'MEPModel') and elem.MEPModel:
            connectors = list(elem.MEPModel.ConnectorManager.Connectors)
            if len(connectors) == 1:
                return connectors[0].Origin
    except:
        pass

    try:
        bbox = elem.get_BoundingBox(view)
        if bbox:
            return (bbox.Min + bbox.Max) / 2.0
    except:
        pass

    return None


def normalize_xy(vec, fallback=None):
    """Normalize an XY vector safely."""
    try:
        v = DB.XYZ(vec.X, vec.Y, 0)
        length = v.GetLength()
        if length > 1e-9:
            return DB.XYZ(v.X / length, v.Y / length, 0)
    except:
        pass
    return fallback if fallback else DB.XYZ.BasisX


def rotate_xy(vec, angle_rad):
    """Rotate an XY vector by angle in radians."""
    c = math.cos(angle_rad)
    s = math.sin(angle_rad)
    return DB.XYZ((vec.X * c) - (vec.Y * s), (vec.X * s) + (vec.Y * c), 0)


def get_tag_axes(direction, rot_mode, extra_rad):
    """Return the tag major axis and side axis in the active view plane."""
    base_dir = normalize_xy(direction, DB.XYZ.BasisX)

    if rot_mode == "In-Line":
        tag_axis = base_dir
    elif rot_mode == "Perpendicular":
        tag_axis = DB.XYZ(-base_dir.Y, base_dir.X, 0)
    elif rot_mode == "Vertical":
        tag_axis = DB.XYZ.BasisY
    else:
        tag_axis = DB.XYZ.BasisX

    if abs(extra_rad) > 1e-9:
        tag_axis = rotate_xy(tag_axis, extra_rad)

    tag_axis = normalize_xy(tag_axis, DB.XYZ.BasisX)
    side_axis = DB.XYZ(-tag_axis.Y, tag_axis.X, 0)
    return tag_axis, side_axis


def compute_free_end_elbow(tag_head, center, tag_axis, shoulder_ft=0.0):
    """
    Compute the right-angle elbow point for a leader.

    The first leg runs from tag_head ALONG tag_axis for shoulder_ft feet.
    The second leg runs perpendicular from the elbow toward center.

    If shoulder_ft is zero or not provided, the elbow is placed at the
    natural projection of center onto the tag axis (preserves old behavior).

        tag_head
           |  <- first leg (parallel to tag axis, length = shoulder_ft)
           |
        elbow
           |  <- second leg (perpendicular, toward element center)
           ↓
        center
    """
    axis = normalize_xy(tag_axis, DB.XYZ.BasisX)

    if abs(shoulder_ft) > 1e-6:
        # Determine sign: which direction along axis gets us closer to center
        delta = DB.XYZ(center.X - tag_head.X, center.Y - tag_head.Y, 0)
        dot = delta.X * axis.X + delta.Y * axis.Y
        sign = 1.0 if dot >= 0 else -1.0
        return DB.XYZ(
            tag_head.X + axis.X * shoulder_ft * sign,
            tag_head.Y + axis.Y * shoulder_ft * sign,
            tag_head.Z
        )
    else:
        # Natural projection fallback
        delta = DB.XYZ(center.X - tag_head.X, center.Y - tag_head.Y, 0)
        along_axis = delta.X * axis.X + delta.Y * axis.Y
        return DB.XYZ(
            tag_head.X + axis.X * along_axis,
            tag_head.Y + axis.Y * along_axis,
            tag_head.Z
        )


# ---------------------------- DATA CLASSES ---------------------------

class TagMappingItem(object):
    """Represents a row in the Tag Configuration list."""
    def __init__(self, category_name, count, assigned_tag_symbol):
        self.Category = category_name
        self.Count = count
        self.TagSymbol = assigned_tag_symbol
        self.TagName = self._get_display_name()
        # "link" style (blue) if assigned, "error" style (red) if missing
        self.State = "link" if assigned_tag_symbol else "error"
    
    def _get_display_name(self):
        if self.TagSymbol:
            fam, typ = get_safe_symbol_info(self.TagSymbol)
            if fam == "Unknown Family":
                return "Valid Tag (Name Read Error)"
            return "{} : {}".format(fam, typ)
        return "[Click to Assign Tag]"


class DeselectItem(object):
    """Represents a row in the Deselect list."""
    def __init__(self, name, count, element_ids=None):
        self.Name = name
        self.Count = "({})".format(count)
        self.ElementIds = element_ids or []

# ---------------------------- MAIN WINDOW ---------------------------

class TaggingWindow(forms.WPFWindow):
    def __init__(self, xaml_file, pre_selection=None):
        forms.WPFWindow.__init__(self, xaml_file)
        
        # Data
        self.elements = pre_selection or []
        self.tag_mapping = {} # {CategoryName: FamilySymbol}
        self.profiles = load_profiles()
        
        # State
        self._next_action = None
        
        # UI Setup
        self._setup_ui()
        self._refresh_analysis()
    
    @property
    def include_nested(self):
        """Returns True if Include Nested Families checkbox is checked"""
        try:
            return self.IncludeNestedChk.IsChecked == True
        except:
            return True  # Default to True if checkbox not found

    def _setup_ui(self):
        # Profiles
        self.ProfileCombo.Items.Add("-- New Profile --")
        for p in self.profiles:
            self.ProfileCombo.Items.Add(p.get("name", "Unnamed"))
        self.ProfileCombo.SelectedIndex = 0
        
        # Rotation
        self.RotationCombo.Items.Add("In-Line")
        self.RotationCombo.Items.Add("Horizontal")
        self.RotationCombo.Items.Add("Vertical")
        self.RotationCombo.Items.Add("Perpendicular")
        self.RotationCombo.SelectedIndex = 0
        
        # Leader options
        self.LeaderCombo.Items.Add("None")
        self.LeaderCombo.Items.Add("Attached End")
        self.LeaderCombo.Items.Add("Free End")
        self.LeaderCombo.SelectedIndex = 0

        # Event wiring
        self.MemberGrid.MouseDoubleClick += self.OnGridDoubleClick

    def _refresh_analysis(self):
        """Analyze selected elements and populate the UI list."""
        if not self.elements:
            self.StatusText.Text = "No elements selected."
            self.MemberGrid.ItemsSource = None
            self.RunBtn.IsEnabled = False
            self._refresh_deselect_list()
            return

        self.StatusText.Text = "Ready. {} elements selected.".format(len(self.elements))
        self.RunBtn.IsEnabled = True

        # Group by Category
        cat_counts = {}
        for e in self.elements:
            try:
                if not e.Category: continue
                cname = e.Category.Name
                cat_counts[cname] = cat_counts.get(cname, 0) + 1
            except: continue

        # Populate List
        items = ObservableCollection[TagMappingItem]()
        
        for cname in sorted(cat_counts.keys()):
            # Check if we already have a mapping for this session
            assigned = self.tag_mapping.get(cname, None)
            
            # If not, try auto-detect
            if not assigned:
                matches = find_tags_like_tag_all(cname)
                if matches:
                    # Default to first match
                    assigned = matches[0]
                    self.tag_mapping[cname] = assigned
            
            items.Add(TagMappingItem(cname, cat_counts[cname], assigned))
            
        self.MemberGrid.ItemsSource = items
        self._refresh_deselect_list()

    # --- UI Events ---

    def OnSelectElements(self, sender, args):
        self._next_action = "pick"
        self.Close()

    def OnDeselectModeChanged(self, sender, args):
        """Handle radio button change for deselect mode."""
        self._refresh_deselect_list()

    def OnDeselect(self, sender, args):
        """Remove selected items based on current deselect mode."""
        if not hasattr(self, 'DeselectList') or not self.DeselectList:
            return
        
        selected_items = list(self.DeselectList.SelectedItems)
        if not selected_items:
            return
        
        # Determine which mode is selected
        if hasattr(self, 'DeselectIndividualRadio') and self.DeselectIndividualRadio.IsChecked:
            self._deselect_by_individual(selected_items)
        elif hasattr(self, 'DeselectFamilyRadio') and self.DeselectFamilyRadio.IsChecked:
            self._deselect_by_family(selected_items)
        elif hasattr(self, 'DeselectTypeRadio') and self.DeselectTypeRadio.IsChecked:
            self._deselect_by_type(selected_items)
        elif hasattr(self, 'DeselectTaggedRadio') and self.DeselectTaggedRadio.IsChecked:
            self._deselect_by_tagged(selected_items)
        else:  # Default to category
            self._deselect_by_category(selected_items)

    def _deselect_by_individual(self, selected_items):
        """Remove specific elements by ID."""
        ids_to_remove = set()
        for item in selected_items:
            for eid in item.ElementIds:
                ids_to_remove.add(str(eid))
        
        self.elements = [e for e in self.elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self._refresh_analysis()

    def _deselect_by_category(self, selected_items):
        """Remove all elements matching selected categories."""
        cats_to_remove = set(item.Name for item in selected_items)
        self.elements = [e for e in self.elements if not (e.Category and e.Category.Name in cats_to_remove)]
        
        # Remove from tag_mapping too
        for cat_name in cats_to_remove:
            if cat_name in self.tag_mapping:
                del self.tag_mapping[cat_name]
        
        self._refresh_analysis()

    def _deselect_by_family(self, selected_items):
        """Remove all elements matching selected families."""
        fams_to_remove = set(item.Name for item in selected_items)
        self.elements = [e for e in self.elements if self._get_element_family(e) not in fams_to_remove]
        self._refresh_analysis()

    def _deselect_by_type(self, selected_items):
        """Remove all elements matching selected types."""
        types_to_remove = set(item.Name for item in selected_items)
        self.elements = [e for e in self.elements if self._get_element_type_name(e) not in types_to_remove]
        self._refresh_analysis()

    def _deselect_by_tagged(self, selected_items):
        """Remove individually selected tagged elements."""
        ids_to_remove = set()
        for item in selected_items:
            for eid in item.ElementIds:
                if eid:  # Skip empty placeholder
                    ids_to_remove.add(str(eid))
        
        if not ids_to_remove:
            return
        
        self.elements = [e for e in self.elements if str(e.Id.IntegerValue) not in ids_to_remove]
        self._refresh_analysis()

    def _get_tagged_element_ids(self):
        """Get set of element IDs that already have tags in the active view."""
        tagged_ids = set()
        view = doc.ActiveView
        
        if not view:
            return tagged_ids
        
        try:
            # Collect all IndependentTags in the view
            tags = DB.FilteredElementCollector(doc, view.Id)\
                     .OfClass(DB.IndependentTag)\
                     .ToElements()
            
            for tag in tags:
                try:
                    # Revit 2018+ method - can tag multiple elements
                    if hasattr(tag, 'GetTaggedLocalElementIds'):
                        for eid in tag.GetTaggedLocalElementIds():
                            tagged_ids.add(eid.IntegerValue)
                    # Older single-element method
                    elif hasattr(tag, 'TaggedLocalElementId'):
                        eid = tag.TaggedLocalElementId
                        if eid and eid != DB.ElementId.InvalidElementId:
                            tagged_ids.add(eid.IntegerValue)
                except:
                    pass
        except:
            pass
        
        return tagged_ids

    def _get_element_family(self, elem):
        """Get family name for an element."""
        try:
            et = doc.GetElement(elem.GetTypeId())
            if et:
                fp = et.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
                if fp and fp.HasValue:
                    return fp.AsString()
        except:
            pass
        return "Unknown Family"

    def _get_element_type_name(self, elem):
        """Get type name for an element."""
        try:
            et = doc.GetElement(elem.GetTypeId())
            if et:
                tp = et.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if tp and tp.HasValue:
                    return tp.AsString()
                if hasattr(et, 'Name'):
                    return et.Name
        except:
            pass
        return "Unknown Type"

    def _refresh_deselect_list(self):
        """Populate the deselect list based on selected mode."""
        if not hasattr(self, 'DeselectList') or not self.DeselectList:
            return
        
        if not self.elements:
            self.DeselectList.ItemsSource = ObservableCollection[DeselectItem]()
            return
        
        # Determine which mode is selected
        if hasattr(self, 'DeselectIndividualRadio') and self.DeselectIndividualRadio.IsChecked:
            self._populate_individual_list()
        elif hasattr(self, 'DeselectFamilyRadio') and self.DeselectFamilyRadio.IsChecked:
            self._populate_family_list()
        elif hasattr(self, 'DeselectTypeRadio') and self.DeselectTypeRadio.IsChecked:
            self._populate_type_list()
        elif hasattr(self, 'DeselectTaggedRadio') and self.DeselectTaggedRadio.IsChecked:
            self._populate_tagged_list()
        else:  # Default to category
            self._populate_category_list()

    def _populate_individual_list(self):
        """Populate list with individual elements."""
        items = ObservableCollection[DeselectItem]()
        for elem in self.elements:
            elem_id = str(elem.Id.IntegerValue)
            category = elem.Category.Name if elem.Category else "Unknown"
            name = "Element {}".format(elem_id)
            try:
                np = elem.get_Parameter(DB.BuiltInParameter.ELEM_NAME_PARAM)
                if np and np.HasValue:
                    name = np.AsString()
            except:
                pass
            display = "ID {} - {} - {}".format(elem_id, category, name)
            items.Add(DeselectItem(display, 1, [elem_id]))
        self.DeselectList.ItemsSource = items

    def _populate_category_list(self):
        """Populate list grouped by category."""
        cat_dict = {}
        for elem in self.elements:
            try:
                if not elem.Category:
                    continue
                cname = elem.Category.Name
                if cname not in cat_dict:
                    cat_dict[cname] = []
                cat_dict[cname].append(elem.Id.IntegerValue)
            except:
                continue
        
        items = ObservableCollection[DeselectItem]()
        for cname in sorted(cat_dict.keys()):
            items.Add(DeselectItem(cname, len(cat_dict[cname]), cat_dict[cname]))
        self.DeselectList.ItemsSource = items

    def _populate_family_list(self):
        """Populate list grouped by family."""
        fam_dict = {}
        for elem in self.elements:
            fn = self._get_element_family(elem)
            if fn not in fam_dict:
                fam_dict[fn] = []
            fam_dict[fn].append(elem.Id.IntegerValue)
        
        items = ObservableCollection[DeselectItem]()
        for fn in sorted(fam_dict.keys()):
            items.Add(DeselectItem(fn, len(fam_dict[fn]), fam_dict[fn]))
        self.DeselectList.ItemsSource = items

    def _populate_type_list(self):
        """Populate list grouped by type."""
        type_dict = {}
        for elem in self.elements:
            tn = self._get_element_type_name(elem)
            if tn not in type_dict:
                type_dict[tn] = []
            type_dict[tn].append(elem.Id.IntegerValue)
        
        items = ObservableCollection[DeselectItem]()
        for tn in sorted(type_dict.keys()):
            items.Add(DeselectItem(tn, len(type_dict[tn]), type_dict[tn]))
        self.DeselectList.ItemsSource = items

    def _populate_tagged_list(self):
        """Populate list showing only already-tagged elements with family:type."""
        tagged_ids = self._get_tagged_element_ids()
        
        items = ObservableCollection[DeselectItem]()
        
        # Build list of tagged elements only
        element_info = []
        for elem in self.elements:
            elem_id = elem.Id.IntegerValue
            if elem_id in tagged_ids:
                family = self._get_element_family(elem)
                type_name = self._get_element_type_name(elem)
                element_info.append({
                    'id': elem_id,
                    'family': family,
                    'type': type_name
                })
        
        # Sort by family:type
        element_info.sort(key=lambda x: (x['family'], x['type']))
        
        for info in element_info:
            display = "{} : {} (ID: {})".format(
                info['family'],
                info['type'],
                info['id']
            )
            items.Add(DeselectItem(display, 1, [info['id']]))
        
        if not element_info:
            items.Add(DeselectItem("No tagged elements in selection", 0, []))
        
        self.DeselectList.ItemsSource = items

    def OnGridDoubleClick(self, sender, args):
        """Handle double click to change tag."""
        if self.MemberGrid.SelectedItems.Count == 0: return
        
        selected_item = self.MemberGrid.SelectedItem
        if not selected_item: return

        cat_name = selected_item.Category
        
        # Find potential tags
        all_tags = find_tags_like_tag_all(cat_name)
        
        if not all_tags:
            # Ask user if they want to load a tag family
            load_family = forms.alert(
                "No loaded tags found for category '{}'.\n\nWould you like to load a tag family?".format(cat_name),
                title="No Tags Found",
                ok=False,
                yes=True,
                no=True
            )
            
            if load_family:
                # Open file picker for RFA files
                rfa_path = forms.pick_file(file_ext='rfa', title="Select Tag Family to Load")
                if rfa_path:
                    try:
                        t = DB.Transaction(doc, "Load Tag Family")
                        t.Start()
                        loaded_family = clr.Reference[DB.Family]()
                        success = doc.LoadFamily(rfa_path, loaded_family)
                        t.Commit()
                        
                        if success:
                            forms.toast("Family loaded successfully!", title="Load Family")
                            # Re-check for tags after loading
                            all_tags = find_tags_like_tag_all(cat_name)
                            if all_tags:
                                # Auto-assign the first match
                                self.tag_mapping[cat_name] = all_tags[0]
                                self._refresh_analysis()
                        else:
                            forms.alert("Family may already be loaded or failed to load.", title="Load Family")
                            # Still refresh in case it was already loaded
                            all_tags = find_tags_like_tag_all(cat_name)
                            if all_tags:
                                self.tag_mapping[cat_name] = all_tags[0]
                                self._refresh_analysis()
                    except Exception as ex:
                        if t.HasStarted():
                            t.RollBack()
                        forms.alert("Error loading family: {}".format(ex), title="Error")
            return

        # Prepare list for selection dialog
        options = {}
        sorted_keys = []
        
        for symbol in all_tags:
            fam, typ = get_safe_symbol_info(symbol)
            key = "{} : {}".format(fam, typ)
            
            # Ensure unique keys in list
            if key in options:
                # Fallback if duplicate names exist (rare)
                key = "{} : {} (Id: {})".format(fam, typ, symbol.Id)
            
            options[key] = symbol
            sorted_keys.append(key)
            
        sorted_keys.sort()

        # Show Dialog
        sel_name = forms.SelectFromList.show(
            sorted_keys,
            title="Select Tag for {}".format(cat_name),
            button_name="Assign Tag"
        )
        
        if sel_name:
            symbol = options[sel_name]
            self.tag_mapping[cat_name] = symbol
            self._refresh_analysis() # Refresh list to show new name

    def OnSaveProfile(self, sender, args):
        name = self.ProfileNameBox.Text
        if not name: return
        
        data = self.GetProfileData()
        data["name"] = name
        save_profile(data)
        
        # Refresh combo
        self.profiles = load_profiles()
        self.ProfileCombo.Items.Clear()
        self.ProfileCombo.Items.Add("-- New Profile --")
        for p in self.profiles:
            self.ProfileCombo.Items.Add(p["name"])
        self.ProfileCombo.SelectedIndex = self.ProfileCombo.Items.IndexOf(name)

    def OnLeaderChanged(self, sender, args):
        """Show/hide leader offset panel based on leader selection."""
        if not hasattr(self, 'LeaderOffsetPanel'):
            return
        if self.LeaderCombo.SelectedItem in ["Attached End", "Free End"]:
            self.LeaderOffsetPanel.Visibility = System.Windows.Visibility.Visible
            # Auto-set Shoulder and Drop defaults
            self._sync_leader_defaults()
        else:
            self.LeaderOffsetPanel.Visibility = System.Windows.Visibility.Collapsed

    def _sync_leader_defaults(self):
        """Set default Shoulder value when leader is enabled."""
        try:
            # Only auto-set if Shoulder is currently 0 (default)
            current_shoulder = float(self.ElbowXBox.Text or 0)
            if abs(current_shoulder) < 0.001:
                self.ElbowXBox.Text = "16"
        except:
            pass

    def OnProfileChanged(self, sender, args):
        if self.ProfileCombo.SelectedIndex < 1: return
        
        name = self.ProfileCombo.SelectedItem
        selected_profile = next((p for p in self.profiles if p["name"] == name), None)
        
        if selected_profile:
            self.ProfileNameBox.Text = selected_profile.get("name", "")
            # Changed key from off_perp to off_updown to match new UI logic
            self.OffsetUpDownBox.Text = str(selected_profile.get("off_updown", 12))
            self.OffsetAlongBox.Text = str(selected_profile.get("off_along", 0))
            self.ExtraAngleBox.Text = str(selected_profile.get("extra_angle", 0))
            
            # Leader type - handle legacy boolean format
            leader_val = selected_profile.get("leader", "None")
            if isinstance(leader_val, bool):
                leader_val = "Attached End" if leader_val else "None"
            if leader_val in ["None", "Attached End", "Free End"]:
                self.LeaderCombo.SelectedItem = leader_val
            else:
                self.LeaderCombo.SelectedIndex = 0
            
            # Shoulder offset
            self.ElbowXBox.Text = str(selected_profile.get("elbow_x", 0))
            
            # Restore Re-tag toggle (inverted: skip_tagged=True means retag=False)
            skip_tagged = selected_profile.get("skip_tagged", True)
            try:
                self.RetagChk.IsChecked = not skip_tagged
            except:
                pass
            
            rot = selected_profile.get("rotation", "In-Line")
            if rot in self.RotationCombo.Items:
                self.RotationCombo.SelectedItem = rot
            
            # Restore tag assignments
            tag_assignments = selected_profile.get("tag_assignments", {})
            if tag_assignments:
                for cat_name, tag_info in tag_assignments.items():
                    family = tag_info.get("family", "")
                    typ = tag_info.get("type", "")
                    if family and typ:
                        symbol = find_tag_by_family_type(family, typ)
                        if symbol:
                            self.tag_mapping[cat_name] = symbol
                # Refresh the grid to show restored tags
                self._refresh_analysis()

    def GetProfileData(self):
        try:
            # Build serializable tag mapping
            tag_assignments = {}
            for cat_name, symbol in self.tag_mapping.items():
                if symbol:
                    fam_name, typ_name = get_safe_symbol_info(symbol)
                    tag_assignments[cat_name] = {
                        "family": fam_name,
                        "type": typ_name
                    }
            
            return {
                "name": self.ProfileNameBox.Text,
                "off_updown": float(self.OffsetUpDownBox.Text or 0),
                "off_along": float(self.OffsetAlongBox.Text or 0),
                "rotation": self.RotationCombo.SelectedItem,
                "extra_angle": float(self.ExtraAngleBox.Text or 0),
                "leader": self.LeaderCombo.SelectedItem or "None",
                "elbow_x": float(self.ElbowXBox.Text or 0),
                "skip_tagged": self.RetagChk.IsChecked != True,
                "tag_assignments": tag_assignments
            }
        except:
            return None

    def OnRun(self, sender, args):
        self._next_action = "run"
        self.Close()

# ---------------------------- LOGIC LOOP ---------------------------

def run_tool(preselected_elements=None):
    # Initial Selection - use preselected if provided, else get from Revit selection
    if preselected_elements:
        selection = preselected_elements
    else:
        selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]
    selection = [e for e in selection if e and e.Category] # Filter valid
    
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, "window.xaml")
    
    # State persistence
    current_mapping = {}
    include_nested = True  # Default to include nested
    
    while True:
        # Expand selection with nested families before showing window
        expanded_selection = collect_nested_families(selection, include_nested)
        
        # Launch Window
        window = TaggingWindow(xaml_path, pre_selection=expanded_selection)
        # Restore mapping if we looped
        if current_mapping:
            window.tag_mapping = current_mapping
            window._refresh_analysis()
            
        window.ShowDialog()
        
        action = window._next_action
        include_nested = window.include_nested  # Get current checkbox state
        
        if action == "pick":
            # Hide and Pick
            try:
                from Autodesk.Revit.UI.Selection import ObjectType
                refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements to tag")
                selection = [doc.GetElement(r.ElementId) for r in refs]
                selection = [e for e in selection if e and e.Category]  # Filter valid
                current_mapping = window.tag_mapping # Persist mapping
            except:
                pass # User cancelled pick
            continue # Loop back to show window
            
        elif action == "run":
            # Execute
            profile = window.GetProfileData()
            mapping = window.tag_mapping
            if not profile:
                forms.alert("Invalid input values (offsets/angles must be numbers).")
                continue

            execute_tagging(window.elements, mapping, profile)
            break  # Done
            
        else:
            break # Close/Cancel

def execute_fix_leaders(tags, leader_type, overwrite, view):
    """
    Repair leader geometry on existing IndependentTags using the same
    compute_free_end_elbow logic as placement. Infers tag axis from the
    tag's current visual orientation. Delete-and-recreate strategy ensures
    SetLeaderElbow sticks (only reliable on freshly created tags).
    """
    use_attached = (leader_type == "Attached End")
    leader_end_cond = DB.LeaderEndCondition.Attached if use_attached else DB.LeaderEndCondition.Free

    count_ok = 0
    count_skip = 0
    count_fail = 0
    errors = []

    # Collect all tag data BEFORE the transaction
    tag_data = []
    for tag in tags:
        try:
            if not overwrite:
                if tag.HasLeader:
                    count_skip += 1
                    continue

            # Resolve tagged element and reference
            tagged_elem = None
            tag_ref = None
            try:
                if hasattr(tag, 'GetTaggedReferences'):
                    refs = list(tag.GetTaggedReferences())
                    if refs:
                        tag_ref = refs[0]
                        tagged_elem = doc.GetElement(tag_ref.ElementId)
            except:
                pass
            if not tagged_elem or not tag_ref:
                count_fail += 1
                continue

            bbox = tagged_elem.get_BoundingBox(view)
            if not bbox:
                count_fail += 1
                continue

            center = get_anchor_point(tagged_elem, view)
            if not center:
                center = (bbox.Min + bbox.Max) / 2.0
            tag_head = tag.TagHeadPosition
            tag_type_id = tag.GetTypeId()
            tag_orient = tag.TagOrientation

            # Infer tag axis from its visual orientation
            # TagOrientation.Horizontal = axis is BasisX
            # TagOrientation.Vertical   = axis is BasisY
            # TagOrientation.Model      = use element direction
            if tag_orient == DB.TagOrientation.Horizontal:
                tag_axis = DB.XYZ.BasisX
            elif tag_orient == DB.TagOrientation.Vertical:
                tag_axis = DB.XYZ.BasisY
            else:  # Model -- infer from element
                direction = normalize_xy(get_element_direction(tagged_elem), DB.XYZ.BasisX)
                tag_axis = direction

            elbow_pt = compute_free_end_elbow(tag_head, center, tag_axis)

            tag_data.append({
                "old_id":    tag.Id,
                "type_id":   tag_type_id,
                "tag_ref":   tag_ref,
                "tag_head":  tag_head,
                "center":    center,
                "elbow_pt":  elbow_pt,
                "orientation": tag_orient,
            })
        except Exception as ex:
            count_fail += 1
            errors.append("Tag {}: {}".format(
                tag.Id.IntegerValue if hasattr(tag, 'Id') else '?', str(ex)))

    if not tag_data:
        forms.toast("Skipped: {} | Errors: {}".format(count_skip, count_fail),
                    title="Fix Leaders")
        return

    t = DB.Transaction(doc, "Fix Tag Leaders")
    t.Start()

    for d in tag_data:
        try:
            doc.Delete(d["old_id"])
            doc.Regenerate()

            sym = doc.GetElement(d["type_id"])
            if sym and hasattr(sym, 'IsActive') and not sym.IsActive:
                sym.Activate()
                doc.Regenerate()

            new_tag = DB.IndependentTag.Create(
                doc, d["type_id"], view.Id,
                d["tag_ref"], True, d["orientation"], d["tag_head"]
            )
            doc.Regenerate()

            try:
                new_tag.TagHeadPosition = d["tag_head"]
                doc.Regenerate()
            except:
                pass

            if not use_attached:
                try:
                    new_tag.LeaderEndCondition = leader_end_cond
                    doc.Regenerate()
                    new_tag.SetLeaderEnd(d["tag_ref"], d["center"])
                    doc.Regenerate()
                    new_tag.SetLeaderElbow(d["tag_ref"], d["elbow_pt"])
                    doc.Regenerate()
                    new_tag.SetLeaderEnd(d["tag_ref"], d["center"])
                    doc.Regenerate()
                except Exception as ex:
                    errors.append("Tag recreate leader: {}".format(str(ex)))

            try:
                new_tag.TagHeadPosition = d["tag_head"]
            except:
                pass

            count_ok += 1

        except Exception as ex:
            count_fail += 1
            errors.append("Tag recreate: {}".format(str(ex)))

    t.Commit()

    msg_parts = ["Fixed: {}".format(count_ok)]
    if count_skip:
        msg_parts.append("Skipped: {}".format(count_skip))
    if count_fail:
        msg_parts.append("Failed: {}".format(count_fail))
    forms.toast(" | ".join(msg_parts), title="Fix Leaders Complete")
    if errors:
        forms.alert("\n".join(errors[:10]), title="Fix Leaders Errors")




def execute_tagging(elements, mapping, profile):
    count_ok = 0
    count_fail = 0
    count_skip = 0
    created_tags = []
    active_view = doc.ActiveView

    off_perp = inches_to_feet(profile["off_updown"])
    off_along = inches_to_feet(profile["off_along"])
    extra_rad = math.radians(profile["extra_angle"])
    leader_type = profile.get("leader", "None")
    if isinstance(leader_type, bool):
        leader_type = "Attached End" if leader_type else "None"
    rot_mode = profile["rotation"]
    skip_tagged = profile.get("skip_tagged", True)
    shoulder_ft = inches_to_feet(profile.get("elbow_x", 0))

    has_leader = leader_type in ["Attached End", "Free End"]

    # Build set of already-tagged element IDs for fast lookup
    already_tagged_ids = set()
    if skip_tagged:
        try:
            existing_tags = DB.FilteredElementCollector(doc, active_view.Id)\
                .OfClass(DB.IndependentTag).ToElements()
            for et in existing_tags:
                try:
                    if hasattr(et, 'GetTaggedReferences'):
                        for ref in et.GetTaggedReferences():
                            already_tagged_ids.add(ref.ElementId.IntegerValue)
                    elif hasattr(et, 'TaggedLocalElementId'):
                        already_tagged_ids.add(et.TaggedLocalElementId.IntegerValue)
                except:
                    pass
        except:
            pass

    t = DB.Transaction(doc, "Place Multi-Category Tags")
    t.Start()

    for elem in elements:
        try:
            if not elem or not elem.Category:
                continue

            if skip_tagged and elem.Id.IntegerValue in already_tagged_ids:
                count_skip += 1
                continue

            cname = elem.Category.Name
            tag_symbol = mapping.get(cname)
            if not tag_symbol:
                count_fail += 1
                continue

            if not tag_symbol.IsActive:
                tag_symbol.Activate()
                doc.Regenerate()

            center = get_anchor_point(elem, active_view)
            if not center:
                count_fail += 1
                continue

            direction = normalize_xy(get_element_direction(elem), DB.XYZ.BasisX)
            perp = DB.XYZ(-direction.Y, direction.X, 0)
            final_pt = center + (perp * off_perp) + (direction * off_along)

            tag_axis, side_axis = get_tag_axes(direction, rot_mode, extra_rad)
            tag_angle = math.atan2(tag_axis.Y, tag_axis.X)

            tag = DB.IndependentTag.Create(
                doc,
                tag_symbol.Id,
                active_view.Id,
                DB.Reference(elem),
                has_leader,
                DB.TagOrientation.Horizontal,
                final_pt
            )

            try:
                if hasattr(tag, 'TagHeadPosition'):
                    tag.TagHeadPosition = final_pt
            except:
                pass

            doc.Regenerate()

            tag_ref = None
            try:
                if hasattr(tag, 'GetTaggedReferences'):
                    refs = tag.GetTaggedReferences()
                    if refs and refs.Count > 0:
                        tag_ref = refs[0]
            except:
                pass
            if not tag_ref:
                tag_ref = DB.Reference(elem)

            if has_leader and leader_type == "Attached End":
                try:
                    tag.LeaderEndCondition = DB.LeaderEndCondition.Attached
                    doc.Regenerate()
                except:
                    pass

            if abs(tag_angle) > 0.001:
                axis = DB.Line.CreateBound(final_pt, final_pt + DB.XYZ.BasisZ)
                DB.ElementTransformUtils.RotateElement(doc, tag.Id, axis, tag_angle)
                doc.Regenerate()

            if has_leader and leader_type == "Free End":
                try:
                    tag.LeaderEndCondition = DB.LeaderEndCondition.Free
                    doc.Regenerate()

                    final_head = final_pt
                    try:
                        if hasattr(tag, 'TagHeadPosition'):
                            final_head = tag.TagHeadPosition
                    except:
                        pass

                    elbow_pt = compute_free_end_elbow(final_head, center, tag_axis, shoulder_ft)

                    try:
                        tag.SetLeaderEnd(tag_ref, center)
                        doc.Regenerate()
                    except:
                        pass

                    try:
                        tag.SetLeaderElbow(tag_ref, elbow_pt)
                        doc.Regenerate()
                    except:
                        pass

                    try:
                        tag.SetLeaderEnd(tag_ref, center)
                        doc.Regenerate()
                    except:
                        pass

                    try:
                        if hasattr(tag, 'TagHeadPosition'):
                            tag.TagHeadPosition = final_head
                            doc.Regenerate()
                    except:
                        pass
                except:
                    pass

            count_ok += 1
            created_tags.append(tag)

        except Exception:
            count_fail += 1

    t.Commit()
    msg_parts = ["Placed: {}".format(count_ok)]
    if count_skip:
        msg_parts.append("Skipped (already tagged): {}".format(count_skip))
    if count_fail:
        msg_parts.append("Failed: {}".format(count_fail))
    forms.toast(" | ".join(msg_parts), title="Tagging Complete")
    return created_tags

if __name__ == "__main__":
    # Check if launched from ReNumber with preselected elements
    preselected = globals().get('_preselected_elements', None)
    run_tool(preselected_elements=preselected)