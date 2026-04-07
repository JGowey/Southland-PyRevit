# -*- coding: utf-8 -*-
"""
NW Shop QA - Coil & Standards (Fabrication Duct)
Read-only analyzer for MEP Fabrication parts vs. NW Division standards.

User Configurable Variables are located in the SHOP_STANDARDS dictionary below.
Runtime toggles (Project specific exceptions) are asked via the XAML Window.
"""

from __future__ import division
from pyrevit import revit, DB, forms, script
import clr
import re
import json
import os
import math

doc = revit.doc
uidoc = revit.uidoc
out = script.get_output()

# .NET List for selection sets
clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

# Config path for default settings
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "user_config.json")

# ==============================================================================
#   PART IDENTIFICATION SYSTEM - Issue #17
#   
#   Priority order for identification:
#   1. CID (Part Pattern Number) - Most definitive, unique per part in database
#   2. Alias parameter - Short codes like "3", "7", "OS", "EC"
#   3. Long-form ProductName - MAJ-style aliases like "Rct_ElbowRadius~LBS"
#   4. ELEM_FAMILY_PARAM - Fallback using part type name
#   
#   Data sources:
#   - ItemCustomId / Part Pattern Number: CID from fabrication database
#   - Alias parameter: Short codes from database
#   - ELEM_FAMILY_PARAM: Returns part type like "Straight", "Elbow"
# ==============================================================================

# Path to shared CID config file in extension root
def find_extension_root():
    """Find the DetailersTools.extension folder"""
    current = os.path.dirname(__file__)
    for _ in range(10):  # Max 10 levels up
        if os.path.basename(current) == "DetailersTools.extension":
            return current
        parent = os.path.dirname(current)
        if parent == current:  # Reached root
            break
        current = parent
    # Fallback to script directory
    return os.path.dirname(__file__)

EXTENSION_ROOT = find_extension_root()
CID_CONFIG_FILE = os.path.join(EXTENSION_ROOT, "config", "cid_config.json")

def load_cid_config():
    """
    Load CID mappings from JSON config file.
    Returns dict mapping CID (as int) to (part_type, is_preferred, description)
    """
    cid_map = {}
    if os.path.exists(CID_CONFIG_FILE):
        try:
            with open(CID_CONFIG_FILE, 'r') as f:
                config = json.load(f)
            for cid_str, entry in config.get("cid_map", {}).items():
                try:
                    cid = int(cid_str)
                    part_type = entry.get("part_type", "Unknown")
                    is_preferred = entry.get("is_preferred", True)
                    description = entry.get("description", "")
                    cid_map[cid] = (part_type, is_preferred, description)
                except (ValueError, TypeError):
                    continue
        except Exception as ex:
            print("Warning: Could not load CID config: {}".format(str(ex)))
    return cid_map

# Load CID map at startup
CID_MAP = load_cid_config()

# SHORT ALIAS CODE MAPPING (from user's alias spreadsheet)
# Maps the Alias parameter value to (PartType, IsPreferred)
# Note: Some aliases map to multiple parts - use ELEM_FAMILY_PARAM to disambiguate
ALIAS_CODE_MAP = {
    # UNAMBIGUOUS PREFERRED PARTS (Rect and Rnd palettes)
    "3":   ("Straight", True),           # Straight
    "6":   ("Tap", True),                # Boot Tap, Boot Tap - VD
    "7":   ("Elbow", True),              # Radius Elbow, Radius Elbow-45, Radius Elbow-60
    "9":   ("Elbow", True),              # Reducing Elbow
    "12":  ("Transition", True),         # Transition, Double Boot Tap, Transition Tap
    "15":  ("Ogee", True),               # Radius Offset
    "17":  ("SquareToRound", True),      # Sq to Rnd Tap, Square to Round
    "DM":  ("Accessory", True),          # DM Frame
    "EC":  ("Cap", True),                # TDC End Cap, Welded End Cap, EndCap, EndCap w-Collar
    "F45": ("Elbow", True),              # Elbow Spiral 45
    "F90": ("Elbow", True),              # Elbow Spiral 90
    "FEC": ("Cap", True),                # Furnace Cap
    "OS":  ("Offset", True),             # Mitered Offset (NOT an elbow!)
    "VD":  ("Damper", True),             # VD Round, VD Round (Standoff 2)
    "WMS": ("Accessory", True),          # Wire Mesh Screen, RND WMS End Cap
    
    # UNAMBIGUOUS NON-PREFERRED PARTS (RectAlternate palette - yellow parts)
    "23":  ("Tee", False),               # Y Branch
    
    # AMBIGUOUS ALIASES - these need ELEM_FAMILY_PARAM to disambiguate
    # "10" = Square Elbow with TV (preferred) OR Square Elbow (non-preferred)
    # "11" = Mitered Elbow (from Rect=preferred) OR Mitered Elbow (from RectAlternate=non-preferred)
    # Handled in identify_part() function below
}

# AMBIGUOUS ALIAS RESOLUTION
# Maps (Alias, ELEM_FAMILY_PARAM pattern) to (PartType, IsPreferred)
AMBIGUOUS_ALIAS_MAP = {
    # Alias 10: Square Elbow variants
    ("10", "with tv"): ("SquareElbow", True),      # Square Elbow with TV = PREFERRED (has turning vanes)
    ("10", "square elbow"): ("SquareElbow", False), # Plain Square Elbow = NON-PREFERRED
    
    # Alias 11: Mitered Elbow variants  
    # Both Rect and RectAlternate have Mitered Elbow with same alias
    # RectAlternate version is non-preferred (yellow in palette)
    ("11", "mitered elbow"): ("MiteredElbow", False),  # Treat as non-preferred by default
    ("11", "mitered elbow-60"): ("MiteredElbow", False),
}

# LONG-FORM ALIAS MAPPING (from MAJ files)
# Maps ProductName/long alias to (PartType, IsPreferred)
LONG_ALIAS_MAP = {
    # === RECT.MAJ - PREFERRED ===
    "Rct_Straight": ("Straight", True),
    "Rct_ElbowRadius": ("Elbow", True),
    "Rct_ElbowTrans": ("Elbow", True),           # Reducing elbow
    "Rct_Transition": ("Transition", True),
    "Rct_Offset": ("Ogee", True),
    "Rct_OffsetRadius": ("Ogee", True),
    "Rct_OffsetRadius2Way": ("Ogee", True),
    "Rct_TapBoot": ("Tap", True),
    "Rct_Tap45": ("Tap", True),
    "Rct_Tap": ("Tap", True),
    "Rct_ElbowDropCheek": ("Elbow", True),
    "Chg_Sq2Rnd": ("SquareToRound", True),
    
    # === RECTALTERNATE.MAJ - NOT PREFERRED ===
    "Rct_ElbowSquare": ("SquareElbow", False),
    "Rct_ElbowAngle": ("MiteredElbow", False),
    "Rct_ElbowSqThRadHeel": ("SquareElbow", False),
    "Rct_OffsetDouble": ("Ogee", False),         # Double offset - non-preferred
    "Rct_Tee": ("Tee", False),
    "Rct_TeeWye": ("Tee", False),
    "Rct_TeePants": ("Tee", False),
    
    # === RD.MAJ - PREFERRED ===
    "Rnd_Elbow": ("Elbow", True),
    "Rnd_Offset": ("Offset", True),              # Offset Mitered Round
    "Rnd_Reducer": ("Transition", True),
    "Rnd_ReducerEccentric": ("Transition", True),
    "Rnd_Coupling": ("Straight", True),
    "Rnd_Straight": ("Straight", True),
    "Rnd_SaddleLateral": ("Tap", True),
    "Rnd_SaddleConical": ("Tap", True),
    "Rnd_SaddleShoe": ("Tap", True),
    "Chg_TapRndLateralOnFlat": ("Tap", True),
    "Chg_TapRndConicalOnFlat": ("Tap", True),
    "Rnd_EndCap": ("Cap", True),
    
    # === RDBARREL.MAJ - NOT PREFERRED (Barrel Tees) ===
    "Rnd_TeeConical": ("BarrelTee", False),
    "Rnd_TeeLateral": ("BarrelTee", False),
    "Rnd_TeeReducerConical": ("BarrelTee", False),
    "Rnd_TeeReducerLateral": ("BarrelTee", False),
    "Rnd_CrossShoe": ("BarrelTee", False),
    "Rnd_CrossReducerShoe": ("BarrelTee", False),
    "Chg_TeeRect90OnRnd": ("BarrelTee", False),
    "Chg_TeeRectLateralOnRnd": ("BarrelTee", False),
    "Chg_TeeRectShoeOnRnd": ("BarrelTee", False),
}

# FAMILY NAME TO PART TYPE MAPPING (from ELEM_FAMILY_PARAM)
# This is the fallback when Alias is not available
FAMILY_NAME_MAP = {
    "straight": "Straight",
    "elbow": "Elbow",
    "transition": "Transition",
    "reducer": "Transition",
    "tee": "Tee",
    "wye": "Tee",
    "tap": "Tap",
    "takeoff": "Tap",
    "offset": "Offset",
    "cap": "Cap",
    "end cap": "Cap",
    "damper": "Damper",
    "square to round": "SquareToRound",
}

def get_alias_param(e):
    """
    Get the short Alias parameter value from a FabricationPart.
    This is the primary identifier (e.g., "3", "7", "OS", "EC").
    
    Returns: The alias string or None
    """
    if not isinstance(e, DB.FabricationPart):
        return None
    
    try:
        p = e.LookupParameter("Alias")
        if p:
            val = p.AsString() or p.AsValueString()
            if val and val.strip():
                return val.strip()
    except:
        pass
    
    return None

def get_cid(e):
    """
    Get the CID (Part Pattern Number / ItemCustomId) from a FabricationPart.
    This is the most definitive identifier - unique per part in the database.
    
    Returns: The CID as an integer, or None
    """
    if not isinstance(e, DB.FabricationPart):
        return None
    
    # Try API property first (most reliable)
    try:
        cid = e.ItemCustomId
        if cid is not None and cid > 0:
            return int(cid)
    except:
        pass
    
    # Try Part Pattern Number parameter
    try:
        p = e.LookupParameter("Part Pattern Number")
        if p:
            val = p.AsInteger() if p.StorageType == DB.StorageType.Integer else None
            if val is None:
                # Try as string and convert
                val_str = p.AsString() or p.AsValueString()
                if val_str and val_str.isdigit():
                    val = int(val_str)
            if val and val > 0:
                return val
    except:
        pass
    
    return None

def get_family_name(e):
    """
    Get the family name from ELEM_FAMILY_PARAM.
    This is used as a fallback identifier.
    
    Returns: The family name string or None
    """
    if not isinstance(e, DB.FabricationPart):
        return None
    
    try:
        p = e.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        if p:
            val = p.AsString() or p.AsValueString()
            if val and val.strip():
                return val.strip()
    except:
        pass
    
    return None

def get_product_name(e):
    """
    Get the ProductName (long-form alias) from a FabricationPart.
    This is the MAJ-style alias like "Rct_ElbowRadius~LBS".
    
    Returns: The product name string or None
    """
    if not isinstance(e, DB.FabricationPart):
        return None
    
    # Try API property first
    try:
        pn = e.ProductName
        if pn and pn.strip():
            return pn.strip()
    except:
        pass
    
    # Try parameter lookup
    for pname in ["Product Name", "ProductName", "Fabrication Product Name"]:
        try:
            p = e.LookupParameter(pname)
            if p:
                val = p.AsString() or p.AsValueString()
                if val and val.strip():
                    return val.strip()
        except:
            continue
    
    return None

def identify_part(e):
    """
    Comprehensive part identification using all available data sources.
    
    Priority order:
    1. CID (Part Pattern Number) - Most definitive, unique per part
    2. Short Alias code (e.g., "3", "7", "OS") 
    3. CID (Part Pattern Number) - auto-generated, fallback
    4. Long-form ProductName (e.g., "Rct_ElbowRadius~LBS")
    5. ELEM_FAMILY_PARAM (e.g., "Straight", "Elbow")
    
    Returns: (part_type, is_preferred) tuple or (None, None) if unknown
    """
    if not isinstance(e, DB.FabricationPart):
        return (None, None)
    
    # Get all identifiers upfront
    cid = get_cid(e)
    alias = get_alias_param(e)
    family_name = get_family_name(e)
    family_lower = family_name.lower() if family_name else ""
    
    # 0. Check API method for taps first - most reliable for taps
    try:
        if e.IsATap():
            # It's a tap - check family name for more specific type
            if "boot" in family_lower:
                return ("Tap", True)
            elif "lateral" in family_lower:
                return ("Tap", True)
            elif "conical" in family_lower:
                return ("Tap", True)
            elif "shoe" in family_lower:
                return ("Tap", True)
            return ("Tap", True)
    except:
        pass
    
    # 1. Check for ambiguous aliases first (need family name to disambiguate)
    if alias in ("10", "11"):
        # Alias 10 with "tv" in name is preferred (has turning vanes)
        if alias == "10" and ("tv" in family_lower or "turning vane" in family_lower):
            return ("SquareElbow", True)
        elif alias == "10":
            return ("SquareElbow", False)
        elif alias == "11":
            return ("MiteredElbow", False)
    
    # Alias 12 is also ambiguous - Transition vs Boot Tap vs Transition Tap
    if alias == "12":
        if "tap" in family_lower or "boot" in family_lower:
            return ("Tap", True)
        # Otherwise it's a transition
        return ("Transition", True)
    
    # 2. Try short Alias code (Southland database - most reliable)
    if alias and alias in ALIAS_CODE_MAP:
        return ALIAS_CODE_MAP[alias]
    
    # 3. Check CID_MAP (auto-generated, may have errors - use as fallback)
    if cid and cid in CID_MAP:
        part_type, is_preferred, _desc = CID_MAP[cid]
        return (part_type, is_preferred)
    
    # 4. Try long-form ProductName
    product_name = get_product_name(e)
    if product_name:
        # Clean up: remove [Category] prefix and ~Suffix
        clean_name = product_name
        if "]" in clean_name:
            clean_name = clean_name.split("]")[-1]
        if "~" in clean_name:
            clean_name = clean_name.split("~")[0]
        
        # Direct lookup
        if clean_name in LONG_ALIAS_MAP:
            return LONG_ALIAS_MAP[clean_name]
        
        # Partial match
        for key, value in LONG_ALIAS_MAP.items():
            if key in clean_name or clean_name in key:
                return value
    
    # 5. Fall back to ELEM_FAMILY_PARAM (already retrieved at top)
    if family_name:
        for key, part_type in FAMILY_NAME_MAP.items():
            if key in family_lower:
                # Family name doesn't tell us if preferred, assume True
                return (part_type, True)
    
    return (None, None)

def fitting_kind_from_alias(e):
    """
    Determine fitting kind from alias/identification.
    Wrapper for backward compatibility.
    
    Returns: Fitting kind string or None
    """
    part_type, _ = identify_part(e)
    return part_type

def is_preferred_part(e):
    """
    Check if a part is from a preferred palette.
    Non-preferred parts are from RectAlternate or RndBarrel.
    
    Returns: True if preferred, False if non-preferred, None if unknown
    """
    _, is_preferred = identify_part(e)
    return is_preferred


# ==============================================================================
#   USER CONFIGURATION / SHOP STANDARDS (DEFAULTS)
# ==============================================================================

SHOP_STANDARDS = {
    # --- Geometric Limits ---
    "max_coil_width": 60.0,
    "max_rect_straight_length": 60.0,   # Hard cap for rect straights
    "max_transition_length": 48.0,
    "min_transition_length": 12.0,
    
    "min_spiral_length": 4.0,
    "max_spiral_length_round": 240.0,
    "max_spiral_length_oval": 120.0,
    
    # --- Straight Duct & Connector Targets ---
    "target_len_tdc": 56.25,
    "target_len_sd": 59.0,
    "target_len_sd_large": 58.5,
    "min_straight_tdc": 9.0,
    
    "min_elbow_throat_length": 9.0,
    
    # --- Extensions (Straight In/Out) ---
    "ext_tdc": 2.0,
    "ext_sd": 1.0,
    "ext_flange": 1.5,
    "ext_ogee_1way": 0.0,
    
    # --- Square Elbow Throats ---
    "sq_elbow_throat_std": 6.0,          # Standard throat (6"/6")
    "sq_elbow_throat_min": 1.5,          # No shorter than 1.5"
    "sq_elbow_throat_max": 24.0,         # No longer than 24"
    "sq_elbow_throat_total_tdc": 9.0,    # Min total throat for TDC/TDC
    
    # --- Square to Round ---
    "sq2rd_body_min": 8.0,               # Min body length
    "sq2rd_body_max": 48.0,              # Max body length
    "sq2rd_ratio_max": 2.0,              # Max dia:LxW ratio (1:2)
    
    # --- Rectangular Offset (Mitered) ---
    "min_offset_length": 12.0,
    "max_offset_length": 48.0,
    
    # --- TDC Tap Requirements ---
    "tap_tdc_straight_in": 1.5,          # TDC taps need 1.5" straight in
    "tap_machine_straight_in": 0.0,      # Machine/inside taps = 0"
    
    # --- Double Wall Round ---
    "max_double_wall_round_length": 144.0,  # Weight limit
    
    # --- Short Joint Seam Threshold ---
    "short_joint_seam_threshold": 7.0,   # <7" = lap & spot, >7" = Pittsburgh OK
    
    # --- Elbows ---
    "elbow_throat_radius": 6.0,
    "elbow_split_threshold_a": 40.0,
    "elbow_split_threshold_b": 60.0,
    
    # --- Dampers ---
    "damper_buyout_height": 12.0,
    "damper_single_blade_max_width": 36.0,   # Single blade width limit
    "damper_single_blade_max_depth": 12.0,   # Single blade depth limit
    "damper_multi_blade_width": 36.0,        # Over this = multi/opposed blade
    "damper_multi_blade_depth": 14.0,        # Over this = multi/opposed blade

    # --- Tolerance ---
    "tolerance": 0.25,

    # --- Liner Widths (for multiples check) ---
    "liner_widths": {
        "spiracoustic": 48,
        "armacell_ap": 48,
        "armacell": 48,
        "ap coilflex": 48,
        "kflex": 60,
        "k-flex": 60,
        "k flex": 60
    },

    # --- Min CLR Table (Angle -> Min CLR) ---
    "elbow_min_radius_table": [
        (60.0,  6.0), (51.0,  7.0), (45.0,  8.0), (40.0,  9.0), (36.0, 10.0),
        (33.0, 11.0), (30.0, 12.0), (28.0, 13.0), (26.0, 14.0), (24.0, 15.0),
        (22.0, 16.0), (21.0, 17.0), (20.0, 18.0), (19.0, 19.0), (18.0, 20.0),
        (17.0, 21.0), (16.0, 22.0), (15.0, 24.0)
    ]
}

# ==============================================================================
#   RUNTIME STATE
# ==============================================================================

RUNTIME = {
    "db_profile": "Enterprise",
    "allow_ogee": False,
    "allow_sq2rd_straight": False,
    "allow_trans_straight": False,
    "allow_bullhead_tees": False,
}

# ==============================================================================
#   RULE TOGGLES (Enable/Disable individual rules)
# ==============================================================================

RULES_ENABLED = {
    # Rectangular Coil Fit
    "rule_coil_panel_width": True,
    "rule_rect_elbow_cheek": True,
    "rule_rect_straight_length": True,
    
    # Rectangular Elbows
    "rule_rect_elbow_min_radius": True,
    "rule_rect_elbow_throat": True,
    "rule_square_elbow_throat": True,
    
    # Round/Oval
    "rule_round_length_range": True,
    "rule_oval_length_max": True,
    "rule_lined_round_multiple": True,
    "rule_round_elbow_segmentation": True,
    "rule_double_wall_round_length": True,
    
    # Transitions & Offsets
    "rule_rect_transition_length": True,
    "rule_rect_offset_length": True,
    "rule_ogee_offset_straight": True,
    
    # Square to Round
    "rule_sq2rd_body_limits": True,
    
    # Taps
    "rule_tdc_tap_straight": True,
    "rule_no_corner_tap_tdc": True,
    
    # Straights
    "rule_rect_tdc_minjoint": True,
    "rule_short_joint_seam": True,
    "rule_fullauto_sameconn": True,
    
    # Dampers
    "rule_damper_buyout": True,
    "rule_damper_blade_sizing": True,
    
    # Policy
    "rule_exposed_note": True,
}

# ==============================================================================
#   CID DATABASE DISPLAY CLASS
# ==============================================================================

class CidEntry(object):
    """Represents a CID entry for display in the DataGrid"""
    def __init__(self, cid, part_type, is_preferred, description, alias="", service_name=""):
        self.CID = cid
        self.PartType = part_type
        self.IsPreferred = is_preferred
        self.PrefDisplay = "✓" if is_preferred else "✗"
        self.Description = description
        self.Alias = alias
        self.ServiceName = service_name

# ==============================================================================
#   WPF WINDOW CLASS
# ==============================================================================

class ShopQAWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, preselection=None):
        self.elements = preselection or []
        self._next_action = None
        self._cid_entries = []  # All CID entries
        self._filtered_cid_entries = []  # Filtered for display
        forms.WPFWindow.__init__(self, xaml_file_name)
        
        # Load Defaults if available
        self._load_defaults_from_disk()
        
        # Initialize UI *after* the window is loaded
        self.populate_ui()
        self.update_selection_info()
        
        # Initialize CID Database tab
        self._load_cid_database()

    def _load_defaults_from_disk(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    data = json.load(f)
                    if "SHOP_STANDARDS" in data:
                        SHOP_STANDARDS.update(data["SHOP_STANDARDS"])
                    if "RUNTIME" in data:
                        RUNTIME.update(data["RUNTIME"])
                    if "RULES_ENABLED" in data:
                        RULES_ENABLED.update(data["RULES_ENABLED"])
            except:
                pass

    def populate_ui(self):
        # Bind SHOP_STANDARDS values to TextBoxes by name
        for key, value in SHOP_STANDARDS.items():
            if isinstance(value, (int, float)):
                control = getattr(self, key, None)
                if control:
                    control.Text = str(value)
        
        # Set Default State from RUNTIME
        if hasattr(self, "rb_enterprise"): self.rb_enterprise.IsChecked = (RUNTIME["db_profile"] == "Enterprise")
        if hasattr(self, "rb_legacy"): self.rb_legacy.IsChecked = (RUNTIME["db_profile"] == "Legacy")
        
        if hasattr(self, "allow_ogee"): self.allow_ogee.IsChecked = RUNTIME["allow_ogee"]
        if hasattr(self, "allow_sq2rd_straight"): self.allow_sq2rd_straight.IsChecked = RUNTIME["allow_sq2rd_straight"]
        if hasattr(self, "allow_trans_straight"): self.allow_trans_straight.IsChecked = RUNTIME["allow_trans_straight"]
        if hasattr(self, "allow_bullhead_tees"): self.allow_bullhead_tees.IsChecked = RUNTIME["allow_bullhead_tees"]
        
        # Bind RULES_ENABLED to CheckBoxes
        for key, value in RULES_ENABLED.items():
            control = getattr(self, key, None)
            if control:
                control.IsChecked = value

    def update_selection_info(self):
        if hasattr(self, "SelectionCount"):
            self.SelectionCount.Text = "{} elements selected".format(len(self.elements))
            # Enable Run button only if selection exists
            if hasattr(self, "RunBtn"): self.RunBtn.IsEnabled = len(self.elements) > 0
        
    def _read_ui_to_settings(self):
        # 1. Update SHOP_STANDARDS from TextBoxes
        for key in SHOP_STANDARDS.keys():
            control = getattr(self, key, None)
            if control:
                try:
                    val = float(control.Text)
                    SHOP_STANDARDS[key] = val
                except ValueError:
                    pass 

        # 2. Update RUNTIME
        if hasattr(self, "rb_legacy"):
            RUNTIME["db_profile"] = "Legacy" if self.rb_legacy.IsChecked else "Enterprise"
        
        if hasattr(self, "allow_ogee"): RUNTIME["allow_ogee"] = self.allow_ogee.IsChecked
        if hasattr(self, "allow_sq2rd_straight"): RUNTIME["allow_sq2rd_straight"] = self.allow_sq2rd_straight.IsChecked
        if hasattr(self, "allow_trans_straight"): RUNTIME["allow_trans_straight"] = self.allow_trans_straight.IsChecked
        if hasattr(self, "allow_bullhead_tees"): RUNTIME["allow_bullhead_tees"] = self.allow_bullhead_tees.IsChecked
        
        # 3. Update RULES_ENABLED from CheckBoxes
        for key in RULES_ENABLED.keys():
            control = getattr(self, key, None)
            if control:
                RULES_ENABLED[key] = control.IsChecked

    # --- Rule Toggle Actions ---
    
    def EnableAll_Click(self, sender, e):
        for key in RULES_ENABLED.keys():
            control = getattr(self, key, None)
            if control:
                control.IsChecked = True
    
    def DisableAll_Click(self, sender, e):
        for key in RULES_ENABLED.keys():
            control = getattr(self, key, None)
            if control:
                control.IsChecked = False
    
    def ResetRules_Click(self, sender, e):
        # Reset all rules to True (default)
        for key in RULES_ENABLED.keys():
            control = getattr(self, key, None)
            if control:
                control.IsChecked = True

    # --- Profile Actions ---
    
    def Import_Click(self, sender, e):
        path = forms.pick_file(file_ext='json')
        if path:
            try:
                with open(path, 'r') as f:
                    data = json.load(f)
                    if "SHOP_STANDARDS" in data:
                        SHOP_STANDARDS.update(data["SHOP_STANDARDS"])
                    if "RUNTIME" in data:
                        RUNTIME.update(data["RUNTIME"])
                    if "RULES_ENABLED" in data:
                        RULES_ENABLED.update(data["RULES_ENABLED"])
                self.populate_ui() # Refresh UI
                forms.alert("Profile imported successfully.")
            except Exception as ex:
                forms.alert("Error importing profile: {}".format(ex))

    def Export_Click(self, sender, e):
        self._read_ui_to_settings() # ensure current UI state is captured
        path = forms.save_file(file_ext='json', default_name='shop_qa_profile')
        if path:
            data = {"SHOP_STANDARDS": SHOP_STANDARDS, "RUNTIME": RUNTIME, "RULES_ENABLED": RULES_ENABLED}
            try:
                with open(path, 'w') as f:
                    json.dump(data, f, indent=4)
                forms.alert("Profile exported successfully.")
            except Exception as ex:
                forms.alert("Error exporting profile: {}".format(ex))

    def SaveDefault_Click(self, sender, e):
        self._read_ui_to_settings()
        data = {"SHOP_STANDARDS": SHOP_STANDARDS, "RUNTIME": RUNTIME, "RULES_ENABLED": RULES_ENABLED}
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(data, f, indent=4)
            forms.alert("Settings saved as default for future runs.")
        except Exception as ex:
            forms.alert("Error saving defaults: {}".format(ex))
        except Exception as ex:
            forms.alert("Error saving defaults: {}".format(ex))

    # --- Navigation Actions ---
    def Pick_Click(self, sender, e):
        self._read_ui_to_settings()
        self._next_action = "pick"
        self.Close()

    def Run_Click(self, sender, e):
        self._read_ui_to_settings()
        self._next_action = "run"
        self.Close()

    def Cancel_Click(self, sender, e):
        self._next_action = None
        self.Close()

    # ==========================================================================
    #   CID DATABASE TAB METHODS
    # ==========================================================================
    
    def _load_cid_database(self):
        """Load CID entries from config file and populate the DataGrid"""
        self._cid_entries = []
        
        # Update config path display
        if hasattr(self, "CidConfigPath"):
            self.CidConfigPath.Text = "Config: {}".format(CID_CONFIG_FILE)
        
        # Load from JSON config
        if os.path.exists(CID_CONFIG_FILE):
            try:
                with open(CID_CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                
                cid_map = config.get("cid_map", {})
                for cid_str, entry in cid_map.items():
                    try:
                        cid = int(cid_str)
                        part_type = entry.get("part_type", "Unknown")
                        is_preferred = entry.get("is_preferred", True)
                        description = entry.get("description", "")
                        alias = entry.get("alias", "")
                        service_name = entry.get("service_name", "")
                        
                        self._cid_entries.append(CidEntry(
                            cid, part_type, is_preferred, description, alias, service_name
                        ))
                    except (ValueError, TypeError):
                        continue
                        
                # Sort by CID
                self._cid_entries.sort(key=lambda x: x.CID)
                
            except Exception as ex:
                print("Error loading CID config: {}".format(str(ex)))
        
        # Update stats and display
        self._update_cid_stats()
        self._apply_cid_filter()
    
    def _update_cid_stats(self):
        """Update the CID count statistics"""
        total = len(self._cid_entries)
        preferred = sum(1 for e in self._cid_entries if e.IsPreferred)
        non_preferred = total - preferred
        
        if hasattr(self, "CidTotalCount"):
            self.CidTotalCount.Text = "{} total".format(total)
        if hasattr(self, "CidPreferredCount"):
            self.CidPreferredCount.Text = "{} preferred".format(preferred)
        if hasattr(self, "CidNonPreferredCount"):
            self.CidNonPreferredCount.Text = "{} non-preferred".format(non_preferred)
    
    def _apply_cid_filter(self):
        """Filter CID entries based on filter box and type dropdown"""
        # Get filter text
        filter_text = ""
        if hasattr(self, "CidFilterBox") and self.CidFilterBox.Text:
            filter_text = self.CidFilterBox.Text.lower().strip()
        
        # Get filter type
        filter_type = "All"
        if hasattr(self, "CidFilterType") and self.CidFilterType.SelectedItem:
            filter_type = self.CidFilterType.SelectedItem.Content
        
        # Apply filters
        filtered = []
        for entry in self._cid_entries:
            # Text filter
            if filter_text:
                searchable = "{} {} {} {} {}".format(
                    entry.CID, entry.PartType, entry.Description, 
                    entry.Alias, entry.ServiceName
                ).lower()
                if filter_text not in searchable:
                    continue
            
            # Type filter
            if filter_type == "Preferred" and not entry.IsPreferred:
                continue
            elif filter_type == "Non-Preferred" and entry.IsPreferred:
                continue
            elif filter_type == "Elbows" and "elbow" not in entry.PartType.lower():
                continue
            elif filter_type == "Straights" and "straight" not in entry.PartType.lower():
                continue
            elif filter_type == "Transitions" and "transition" not in entry.PartType.lower():
                continue
            elif filter_type == "Taps" and "tap" not in entry.PartType.lower():
                continue
            elif filter_type == "Tees" and "tee" not in entry.PartType.lower():
                continue
            
            filtered.append(entry)
        
        self._filtered_cid_entries = filtered
        
        # Update DataGrid
        if hasattr(self, "CidDataGrid"):
            self.CidDataGrid.ItemsSource = filtered
    
    def RefreshCid_Click(self, sender, e):
        """Reload CID database from file"""
        # Also reload the global CID_MAP
        global CID_MAP
        CID_MAP = load_cid_config()
        self._load_cid_database()
    
    def OpenCidFile_Click(self, sender, e):
        """Open the CID config file in default editor"""
        import subprocess
        if os.path.exists(CID_CONFIG_FILE):
            try:
                os.startfile(CID_CONFIG_FILE)
            except:
                try:
                    subprocess.Popen(['notepad.exe', CID_CONFIG_FILE])
                except:
                    forms.alert("Could not open file:\n{}".format(CID_CONFIG_FILE))
        else:
            forms.alert("CID config file not found:\n{}\n\nRun 'Dump FabPart Info' on some parts to create it.".format(CID_CONFIG_FILE))
    
    def CidFilter_Changed(self, sender, e):
        """Handle filter text changed"""
        self._apply_cid_filter()
    
    def CidFilterType_Changed(self, sender, e):
        """Handle filter type dropdown changed"""
        self._apply_cid_filter()

# ==============================================================================
#   HELPER FUNCTIONS
# ==============================================================================

NOTE_PARAM_CANDIDATES = ["Fabrication Notes", "Fab Notes", "Notes", "Comment", "Comments"]
EXPOSED_FLAG_CANDIDATES = ["Exposed", "Is Exposed", "Exposed?"]
BOUGHT_OUT_CANDIDATES = ["Bought Out", "Is Bought Out", "Purchased"]

def feet_to_inches(x):
    try:
        return round(float(x) * 12.0, 3)
    except Exception:
        return None

def safe_float(x):
    try:
        return float(x)
    except Exception:
        return None

def get_property(el, name):
    try:
        pi = el.GetType().GetProperty(name)
        if pi and pi.CanRead:
            return pi.GetValue(el, None)
    except Exception:
        pass
    return None

def get_parameter(el, name):
    try:
        p = el.LookupParameter(name)
        return p if (p and p.HasValue) else None
    except Exception:
        return None

def get_fab_dim_val(el, name):
    """Get dimension value from FabricationPart by dimension name."""
    if not isinstance(el, DB.FabricationPart):
        return None
    try:
        # FabricationPart dimensions accessed via GetDimensions() + GetDimensionValue()
        dims = el.GetDimensions()
        if dims:
            for dim_def in dims:
                if dim_def.Name == name:
                    return el.GetDimensionValue(dim_def)
    except Exception:
        pass
    # Fallback to standard parameter
    try:
        p = el.LookupParameter(name)
        if p and p.HasValue and p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
    except:
        pass
    return None

def debug_fab_dims(el):
    """Debug: Print all fabrication dimensions for an element."""
    results = []
    try:
        # Use GetDimensions() to get definitions, GetDimensionValue() to get values
        dims = el.GetDimensions()
        if dims:
            for dim_def in dims:
                val = el.GetDimensionValue(dim_def)
                results.append("{}: {}".format(dim_def.Name, val))
            return " | ".join(results)
    except Exception as ex:
        results.append("GetDimensions error: {}".format(ex))
    
    # Fallback: list parameters with ext/straight/length
    try:
        params = el.Parameters
        for p in params:
            if p.HasValue:
                name = p.Definition.Name
                if any(x in name.lower() for x in ["ext", "straight", "length"]):
                    try:
                        val_str = p.AsValueString() or str(p.AsDouble() if p.StorageType == DB.StorageType.Double else p.AsString())
                        results.append("{}: {}".format(name, val_str))
                    except:
                        pass
    except:
        pass
    return " | ".join(results) if results else "No dims found"

def get_param_val(el, name):
    # First try standard parameter lookup
    p = get_parameter(el, name)
    if p:
        st = p.StorageType
        try:
            if st == DB.StorageType.String:
                return p.AsString()
            if st == DB.StorageType.Integer:
                return p.AsInteger()
            if st == DB.StorageType.Double:
                return p.AsDouble()
            if st == DB.StorageType.ElementId:
                eid = p.AsElementId()
                return eid.IntegerValue if eid else None
        except Exception:
            pass
    
    # For FabricationParts, also try fabrication dimension API
    fab_val = get_fab_dim_val(el, name)
    if fab_val is not None:
        return fab_val
    
    return None

def get_param_str(el, name):
    p = get_parameter(el, name)
    if not p:
        return None
    try:
        s = p.AsValueString()
        if s:
            return s
        if p.StorageType == DB.StorageType.String:
            return p.AsString()
    except Exception:
        return None
    return None

def try_names(el, names, getter):
    for nm in names:
        v = getter(el, nm)
        if v not in (None, ""):
            return v
    return None

# --- Classification Helpers ---

def category_of(e):
    if isinstance(e, DB.FabricationPart):
        return "FabPart"
    try:
        cat = e.Category.Name if e.Category else ""
        if "Fitting" in cat: return "DuctFitting"
        if "Accessory" in cat: return "DuctAccessory"
        if "Duct" in cat: return "DuctCurves"
    except Exception:
        pass
    return "Unknown"

def system_name(e):
    try:
        ms = getattr(e, "MEPSystem", None)
        if ms:
            return str(ms.Name)
    except Exception:
        pass
    return try_names(e, ["System Name", "FabricationSystemName", "Service Name", "Service"], get_param_str) or "Unknown"

def part_type_str(e):
    try:
        s = str(e.PartType)
        if s:
            return s
    except Exception:
        pass
    return get_param_str(e, "eM_Fitting Type") or "Unknown"

def is_round(e):
    shp = get_property(e, "Shape")
    if shp and "ROUND" in str(shp).upper():
        return True
    s = (try_names(e, ["Size", "OverallSize", "FreeSize"], get_param_str) or "").upper()
    if "DIA" in s or "Ø" in s or " DIA." in s:
        return True
    dia = try_names(e, ["Diameter", "Nominal Diameter", "Duct Diameter"], get_param_val)
    return dia is not None

def is_oval(e):
    shp = get_property(e, "Shape")
    if shp and "OVAL" in str(shp).upper():
        return True
    s = (try_names(e, ["Size", "OverallSize", "FreeSize"], get_param_str) or "").upper()
    return "OVAL" in s

def is_rect(e):
    if get_rect_pairs(e):
        return True
    shp = get_property(e, "Shape")
    if shp and "RECT" in str(shp).upper():
        return True
    w = try_names(e, ["NominalWidth", "Width"], get_property)
    h = try_names(e, ["NominalHeight", "Height"], get_property)
    return (w is not None and h is not None)

def shape_name(e):
    if is_rect(e):
        return "Rectangular"
    if is_round(e):
        return "Round"
    if is_oval(e):
        return "Oval"
    return "Unknown"

RECT_PAIR_RE = re.compile(r'([0-9.]+)"\s*x\s*([0-9.]+)"', re.IGNORECASE)

def _parse_rect_pairs_from_string(s):
    if not s:
        return []
    return [(safe_float(a), safe_float(b)) for (a, b) in RECT_PAIR_RE.findall(s)]

def _nominal_rect_from_props(e):
    w = get_property(e, "NominalWidth")
    h = get_property(e, "NominalHeight")
    if w is not None and h is not None:
        return [(feet_to_inches(w), feet_to_inches(h))]
    return []

def get_rect_pairs(e):
    pairs = _nominal_rect_from_props(e)
    if not pairs:
        s = try_names(e, ["Size", "OverallSize", "FreeSize"], get_param_str)
        if s:
            pairs = _parse_rect_pairs_from_string(s)
    pairs = [(w, h) for (w, h) in pairs if w and h]
    return pairs

def compute_panel_width_in(e):
    pairs = get_rect_pairs(e)
    if not pairs:
        return None
    mins = [min(w, h) for (w, h) in pairs]
    return max(mins) if mins else None

def centerline_length_in(e):
    cl = get_property(e, "CenterlineLength")
    if cl is not None and cl > 0:
        return feet_to_inches(cl)
    for nm in ["CenterlineLength", "Centerline Length", "CL Length", "Length", "eM_CenterlineLength"]:
        val = get_param_val(e, nm)
        if val is not None and val > 0:
            return val if val > 2.0 else feet_to_inches(val)
    return None

def diameter_in(e):
    d = try_names(e, ["Diameter", "Nominal Diameter", "Duct Diameter"], get_param_val)
    if d is None:
        s = try_names(e, ["Size", "OverallSize", "FreeSize"], get_param_str) or ""
        m = re.search(r'([0-9.]+)"\s*(DIA|Ø)', s.upper())
        if m:
            return safe_float(m.group(1))
    return d if d is None or d > 2.0 else feet_to_inches(d)

def get_connectors(e):
    try:
        if hasattr(e, "ConnectorManager") and e.ConnectorManager:
            return list(e.ConnectorManager.Connectors)
    except:
        pass
    return []

def connector_count(e):
    return len(get_connectors(e))

def calculate_angle_from_connectors(e):
    conns = get_connectors(e)
    if len(conns) != 2:
        return None
    try:
        v1 = conns[0].CoordinateSystem.BasisZ
        v2 = conns[1].CoordinateSystem.BasisZ
        angle_rad = v1.AngleTo(v2)
        angle_deg = math.degrees(angle_rad)
        diff = abs(180.0 - angle_deg)
        return diff
    except:
        return None

def elbow_angle_deg(e):
    # 1. Try fabrication dimensions first
    fab_angle = get_fab_dim_val(e, "Angle")
    if fab_angle is not None and fab_angle > 0.01:
        # Fab dims return radians, convert to degrees
        deg = math.degrees(fab_angle)
        if deg > 0.5:
            return deg
    
    # Also try "Round Angle" for some elbow types
    fab_angle = get_fab_dim_val(e, "Round Angle")
    if fab_angle is not None and fab_angle > 0.01:
        deg = math.degrees(fab_angle)
        if deg > 0.5:
            return deg
    
    # 2. Try standard parameters (these typically return degrees directly)
    for nm in ["Angle", "Elbow Angle", "Bend Angle", "Fabrication Angle", "Round Angle"]:
        val = get_param_val(e, nm)
        if val is not None and val > 0.1:
            # If value is small, it might be radians
            if val < 6.3:  # Less than 2*pi, likely radians
                return math.degrees(val)
            return val
    
    # 3. Try Geometry (Vectors)
    calc_ang = calculate_angle_from_connectors(e)
    if calc_ang is not None and calc_ang > 1.0:
        return calc_ang
        
    return None

def clr_in(e):
    # CLR = Centerline Radius = the radius to the center of the duct
    # For fabrication parts with "Inner Radius", CLR = Inner Radius + Width/2
    
    # First try to get actual CLR parameter
    for nm in ["CLR", "Centerline Radius", "CenterlineRadius"]:
        v = get_fab_dim_val(e, nm)
        if v is not None and v > 0.001:
            return v if v > 2.0 else feet_to_inches(v)
        v = get_param_val(e, nm)
        if v is not None and v > 0.001:
            return v if v > 2.0 else feet_to_inches(v)
    
    # If only Inner Radius available, calculate CLR = Inner Radius + Width/2
    inner_rad = get_fab_dim_val(e, "Inner Radius")
    if inner_rad is not None:
        if inner_rad < 5.0:
            inner_rad = feet_to_inches(inner_rad)
        # Get width to calculate CLR
        w = get_fab_dim_val(e, "Width")
        if w is not None and w < 5.0:
            w = feet_to_inches(w)
        if w is None:
            w = try_names(e, ["Width", "NominalWidth"], get_param_val)
            if w is not None and w < 10.0:
                w = feet_to_inches(w)
        if w is not None:
            return inner_rad + (w / 2.0)
        return inner_rad  # Fallback if no width
    
    # Try generic "Radius" as last resort
    for nm in ["Radius"]:
        v = get_fab_dim_val(e, nm)
        if v is not None and v > 0.001:
            return v if v > 2.0 else feet_to_inches(v)
        v = get_param_val(e, nm)
        if v is not None and v > 0.001:
            return v if v > 2.0 else feet_to_inches(v)
    
    return None

def inner_radius_in(e):
    """Get the inner (throat) radius of an elbow."""
    v = get_fab_dim_val(e, "Inner Radius")
    if v is not None:
        return v if v > 2.0 else feet_to_inches(v)
    # Fallback: calculate from CLR - Width/2
    clr = None
    for nm in ["CLR", "Centerline Radius", "CenterlineRadius"]:
        clr = get_fab_dim_val(e, nm) or get_param_val(e, nm)
        if clr is not None:
            break
    if clr is not None:
        if clr < 5.0:
            clr = feet_to_inches(clr)
        w = try_names(e, ["Width", "NominalWidth"], get_param_val)
        if w is not None and w < 10.0:
            w = feet_to_inches(w)
        if w is not None:
            return max(0, clr - (w / 2.0))
    return None

def normalize_conn(s):
    if not s:
        return None
    up = s.upper().strip()
    if "TDC" in up or ("DRIVE" in up and "CLEAT" in up):
        return "TDC"
    if "SLIP" in up or "S&D" in up or "S AND D" in up:
        return "S&D"
    if "FLANGE" in up:
        return "FLANGE"
    if "HEAD" in up:
        return "HEAD"
    return up

def find_connector_by_brute_force(e):
    """
    Scans ALL string parameters for known connector keywords.
    Returns (ConnA, ConnB) if found, or None.
    Use this if specific parameter names fail.
    """
    found = []
    keywords = ["TDC", "S&D", "SLIP", "FLANGE", "RAW", "CAP"]
    
    try:
        iterator = e.Parameters.GetEnumerator()
        while iterator.MoveNext():
            p = iterator.Current
            if p.StorageType == DB.StorageType.String and p.HasValue:
                val = p.AsString() or ""
                val_up = val.upper()
                # Check if value matches a connector type
                for kw in keywords:
                    if kw in val_up and len(val_up) < 20: # Sanity check length
                        # Avoid finding "TDC" in "TDC Machine" or notes
                        # We want short connector codes
                        found.append(normalize_conn(val))
                        break
            if len(found) >= 2: break
    except:
        pass
    
    if len(found) == 1: return (found[0], found[0]) # Assume same if only 1 found
    if len(found) >= 2: return (found[0], found[1])
    return (None, None)

def end_treatment_names(e):
    # 1. Try Specific Names
    names_a = ["Item Connector 1", "Connector 1", "Connector 1 Name", "C1 Name", "C1", "End A", "EndA", "End 1", "Connector A"]
    names_b = ["Item Connector 2", "Connector 2", "Connector 2 Name", "C2 Name", "C2", "End B", "EndB", "End 2", "Connector B"]
    
    a = normalize_conn(try_names(e, names_a, get_param_str))
    b = normalize_conn(try_names(e, names_b, get_param_str))
    
    if a and b: return (a, b)
    
    # 2. Fallback: Brute Force Scan
    bf_a, bf_b = find_connector_by_brute_force(e)
    return (a or bf_a, b or bf_b)

def is_lined_round(e):
    if not is_round(e):
        return False
    if liner_brand_norm(e):
        return True
    lt = try_names(e, ["LiningType", "Liner Type"], get_param_str)
    return bool(lt)

def liner_brand_norm(e):
    raw = try_names(e, ["Liner Brand", "LinerBrand", "Liner Manufacturer", "Liner"], get_param_str)
    if not raw:
        return None
    key = raw.strip().lower()
    for brand_key in SHOP_STANDARDS["liner_widths"]:
        if brand_key in key:
            return brand_key
    return None

def joint_length_in(e):
    jl = try_names(e, ["DuctJointLength", "Joint Length", "JointLength"], get_param_val)
    if jl:
        return jl if jl > 2.0 else feet_to_inches(jl)
    return centerline_length_in(e)

def get_note_text(e):
    txt = try_names(e, NOTE_PARAM_CANDIDATES, get_param_str)
    return (txt or "").strip()

def get_bool_flag(e, names):
    raw = try_names(e, names, get_param_val)
    if isinstance(raw, bool):
        return raw
    if isinstance(raw, int):
        return raw != 0
    s = try_names(e, names, get_param_str)
    if s:
        up = s.strip().upper()
        if up in ("Y", "YES", "TRUE", "1", "ON"):
            return True
        if up in ("N", "NO", "FALSE", "0", "OFF"):
            return False
    return None

def min_radius_for_angle(angle):
    if angle is None:
        return None
    table = SHOP_STANDARDS["elbow_min_radius_table"]
    for thresh, req in table:
        if angle >= thresh:
            return req
    return table[-1][1]

def get_extensions(e):
    # Returns (Extension 1, Extension 2) - various names depending on fitting type:
    # - Transitions: Extension In / Extension Out
    # - Elbows: Top Extension / Bottom Extension
    # - Sq2Rd: Extension (single, on rect end)
    # - Round: Left Extension / Right Extension (collars)
    
    # Try fabrication dimensions first (for FabricationParts)
    si = get_fab_dim_val(e, "Extension In")
    so = get_fab_dim_val(e, "Extension Out")
    
    # Try Top/Bottom Extension for elbows
    if si is None:
        si = get_fab_dim_val(e, "Top Extension")
    if so is None:
        so = get_fab_dim_val(e, "Bottom Extension")
    
    # Try Left/Right Extension for round fittings
    if si is None:
        si = get_fab_dim_val(e, "Left Extension")
    if so is None:
        so = get_fab_dim_val(e, "Right Extension")
    
    # Fallback to standard parameters if not found
    if si is None:
        si = try_names(e, ["Extension In", "Top Extension", "Straight In", "StraightIn", "Extension 1", "Left Extension"], get_param_val)
    if so is None:
        so = try_names(e, ["Extension Out", "Bottom Extension", "Straight Out", "StraightOut", "Extension 2", "Right Extension"], get_param_val)
    
    # For Sq2Rd with single "Extension" param (on the rect end)
    if si is None and so is None:
        ext_single = get_fab_dim_val(e, "Extension")
        if ext_single is None:
            ext_single = try_names(e, ["Extension"], get_param_val)
        if ext_single is not None:
            si = ext_single
            so = 0.0
    
    si = si or 0.0
    so = so or 0.0
    
    # Normalize from feet to inches if needed (Revit API returns feet)
    if si < 5.0: si = feet_to_inches(si)
    if so < 5.0: so = feet_to_inches(so)
    return (si, so)

def fitting_kind(e):
    """
    Determine fitting kind using CID alias first, then API methods and geometric analysis.
    
    Priority order (Issue fixes):
    0. CID Alias (Issue #17): Definitive identification from MAJ database
    1. API methods: IsATap(), IsAStraight() - most reliable API calls
    2. Name/string matching for specific types (pants, offset, etc.)
    3. Multi-connector check for Tees
    4. Size change detection for Transitions (BEFORE angle check)
    5. Angle-based elbow detection (ONLY if same size both ends)
    6. Mixed shape detection for SquareToRound (ONLY if not a tap)
    7. Fallback to CutType and string matching
    """
    is_fab = isinstance(e, DB.FabricationPart)
    
    # --- 0. CID ALIAS CHECK (Issue #17 - Most Definitive) ---
    # The CID alias from the MAJ database definitively identifies fitting type
    if is_fab:
        alias_kind = fitting_kind_from_alias(e)
        if alias_kind:
            return alias_kind
    
    # --- 1. USE API METHODS FIRST (Most Reliable) ---
    # Issue #7, #8: Taps with mixed connectors were misidentified as SquareToRound
    if is_fab:
        try:
            if e.IsATap():
                return "Tap"
        except:
            pass
        try:
            if e.IsAStraight():
                return "Straight"
        except:
            pass
    
    # --- 2. STRING MATCHING FOR SPECIFIC TYPES ---
    # Check name/type strings early for pants, wye, offset, etc.
    pt = (part_type_str(e) or "").lower()
    name = (e.Name or "").lower()
    signals = (pt + " " + name).strip()
    
    # Issue #6: Pants were misidentified as Transition
    if "pants" in signals or "wye" in signals or "trouser" in signals:
        return "Tee"  # Pants/Wye are branching fittings
    
    # Issue #5: Mitered offset misidentified as Elbow
    if "offset" in signals and "mitered" in signals:
        return "Ogee"
    if "ogee" in signals:
        return "Ogee"
    
    # Tap detection by name
    if "tap" in signals:
        return "Tap"
    
    # Damper detection
    if "damper" in signals:
        return "Damper"
    
    # --- 3. MULTI-CONNECTOR CHECK (Tees/Wyes/Pants) ---
    num_conns = connector_count(e)
    if num_conns > 2:
        return "Tee"
    
    # --- 4. SIZE CHANGE DETECTION (Transitions) - BEFORE angle check ---
    # Issue #4: Transition was misidentified as Elbow because angle was checked first
    conns = get_connectors(e)
    
    if is_rect(e):
        pairs = get_rect_pairs(e)
        if len(pairs) >= 2:
            w1, h1 = pairs[0]
            w2, h2 = pairs[1]
            # Check if BOTH dimensions are different (transition)
            width_diff = abs(w1 - w2) > 0.5
            height_diff = abs(h1 - h2) > 0.5
            
            if width_diff and height_diff:
                # Both dims change = Transition
                return "Transition"
            elif width_diff or height_diff:
                # Issue #2, #3: One dim same, one different = Tap (not elbow!)
                # This catches angled taps that were misidentified as elbows
                return "Tap"
    
    elif is_round(e) and conns and len(conns) >= 2:
        try:
            rads = [c.Radius for c in conns if hasattr(c, 'Radius')]
            if len(rads) >= 2 and abs(rads[0] - rads[1]) > 0.01:
                return "Transition"
        except:
            pass
    
    # --- 5. SQUARE TO ROUND CHECK (Mixed shapes) ---
    # Issue #7, #8: Only check AFTER confirming it's not a tap
    # (Taps were already caught by IsATap() or name matching above)
    if conns and len(conns) >= 2:
        has_rect = False
        has_round = False
        for c in conns:
            try:
                shape = c.Shape
                if shape == DB.ConnectorProfileType.Rectangular:
                    has_rect = True
                elif shape == DB.ConnectorProfileType.Round:
                    has_round = True
            except:
                pass
        if has_rect and has_round:
            # Double-check it's not a tap by name
            if "tap" not in signals and "boot" not in signals:
                return "SquareToRound"
            else:
                return "Tap"
    
    # --- 6. ANGLE-BASED DETECTION (Elbows vs Offsets) ---
    ang = elbow_angle_deg(e)
    if ang and ang > 15.0:  # Must be significant angle (>15°) to be an elbow
        # Has an angle - but is it an elbow or an offset?
        # Issue #5: Mitered offsets have angles but are NOT elbows
        
        # Check if it's an offset (same size both ends, shifts position)
        if "offset" in signals:
            return "Ogee"
        
        # For rect fittings, verify same size both ends for true elbow
        if is_rect(e):
            pairs = get_rect_pairs(e)
            if pairs and len(set(pairs)) == 1:
                # Same size both ends + angle = Elbow
                return "Elbow"
            elif len(pairs) >= 2:
                # Different sizes + angle = probably Tap or Transition, not Elbow
                # Already handled above, but double-check
                return "Tap"
        else:
            # Round/Oval with angle = Elbow
            return "Elbow"
    
    # --- 7. FALLBACK TO CUTTYPE ---
    if is_fab:
        try:
            ct = int(e.CutType)
            if ct == 3:
                return "Straight"
            
            # Rect with no angle, no size change = Straight
            if is_rect(e) and (ang is None or ang < 1.0):
                pairs = get_rect_pairs(e)
                if len(set(pairs)) == 1:
                    return "Straight"
            
            if ct == 2 and is_rect(e):
                return "Straight"
            if ct == 1:
                return "Fitting"
        except:
            pass
    
    # --- 8. REMAINING STRING MATCHES ---
    if "elbow" in signals or "bend" in signals:
        return "Elbow"
    if "transition" in signals or "taper" in signals:
        return "Transition"
    if "square" in signals and "round" in signals:
        return "SquareToRound"
    if "straight" in signals or "pipe" in signals or "duct" in signals:
        return "Straight"
    
    # --- 9. FINAL FALLBACK ---
    l = centerline_length_in(e)
    if l and l > 0.1:
        return "Straight"
    
    return pt.title() if pt and "default" not in pt else "Unknown"

def body_length_in(e):
    """
    Get the body length of a fitting in inches.
    
    Issue #6 Fix: For multi-branch fittings (Pants, Tees, Wyes), 
    CenterlineLength includes branch paths which is WRONG.
    Use the "Length" parameter first which gives actual body length.
    
    Priority:
    1. "Body Length" / "BodyLength" - explicit body length param
    2. "Length" parameter - for Pants/Tees (not CenterlineLength!)
    3. CenterlineLength - fallback for straights and simple fittings
    """
    # First try explicit body length params
    for nm in ["Body Length", "BodyLength"]:
        v = get_param_val(e, nm)
        if v is not None:
            return v if v > 2.0 else feet_to_inches(v)
    
    # Issue #6: For multi-branch fittings, use "Length" parameter
    # CenterlineLength for Pants = 56.91" (wrong, includes branches)
    # Length parameter for Pants = 24" (correct, actual body)
    kind = fitting_kind(e)
    if kind in ("Tee", "Transition"):
        # Try Length parameter first (it's the actual body, not centerline sum)
        for nm in ["Length"]:
            v = get_param_val(e, nm)
            if v is not None and v > 0:
                # Convert from feet if needed
                return v if v > 2.0 else feet_to_inches(v)
    
    # Fallback to centerline for straights and simple fittings
    return centerline_length_in(e)

def is_square_elbow(e):
    """Returns True if elbow has no CLR (square elbow) or CLR = 0."""
    if fitting_kind(e) != "Elbow":
        return False
    clr = clr_in(e)
    # Square elbow has no CLR or CLR = 0
    if clr is None or clr < 0.5:
        return True
    # Also check family/type name
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    if "square" in name and "round" not in name:
        return True
    return False

def is_tap(e):
    """
    Returns True if element is a tap fitting.
    
    Uses API method IsATap() first for reliability (Issue #7, #8 fix).
    """
    # Use API method first - most reliable
    is_fab = isinstance(e, DB.FabricationPart)
    if is_fab:
        try:
            if e.IsATap():
                return True
        except:
            pass
    
    # Fallback to kind detection
    kind = fitting_kind(e)
    if kind == "Tap":
        return True
    
    # Name-based detection
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    return "tap" in name

def is_double_wall(e):
    """Returns True if element is double wall duct."""
    # Check spec or type name for double wall indicators
    dw = try_names(e, ["Double Wall", "DoubleWall", "DW", "Is Double Wall"], get_param_val)
    if dw:
        return True
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    if "double" in name and "wall" in name:
        return True
    # Check insulation type for internal liner
    spec = try_names(e, ["Specification", "Insulation Type", "Lining Type"], get_param_str) or ""
    if "double" in spec.lower():
        return True
    return False

def get_round_end_diameter(e):
    """For Sq2Rd fittings, get the round end diameter."""
    # Try specific parameter names
    for nm in ["Round Diameter", "Diameter 2", "End 2 Diameter", "Collar Diameter"]:
        v = get_param_val(e, nm)
        if v is not None:
            return v if v > 2.0 else feet_to_inches(v)
    # Fallback: get diameter from connectors
    conns = get_connectors(e)
    for c in conns:
        try:
            if hasattr(c, 'Shape') and str(c.Shape).upper() == "ROUND":
                return feet_to_inches(c.Radius * 2)
        except:
            pass
    return diameter_in(e)

def tap_type(e):
    """Returns tap type: 'tdc', 'machine', 'inside', or 'unknown'."""
    if not is_tap(e):
        return None
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    if "machine" in name:
        return "machine"
    if "inside" in name:
        return "inside"
    # Check connector type
    connA, _ = end_treatment_names(e)
    if connA == "TDC":
        return "tdc"
    return "unknown"

def is_corner_tap(e):
    """Returns True if tap is positioned at a corner."""
    if not is_tap(e):
        return False
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    if "corner" in name:
        return True
    # Check position parameter if available
    pos = try_names(e, ["Tap Position", "Position", "Location"], get_param_str) or ""
    return "corner" in pos.lower()

def ogee_offset_dirs(e):
    v = try_names(e, ["Offset Directions", "OffsetDirs"], get_param_val)
    if v is not None:
        try:
            return int(v)
        except Exception:
            pass
    s = try_names(e, ["Offset Type", "Offset"], get_param_str)
    if s:
        up = s.upper()
        if "2" in up:
            return 2
        if "1" in up:
            return 1
    return None

def elbow_segment_count(e):
    v = try_names(e, ["Segments", "Segment Count", "Gores"], get_param_val)
    if v is not None:
        try:
            return int(v)
        except Exception:
            pass
    return None

# ==============================================================================
#   RESULTS AGGREGATION
# ==============================================================================

RESULTS = {
    "run_summary": {"scanned": 0, "violations": 0, "unknown": 0, "warnings": []},
    "items": []
}

def add_item(e, rule_id, rule_name, status, details, measured, required, fix, missing):
    itm = {
        "element_id": e.Id.IntegerValue,
        "category": category_of(e),
        "system": system_name(e),
        "shape": shape_name(e),
        "rule_id": rule_id,
        "rule_name": rule_name,
        "status": status,
        "details": details or "",
        "measured": measured or {},
        "required": required or {},
        "suggested_fix": fix or "",
        "missing_params": missing or []
    }
    RESULTS["items"].append(itm)
    if status == "fail":
        RESULTS["run_summary"]["violations"] += 1
    elif status == "unknown":
        RESULTS["run_summary"]["unknown"] += 1

# ==============================================================================
#   RULES IMPLEMENTATION
# ==============================================================================

def r_coil_panel_width(e):
    if not is_rect(e):
        return
    pw = compute_panel_width_in(e)
    L = centerline_length_in(e)
    limit = SHOP_STANDARDS["max_coil_width"]
    tol = SHOP_STANDARDS["tolerance"]

    if pw is None:
        add_item(e, "R-COIL-PANEL-WIDTH", "Rectangular coil-fit: across-coil check",
                 "unknown", "Missing rectangular sizes.", {}, {}, "Expose sizes.", ["rect_pairs"])
        return
    if pw > limit + tol:
        details = "Panel width %.2f in > %.2f in" % (pw, limit)
        if L and L > limit + tol:
            details = "Panel larger than coil limit (width %.2f, length %.2f)" % (pw, L)
        add_item(e, "R-COIL-PANEL-WIDTH", "Rectangular coil-fit check",
                 "fail", details, {"panel_width": pw, "length": L}, {"max_panel_width": limit},
                 "Reduce across-coil dimension (rotate or split).", [])
    else:
        add_item(e, "R-COIL-PANEL-WIDTH", "Rectangular coil-fit check",
                 "pass", "Fits across coil.", {"panel_width": pw}, {"max_panel_width": limit}, "", [])

def r_rect_elbow_cheek_coil(e):
    if fitting_kind(e) != "Elbow" or not is_rect(e):
        return
    pairs = get_rect_pairs(e)
    if not pairs: return
    max_h = max(h for (w, h) in pairs)
    limit = SHOP_STANDARDS["max_coil_width"]
    tol = SHOP_STANDARDS["tolerance"]

    if max_h > limit + tol:
        add_item(e, "R-RECT-ELBOW-CHEEK-COIL", "Rect radius elbows: cheek fit",
                 "fail", "Cheek %.2f in > %.2f in" % (max_h, limit),
                 {"max_height": max_h}, {"max": limit}, "Reduce height or split cheek.", [])
    else:
        add_item(e, "R-RECT-ELBOW-CHEEK-COIL", "Rect radius elbows: cheek fit",
                 "pass", "Cheek fits.", {"max_height": max_h}, {"max": limit}, "", [])

def r_round_length_range(e):
    """
    Round/spiral length range check.
    
    Issue #1 Fix: Only apply spiral length limits to STRAIGHT round pipe.
    Collars and round fittings are made with lock seam rolled metal, NOT spiral.
    The 4" minimum applies to spiral pipe only, not collars.
    """
    if not is_round(e):
        return
    
    # Only apply to STRAIGHTS - not fittings/collars
    # Issue #1: Collars (short round pieces on elbows/transitions) are NOT spiral
    kind = fitting_kind(e)
    if kind != "Straight":
        return  # Skip elbows, transitions, taps, etc. - they use lock seam, not spiral
    
    # Also check if it's actually a straight via API
    is_fab = isinstance(e, DB.FabricationPart)
    if is_fab:
        try:
            if not e.IsAStraight():
                return  # Not a straight, skip spiral length check
        except:
            pass
    
    L = centerline_length_in(e)
    min_l = SHOP_STANDARDS["min_spiral_length"]
    max_l = SHOP_STANDARDS["max_spiral_length_round"]
    tol = SHOP_STANDARDS["tolerance"]

    if L is None:
        return
    if L + 1e-6 < min_l or L > max_l + tol:
        add_item(e, "R-ROUND-LENGTH-RANGE", "Round/spiral length range",
                 "fail", "Length %.2f in outside %.0f-%.0f in." % (L, min_l, max_l),
                 {"length": L}, {"min": min_l, "max": max_l}, "Adjust length within range.", [])
    else:
        add_item(e, "R-ROUND-LENGTH-RANGE", "Round/spiral length range",
                 "pass", "Within range.", {"length": L}, {"min": min_l, "max": max_l}, "", [])

def r_oval_length_max(e):
    if not is_oval(e):
        return
    L = centerline_length_in(e)
    max_l = SHOP_STANDARDS["max_spiral_length_oval"]
    tol = SHOP_STANDARDS["tolerance"]

    if L is None:
        return
    if L > max_l + tol:
        add_item(e, "R-OVAL-LENGTH-MAX", "Oval length max",
                 "fail", "Length %.2f in > %.0f in." % (L, max_l),
                 {"length": L}, {"max": max_l}, "Split into shorter pieces.", [])
    else:
        add_item(e, "R-OVAL-LENGTH-MAX", "Oval length max",
                 "pass", "Within limits.", {"length": L}, {"max": max_l}, "", [])

def r_lined_round_multiple(e):
    if not (is_round(e) and is_lined_round(e)):
        return
    brand = liner_brand_norm(e)
    lw = SHOP_STANDARDS["liner_widths"].get(brand or "", None)
    jl = joint_length_in(e)
    
    if brand is None or lw is None: return 
    if jl is None: return

    k = jl / float(lw)
    if abs(k - round(k)) > 1e-4:
        add_item(e, "R-ROUND-LINER-MULTIPLE", "Round lined: joint multiple of liner",
                 "fail", "Joint %.2f in not multiple of %d in" % (jl, lw),
                 {"joint_length": jl, "liner_width": lw}, {"multiple_of": lw},
                 "Adjust joint length to a %d inch multiple." % lw, [])
    else:
        add_item(e, "R-ROUND-LINER-MULTIPLE", "Round lined: joint multiple of liner",
                 "pass", "Aligned.", {"joint_length": jl}, {"multiple_of": lw}, "", [])

def r_rect_transition_length(e):
    if fitting_kind(e) != "Transition" or not is_rect(e):
        return
    
    # Skip taps that may be misidentified as transitions (e.g., Double Boot Tap with Alias 12)
    # Check API method first
    try:
        if e.IsATap():
            return
    except:
        pass
    
    # Also check family name for tap patterns
    family_name = get_family_name(e) or ""
    fn_lower = family_name.lower()
    if "tap" in fn_lower or "boot" in fn_lower:
        return
    
    L = body_length_in(e)
    min_l = SHOP_STANDARDS["min_transition_length"]
    max_l = SHOP_STANDARDS["max_transition_length"]
    tol = SHOP_STANDARDS["tolerance"]

    if L is None: return

    if min_l - tol <= L <= max_l + tol:
        add_item(e, "R-RECT-TRANS-LENGTH", "Rect transition length limits",
                 "pass", "OK.", {"length": L}, {"min": min_l, "max": max_l}, "", [])
    else:
        add_item(e, "R-RECT-TRANS-LENGTH", "Rect transition length limits",
                 "fail", "Length %.2f in outside %.0f-%.0f in." % (L, min_l, max_l), 
                 {"length": L}, {"min": min_l, "max": max_l}, "Resize body length.", [])

def r_rect_elbow_min_radius(e):
    if fitting_kind(e) != "Elbow" or not is_rect(e):
        return
    ang = elbow_angle_deg(e)
    r = clr_in(e)
    tol = SHOP_STANDARDS["tolerance"]

    if ang is None or r is None: return
    req = min_radius_for_angle(ang)
    
    if r + tol < req:
        add_item(e, "R-ELBOW-MIN-RADIUS", "Rect/radius elbow min CLR",
                 "fail", "CLR %.2f in < %.2f in min" % (r, req),
                 {"angle": ang, "clr": r}, {"min_clr": req}, "Increase CLR.", [])
    else:
        add_item(e, "R-ELBOW-MIN-RADIUS", "Rect/radius elbow min CLR",
                 "pass", "OK.", {"clr": r}, {"min_clr": req}, "", [])

def r_rect_elbow_throat(e):
    """
    Rule: Rectangular elbows with TDC on BOTH ends must satisfy minimum throat length (9").
    
    Issue #2-5 Fixes:
    - Only apply to TRUE elbows (not taps, transitions, offsets)
    - Only apply when BOTH connectors are TDC
    - Taps with one TDC end don't need 9" throat
    - Mitered offsets with angles are NOT elbows
    - Skip very shallow angles (< 15°) - those aren't real elbows
    """
    kind = fitting_kind(e)
    
    # Must be a true Elbow AND rectangular
    if kind != "Elbow" or not is_rect(e):
        return
    
    # Issue #2-5: Check that BOTH connectors are TDC
    # The 9" minimum is only for TDC-to-TDC connections
    # Issue #18: For rectangular parts, assume TDC if connector type unknown
    connA, connB = end_treatment_names(e)
    # Treat unknown ("-" or None) as TDC for rectangular - it's the default
    connA_is_tdc = connA == "TDC" or connA in (None, "-", "")
    connB_is_tdc = connB == "TDC" or connB in (None, "-", "")
    # Only skip if we KNOW it's not TDC (e.g., slip, flange, raw)
    if not connA_is_tdc or not connB_is_tdc:
        return  # Skip if explicitly non-TDC on either end
    
    angle = elbow_angle_deg(e)
    if angle is None:
        return
    
    # Skip very shallow angles - not real elbows
    # At < 15°, this is essentially a straight piece with minor deflection
    if angle < 15.0:
        return
    
    si, so = get_extensions(e)
    
    w = try_names(e, ["Width", "NominalWidth"], get_param_val)
    if w is not None and w < 10.0: w = feet_to_inches(w)
    if w is None:
        pairs = get_rect_pairs(e)
        if pairs: w = pairs[0][0]
    
    if w is None: return
    
    # Get Inner Radius (throat radius) - this is the inside edge radius
    inner_rad = inner_radius_in(e)
    if inner_rad is None:
        inner_rad = 0.0  # Square elbow has 0 inner radius
    
    # Arc Length along inside throat = r * theta (where theta is in radians)
    arc_len = inner_rad * math.radians(angle)
    
    total_throat = si + so + arc_len
    limit = SHOP_STANDARDS["min_elbow_throat_length"]
    tol = SHOP_STANDARDS["tolerance"]
    
    if total_throat + tol < limit:
        add_item(e, "R-RECT-ELBOW-THROAT", "Rect Elbow throat length >= %.1f\" (TDC/TDC)" % limit,
                 "fail", "Throat total %.2f in < %.1f in (Ext: %.1f+%.1f, Arc: %.1f, InnerR: %.1f)" % (total_throat, limit, si, so, arc_len, inner_rad or 0),
                 {"throat_total": total_throat}, {"min_throat": limit}, "Increase straights or radius.", [])
    else:
        add_item(e, "R-RECT-ELBOW-THROAT", "Rect Elbow throat length >= %.1f\" (TDC/TDC)" % limit,
                 "pass", "OK.", {"throat_total": total_throat}, {"min_throat": limit}, "", [])

def r_round_elbow_segmentation(e):
    if fitting_kind(e) != "Elbow" or not is_round(e):
        return
    d = diameter_in(e)
    ang = elbow_angle_deg(e)
    segs = elbow_segment_count(e)
    tol = SHOP_STANDARDS["tolerance"]

    if d is None or ang is None or abs(ang - 90.0) > 1.0:
        return
    if segs is None: return

    thresh_a = SHOP_STANDARDS["elbow_split_threshold_a"]
    thresh_b = SHOP_STANDARDS["elbow_split_threshold_b"]

    need = 1
    if thresh_a - tol <= d < thresh_b - tol:
        need = 2
    elif d >= thresh_b - tol:
        need = 3
    
    if need > 1 and segs < need:
        add_item(e, "R-ELBOW-SEGMENTATION", "Round 90 deg elbow segmentation",
                 "fail", "Dia %.0f in requires %d segments, has %d." % (d, need, segs),
                 {"diameter": d, "segments": segs}, {"required_segments": need}, "Use required segmentation.", [])
    else:
        add_item(e, "R-ELBOW-SEGMENTATION", "Round 90 deg elbow segmentation",
                 "pass", "OK.", {"segments": segs}, {"required_segments": need}, "", [])

def r_ogee_offset_straight(e):
    if fitting_kind(e) not in ("Ogee", "Offset") or not is_rect(e):
        return
    dirs = ogee_offset_dirs(e)
    connA, connB = end_treatment_names(e)
    
    si, so = get_extensions(e)
    
    tol = SHOP_STANDARDS["tolerance"]

    if dirs is None: return
    
    if int(dirs) == 1:
        req = SHOP_STANDARDS["ext_ogee_1way"]
        if (si is not None and si < req - tol) or (so is not None and so < req - tol):
             add_item(e, "R-OGEE-OFFSET-STRAIGHT", "OGEE straight requirement",
                 "fail", "1-Way requires %.1f straight." % req, {}, {}, "", [])
        else:
             add_item(e, "R-OGEE-OFFSET-STRAIGHT", "OGEE straight requirement",
                 "pass", "OK.", {}, {}, "", [])
        return

    def get_req(conn):
        if conn == "TDC": return SHOP_STANDARDS["ext_tdc"]
        if conn == "S&D": return SHOP_STANDARDS["ext_sd"]
        return 0.0

    needA = get_req(connA)
    needB = get_req(connB)

    ok = (si is not None and si + tol >= needA) and (so is not None and so + tol >= needB)
    if not ok:
        add_item(e, "R-OGEE-OFFSET-STRAIGHT", "OGEE straight requirement",
                 "fail", "Requires %.1f/%.1f in straight (A/B)." % (needA, needB),
                 {"straight_in": si, "straight_out": so}, {"req_A": needA, "req_B": needB},
                 "Add straight per connection rule.", [])
    else:
        add_item(e, "R-OGEE-OFFSET-STRAIGHT", "OGEE straight requirement",
                 "pass", "OK.", {}, {}, "", [])

def r_rect_tdc_minjoint(e):
    if not is_rect(e): return
    if fitting_kind(e) in ("Elbow", "Transition", "Ogee", "SquareToRound"): return
    
    connA, connB = end_treatment_names(e)
    if connA == "TDC" and connB == "TDC":
        L = centerline_length_in(e)
        min_l = SHOP_STANDARDS["min_straight_tdc"]
        if L is not None and L + 1e-6 < min_l:
             add_item(e, "R-RECT-TDC-MINJOINT", "Rect TDC/TDC straight min length",
                     "fail", "Length %.2f in < %.1f in." % (L, min_l),
                     {"length": L}, {"min_length": min_l}, "Increase straight length.", [])

def r_rect_straight_length_limit(e):
    if not is_rect(e): return
    if fitting_kind(e) != "Straight": return

    L = centerline_length_in(e)
    limit = SHOP_STANDARDS["max_rect_straight_length"]
    tol = SHOP_STANDARDS["tolerance"]

    if L is None: return

    if L > limit + tol:
         add_item(e, "R-RECT-LENGTH-MAX", "Rect straight max length",
                 "fail", "Length %.2f in > %.2f in cap." % (L, limit),
                 {"length": L}, {"max": limit}, "Split straight to fit coil standard.", [])
    else:
         add_item(e, "R-RECT-LENGTH-MAX", "Rect straight max length",
                 "pass", "OK.", {"length": L}, {"max": limit}, "", [])

def r_no_straight_sq2rd_trans(e):
    """
    Check for straights on transitions and square-to-round fittings.
    
    Issue #8 Fix: Skip taps - they're not Sq2Rd even if they have
    mixed round/rect connectors. Taps with extensions (like MVD boot taps)
    are perfectly valid.
    """
    kind = fitting_kind(e)
    if kind not in ("Transition", "SquareToRound"): return
    
    # Issue #7, #8: Double-check it's not actually a tap
    # Taps with mixed connectors were misidentified as SquareToRound
    if is_tap(e):
        return  # Skip taps - they can have extensions (MVD, etc.)
    
    # Round transitions use collars (Left/Right Extension) which are allowed
    # Only flag RECTANGULAR transitions
    if kind == "Transition" and is_round(e):
        return  # Round transitions with collars are OK
    
    si, so = get_extensions(e)
    has_straight = (si or 0.0) > 0.01 or (so or 0.0) > 0.01

    if kind == "SquareToRound":
        # Only check the rectangular end extension, not the round end collar
        # The single "Extension" parameter is on the rect end
        # Round end collar is typically "Collar" parameter, not flagged
        allowed = RUNTIME["allow_sq2rd_straight"]
        if has_straight and not allowed:
             add_item(e, "R-NO-STRAIGHT-SQ2RD", "No straight on sq2rd (rect end)",
                     "fail", "Straight present on rect end, not allowed.", {}, {"allowed": False}, "Remove straight.", [])
    elif kind == "Transition":
        # Only rectangular transitions reach here (round filtered above)
        allowed = RUNTIME["allow_trans_straight"]
        if has_straight and not allowed:
             add_item(e, "R-NO-STRAIGHT-TRANS", "No straight on rect transitions",
                     "fail", "Straight present, not allowed.", {}, {"allowed": False}, "Remove straight.", [])

def r_no_ogee_unless_allowed(e):
    """
    Only flag TRUE radius ogees (S-curve offsets using radius elbows).
    
    TRUE Ogees to block:
    - Alias "15" = Radius Offset
    - CID 330 = Radius 2 Way Offset
    - ProductName containing "OffsetRadius" or "Rct_OffsetRadius"
    - Family name containing "Radius 2 Way" or "Radius Offset"
    
    NOT Ogees (should NOT be flagged):
    - Mitered Offsets (Alias "OS") 
    - Reducer Offsets (different category)
    - Square Offsets (use square elbows)
    - Round Mitered Offsets
    """
    if RUNTIME["allow_ogee"]:
        return
    
    # Check Alias - only "15" is true radius ogee
    alias = get_alias_param(e)
    if alias == "15":
        add_item(e, "R-NO-OGEE", "No radius ogee offsets unless allowed",
                 "fail", "Radius ogee offset (Alias 15).", {}, {"allowed": False}, "Replace with mitered offset.", [])
        return
    
    # Check CID for known radius ogee CIDs
    cid = get_cid(e)
    if cid in (9, 330):  # 9 = Radius Offset (Alias 15), 330 = Radius 2 Way Offset
        add_item(e, "R-NO-OGEE", "No radius ogee offsets unless allowed",
                 "fail", "Radius ogee offset (CID {}).".format(cid), {}, {"allowed": False}, "Replace with mitered offset.", [])
        return
    
    # Check family name for radius offset patterns
    family_name = get_family_name(e)
    if family_name:
        fn_lower = family_name.lower()
        if "radius 2 way" in fn_lower or "radius offset" in fn_lower:
            add_item(e, "R-NO-OGEE", "No radius ogee offsets unless allowed",
                     "fail", "Radius ogee offset detected.", {}, {"allowed": False}, "Replace with mitered offset.", [])
            return
    
    # Check ProductName for radius offset patterns
    product_name = get_product_name(e)
    if product_name:
        pn_lower = product_name.lower()
        # Only flag if explicitly a RADIUS offset (not just any offset)
        if "offsetradius" in pn_lower or "radius offset" in pn_lower or "radius 2 way" in pn_lower:
            add_item(e, "R-NO-OGEE", "No radius ogee offsets unless allowed",
                     "fail", "Radius ogee offset detected.", {}, {"allowed": False}, "Replace with mitered offset.", [])

def is_barrel_tee(e):
    """
    Returns True if element is a barrel tee (bullhead) from RdBarrel.MAJ palette.
    
    Issue #17: Now uses CID alias as definitive identification first.
    
    Barrel tees have:
    - CID alias containing: Rnd_Tee*, Chg_TeeRect*, Rnd_Cross* patterns
    - 2 round connectors that are COAXIAL (the straight-through barrel)
    - These can be different sizes (reducing barrel)
    - A TAP coming off the barrel
    
    Trousers/Y Gores have:
    - NO two connectors that are coaxial - all branches spread apart
    
    Detection Priority:
    1. CID alias matching barrel tee patterns (most definitive)
    2. Geometric: Check if ANY two round connectors share the same centerline
    """
    # --- 1. PRODUCT NAME CHECK (Issue #17 - Most Definitive) ---
    # The ProductName from RdBarrel.MAJ definitively identifies barrel tees
    product_name = get_product_name(e)
    if product_name:
        # Extract the core name without category prefix and suffix
        core_alias = product_name
        if "]" in core_alias:
            core_alias = core_alias.split("]")[-1]
        if "~" in core_alias:
            core_alias = core_alias.split("~")[0]
        
        # Barrel tee patterns from RdBarrel.MAJ
        barrel_tee_patterns = [
            "Rnd_Tee",           # Rnd_TeeConical, Rnd_TeeLateral, etc.
            "Chg_TeeRect",       # Chg_TeeRect90OnRnd, Chg_TeeRectLateralOnRnd, etc.
            "Rnd_Cross",         # Rnd_CrossShoe, Rnd_CrossReducerShoe
            "Rnd_TeeReducer",    # Rnd_TeeReducerConical, Rnd_TeeReducerLateral
        ]
        
        for pattern in barrel_tee_patterns:
            if pattern in core_alias:
                return True
        
        # Saddle taps are NOT barrel tees
        saddle_patterns = ["Saddle", "SaddleLateral", "SaddleConical", "SaddleShoe"]
        for pattern in saddle_patterns:
            if pattern in core_alias:
                return False
    
    # --- 2. GEOMETRIC ANALYSIS (Fallback) ---
    conns = get_connectors(e)
    if len(conns) < 3:
        return False  # Must have 3+ connectors to be a tee with tap
    
    # Get all round connectors
    round_conns = []
    for c in conns:
        try:
            if hasattr(c, 'Shape') and "ROUND" in str(c.Shape).upper():
                round_conns.append(c)
        except:
            pass
    
    if len(round_conns) < 2:
        return False  # Need at least 2 round connectors for barrel
    
    # Check if ANY two round connectors are coaxial (share same centerline)
    for i in range(len(round_conns)):
        for j in range(i + 1, len(round_conns)):
            try:
                c1 = round_conns[i]
                c2 = round_conns[j]
                
                # Get direction vectors (BasisZ points outward from connector)
                v1 = c1.CoordinateSystem.BasisZ
                v2 = c2.CoordinateSystem.BasisZ
                
                # Check if anti-parallel (facing opposite directions)
                dir_dot = v1.X * v2.X + v1.Y * v2.Y + v1.Z * v2.Z
                if dir_dot > -0.95:  # Not anti-parallel
                    continue
                
                # Get vector between connector origins
                o1 = c1.Origin
                o2 = c2.Origin
                dx = o2.X - o1.X
                dy = o2.Y - o1.Y
                dz = o2.Z - o1.Z
                
                # Normalize the origin-to-origin vector
                dist = (dx*dx + dy*dy + dz*dz) ** 0.5
                if dist < 0.001:  # Origins too close
                    continue
                dx /= dist
                dy /= dist
                dz /= dist
                
                # Check if origin-to-origin vector is parallel to connector direction
                # If coaxial, the line between origins aligns with connector directions
                align_dot = abs(dx * v1.X + dy * v1.Y + dz * v1.Z)
                
                if align_dot > 0.95:  # Coaxial = barrel tee
                    return True
                    
            except:
                continue
    
    return False  # No coaxial pair found = Trouser / Y Gore type

def r_no_bullhead_tees(e):
    """
    Flag barrel tees (bullheads) from RdBarrel.MAJ palette.
    These have a round barrel (2 coaxial round connectors) with a tap.
    
    Includes:
    - Parts with Kind="Tee" that are barrel tees
    - Parts with Kind="Tap" but family name contains "Tee" (e.g., "Tee Rect Boot")
    
    Standard round tees (Rd.MAJ: Y Gore, Trouser) are acceptable - branches spread.
    Standard rectangular tees (Rect.MAJ) are acceptable.
    GRD Cans / Plenum Boxes are acceptable (< 3 connectors).
    """
    if RUNTIME["allow_bullhead_tees"]: return
    
    kind = fitting_kind(e)
    family_name = get_family_name(e) or ""
    family_lower = family_name.lower()
    
    # Check if this is a barrel tee configuration
    is_tee_kind = (kind == "Tee")
    is_tap_with_tee_name = (kind == "Tap" and "tee" in family_lower)
    
    if not is_tee_kind and not is_tap_with_tee_name:
        return
    
    # Must have 3+ connectors to be an actual tee
    if connector_count(e) < 3: return
    
    # Check if it's a barrel tee (any two round connectors are coaxial)
    # OR if family name explicitly indicates barrel tee pattern
    is_barrel = is_barrel_tee(e)
    
    # Also catch by family name patterns (Tee Rect Boot, Tee wReducer, etc.)
    barrel_family_patterns = [
        "tee rect boot",
        "tee wreducer rect boot", 
        "tee rect 90",
        "tee rect lateral",
        "tee conical",
        "tee wreducer conical",
        "tee straight",
        "tee wreducer straight",
        "tee shoe",
        "tee wreducer shoe",
        "tee lateral",
        "tee wreducer lateral",
    ]
    
    for pattern in barrel_family_patterns:
        if pattern in family_lower:
            is_barrel = True
            break
    
    if is_barrel:
        add_item(e, "R-NO-BULLHEAD-TEES", "No Bullhead Tees/Pants",
                 "fail", "Barrel tee (round straight-through with tap) detected.", {}, {"allowed": False}, 
                 "Replace with Y Gore or standard configuration.", [])

def r_damper_buyout(e):
    if fitting_kind(e) != "Damper": return
    
    h = try_names(e, ["Height", "NominalHeight", "Diameter", "Nominal Diameter"], get_param_val)
    if h is None: return
    h_in = h if h > 2.0 else feet_to_inches(h)
    
    limit = SHOP_STANDARDS["damper_buyout_height"]
    if h_in > limit:
        is_bought = get_bool_flag(e, BOUGHT_OUT_CANDIDATES)
        if not is_bought:
             add_item(e, "R-DAMPER-BUYOUT", "Damper Buyout Check",
                 "fail", "Damper > %.1f in must be Bought Out." % limit, 
                 {"height": h_in}, {"buy_out_required": True}, "Set 'Bought Out' parameter to Yes.", [])
        else:
             add_item(e, "R-DAMPER-BUYOUT", "Damper Buyout Check",
                 "pass", "Marked as Bought Out.", {}, {}, "", [])

def r_exposed_note(e):
    flag = get_bool_flag(e, EXPOSED_FLAG_CANDIDATES)
    if flag is None or flag is False: return
    note = get_note_text(e)
    if "EXPOSED" not in (note or "").upper():
        add_item(e, "R-EXPOSED-NOTE", "Exposed pieces require note",
                 "fail", "Missing 'Exposed' in notes.", {}, {"notes_contains": "Exposed"}, "Add note.", [])

def r_fullauto_sameconn(e):
    if part_type_str(e) and "Straight" not in part_type_str(e) and fitting_kind(e) != "Unknown": return
    connA, connB = end_treatment_names(e)
    if connA and connB and connA != connB:
        add_item(e, "R-FULLAUTO-SAMECONN", "Full (Auto) joints same connection",
                 "fail", "Ends differ: %s / %s" % (connA, connB), {}, {"same": True}, "Match connections.", [])

# ==============================================================================
#   NEW CONSTRUCTABILITY RULES (NW SMS Standards)
# ==============================================================================

def r_sq2rd_body_limits(e):
    """
    Square to Round: Body length 8"-48", dia:LxW ratio max 1:2
    
    Issue #7, #8 Fix: Skip taps - they're not Sq2Rd even if they have
    mixed round/rect connectors. Use is_tap() to filter them out.
    """
    if fitting_kind(e) != "SquareToRound":
        return
    
    # Issue #7, #8: Double-check it's not actually a tap
    # Taps with mixed connectors were misidentified as SquareToRound
    if is_tap(e):
        return  # Skip taps
    
    L = body_length_in(e)
    min_l = SHOP_STANDARDS["sq2rd_body_min"]
    max_l = SHOP_STANDARDS["sq2rd_body_max"]
    ratio_max = SHOP_STANDARDS["sq2rd_ratio_max"]
    tol = SHOP_STANDARDS["tolerance"]
    
    if L is None:
        return
    
    # Check body length limits
    if L < min_l - tol:
        add_item(e, "R-SQ2RD-BODY-MIN", "Sq2Rd body length min",
                 "fail", "Body %.2f in < %.0f in min." % (L, min_l),
                 {"body_length": L}, {"min": min_l}, "Increase body length.", [])
    elif L > max_l + tol:
        add_item(e, "R-SQ2RD-BODY-MAX", "Sq2Rd body length max",
                 "fail", "Body %.2f in > %.0f in max." % (L, max_l),
                 {"body_length": L}, {"max": max_l}, "Reduce body length or split.", [])
    else:
        add_item(e, "R-SQ2RD-BODY-LENGTH", "Sq2Rd body length",
                 "pass", "OK.", {"body_length": L}, {"min": min_l, "max": max_l}, "", [])
    
    # Check dia:LxW ratio (round diameter should not be too small relative to rect)
    dia = get_round_end_diameter(e)
    pairs = get_rect_pairs(e)
    if dia and pairs:
        # Get rect dimensions (use first pair - rect end)
        w, h = pairs[0]
        rect_area = w * h
        # Ratio check: dia^2 should be >= (LxW) / ratio_max^2
        # i.e., dia:sqrt(LxW) should not exceed 1:2
        min_dia = min(w, h) / ratio_max
        if dia < min_dia - tol:
            add_item(e, "R-SQ2RD-RATIO", "Sq2Rd dia:rect ratio",
                     "fail", "Dia %.1f in too small for %.0fx%.0f rect (1:%.0f max)." % (dia, w, h, ratio_max),
                     {"diameter": dia, "rect_w": w, "rect_h": h}, {"ratio_max": ratio_max},
                     "Increase round diameter or reduce rect size.", [])

def r_square_elbow_throat(e):
    """Square elbows: Standard 6"/6" throats, min 1.5", max 24", 9" total for TDC/TDC"""
    if not is_square_elbow(e) or not is_rect(e):
        return
    
    # Skip very shallow angles - not real elbows
    angle = elbow_angle_deg(e)
    if angle is not None and angle < 15.0:
        return
    
    si, so = get_extensions(e)
    connA, connB = end_treatment_names(e)
    
    throat_min = SHOP_STANDARDS["sq_elbow_throat_min"]
    throat_max = SHOP_STANDARDS["sq_elbow_throat_max"]
    throat_total_tdc = SHOP_STANDARDS["sq_elbow_throat_total_tdc"]
    tol = SHOP_STANDARDS["tolerance"]
    
    total_throat = (si or 0.0) + (so or 0.0)
    
    # Check individual throat limits
    for throat, name in [(si, "Straight In"), (so, "Straight Out")]:
        if throat is not None:
            if throat > 0 and throat < throat_min - tol:
                add_item(e, "R-SQ-ELBOW-THROAT-MIN", "Square elbow throat min",
                         "fail", "%s %.2f in < %.1f in min." % (name, throat, throat_min),
                         {"throat": throat}, {"min": throat_min}, "Increase throat to min %.1f in." % throat_min, [])
            elif throat > throat_max + tol:
                add_item(e, "R-SQ-ELBOW-THROAT-MAX", "Square elbow throat max",
                         "fail", "%s %.2f in > %.1f in max." % (name, throat, throat_max),
                         {"throat": throat}, {"max": throat_max}, "Reduce throat or use radius elbow.", [])
    
    # Check TDC/TDC total throat requirement
    if connA == "TDC" and connB == "TDC":
        if total_throat < throat_total_tdc - tol:
            add_item(e, "R-SQ-ELBOW-THROAT-TDC", "Square elbow TDC/TDC min total throat",
                     "fail", "Total throat %.2f in < %.1f in (TDC/TDC requires 9\" min)." % (total_throat, throat_total_tdc),
                     {"throat_in": si, "throat_out": so, "total": total_throat}, {"min_total": throat_total_tdc},
                     "Increase throats (e.g., 3/6, 4.5/4.5, 8/1).", [])
        else:
            add_item(e, "R-SQ-ELBOW-THROAT-TDC", "Square elbow TDC/TDC total throat",
                     "pass", "OK.", {"total_throat": total_throat}, {"min_total": throat_total_tdc}, "", [])

def r_rect_offset_length(e):
    """Rectangular mitered offsets: 12"-48" length limits (same as transitions)"""
    kind = fitting_kind(e)
    # Only apply to mitered offsets, not ogee
    if kind not in ("Offset",):
        return
    # Skip if ogee (radius offset)
    name = ((e.Name or "") + " " + (part_type_str(e) or "")).lower()
    if "ogee" in name or "radius" in name:
        return
    if not is_rect(e):
        return
    
    L = body_length_in(e)
    min_l = SHOP_STANDARDS["min_offset_length"]
    max_l = SHOP_STANDARDS["max_offset_length"]
    tol = SHOP_STANDARDS["tolerance"]
    
    if L is None:
        return
    
    if L < min_l - tol:
        add_item(e, "R-RECT-OFFSET-MIN", "Rect mitered offset min length",
                 "fail", "Length %.2f in < %.0f in min." % (L, min_l),
                 {"length": L}, {"min": min_l}, "Increase offset length.", [])
    elif L > max_l + tol:
        add_item(e, "R-RECT-OFFSET-MAX", "Rect mitered offset max length",
                 "fail", "Length %.2f in > %.0f in max." % (L, max_l),
                 {"length": L}, {"max": max_l}, "Reduce offset length.", [])
    else:
        add_item(e, "R-RECT-OFFSET-LENGTH", "Rect mitered offset length",
                 "pass", "OK.", {"length": L}, {"min": min_l, "max": max_l}, "", [])

def r_tdc_tap_straight(e):
    """TDC taps require 1.5" straight in. No corner taps on TDC."""
    if not is_tap(e):
        return
    
    tt = tap_type(e)
    si, so = get_extensions(e)
    straight_in = si or 0.0
    tol = SHOP_STANDARDS["tolerance"]
    
    # TDC tap straight requirement
    if tt == "tdc":
        req = SHOP_STANDARDS["tap_tdc_straight_in"]
        if straight_in < req - tol:
            add_item(e, "R-TDC-TAP-STRAIGHT", "TDC tap straight in",
                     "fail", "Straight in %.2f in < %.1f in required." % (straight_in, req),
                     {"straight_in": straight_in}, {"required": req}, "Increase straight in to %.1f in." % req, [])
        else:
            add_item(e, "R-TDC-TAP-STRAIGHT", "TDC tap straight in",
                     "pass", "OK.", {"straight_in": straight_in}, {"required": req}, "", [])
    
    # Machine/inside taps = 0" straight OK (just note)
    elif tt in ("machine", "inside"):
        add_item(e, "R-TAP-TYPE", "Tap type detection",
                 "pass", "%s tap (0\" straight OK)." % tt.title(), {}, {}, "", [])
    
    # No corner taps on TDC
    if is_corner_tap(e):
        connA, _ = end_treatment_names(e)
        if connA == "TDC":
            add_item(e, "R-NO-CORNER-TAP-TDC", "No corner taps on TDC",
                     "fail", "Corner tap with TDC not allowed.",
                     {}, {}, "Reposition tap away from corner.", [])

def r_double_wall_round_length(e):
    """Double wall round duct max 144" (weight limit)"""
    if not is_round(e):
        return
    if not is_double_wall(e):
        return
    
    L = centerline_length_in(e)
    max_l = SHOP_STANDARDS["max_double_wall_round_length"]
    tol = SHOP_STANDARDS["tolerance"]
    
    if L is None:
        return
    
    if L > max_l + tol:
        add_item(e, "R-DW-ROUND-LENGTH", "Double wall round max length",
                 "fail", "Length %.2f in > %.0f in (weight limit)." % (L, max_l),
                 {"length": L}, {"max": max_l}, "Split into shorter sections.", [])
    else:
        add_item(e, "R-DW-ROUND-LENGTH", "Double wall round length",
                 "pass", "OK.", {"length": L}, {"max": max_l}, "", [])

def r_damper_blade_sizing(e):
    """Volume damper blade sizing: Single blade up to 36"x12", multi/opposed over 36" or 14" depth"""
    if fitting_kind(e) != "Damper":
        return
    if not is_rect(e):
        return
    
    pairs = get_rect_pairs(e)
    if not pairs:
        return
    
    w, h = pairs[0]  # Width x Height
    max_w = SHOP_STANDARDS["damper_single_blade_max_width"]
    max_d = SHOP_STANDARDS["damper_single_blade_max_depth"]
    multi_w = SHOP_STANDARDS["damper_multi_blade_width"]
    multi_d = SHOP_STANDARDS["damper_multi_blade_depth"]
    tol = SHOP_STANDARDS["tolerance"]
    
    # Determine blade type needed
    needs_multi = (w > multi_w + tol) or (h > multi_d + tol)
    exceeds_single = (w > max_w + tol) or (h > max_d + tol)
    
    if needs_multi:
        # Flag as needing multi/opposed blade (hand-make)
        add_item(e, "R-DAMPER-MULTI-BLADE", "Damper multi/opposed blade required",
                 "fail", "Size %.0fx%.0f requires multi/opposed blade (hand-make)." % (w, h),
                 {"width": w, "height": h}, {"single_max_w": max_w, "single_max_d": max_d},
                 "Verify multi-blade construction or buy out.", [])
    elif exceeds_single:
        add_item(e, "R-DAMPER-BLADE-SIZE", "Damper blade sizing warning",
                 "fail", "Size %.0fx%.0f exceeds single blade (%.0fx%.0f)." % (w, h, max_w, max_d),
                 {"width": w, "height": h}, {"single_max_w": max_w, "single_max_d": max_d},
                 "Review blade configuration.", [])
    else:
        add_item(e, "R-DAMPER-BLADE-SIZE", "Damper blade sizing",
                 "pass", "Single blade OK.", {"width": w, "height": h}, {}, "", [])

def r_short_joint_seam(e):
    """Short joints (<7"): Use lap & spot weld, not Pittsburgh seam."""
    if not is_rect(e):
        return
    if fitting_kind(e) not in ("Straight",):
        return
    
    L = centerline_length_in(e)
    threshold = SHOP_STANDARDS["short_joint_seam_threshold"]
    
    if L is None:
        return
    
    if L < threshold:
        # Advisory: short joint should use lap & spot
        add_item(e, "R-SHORT-JOINT-SEAM", "Short joint seam type",
                 "pass", "Short joint (%.2f in < %.0f in) - use lap & spot weld seam." % (L, threshold),
                 {"length": L}, {"threshold": threshold}, "", [])



def print_debug_manifest(parts):
    out.print_md("### Debug: Selection Manifest")
    out.print_md("*CID Config: {} entries loaded from `cid_config.json`*".format(len(CID_MAP)))
    out.print_md("")
    out.print_md("| ID | CID | Alias | Kind | Family | Pref | Dims | Length | Connectors | Angle | Ext | Throat |")
    out.print_md("|--- |--- |--- |--- |--- |--- |--- |--- |--- |--- |--- |--- |")

    for e in parts:
        eid = e.Id.IntegerValue
        
        # Get identification info
        cid = get_cid(e)
        cid_str = str(cid) if cid else "-"
        alias = get_alias_param(e) or "-"
        kind, is_pref = identify_part(e)
        kind = kind or fitting_kind(e) or "Unknown"
        family_name = get_family_name(e) or "-"
        pref_str = "✅" if is_pref == True else ("❌" if is_pref == False else "?")
        
        dims = ""
        if is_rect(e):
            pairs = get_rect_pairs(e)
            dims = ", ".join(["{0:.1f}x{1:.1f}".format(w,h) for w,h in pairs])
        elif is_round(e):
            d = diameter_in(e) or 0
            dims = "Ø{0:.1f}".format(d)
        elif is_oval(e):
            dims = "Oval" 
        
        l = centerline_length_in(e)
        l_str = "{0:.2f}".format(l) if l else "-"
        
        ca, cb = end_treatment_names(e)
        conn_str = "{0} / {1}".format(ca or "-", cb or "-")
        
        ang = elbow_angle_deg(e)
        ang_str = "{0:.1f}°".format(ang) if ang else "-"
        
        si, so = get_extensions(e)
        ext_str = "{0:.1f}/{1:.1f}".format(si, so)
        
        # Calc Throat for debug
        throat_str = "-"
        if kind == "Elbow" and is_rect(e) and ang:
            inner_rad = inner_radius_in(e) or 0.0
            arc = inner_rad * math.radians(ang)
            tot = si + so + arc
            throat_str = "{0:.2f}".format(tot)

        row = "| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} | {8} | {9} | {10} | {11} |".format(
            eid, cid_str, alias, kind, family_name, pref_str, dims, l_str, conn_str, ang_str, ext_str, throat_str
        )
        out.print_md(row)
    out.print_md("---")

# ==============================================================================
#   MAIN EXECUTION
# ==============================================================================

def analyze(parts):
    for e in parts:
        # --- Rectangular Coil Fit ---
        if RULES_ENABLED.get("rule_coil_panel_width", True):
            r_coil_panel_width(e)
        if RULES_ENABLED.get("rule_rect_elbow_cheek", True):
            r_rect_elbow_cheek_coil(e)
        
        # --- Rectangular Elbows ---
        if RULES_ENABLED.get("rule_rect_elbow_min_radius", True):
            r_rect_elbow_min_radius(e)
        if RULES_ENABLED.get("rule_rect_elbow_throat", True):
            r_rect_elbow_throat(e)
        if RULES_ENABLED.get("rule_square_elbow_throat", True):
            r_square_elbow_throat(e)
        
        # --- Round/Oval Duct ---
        if RULES_ENABLED.get("rule_round_length_range", True):
            r_round_length_range(e)
        if RULES_ENABLED.get("rule_oval_length_max", True):
            r_oval_length_max(e)
        if RULES_ENABLED.get("rule_lined_round_multiple", True):
            r_lined_round_multiple(e)
        if RULES_ENABLED.get("rule_round_elbow_segmentation", True):
            r_round_elbow_segmentation(e)
        if RULES_ENABLED.get("rule_double_wall_round_length", True):
            r_double_wall_round_length(e)
        
        # --- Rectangular Fittings ---
        if RULES_ENABLED.get("rule_rect_transition_length", True):
            r_rect_transition_length(e)
        if RULES_ENABLED.get("rule_rect_offset_length", True):
            r_rect_offset_length(e)
        if RULES_ENABLED.get("rule_rect_tdc_minjoint", True):
            r_rect_tdc_minjoint(e)
        if RULES_ENABLED.get("rule_rect_straight_length", True):
            r_rect_straight_length_limit(e)
        if RULES_ENABLED.get("rule_short_joint_seam", True):
            r_short_joint_seam(e)
        
        # --- Square to Round ---
        if RULES_ENABLED.get("rule_sq2rd_body_limits", True):
            r_sq2rd_body_limits(e)
        
        # --- Ogee/Offset ---
        if RULES_ENABLED.get("rule_ogee_offset_straight", True):
            r_ogee_offset_straight(e)
        
        # --- Taps ---
        if RULES_ENABLED.get("rule_tdc_tap_straight", True):
            r_tdc_tap_straight(e)
        
        # --- Dampers ---
        if RULES_ENABLED.get("rule_damper_buyout", True):
            r_damper_buyout(e)
        if RULES_ENABLED.get("rule_damper_blade_sizing", True):
            r_damper_blade_sizing(e)
        
        # --- Policy Checks ---
        if RULES_ENABLED.get("rule_exposed_note", True):
            r_exposed_note(e)
        if RULES_ENABLED.get("rule_fullauto_sameconn", True):
            r_fullauto_sameconn(e)
        
        # --- Toggleable Policy ---
        r_no_straight_sq2rd_trans(e)
        r_no_ogee_unless_allowed(e)
        r_no_bullhead_tees(e)

class FabPartFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            if elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_FabricationDuctwork):
                return True
        except:
            pass
        return False
    def AllowReference(self, reference, position):
        return False

def main():
    xaml_path = os.path.join(os.path.dirname(__file__), "window.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("Could not find window.xaml", title="Error")
        return

    sel = uidoc.Selection.GetElementIds()
    preselection = [doc.GetElement(eid) for eid in sel]
    preselection = [p for p in preselection if category_of(p) in ("DuctFitting", "DuctCurves", "DuctAccessory", "FabPart")]

    while True:
        win = ShopQAWindow(xaml_path, preselection)
        win.ShowDialog()
        
        action = getattr(win, "_next_action", None)
        
        if not action:
            break
            
        if action == "pick":
            try:
                refs = uidoc.Selection.PickObjects(ObjectType.Element, FabPartFilter(), "Select Fabrication Parts")
                preselection = [doc.GetElement(r.ElementId) for r in refs]
            except Exception:
                pass
            continue
            
        if action == "run":
            if not preselection:
                forms.alert("No elements selected.", title="Error")
                continue
                
            print_debug_manifest(preselection)
            
            out.print_md("## NW Shop QA Compliance")
            out.print_md("**Profile:** {0}".format(RUNTIME["db_profile"]))
            out.print_md("**Exceptions:** Ogee={0}, Sq2RdStr={1}, TransStr={2}, Tees={3}".format(
                RUNTIME["allow_ogee"], RUNTIME["allow_sq2rd_straight"], 
                RUNTIME["allow_trans_straight"], RUNTIME["allow_bullhead_tees"]))
            
            analyze(preselection)
            
            v = RESULTS["run_summary"]["violations"]
            out.print_md("---")
            out.print_md("**Violations Found:** {0}".format(v))
            
            # --- SELECTION & REPORTING LOGIC ---
            failing_ids = []
            passing_items = []
            
            seen_fail = set()
            seen_pass = set()
            
            for it in RESULTS["items"]:
                eid = it["element_id"]
                if it["status"] == "fail":
                    failing_ids.append(DB.ElementId(eid))
                    seen_fail.add(eid)
                else:
                    seen_pass.add(eid)
            
            if failing_ids:
                try:
                    uidoc.Selection.SetElementIds(List[DB.ElementId](failing_ids))
                except Exception:
                    pass
            else:
                # Clear selection when everything passes
                try:
                    uidoc.Selection.SetElementIds(List[DB.ElementId]([]))
                except Exception:
                    pass

            if seen_fail:
                out.print_md("### ❌ FAILED ITEMS")
                for it in RESULTS["items"]:
                     if it["status"] == "fail":
                         out.print_md("- **FAIL** [{0}] {1}: {2}".format(it["element_id"], it["rule_name"], it["details"]))
            
            fully_passing = [eid for eid in seen_pass if eid not in seen_fail]
            
            if fully_passing:
                out.print_md("### ✅ PASSING ITEMS")
                if len(fully_passing) > 10:
                    out.print_md("- {0} items passed all checks.".format(len(fully_passing)))
                    out.print_md("- IDs: " + ", ".join([str(i) for i in fully_passing]))
                else:
                    for eid in fully_passing:
                        out.print_md("- **PASS** [{0}] - All checks passed.".format(eid))

            json_text = json.dumps(RESULTS, indent=2)
            try:
                from System.Windows.Forms import Clipboard
                Clipboard.SetText(json_text)
                out.print_md("*Full results copied to clipboard (JSON).*")
            except:
                pass
            break

if __name__ == "__main__":
    main()