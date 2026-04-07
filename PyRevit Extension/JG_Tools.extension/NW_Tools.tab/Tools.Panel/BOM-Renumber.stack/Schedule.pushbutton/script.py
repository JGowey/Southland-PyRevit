# -*- coding: utf-8 -*-
"""
Southland NW VDC - BOM Parameter Update Tool
=============================================
Version: 1.0 (BOM Update Only)
Author: Southland NW VDC Team
Date: 2026-02-17

OVERVIEW:
    Updates BOM parameters for MEP and structural elements WITHOUT changing
    item numbers. Identical pipeline to the full renumber script — selection,
    classification, and parameter extraction are unchanged. Renumbering
    (Steps 5 and 6) has been removed.

WHAT THIS SCRIPT UPDATES:
    ✓ CP_BOM_Description
    ✓ CP_Size
    ✓ CP_Length (straights and fishmouths only)
    ✓ CP_Material & Coating (fabrication)
    ✓ CP_Fabrication_Connector_Name_1/2 (fabrication)

WHAT THIS SCRIPT DOES NOT TOUCH:
    ✗ CP_Item Number
    ✗ Item Number
    ✗ ItemNumber

SUPPORTED ELEMENTS:
    • MEP Fabrication Parts (Pipework, Ductwork, Hangers)
    • Native Revit MEP (Pipes, Ducts, Flex, Fittings)
    • Family Instances (Structural Framing, Conduit Fittings, Accessories)

WORKFLOW:
    1. User selects elements
    2. Elements classified by type (Fabrication / Native MEP / RFA)
    3. BOM parameters populated
    4. Changes committed to model
"""

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    Transaction, FabricationPart, FamilyInstance, StorageType,
    BuiltInCategory, BuiltInParameter, ElementId,
    CategoryType, FabricationConfiguration
)
from Autodesk.Revit.UI.Selection import ObjectType

doc = revit.doc
uidoc = revit.uidoc


# =============================================================================
# CONFIGURATION
# =============================================================================

# BOM Parameters
DESC_TARGETS = ("CP_BOM_Description",)
MAT_TARGETS = ("CP_Material & Coating", "CP_Meterial & Coating")
SIZE_TARGETS = ("CP_Size",)
LENGTH_TARGETS = ("CP_Length",)
CONN1_PARAM = "CP_Fabrication_Connector_Name_1"
CONN2_PARAM = "CP_Fabrication_Connector_Name_2"

# Detection Parameters
DIAM_PARAM = "Main Primary Diameter"
FAMILY_PARAM = "Family"
NESTED_FLAG_PARAM = "CP_Nested Family"

# Filtering
SKIP_GTP_FAMILY_NAME = "GTP"
EXCLUDE_FAMILY_SUBSTR = "gtp"
ALLOWED_FAB_CATEGORIES = (
    "Fabrication Pipework",
    "Fabrication Ductwork",
    "Fabrication Hangers",
)

# Native MEP Category IDs
NATIVE_MEP_CATEGORIES = {
    -2008044,  # Pipes
    -2008049,  # Pipe Fittings
    -2008000,  # Ducts
    -2008010,  # Duct Fittings
    -2008050,  # Flex Pipes
    -2008020,  # Flex Ducts
}

# Fabrication Constants
FAB_PIPEWORK_CATEGORY = "MEP Fabrication Pipework"
PIPE_LENGTH_THRESHOLD_FT = 21.0
PIPE_PREFIX_SHORT = "21FT-"
PIPE_PREFIX_LONG = "40FT-"
FISHMOUTH_CID = 2875
WELD_EXCLUDES = ("weld neck", "weld-neck", "weldneck", "wn flange", "wnflange", "w/n flange")

# Grouping Tolerance
LENGTH_TOLERANCE_FT = 0.005208333  # 1/16"

# Global Cache
_connector_name_cache = {}


# =============================================================================
# PHASE CLASSIFICATION
# =============================================================================

class Phase:
    """Fabrication part phase enumeration."""
    STRAIGHT = 1
    FISHMOUTH = 2
    FITTING = 3
    TAP = 4
    WELD = 5
    FLEX = 6


def classify_fab_phase(fp):
    """
    Determines fabrication part phase using priority order.
    
    Priority (first match wins):
    1. Flex (by name/category)
    2. Weld (by description, excluding weld neck)
    3. Fishmouth (by CustomItemId)
    4. Tap (by IsATap)
    5. Straight (by IsAStraight)
    6. Fitting (default)
    
    Args:
        fp: FabricationPart element
        
    Returns:
        Phase enum value (1-6)
    """
    # Check flex
    try:
        text = ((getattr(fp, "ItemName", "") or "") + " " + (fp.ProductLongDescription or "")).lower()
        cat = (fp.Category.Name or "").lower()
        if "flex" in text or "flex" in cat:
            return Phase.FLEX
    except:
        pass
    
    # Check weld (exclude weld neck flanges)
    try:
        desc = " ".join([
            str(getattr(fp, "ItemName", "") or ""),
            str(fp.ProductLongDescription or ""),
            str(get_param_str(fp, FAMILY_PARAM))
        ]).lower()
        
        is_weld_neck = any(exclude in desc for exclude in WELD_EXCLUDES)
        if not is_weld_neck and "weld" in desc:
            return Phase.WELD
    except:
        pass
    
    # Check fishmouth (must be before tap - fishmouths may also be taps)
    try:
        if int(getattr(fp, "ItemCustomId", 0)) == FISHMOUTH_CID:
            return Phase.FISHMOUTH
    except:
        pass
    
    # Check tap
    try:
        val = fp.IsATap
        if callable(val) and val() or not callable(val) and bool(val):
            return Phase.TAP
    except:
        pass
    
    # Check straight
    try:
        val = fp.IsAStraight
        if callable(val) and val() or not callable(val) and bool(val):
            return Phase.STRAIGHT
    except:
        pass
    
    # Default: fitting
    return Phase.FITTING


# =============================================================================
# PARAMETER UTILITIES
# =============================================================================

def get_param(el, name):
    """Gets parameter by name, returns None if not found."""
    try:
        return el.LookupParameter(name)
    except:
        return None


def get_first_param(el, names):
    """Gets first available parameter from list."""
    for name in names:
        p = get_param(el, name)
        if p:
            return p
    return None


def get_param_str(el, name):
    """Gets parameter value as string."""
    p = get_param(el, name)
    if not p:
        return ""
    try:
        return (p.AsString() or "").strip()
    except:
        return ""


def get_param_double(el, name, default=0.0):
    """Gets parameter value as double."""
    p = get_param(el, name)
    if not p:
        return default
    try:
        return p.AsDouble()
    except:
        return default


def set_param_value(p, value):
    """Sets parameter value with automatic type handling."""
    if not p or p.IsReadOnly:
        return False
    
    try:
        st = p.StorageType
        
        if st == StorageType.String:
            p.Set("" if value is None else str(value))
            return True
        
        if st == StorageType.Integer:
            if value is None or str(value).strip() == "":
                p.Set(0)
            else:
                p.Set(int(float(value)))
            return True
        
        if st == StorageType.Double:
            if value is None or str(value).strip() == "":
                return False
            p.Set(float(value))
            return True
    except:
        return False
    
    return False


def set_bom_param(el, candidates, value):
    """Sets BOM parameter using first available candidate."""
    p = get_first_param(el, candidates)
    return set_param_value(p, value)


def is_param_truthy(p):
    """Checks if parameter has truthy value."""
    if not p:
        return False
    try:
        if p.StorageType == StorageType.Integer:
            return p.AsInteger() == 1
        val = (p.AsValueString() or "").strip().lower()
        return val in ("yes", "true", "1")
    except:
        return False


def normalize_length(length):
    """Normalizes length to tolerance (1/16")."""
    try:
        return round(float(length) / LENGTH_TOLERANCE_FT) * LENGTH_TOLERANCE_FT
    except:
        return None


# =============================================================================
# ELEMENT CLASSIFICATION
# =============================================================================

def is_valid_model_element(el):
    """Checks if element is a valid model element."""
    if not el:
        return False
    try:
        cat = el.Category
        if not cat or cat.CategoryType != CategoryType.Model:
            return False
        if el.ViewSpecific:
            return False
        return True
    except:
        return False


def is_excluded_family(el):
    """Checks if family should be excluded (contains 'gtp')."""
    try:
        tid = el.GetTypeId()
        te = doc.GetElement(tid)
        if not te:
            return False
        
        # Try FamilyName property
        fam = getattr(te, "FamilyName", "")
        
        # Fallback to parameter
        if not fam:
            p = te.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            fam = p.AsString() if p else ""
        
        return EXCLUDE_FAMILY_SUBSTR in fam.lower()
    except:
        return False


def classify_element(el):
    """
    Classifies element type.
    
    Returns:
        tuple: (type_str, element) where type_str is one of:
               'fabrication', 'native_mep', 'rfa', 'skip'
    """
    # Fabrication parts
    if isinstance(el, FabricationPart):
        try:
            cat_name = el.Category.Name if el.Category else ""
            if any(substr in cat_name for substr in ALLOWED_FAB_CATEGORIES):
                return ('fabrication', el)
        except:
            pass
        return ('skip', el)
    
    # Native MEP
    try:
        if el.Category and el.Category.Id.IntegerValue in NATIVE_MEP_CATEGORIES:
            return ('native_mep', el)
    except:
        pass
    
    # Family instances
    if isinstance(el, FamilyInstance):
        return ('rfa', el)
    
    return ('skip', el)


# =============================================================================
# FABRICATION DATA EXTRACTION
# =============================================================================

def get_fab_length(fp):
    """Gets fabrication part length in feet."""
    try:
        return float(fp.CenterlineLength)
    except:
        return 0.0


def get_fab_material(fp):
    """Gets fabrication part material description."""
    # Try property
    try:
        val = getattr(fp, "ProductMaterialDescription", None)
        if val:
            return str(val).strip()
    except:
        pass
    
    # Try parameter
    try:
        p = fp.LookupParameter("ProductMaterialDescription")
        if p:
            val = (p.AsString() or "").strip()
            if val:
                return val
            return (p.AsValueString() or "").strip()
    except:
        pass
    
    return ""


def get_fab_size(fp):
    """Gets fabrication part size description."""
    try:
        return (fp.ProductSizeDescription or "").strip()
    except:
        return ""


def get_fab_description(fp):
    """Gets fabrication part description with fallbacks."""
    # Priority 1: ProductLongDescription
    desc = fp.ProductLongDescription or ""
    if desc:
        return desc.strip()
    
    # Priority 2: Family and Type parameter (ductwork)
    desc = get_param_str(fp, "Family and Type").strip()
    if desc:
        return desc
    
    # Priority 3: Family parameter
    desc = get_param_str(fp, "Family").strip()
    if desc:
        return desc
    
    # Priority 4: Fabrication Fitting Description (type)
    try:
        te = doc.GetElement(fp.GetTypeId())
        if te:
            desc = get_param_str(te, "Fabrication Fitting Description").strip()
            if desc:
                return desc
    except:
        pass
    
    # Priority 5: Name property
    try:
        desc = fp.Name or ""
        if desc:
            return desc.strip()
    except:
        pass
    
    # Fallback: Category name
    try:
        return fp.Category.Name if fp.Category else "Fabrication Item"
    except:
        return "Fabrication Item"


def is_fab_pipe(fp):
    """Checks if fabrication part is a pipe (not fitting)."""
    # Must be in pipework category
    try:
        if FAB_PIPEWORK_CATEGORY not in (fp.Category.Name or ""):
            return False
    except:
        return False
    
    # Straights in pipework are pipes
    try:
        val = fp.IsAStraight
        if callable(val) and val() or not callable(val) and bool(val):
            return True
    except:
        pass
    
    # Check ProductDataRange as fallback
    for param in ("ProductDataRange", "Product Data Range"):
        val = get_param_str(fp, param).lower()
        if val and (val == "pipe" or val.endswith(".pipe") or val.endswith(":pipe")):
            return True
    
    return False


def add_pipe_prefix(fp, description, length_ft):
    """Adds 21FT- or 40FT- prefix to pipe descriptions."""
    if not is_fab_pipe(fp):
        return description
    
    # Don't double-apply
    if description.startswith(PIPE_PREFIX_SHORT) or description.startswith(PIPE_PREFIX_LONG):
        return description
    
    prefix = PIPE_PREFIX_SHORT if length_ft <= PIPE_LENGTH_THRESHOLD_FT else PIPE_PREFIX_LONG
    return prefix + description


def get_fab_connectors(fp, cfg):
    """Gets connector names sorted by local coordinates."""
    pairs = []
    
    # Get transform
    try:
        transform = fp.GetTransform()
        inv_transform = transform.Inverse if transform else None
    except:
        inv_transform = None
    
    # Get connectors
    try:
        connectors = list(fp.ConnectorManager.Connectors)
    except:
        return []
    
    # Extract names and positions
    for conn in connectors:
        try:
            fci = conn.GetFabricationConnectorInfo()
            if not fci:
                continue
            
            cid = fci.BodyConnectorId
            if not cid:
                continue
            
            # Get name (cached)
            key = cid.IntegerValue if hasattr(cid, 'IntegerValue') else str(cid)
            if key not in _connector_name_cache:
                try:
                    _connector_name_cache[key] = cfg.GetFabricationConnectorName(cid) or ""
                except:
                    _connector_name_cache[key] = ""
            
            name = _connector_name_cache[key]
            if not name:
                continue
            
            # Get local position
            origin = conn.Origin
            if origin and inv_transform:
                local = inv_transform.OfPoint(origin)
                pos = (round(local.X, 6), round(local.Y, 6), round(local.Z, 6))
            elif origin:
                pos = (round(origin.X, 6), round(origin.Y, 6), round(origin.Z, 6))
            else:
                pos = (0.0, 0.0, 0.0)
            
            pairs.append((pos, name))
        except:
            continue
    
    # Sort by position
    pairs.sort(key=lambda x: x[0])
    return [name for _, name in pairs]


# =============================================================================
# NATIVE MEP DATA EXTRACTION
# =============================================================================

def get_native_description(el):
    """
    Gets description for native MEP elements.
    
    Format:
        Fittings: "Family - Type"
        Curves:   "Category - Type"
    """
    try:
        cat_id = el.Category.Id.IntegerValue if el.Category else 0
        cat_name = el.Category.Name if el.Category else "MEP"
    except:
        cat_id = 0
        cat_name = "MEP"
    
    # Get type element
    try:
        type_elem = doc.GetElement(el.GetTypeId())
        if not type_elem:
            return cat_name
    except:
        return cat_name
    
    # Check if fitting
    is_fitting = cat_id in (-2008049, -2008010)
    
    if is_fitting:
        # Fittings: Family - Type
        try:
            fam = type_elem.get_Parameter(BuiltInParameter.ALL_MODEL_FAMILY_NAME).AsString() or ""
        except:
            fam = ""
        
        try:
            typ = type_elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or ""
        except:
            typ = ""
        
        if fam and typ:
            return "{} - {}".format(fam, typ)
        return fam or typ or cat_name
    
    else:
        # Curves: Category - Type
        type_name = ""
        
        # Try ELEM_TYPE_PARAM
        if not type_name:
            try:
                p = type_elem.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
                if p:
                    type_name = p.AsValueString() or p.AsString() or ""
            except:
                pass
        
        # Try ALL_MODEL_TYPE_NAME
        if not type_name:
            try:
                p = type_elem.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if p:
                    type_name = p.AsString() or ""
            except:
                pass
        
        # Try Name property
        if not type_name:
            try:
                type_name = type_elem.Name or ""
            except:
                pass
        
        # Format result
        if type_name and type_name.strip() and type_name.strip() != "N/A":
            return "{} - {}".format(cat_name, type_name.strip())
        return cat_name


def get_native_size(el):
    """Gets size for native MEP elements."""
    try:
        # Try diameter
        for bip in (BuiltInParameter.RBS_PIPE_DIAMETER_PARAM, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM):
            p = el.get_Parameter(bip)
            if p and p.HasValue:
                return p.AsValueString() or ""
        
        # Try height x width
        h = el.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
        w = el.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
        if h and w:
            h_val = h.AsValueString() or ""
            w_val = w.AsValueString() or ""
            if h_val and w_val:
                return "{} x {}".format(h_val, w_val)
    except:
        pass
    
    return ""


def get_native_length(el):
    """Gets length for native MEP curves (not fittings)."""
    try:
        cat_id = el.Category.Id.IntegerValue if el.Category else 0
        
        # Fittings don't get length
        if cat_id in (-2008049, -2008010):
            return None
        
        p = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)
        if p and p.HasValue:
            return p.AsDouble()
    except:
        pass
    
    return None


# =============================================================================
# RFA DATA EXTRACTION
# =============================================================================

def get_rfa_family(fi):
    """Gets family name."""
    try:
        sym = fi.Symbol
        if sym and sym.Family:
            return sym.Family.Name or ""
    except:
        pass
    return ""


def get_rfa_type(fi):
    """Gets type name with multiple fallbacks (Symbol.Name can throw errors)."""
    try:
        sym = fi.Symbol
        if not sym:
            return ""
        
        # Method 1: Symbol.Name (can throw AttributeError)
        try:
            if sym.Name:
                return sym.Name.strip()
        except:
            pass  # Symbol.Name not accessible, try other methods
        
        # Method 2: Built-in SYMBOL_NAME_PARAM
        try:
            p = sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if p:
                val = (p.AsString() or "").strip()
                if val:
                    return val
        except:
            pass
        
        # Method 3: Type element ALL_MODEL_TYPE_NAME
        try:
            te = doc.GetElement(fi.GetTypeId())
            if te:
                p2 = te.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                if p2:
                    val2 = (p2.AsString() or "").strip()
                    if val2:
                        return val2
        except:
            pass
        
        # Method 4: Type element Name property
        try:
            te = doc.GetElement(fi.GetTypeId())
            if te and hasattr(te, "Name") and te.Name:
                return te.Name.strip()
        except:
            pass
    except:
        pass
    
    return ""


def get_rfa_type_id(fi):
    """Gets type ID as integer."""
    try:
        tid = fi.GetTypeId()
        return tid.IntegerValue if tid else -1
    except:
        return -1


def get_rfa_category(fi):
    """Gets category name."""
    try:
        return fi.Category.Name if fi.Category else ""
    except:
        return ""


def is_gtp_family(fi):
    """Checks if GTP generic model (to skip)."""
    try:
        is_generic = fi.Category and fi.Category.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel)
        if is_generic and get_rfa_family(fi).strip().upper() == SKIP_GTP_FAMILY_NAME:
            return True
    except:
        pass
    return False


def is_nested_only(fi):
    """Checks if flagged as nested-only."""
    return is_param_truthy(get_param(fi, NESTED_FLAG_PARAM))


def get_rfa_length(fi, param_name):
    """Gets length parameter normalized to tolerance, with Revit Length fallback."""
    # Try custom parameter first
    p = get_param(fi, param_name)
    if p:
        try:
            if p.StorageType == StorageType.Double:
                val = p.AsDouble()
                if val is not None and val > 0:
                    return normalize_length(val)
        except:
            pass
    
    # Fallback: Try Revit's built-in Length parameter (for structural framing, beams, etc.)
    try:
        length_param = fi.get_Parameter(BuiltInParameter.INSTANCE_LENGTH_PARAM)
        if length_param and length_param.HasValue:
            val = length_param.AsDouble()
            if val is not None and val > 0:
                return normalize_length(val)
    except:
        pass
    
    # Fallback: Try STRUCTURAL_FRAME_CUT_LENGTH
    try:
        cut_length = fi.get_Parameter(BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH)
        if cut_length and cut_length.HasValue:
            val = cut_length.AsDouble()
            if val is not None and val > 0:
                return normalize_length(val)
    except:
        pass
    
    return None


def get_rfa_description(fi):
    """Gets description for RFA with smart CP_Type Name handling."""
    fam = (get_rfa_family(fi) or "").strip()
    typ = (get_rfa_type(fi) or "").strip()
    revit_ft = "{} - {}".format(fam, typ) if fam and typ else (typ or fam or "")
    
    # Priority 1: CP_Type Name (instance)
    val = get_param_str(fi, "CP_Type Name").strip()
    if val:
        # If CP_Type Name is basically just the family name, ignore it and use Revit fallback
        if fam and val.replace(" ", "").lower() == fam.replace(" ", "").lower():
            return revit_ft or val
        return val
    
    # Priority 2: CP_Type Name (type)
    try:
        te = doc.GetElement(fi.GetTypeId())
        if te:
            val = get_param_str(te, "CP_Type Name").strip()
            if val:
                # Same check - if it's just the family name, use Revit fallback
                if fam and val.replace(" ", "").lower() == fam.replace(" ", "").lower():
                    return revit_ft or val
                return val
    except:
        pass
    
    # Priority 3: "Family and Type" parameter
    val = get_param_str(fi, "Family and Type").strip()
    if val:
        return val
    
    # Priority 4: Revit API fallback
    return revit_ft or "Unknown"


def collect_nested(host, out_ids):
    """Recursively collects nested subcomponents."""
    try:
        sub_ids = host.GetSubComponentIds()
    except:
        return
    
    for sid in sub_ids:
        try:
            sub = doc.GetElement(sid)
            if isinstance(sub, FamilyInstance):
                eid = sub.Id.IntegerValue
                if eid not in out_ids:
                    out_ids.add(eid)
                    collect_nested(sub, out_ids)
        except:
            pass


def collect_by_supercomponent(instances, host_ids, out_ids):
    """Collects children by SuperComponent relationship."""
    if not host_ids:
        return
    
    for fi in instances:
        try:
            sc = fi.SuperComponent
            if sc and sc.Id and sc.Id.IntegerValue in host_ids:
                out_ids.add(fi.Id.IntegerValue)
        except:
            pass


def build_rfa_list(instances):
    """Builds RFA list with nested support."""
    rfa_ids = set()
    host_ids = set()
    
    # Collect hosts and nested
    for fi in instances:
        host_ids.add(fi.Id.IntegerValue)
        
        # Collect nested
        nested = set()
        collect_nested(fi, nested)
        rfa_ids.update(nested)
        
        # Skip GTP
        if is_gtp_family(fi):
            continue
        
        # Skip nested-only hosts
        if is_nested_only(fi):
            continue
        
        # Skip if no BOM description parameter to update
        if not get_param(fi, DESC_TARGETS[0]):
            continue
        
        rfa_ids.add(fi.Id.IntegerValue)
    
    # Collect by SuperComponent
    collect_by_supercomponent(instances, host_ids, rfa_ids)
    
    # Build final list
    rfas = []
    for eid in rfa_ids:
        fi = doc.GetElement(ElementId(eid))
        if isinstance(fi, FamilyInstance) and not is_gtp_family(fi) and get_param(fi, DESC_TARGETS[0]):
            rfas.append(fi)
    
    return rfas


# =============================================================================
# MAIN EXECUTION
# =============================================================================

# Step 1: Get Selection
try:
    refs = uidoc.Selection.PickObjects(ObjectType.Element, "Select elements for BOM and numbering")
except:
    script.exit()

picked = [doc.GetElement(r) for r in refs if r]
picked = [el for el in picked if el is not None]

# Step 2: Classify
fabrication = []
native_mep = []
rfas = []

for el in picked:
    if not is_valid_model_element(el):
        continue
    if is_excluded_family(el):
        continue
    
    el_type, el = classify_element(el)
    
    if el_type == 'fabrication':
        fabrication.append(el)
    elif el_type == 'native_mep':
        native_mep.append(el)
    elif el_type == 'rfa':
        rfas.append(el)

# Exit if nothing
if not fabrication and not native_mep and not rfas:
    script.exit()

# Step 3: Execute BOM Update
with Transaction(doc, "Update BOM Parameters") as t:
    t.Start()
    
    cfg = FabricationConfiguration.GetFabricationConfiguration(doc)
    
    # Fabrication
    for fp in fabrication:
        length = get_fab_length(fp)
        desc = get_fab_description(fp)
        desc = add_pipe_prefix(fp, desc, length)
        
        set_bom_param(fp, DESC_TARGETS, desc)
        
        # Connectors
        conns = get_fab_connectors(fp, cfg)
        if len(conns) > 0:
            set_bom_param(fp, (CONN1_PARAM,), conns[0])
        if len(conns) > 1:
            set_bom_param(fp, (CONN2_PARAM,), conns[1])
        
        # Material
        mat = get_fab_material(fp)
        if mat:
            set_bom_param(fp, MAT_TARGETS, mat)
        
        # Size
        size = get_fab_size(fp)
        if size:
            set_bom_param(fp, SIZE_TARGETS, size)
        
        # Length (straights and fishmouths only)
        phase = classify_fab_phase(fp)
        if phase in (Phase.STRAIGHT, Phase.FISHMOUTH):
            set_bom_param(fp, LENGTH_TARGETS, float(length))
    
    # Native MEP
    for el in native_mep:
        desc = get_native_description(el)
        set_bom_param(el, DESC_TARGETS, desc)
        
        size = get_native_size(el)
        if size:
            set_bom_param(el, SIZE_TARGETS, size)
        
        length = get_native_length(el)
        if length:
            set_bom_param(el, LENGTH_TARGETS, float(length))
    
    # RFAs
    for fi in build_rfa_list(rfas):
        desc = get_rfa_description(fi)
        set_bom_param(fi, DESC_TARGETS, desc)
    
    t.Commit()
