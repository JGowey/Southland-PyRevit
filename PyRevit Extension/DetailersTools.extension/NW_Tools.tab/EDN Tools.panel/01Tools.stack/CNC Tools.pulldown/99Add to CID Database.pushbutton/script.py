# -*- coding: utf-8 -*-
"""
Add FabricationParts to CID Database
Select parts to add their CID mappings to the shared config file.
"""

from pyrevit import revit, DB, forms, UI
import json
import os

doc = revit.doc
uidoc = revit.uidoc

# =============================================================================
#   CONFIG FILE LOCATION
# =============================================================================

def find_extension_root():
    """Find the DetailersTools.extension folder"""
    current = os.path.dirname(__file__)
    for _ in range(10):
        if os.path.basename(current) == "DetailersTools.extension":
            return current
        parent = os.path.dirname(current)
        if parent == current:
            break
        current = parent
    return os.path.dirname(__file__)

EXTENSION_ROOT = find_extension_root()
CONFIG_DIR = os.path.join(EXTENSION_ROOT, "config")
CID_CONFIG_FILE = os.path.join(CONFIG_DIR, "cid_config.json")

def ensure_config_dir():
    if not os.path.exists(CONFIG_DIR):
        try:
            os.makedirs(CONFIG_DIR)
        except:
            pass

def load_existing_config():
    if os.path.exists(CID_CONFIG_FILE):
        try:
            with open(CID_CONFIG_FILE, 'r') as f:
                return json.load(f)
        except:
            pass
    return {"cid_map": {}, "notes": "CID mappings for NW Shop QA"}

def save_config(config):
    try:
        ensure_config_dir()
        with open(CID_CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2, sort_keys=True)
        return True
    except:
        return False

# =============================================================================
#   PART TYPE DETECTION
# =============================================================================

def determine_part_type(family_name):
    """Determine standardized part type from family name"""
    fl = family_name.lower()
    if "straight" in fl:
        return "Straight"
    elif "elbow" in fl and "square" in fl:
        return "SquareElbow"
    elif "elbow" in fl and "mitered" in fl:
        return "MiteredElbow"
    elif "elbow" in fl:
        return "Elbow"
    elif "transition" in fl:
        return "Transition"
    elif "offset" in fl and "mitered" in fl:
        return "Offset"
    elif "offset" in fl:
        return "Ogee"
    elif "tap" in fl or "boot" in fl:
        return "Tap"
    elif "tee" in fl or "wye" in fl or "branch" in fl:
        return "Tee"
    elif "cap" in fl:
        return "Cap"
    elif "damper" in fl:
        return "Damper"
    elif "square to round" in fl or "sq to rnd" in fl:
        return "SquareToRound"
    else:
        return family_name

def is_preferred_part(alias, service_name, family_name=""):
    """Determine if part is preferred based on alias, service, and family name"""
    family_lower = family_name.lower() if family_name else ""
    
    # Alias 10: Square Elbow with TV = preferred, without TV = non-preferred
    if alias == "10":
        return "tv" in family_lower or "turning vane" in family_lower
    
    # Alias 11, 23 are always non-preferred
    if alias in ("11", "23"):
        return False
    
    # No alias - check service name for alternate/barrel palettes
    if not alias:
        sn_lower = service_name.lower() if service_name else ""
        if "alternate" in sn_lower or "barrel" in sn_lower:
            return False
    
    return True

# =============================================================================
#   MAIN
# =============================================================================

# Get current selection
selection = list(revit.get_selection())
fab_parts = [e for e in selection if isinstance(e, DB.FabricationPart)]

# If no FabricationParts selected, prompt to pick
if not fab_parts:
    try:
        picked = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element,
            "Select FabricationParts to add to CID database (ESC to cancel)"
        )
        if picked:
            fab_parts = [doc.GetElement(ref.ElementId) for ref in picked]
            fab_parts = [e for e in fab_parts if isinstance(e, DB.FabricationPart)]
    except:
        # User cancelled
        pass

if not fab_parts:
    forms.alert("No FabricationParts selected.", exitscript=True)

# Process parts
config = load_existing_config()
cid_map = config.get("cid_map", {})

new_count = 0
updated_count = 0
skipped_count = 0
new_parts = []

for fp in fab_parts:
    try:
        cid = fp.ItemCustomId
        if not cid or cid <= 0:
            continue
        
        cid_str = str(cid)
        
        # Get data
        alias_p = fp.LookupParameter("Alias")
        alias = (alias_p.AsString() or alias_p.AsValueString() or "") if alias_p else ""
        
        fam_p = fp.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_PARAM)
        family = (fam_p.AsString() or fam_p.AsValueString() or "Unknown") if fam_p else "Unknown"
        
        service_name = ""
        try:
            service_name = fp.ServiceName or ""
        except:
            pass
        
        part_type = determine_part_type(family)
        is_pref = is_preferred_part(alias, service_name, family)
        
        entry = {
            "part_type": part_type,
            "is_preferred": is_pref,
            "description": family,
            "alias": alias,
            "service_name": service_name
        }
        
        if cid_str in cid_map:
            existing = cid_map[cid_str]
            if existing.get("part_type") == part_type and existing.get("is_preferred") == is_pref:
                skipped_count += 1
            else:
                cid_map[cid_str] = entry
                updated_count += 1
        else:
            cid_map[cid_str] = entry
            new_count += 1
            pref_icon = "+" if is_pref else "-"
            new_parts.append("  CID {}: {} [{}] {}".format(cid, family, alias or "-", pref_icon))
            
    except Exception as ex:
        pass

# Save
config["cid_map"] = cid_map
saved = save_config(config)

# Build summary message
total = len(cid_map)
msg_lines = [
    "Processed {} FabricationParts".format(len(fab_parts)),
    "",
    "New: {}".format(new_count),
    "Updated: {}".format(updated_count),
    "Unchanged: {}".format(skipped_count),
    "",
    "Total CIDs in database: {}".format(total),
]

if new_parts and len(new_parts) <= 15:
    msg_lines.append("")
    msg_lines.append("Added:")
    msg_lines.extend(new_parts)
elif new_parts:
    msg_lines.append("")
    msg_lines.append("Added {} new entries".format(len(new_parts)))

if saved:
    msg_lines.append("")
    msg_lines.append("Saved to config file")
else:
    msg_lines.append("")
    msg_lines.append("ERROR: Failed to save!")

forms.alert("\n".join(msg_lines), title="CID Database Updated")
