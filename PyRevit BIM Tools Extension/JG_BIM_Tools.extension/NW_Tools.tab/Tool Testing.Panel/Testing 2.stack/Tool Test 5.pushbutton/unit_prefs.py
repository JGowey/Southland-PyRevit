# -*- coding: utf-8 -*-
"""
unit_prefs.py - Unit Preference System
=======================================
Single source of truth for all unit display preferences.

Responsibilities:
    - Define all unit categories and their available options
    - Map SpecTypeId strings to categories
    - Load/save user preferences to JSON
    - Provide convert_to_display() and convert_to_internal() via UnitUtils

Prefs are stored per-user in %APPDATA%/ParamBinder/unit_prefs.json
so they persist across sessions and are per-workstation.

Usage:
    import unit_prefs as up

    # Get display value
    display_val = up.to_display(internal_val, spec_type_id)
    display_str = up.format_value(internal_val, spec_type_id)

    # Get internal value from display
    internal_val = up.to_internal(display_val, spec_type_id)

    # Get label for column header
    label = up.get_label(spec_type_id)   # e.g. "lbs", "in", "deg"

    # Open prefs dialog
    up.show_prefs_dialog(parent_form)
"""

import os
import json
import re

# =============================================================================
# UNIT CATEGORIES
# Each category has:
#   display_name:  shown in the prefs dialog
#   options:       [(UnitTypeId_attr, label), ...]  first = default
#   default:       index into options (usually 0)
# =============================================================================

CATEGORIES = {
    "Length": {
        "display": "Length",
        "options": [
            ("Inches",                  "Decimal Inches",       "in"),
            ("Feet",                    "Decimal Feet",         "ft"),
            ("FractionalInches",        "Fractional Inches",    "in (frac)"),
            ("FeetFractionalInches",    "Feet & Frac. Inches",  "ft-in"),
            ("Millimeters",             "Millimeters",          "mm"),
            ("Centimeters",             "Centimeters",          "cm"),
            ("Meters",                  "Meters",               "m"),
        ],
        "default": 0,  # Decimal Inches
    },
    "Mass": {
        "display": "Mass / Weight",
        "options": [
            ("PoundsMass",              "Pounds (lbm)",         "lbs"),
            ("Kilograms",               "Kilograms",            "kg"),
            ("PoundsForce",             "Pounds Force (lbf)",   "lbf"),
        ],
        "default": 0,  # PoundsMass
    },
    "Force": {
        "display": "Force",
        "options": [
            ("PoundsForce",             "Pounds Force",         "lbf"),
            ("Kips",                    "Kips",                 "kip"),
            ("Newtons",                 "Newtons",              "N"),
            ("Kilonewtons",             "Kilonewtons",          "kN"),
        ],
        "default": 0,
    },
    "LinearForce": {
        "display": "Linear Force (per length)",
        "options": [
            ("PoundsForcePerFoot",      "lbf/ft",               "lbf/ft"),
            ("KipsPerFoot",             "kips/ft",              "kip/ft"),
            ("NewtonsPerMeter",         "N/m",                  "N/m"),
            ("KilonewtonsPerMeter",     "kN/m",                 "kN/m"),
        ],
        "default": 0,
    },
    "AreaForce": {
        "display": "Area Force (per area)",
        "options": [
            ("PoundsForcePerSquareFoot","lbf/ft2",              "lbf/ft2"),
            ("PoundsForcePerSquareInch","psi",                  "psi"),
            ("KipsPerSquareFoot",       "ksf",                  "ksf"),
            ("KipsPerSquareInch",       "ksi",                  "ksi"),
            ("NewtonsPerSquareMeter",   "N/m2 (Pa)",            "Pa"),
            ("KilonewtonsPerSquareMeter","kN/m2 (kPa)",         "kPa"),
        ],
        "default": 0,
    },
    "Moment": {
        "display": "Moment / Torque",
        "options": [
            ("PoundForceFeet",          "lbf-ft",               "lbf-ft"),
            ("KipFeet",                 "kip-ft",               "kip-ft"),
            ("NewtonMeters",            "N-m",                  "N-m"),
            ("KilonewtonMeters",        "kN-m",                 "kN-m"),
        ],
        "default": 0,
    },
    "Stress": {
        "display": "Stress / Pressure (Structural)",
        "options": [
            ("PoundsForcePerSquareFoot","lbf/ft2",              "lbf/ft2"),
            ("PoundsForcePerSquareInch","psi",                  "psi"),
            ("KipsPerSquareInch",       "ksi",                  "ksi"),
            ("Pascals",                 "Pa",                   "Pa"),
            ("Kilopascals",             "kPa",                  "kPa"),
            ("Megapascals",             "MPa",                  "MPa"),
        ],
        "default": 1,  # psi
    },
    "Area": {
        "display": "Area",
        "options": [
            ("SquareFeet",              "Square Feet",          "ft2"),
            ("SquareInches",            "Square Inches",        "in2"),
            ("SquareMeters",            "Square Meters",        "m2"),
            ("SquareMillimeters",       "Square Millimeters",   "mm2"),
            ("SquareCentimeters",       "Square Centimeters",   "cm2"),
        ],
        "default": 0,
    },
    "Volume": {
        "display": "Volume",
        "options": [
            ("CubicFeet",               "Cubic Feet",           "ft3"),
            ("CubicInches",             "Cubic Inches",         "in3"),
            ("CubicMeters",             "Cubic Meters",         "m3"),
            ("Gallons",                 "Gallons (US)",         "gal"),
            ("Liters",                  "Liters",               "L"),
        ],
        "default": 0,
    },
    "Angle": {
        "display": "Angle",
        "options": [
            ("Degrees",                 "Degrees",              "deg"),
            ("Radians",                 "Radians",              "rad"),
            ("Gradians",                "Gradians",             "grad"),
        ],
        "default": 0,
    },
    "Slope": {
        "display": "Slope",
        "options": [
            ("RatioTo12",               "Rise:12 (x:12)",       ":12"),
            ("RatioTo1",                "Rise:1  (x:1)",        ":1"),
            ("PercentSlope",            "Percent (%)",          "%"),
            ("DegreesOfArc",            "Degrees",              "deg"),
        ],
        "default": 0,
    },
    "Temperature": {
        "display": "Temperature",
        "options": [
            ("Fahrenheit",              "Fahrenheit",           "F"),
            ("Celsius",                 "Celsius",              "C"),
            ("Kelvin",                  "Kelvin",               "K"),
        ],
        "default": 0,
    },
    "Velocity": {
        "display": "Velocity / Speed",
        "options": [
            ("FeetPerSecond",           "Feet/Second",          "ft/s"),
            ("FeetPerMinute",           "Feet/Minute",          "ft/min"),
            ("MetersPerSecond",         "Meters/Second",        "m/s"),
            ("MilesPerHour",            "Miles/Hour",           "mph"),
        ],
        "default": 0,
    },
    "HvacAirflow": {
        "display": "Airflow (HVAC)",
        "options": [
            ("CubicFeetPerMinute",      "CFM",                  "CFM"),
            ("LitersPerSecond",         "L/s",                  "L/s"),
            ("CubicMetersPerHour",      "m3/h",                 "m3/h"),
        ],
        "default": 0,
    },
    "HvacPressure": {
        "display": "Pressure (HVAC)",
        "options": [
            ("InchesOfWater",           "Inches w.g.",          "in w.g."),
            ("PascalsMm",               "Pascals",              "Pa"),
            ("PoundsForcePerSquareInch","psi",                  "psi"),
        ],
        "default": 0,
    },
    "PipingFlow": {
        "display": "Flow (Piping)",
        "options": [
            ("GallonsPerMinute",        "GPM",                  "GPM"),
            ("LitersPerMinute",         "L/min",                "L/min"),
            ("CubicMetersPerHour",      "m3/h",                 "m3/h"),
        ],
        "default": 0,
    },
    "PipingPressure": {
        "display": "Pressure (Piping)",
        "options": [
            ("PoundsForcePerSquareInch","psi",                  "psi"),
            ("Kilopascals",             "kPa",                  "kPa"),
            ("Bar",                     "bar",                  "bar"),
            ("Pascals",                 "Pa",                   "Pa"),
        ],
        "default": 0,
    },
    "DuctSize": {
        "display": "Duct / Pipe / Conduit Size",
        "options": [
            ("Inches",                  "Inches",               "in"),
            ("Millimeters",             "Millimeters",          "mm"),
        ],
        "default": 0,
    },
    "ElectricalPower": {
        "display": "Electrical Power",
        "options": [
            ("Watts",                   "Watts",                "W"),
            ("Kilowatts",               "Kilowatts",            "kW"),
        ],
        "default": 0,
    },
    "ApparentPower": {
        "display": "Apparent Power",
        "options": [
            ("VoltAmperes",             "VA",                   "VA"),
            ("KilovoltAmperes",         "kVA",                  "kVA"),
        ],
        "default": 0,
    },
    "Current": {
        "display": "Electrical Current",
        "options": [
            ("Amperes",                 "Amperes",              "A"),
        ],
        "default": 0,
    },
    "Voltage": {
        "display": "Voltage",
        "options": [
            ("Volts",                   "Volts",                "V"),
            ("Kilovolts",               "Kilovolts",            "kV"),
        ],
        "default": 0,
    },
    "Illuminance": {
        "display": "Illuminance",
        "options": [
            ("FootCandles",             "Foot-candles",         "fc"),
            ("Lux",                     "Lux",                  "lux"),
        ],
        "default": 0,
    },
    "LuminousFlux": {
        "display": "Luminous Flux",
        "options": [
            ("Lumens",                  "Lumens",               "lm"),
        ],
        "default": 0,
    },
    "ColorTemperature": {
        "display": "Color Temperature",
        "options": [
            ("Kelvin",                  "Kelvin",               "K"),
        ],
        "default": 0,
    },
    "MassDensity": {
        "display": "Mass Density",
        "options": [
            ("PoundsMassPerCubicFoot",  "lbm/ft3",              "lbm/ft3"),
            ("KilogramsPerCubicMeter",  "kg/m3",                "kg/m3"),
        ],
        "default": 0,
    },
    "Currency": {
        "display": "Currency",
        "options": [
            ("Currency",                "Currency ($)",         "$"),
        ],
        "default": 0,
    },
    "Number": {
        "display": "Number (dimensionless)",
        "options": [],   # no conversion — raw value
        "default": 0,
    },
}

# =============================================================================
# SPEC TYPE ID -> CATEGORY MAPPING
# Key: extracted spec key (after stripping "autodesk.spec." and version)
# =============================================================================

SPEC_TO_CATEGORY = {
    # Length
    "aec:length":                       "Length",
    # Mass / Weight
    "aec.structural:mass":              "Mass",
    # Force
    "aec:force":                        "Force",
    "aec:linearForce":                  "LinearForce",
    "aec:areaForce":                    "AreaForce",
    "aec:moment":                       "Moment",
    "aec:stress":                       "Stress",
    # Area / Volume
    "aec:area":                         "Area",
    "aec:volume":                       "Volume",
    # Angle / Slope
    "aec:angle":                        "Angle",
    "aec:slope":                        "Slope",
    # HVAC
    "aec:hvacAirflow":                  "HvacAirflow",
    "aec:hvacPressure":                 "HvacPressure",
    "aec:hvacTemperature":              "Temperature",
    "aec:hvacVelocity":                 "Velocity",
    "aec:hvacDuctSize":                 "DuctSize",
    "aec:crossSection":                 "DuctSize",
    "aec:hvacRoughness":                "Length",
    # Piping
    "aec:pipingFlow":                   "PipingFlow",
    "aec:pipingPressure":               "PipingPressure",
    "aec:pipingTemperature":            "Temperature",
    "aec:pipingVelocity":               "Velocity",
    "aec:pipingSize":                   "DuctSize",
    "aec:pipingRoughness":              "Length",
    # Electrical
    "aec.electrical:conduitSize":       "DuctSize",
    "aec.electrical:wireSize":          "DuctSize",
    "aec.electrical:electricalPower":   "ElectricalPower",
    "aec.electrical:apparentPower":     "ApparentPower",
    "aec.electrical:current":           "Current",
    "aec.electrical:voltage":           "Voltage",
    "aec.electrical:illuminance":       "Illuminance",
    "aec.electrical:luminousFlux":      "LuminousFlux",
    "aec.electrical:colorTemperature":  "ColorTemperature",
    "aec.electrical:luminousIntensity": "Illuminance",
    "aec.electrical:luminance":         "Illuminance",
    "aec.electrical:efficacy":          "Number",
    "aec.electrical:frequency":         "Number",
    "aec.electrical:powerDensity":      "ElectricalPower",
    # Structural
    "aec.structural:massDensity":       "MassDensity",
    "aec.structural:unitWeight":        "AreaForce",
    # Speed / Misc
    "aec:speed":                        "Velocity",
    "aec:massDensity":                  "MassDensity",
    # Currency
    "measurable:currency":              "Currency",
    # Number / dimensionless / YesNo
    "aec:number":                       "Number",
    "measurable:number":                "Number",
    "spec.bool":                        "Number",   # autodesk.spec:spec.bool = YesNo
    "spec:spec.bool":                   "Number",   # fallback
    "aec:heatGain":                     "Number",
    "aec:factor":                       "Number",
    "aec:thermalResistance":            "Number",
    "aec:thermalMass":                  "Number",
}


# Direct ParameterType string -> category (for Revit 2022/23 fallback)
_PARAM_TYPE_TO_CAT = {
    "mass":             "Mass",
    "weight":           "Mass",
    "force":            "Force",
    "linearforce":      "LinearForce",
    "areaforce":        "AreaForce",
    "moment":           "Moment",
    "stress":           "Stress",
    "length":           "Length",
    "area":             "Area",
    "volume":           "Volume",
    "angle":            "Angle",
    "slope":            "Slope",
    "hvacflow":         "HvacAirflow",
    "hvacpressure":     "HvacPressure",
    "hvactemperature":  "Temperature",
    "hvacvelocity":     "Velocity",
    "pipingflow":       "PipingFlow",
    "pipingpressure":   "PipingPressure",
    "pipingtemperature":"Temperature",
    "pipingvelocity":   "Velocity",
    "electricalpower":  "ElectricalPower",
    "electricalcurrent":"Current",
    "electricalvoltage":"Voltage",
    "currency":         "Currency",
    "number":           "Number",
    "integer":          "Number",
    "yesno":            "Number",
    "text":             "Number",
}

def _extract_key(spec_type_id_str):
    """Strip autodesk.spec prefix variants and version suffix."""
    s = spec_type_id_str or ""
    for prefix in ("autodesk.spec.", "autodesk.spec:", "autodesk."):
        if s.startswith(prefix):
            s = s[len(prefix):]
            break
    s = re.sub(r"-\d+\.\d+\.\d+$", "", s)
    return s


def get_category(spec_type_id_str):
    """Returns the category name for a spec_type_id string, or None."""
    if not spec_type_id_str:
        return None
    key = _extract_key(spec_type_id_str)
    # Direct ForgeTypeId lookup
    cat = SPEC_TO_CATEGORY.get(key)
    if cat:
        return cat
    # Partial ForgeTypeId match
    for map_key, map_cat in SPEC_TO_CATEGORY.items():
        if key and (key in map_key or map_key in key):
            return map_cat
    # ParameterType string fallback (Revit 2022/23)
    key_lower = key.lower().replace(" ", "").replace("_", "")
    for pt_key, pt_cat in _PARAM_TYPE_TO_CAT.items():
        if pt_key in key_lower or key_lower in pt_key:
            return pt_cat
    return None


# =============================================================================
# USER PREFERENCES
# =============================================================================

_PREFS_FILE = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "ParamBinder", "unit_prefs.json"
)

# Runtime cache: {category: option_index}
_prefs = {}

def _default_prefs():
    return {cat: CATEGORIES[cat]["default"] for cat in CATEGORIES}

def load_prefs():
    """Loads prefs from JSON, filling defaults for missing keys."""
    global _prefs
    _prefs = _default_prefs()
    try:
        if os.path.exists(_PREFS_FILE):
            with open(_PREFS_FILE, "r") as f:
                saved = json.load(f)
            for cat, idx in saved.items():
                if cat in _prefs and isinstance(idx, int):
                    opts = CATEGORIES[cat]["options"]
                    if 0 <= idx < len(opts):
                        # Validate the attr actually exists in UnitTypeId
                        attr = opts[idx][0]
                        uid = _get_unit_id(attr)
                        if uid is not None:
                            _prefs[cat] = idx
                        # else: attr missing in this Revit version, keep default
    except:
        pass

def save_prefs():
    """Saves current prefs to JSON."""
    try:
        folder = os.path.dirname(_PREFS_FILE)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with open(_PREFS_FILE, "w") as f:
            json.dump(_prefs, f, indent=2)
    except:
        pass

def get_pref_index(category):
    """Returns current option index for a category."""
    if not _prefs:
        load_prefs()
    return _prefs.get(category, CATEGORIES.get(category, {}).get("default", 0))

def set_pref_index(category, idx):
    """Sets option index for a category (call save_prefs() to persist)."""
    if not _prefs:
        load_prefs()
    _prefs[category] = idx

def get_current_option(category):
    """Returns (unit_attr, display_name, label) for current pref."""
    opts = CATEGORIES.get(category, {}).get("options", [])
    if not opts:
        return (None, "", "")
    idx = get_pref_index(category)
    idx = max(0, min(idx, len(opts) - 1))
    return opts[idx]


# =============================================================================
# CONVERSION  (via UnitUtils)
# =============================================================================

def _get_unit_id(unit_attr):
    """Returns the UnitTypeId object for a given attr name, or None."""
    if not unit_attr:
        return None
    try:
        from Autodesk.Revit.DB import UnitTypeId
        uid = getattr(UnitTypeId, unit_attr, None)
        return uid  # returns None if attr doesn't exist
    except:
        return None


def to_display(internal_val, spec_type_id_str):
    """
    Converts a Revit internal value to display units based on current prefs.
    Returns float. Returns internal_val unchanged if no conversion available.
    """
    category = get_category(spec_type_id_str)
    if not category or category == "Number":
        return internal_val

    unit_attr, _, _ = get_current_option(category)
    uid = _get_unit_id(unit_attr)
    if uid is None:
        return internal_val

    try:
        from Autodesk.Revit.DB import UnitUtils
        result = UnitUtils.ConvertFromInternalUnits(float(internal_val), uid)
        # Sanity check: result shouldn't be wildly different from input
        # (catches cases where wrong unit type is applied to wrong spec)
        return result
    except:
        return internal_val  # fallback: show raw internal value


def to_internal(display_val, spec_type_id_str):
    """
    Converts a display value back to Revit internal units based on current prefs.
    Returns float.
    """
    category = get_category(spec_type_id_str)
    if not category or category == "Number":
        return display_val

    unit_attr, _, _ = get_current_option(category)
    uid = _get_unit_id(unit_attr)
    if uid is None:
        return display_val

    try:
        from Autodesk.Revit.DB import UnitUtils
        return UnitUtils.ConvertToInternalUnits(float(display_val), uid)
    except:
        return display_val


def get_label(spec_type_id_str):
    """
    Returns the short unit label for a spec_type_id string.
    e.g. "lbs", "in", "deg", "GPM"
    Returns "" for dimensionless / unknown.
    """
    category = get_category(spec_type_id_str)
    if not category or category == "Number":
        return ""
    _, _, label = get_current_option(category)
    return label


def format_value(internal_val, spec_type_id_str, decimals=4):
    """Converts and formats an internal value as a string for display."""
    try:
        display = to_display(internal_val, spec_type_id_str)
        return "{:.{}f}".format(float(display), decimals)
    except:
        try:
            return "{:.{}f}".format(float(internal_val), decimals)
        except:
            return ""


def is_length(spec_type_id_str):
    """Returns True if this spec type is a length (special fractional handling)."""
    return get_category(spec_type_id_str) == "Length"


# Load prefs on import
load_prefs()