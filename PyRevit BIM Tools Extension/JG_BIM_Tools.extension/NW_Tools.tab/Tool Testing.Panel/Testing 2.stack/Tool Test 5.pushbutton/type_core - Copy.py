# -*- coding: utf-8 -*-
"""
RFA Type Manager - Core Backend (type_core.py)
===============================================
Handles all Revit API interactions for the RFA Type Manager tool:
    - Snapshotting all types + parameter values from an open family doc
    - Applying a TypeSnapshot back to the family (create / update / delete)
    - Adding and removing family parameters (non-shared)
    - Setting / clearing formulas on parameters
    - Detecting CPython for the Excel bridge subprocess

All value storage uses Revit internal units (decimal feet for length).
Unit conversion for display lives in excel_bridge.py.

Requires an open family document (__revit__.ActiveUIDocument.Document).
"""

import os
import sys
import json
import subprocess
import tempfile

import clr
clr.AddReference("RevitAPI")
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    Transaction, StorageType, ElementId,
    BuiltInParameter
)

# =============================================================================
# REVIT VERSION DETECTION  (re-implemented here so type_core is self-contained)
# =============================================================================

def _rev_year():
    """Returns running Revit major version as int (e.g. 2022, 2024)."""
    try:
        return int(__revit__.Application.VersionNumber)
    except:
        try:
            return int(__revit__.Application.VersionName.split()[1])
        except:
            return 2024


# =============================================================================
# STORAGE TYPE HELPERS
# =============================================================================

# Map StorageType enum value -> stable string token used in JSON / Excel meta
_ST_TO_STR = {
    StorageType.String:    "String",
    StorageType.Integer:   "Integer",
    StorageType.Double:    "Double",
    StorageType.ElementId: "ElementId",
    getattr(StorageType,"None"): "None",
}

_STR_TO_ST = {v: k for k, v in _ST_TO_STR.items()}

def _st_str(storage_type):
    return _ST_TO_STR.get(storage_type, "None")

def _st_from_str(s):
    return _STR_TO_ST.get(s or "None", getattr(StorageType,"None"))


# =============================================================================
# PARAMETER VALUE READ / WRITE
# =============================================================================

def _read_param_value(fp):
    """
    Reads the current value of a FamilyParameter from FamilyManager.
    Returns a Python-native value (str, int, float, or None).
    ElementId params are returned as int.
    """
    try:
        st = fp.StorageType
        if st == StorageType.String:
            return fp.AsString() or ""
        elif st == StorageType.Integer:
            return fp.AsInteger()
        elif st == StorageType.Double:
            v = fp.AsDouble()
            return v if v is not None else 0.0
        elif st == StorageType.ElementId:
            eid = fp.AsElementId()
            return eid.IntegerValue if eid else -1
    except:
        pass
    return None


def _write_param_value(fm, fp, value):
    """
    Sets a parameter value via FamilyManager.Set().
    Converts Python-native values to the correct type for the API call.
    Returns True on success, False on failure.
    """
    if value is None:
        return True  # nothing to write
    try:
        st = fp.StorageType
        if st == StorageType.String:
            fm.Set(fp, str(value) if value is not None else "")
        elif st == StorageType.Integer:
            fm.Set(fp, int(value))
        elif st == StorageType.Double:
            fm.Set(fp, float(value))
        elif st == StorageType.ElementId:
            fm.Set(fp, ElementId(int(value)))
        else:
            return False
        return True
    except Exception:
        return False


# =============================================================================
# SNAPSHOT  (Revit -> Python dict)
# =============================================================================

def snapshot_family(doc):
    """
    Reads every family type and all parameter values from an open family document.

    Returns a TypeSnapshot dict:
    {
        "family_name": str,
        "revit_year": int,
        "parameters": [
            {
                "name": str,
                "storage_type": "Double" | "String" | "Integer" | "ElementId",
                "is_shared": bool,
                "is_instance": bool,         # True = instance param
                "has_formula": bool,
                "read_only": bool,           # True if formula-driven or built-in read-only
                "guid": str | None,          # GUID for shared params, else None
            }, ...
        ],
        "types": [
            {
                "name": str,
                "values": { param_name: value, ... }
                           # formula-driven params are omitted from values
            }, ...
        ]
    }

    Formula-driven parameters are included in the "parameters" list (with
    has_formula=True / read_only=True) so the UI can show them as gray/locked,
    but their values are NOT written to the "types" list — they are never
    exported to Excel.
    """
    if doc is None or not doc.IsFamilyDocument:
        raise ValueError("No open family document.")

    fm = doc.FamilyManager
    family_name = ""
    try:
        family_name = doc.Title.replace(".rfa", "")
    except:
        pass

    # --- Build parameter definitions list ---
    param_defs = []
    for fp in fm.Parameters:
        try:
            defn = fp.Definition
            name = defn.Name
            st   = fp.StorageType
            is_shared  = fp.IsShared
            is_inst    = fp.IsInstance
            has_formula = fp.IsDeterminedByFormula
            read_only   = has_formula  # may extend with other read-only checks

            guid    = None
            formula = ""
            if is_shared:
                try:
                    guid = str(fp.GUID)
                except:
                    pass
            # formula_expr will be filled inside the transaction below
            formula = ""

            param_defs.append({
                "name":         name,
                "storage_type": _st_str(st),
                "is_shared":    bool(is_shared),
                "is_instance":  bool(is_inst),
                "has_formula":  bool(has_formula),
                "formula_expr": formula,   # the expression string, e.g. "Width * 2"
                "read_only":    bool(read_only),
                "guid":         guid,
            })
        except:
            continue

    # Sort: writable first (alphabetical), then read-only (alphabetical)
    param_defs.sort(key=lambda p: (1 if p["read_only"] else 0, p["name"].lower()))

    # Build a fast lookup: name -> FamilyParameter object
    fp_by_name = {}
    for fp in fm.Parameters:
        try:
            fp_by_name[fp.Definition.Name] = fp
        except:
            pass

    # Writable param names only
    writable_names = [p["name"] for p in param_defs if not p["read_only"]]

    # --- Read formula expressions from FamilyParameter.Formula property ---
    # fp.Formula works without a transaction — it is a definition-level property
    formula_exprs = {}
    for fp in fm.Parameters:
        try:
            expr = fp.Formula
            formula_exprs[fp.Definition.Name] = expr if expr is not None else ""
        except:
            try:
                formula_exprs[fp.Definition.Name] = ""
            except:
                pass

    # Back-fill formula expressions into param_defs
    for p in param_defs:
        expr = formula_exprs.get(p["name"], "")
        p["formula_expr"] = expr
        if expr and not p.get("has_formula"):
            p["has_formula"] = True
            p["read_only"]   = True

    # Rebuild writable_names after formula backfill (some may now be formula-driven)
    writable_names = [p["name"] for p in param_defs if not p["read_only"]]

    # Build FamilyParameter lookup by name
    fp_by_name = {}
    for fp in fm.Parameters:
        try:
            fp_by_name[fp.Definition.Name] = fp
        except:
            pass

    # --- Read per-type values ---
    # Correct API: ft.AsDouble(fp) / ft.AsString(fp) / ft.AsInteger(fp)
    # where ft is a FamilyType and fp is a FamilyParameter.
    # No transaction required — this is a pure read.
    type_rows = []
    for ft in fm.Types:
        try:
            type_name = ft.Name
        except:
            type_name = "(unnamed)"

        values = {}
        for name in writable_names:
            fp = fp_by_name.get(name)
            if fp is None:
                continue
            try:
                st = fp.StorageType
                if st == StorageType.String:
                    v = ft.AsString(fp)
                    values[name] = v if v is not None else ""
                elif st == StorageType.Integer:
                    values[name] = ft.AsInteger(fp)
                elif st == StorageType.Double:
                    values[name] = ft.AsDouble(fp)
                elif st == StorageType.ElementId:
                    eid = ft.AsElementId(fp)
                    values[name] = eid.IntegerValue if eid else -1
                else:
                    values[name] = None
            except:
                values[name] = None

        type_rows.append({"name": type_name, "values": values})

    # Sort types alphabetically for predictable output
    type_rows.sort(key=lambda t: t["name"].lower())

    return {
        "family_name":  family_name,
        "revit_year":   _rev_year(),
        "parameters":   param_defs,
        "types":        type_rows,
    }


# =============================================================================
# DIAGNOSTIC SNAPSHOT  (use to debug value reading issues)
# =============================================================================

def diagnose_snapshot(doc):
    """
    Returns a diagnostic string describing exactly what the API returns
    for the first type in the family. Call from script.py to debug.
    """
    lines = []
    try:
        fm = doc.FamilyManager
        types = list(fm.Types)
        params = list(fm.Parameters)
        lines.append("Types found: " + str(len(types)))
        lines.append("Params found: " + str(len(params)))

        if not types:
            lines.append("ERROR: No types")
            return "\n".join(lines)

        ft = types[0]
        lines.append("First type: " + str(ft.Name))
        lines.append("doc.IsModifiable before tx: " + str(doc.IsModifiable))

        t = Transaction(doc, "_diag")
        t.Start()
        lines.append("Transaction started. doc.IsModifiable: " + str(doc.IsModifiable))

        try:
            fm.CurrentType = ft
            lines.append("CurrentType set OK")
        except Exception as e:
            lines.append("CurrentType FAILED: " + str(e))
            t.RollBack()
            return "\n".join(lines)

        # Try reading first 3 params
        for fp in list(fm.Parameters)[:5]:
            try:
                name = fp.Definition.Name
                st = fp.StorageType
                st_name = str(st)
                is_formula = fp.IsDeterminedByFormula
                try:
                    if st == StorageType.String:
                        val = fp.AsString()
                    elif st == StorageType.Integer:
                        val = fp.AsInteger()
                    elif st == StorageType.Double:
                        val = fp.AsDouble()
                    else:
                        val = "(ElementId/other)"
                    lines.append("  " + name + " [" + st_name + "] formula=" + str(is_formula) + " -> " + repr(val))
                except Exception as e:
                    lines.append("  " + name + " READ ERROR: " + str(e))
            except Exception as e:
                lines.append("  PARAM ITER ERROR: " + str(e))

        t.Commit()
        lines.append("Transaction committed OK")

    except Exception as e:
        lines.append("OUTER ERROR: " + str(e))
        import traceback
        lines.append(traceback.format_exc())

    return "\n".join(lines)


# =============================================================================
# APPLY SNAPSHOT  (Python dict -> Revit)
# =============================================================================

def apply_snapshot(doc, snapshot, options=None):
    """
    Applies a TypeSnapshot (from Excel import or in-memory edit) to the
    open family document.

    options dict (all default True/False as noted):
        update_existing  (True)  - overwrite values on types already in the family
        create_new       (True)  - create types that appear in snapshot but not family
        delete_missing   (False) - delete family types not present in snapshot

    Returns:
        {
            "added":   [type_name, ...],
            "updated": [type_name, ...],
            "deleted": [type_name, ...],
            "skipped": [(type_name, reason), ...],
            "errors":  [(type_name, param_name, error_str), ...],
        }
    """
    if options is None:
        options = {}
    do_update = options.get("update_existing", True)
    do_create = options.get("create_new",      True)
    do_delete = options.get("delete_missing",  False)

    fm = doc.FamilyManager

    # Current state of the family
    existing_types = {}  # name -> FamilyType object
    for ft in fm.Types:
        try:
            existing_types[ft.Name] = ft
        except:
            pass

    # Writable FamilyParameter objects by name
    fp_by_name = {}
    for fp in fm.Parameters:
        try:
            if not fp.IsDeterminedByFormula:
                fp_by_name[fp.Definition.Name] = fp
        except:
            pass

    incoming_names = set(t["name"] for t in snapshot.get("types", []))

    added   = []
    updated = []
    deleted = []
    skipped = []
    errors  = []

    t = Transaction(doc, "RFA Type Manager - Apply Types")
    t.Start()
    try:
        # --- Delete types not in snapshot ---
        if do_delete:
            to_delete = [n for n in existing_types if n not in incoming_names]
            for name in to_delete:
                try:
                    ft = existing_types[name]
                    fm.CurrentType = ft
                    fm.DeleteCurrentType()
                    deleted.append(name)
                    del existing_types[name]
                except Exception as e:
                    skipped.append((name, "delete failed: " + str(e)))

        # --- Create new types first (need them to exist before setting values) ---
        newly_created = {}  # name -> FamilyType
        if do_create:
            for row in snapshot.get("types", []):
                name = row["name"]
                if name not in existing_types:
                    try:
                        ft = fm.NewType(name)
                        newly_created[name] = ft
                        existing_types[name] = ft
                        added.append(name)
                    except Exception as e:
                        skipped.append((name, "create failed: " + str(e)))

        # --- Set values ---
        for row in snapshot.get("types", []):
            name   = row["name"]
            values = row.get("values", {})

            ft = existing_types.get(name)
            if ft is None:
                continue  # was skipped during create

            is_new = name in newly_created
            if not is_new and not do_update:
                continue

            fm.CurrentType = ft
            for param_name, value in values.items():
                fp = fp_by_name.get(param_name)
                if fp is None:
                    # param doesn't exist in family — skip with warning
                    errors.append((name, param_name, "parameter not found in family"))
                    continue
                ok = _write_param_value(fm, fp, value)
                if not ok:
                    errors.append((name, param_name, "write failed"))

            if not is_new:
                updated.append(name)

        t.Commit()

    except Exception as e:
        t.RollBack()
        raise RuntimeError("Transaction failed: " + str(e))

    return {
        "added":   added,
        "updated": updated,
        "deleted": deleted,
        "skipped": skipped,
        "errors":  errors,
    }


# =============================================================================
# PARAMETER MANAGEMENT
# =============================================================================

def add_family_parameter(doc, name, param_type_token, group_key, is_instance,
                         formula=None, discipline=None):
    """
    Adds a family-owned (non-shared) parameter to the open family document.

    param_type_token: token from _PARAM_TYPE_MAP (e.g. "Length", "Mass", "Boolean")
    group_key: stable token (e.g. "Construction", "Geometry", "IdentityData")
    is_instance: True = Instance, False = Type
    formula: optional formula expression string
    discipline: ignored in 2022/23 (ParameterType covers it); used in 2024+ for
                SpecTypeId lookup context only
    """
    import binder_core as core

    fm  = doc.FamilyManager
    rev = _rev_year()
    grp = core.resolve_group_for_json(group_key, rev)

    t = Transaction(doc, "Add Family Parameter: " + name)
    t.Start()
    try:
        if rev >= 2024:
            spec = _spec_type_id_from_token(param_type_token)
            opts = DB.FamilyParameterCreationOptions(name, spec)
            opts.IsInstance  = bool(is_instance)
            opts.GroupTypeId = grp
            fp = fm.AddParameter(opts)
        else:
            pt = _param_type_from_token_2022(param_type_token)
            fp = fm.AddParameter(name, grp, pt, bool(is_instance))

        if formula:
            try:
                fm.SetFormula(fp, formula)
            except:
                pass

        t.Commit()
        return fp
    except Exception as e:
        t.RollBack()
        raise RuntimeError("Add parameter failed: " + str(e))


def remove_parameter(doc, param_name):
    """
    Removes a parameter from the open family document.
    Raises ValueError if the parameter doesn't exist or can't be removed.
    """
    fm = doc.FamilyManager
    fp = None
    for p in fm.Parameters:
        try:
            if p.Definition.Name == param_name:
                fp = p
                break
        except:
            pass

    if fp is None:
        raise ValueError("Parameter not found: " + param_name)

    t = Transaction(doc, "Remove Parameter: " + param_name)
    t.Start()
    try:
        fm.RemoveParameter(fp)
        t.Commit()
    except Exception as e:
        t.RollBack()
        raise RuntimeError("Remove failed: " + str(e))


def set_formula(doc, param_name, formula_str):
    """
    Sets or clears the formula on a named parameter.
    Pass formula_str="" to clear an existing formula.
    """
    fm = doc.FamilyManager
    fp = None
    for p in fm.Parameters:
        try:
            if p.Definition.Name == param_name:
                fp = p
                break
        except:
            pass

    if fp is None:
        raise ValueError("Parameter not found: " + param_name)

    t = Transaction(doc, "Set Formula: " + param_name)
    t.Start()
    try:
        fm.SetFormula(fp, formula_str or "")
        t.Commit()
    except Exception as e:
        t.RollBack()
        raise RuntimeError("Set formula failed: " + str(e))


# =============================================================================
# PARAM TYPE TOKEN RESOLUTION  (2024+ SpecTypeId vs 2022/23 ParameterType)
# =============================================================================

# Common tokens used in the Add Parameter dialog, mapped to both API paths.
# Extend this table as needed.
# Discipline -> list of (display_label, token) matching Revit's parameter dialog
PARAM_DISCIPLINES = [
    "Common",
    "Structural",
    "HVAC",
    "Electrical",
    "Piping",
    "Energy",
]

# Data types per discipline (label shown in UI, token used internally)
# Token maps to SpecTypeId attr (2024+) / ParameterType name (2022/23)
PARAM_DATA_TYPES = {
    "Common": [
        ("Length",           "Length",           "Length"),
        ("Area",             "Area",             "Area"),
        ("Volume",           "Volume",           "Volume"),
        ("Angle",            "Angle",            "Angle"),
        ("Slope",            "Slope",            "Slope"),
        ("Number",           "Number",           "Number"),
        ("Integer",          "Int64",            "Integer"),
        ("Yes/No",           "Boolean",          "YesNo"),
        ("Text",             "String",           "Text"),
        ("URL",              "URL",              "URL"),
        ("Material",         "Material",         "Material"),
        ("Image",            "Image",            "Image"),
        ("Family Type",      "FamilyType",       "FamilyType"),
        ("Currency",         "Currency",         "Currency"),
        ("Mass",             "Mass",             "Mass"),
        ("Time",             "Time",             "TimeInterval"),
        ("Speed",            "Speed",            "Speed"),
        ("Rotation",         "Angle",            "Angle"),
    ],
    "Structural": [
        ("Length",           "Length",           "Length"),
        ("Area",             "Area",             "Area"),
        ("Volume",           "Volume",           "Volume"),
        ("Angle",            "Angle",            "Angle"),
        ("Number",           "Number",           "Number"),
        ("Force",            "Force",            "Force"),
        ("Linear Force",     "LinearForce",      "LinearForce"),
        ("Area Force",       "AreaForce",        "AreaForce"),
        ("Moment",           "Moment",           "Moment"),
        ("Stress",           "Stress",           "Stress"),
        ("Mass",             "Mass",             "Mass"),
        ("Mass per Unit Length", "MassPerUnitLength", "MassPerUnitLength"),
        ("Weight",           "Weight",           "Weight"),
        ("Weight per Unit Length", "LinearForce","LinearForce"),
        ("Unit Weight",      "SpecificWeight",   "SpecificWeight"),
        ("Bar Diameter",     "ReinforcementBarDiameter", "RebarBarDiameter"),
        ("Rebar Area",       "ReinforcementArea","RebarArea"),
        ("Rebar Volume",     "ReinforcementVolume","RebarVolume"),
        ("Rebar Cover",      "ReinforcementCover","RebarCover"),
        ("Rebar Spacing",    "ReinforcementSpacing","RebarSpacing"),
        ("Period",           "Period",           "StructuralPeriod"),
        ("Energy",           "Energy",           "Energy"),
    ],
    "HVAC": [
        ("Air Flow",         "AirFlow",          "HVACAirflow"),
        ("Cross Section",    "CrossSection",     "HVACCrossSection"),
        ("Duct Insulation Thickness", "DuctInsulationThickness", "HVACDuctInsulationThickness"),
        ("Duct Lining Thickness", "DuctLiningThickness", "HVACDuctLiningThickness"),
        ("Duct Size",        "DuctSize",         "HVACDuctSize"),
        ("Factor",           "Factor",           "HVACFactor"),
        ("Flow",             "Flow",             "HVACFlow"),
        ("Heat Gain",        "HeatGain",         "HVACHeatGain"),
        ("Power",            "HVACPower",        "HVACPower"),
        ("Pressure",         "Pressure",         "HVACPressure"),
        ("Roughness",        "Roughness",        "HVACRoughness"),
        ("Temperature",      "HVACTemperature",  "HVACTemperature"),
        ("Thermal Resistance","ThermalResistance","HVACThermalResistance"),
        ("Thermal Mass",     "ThermalMass",      "HVACThermalMass"),
        ("Velocity",         "Velocity",         "HVACVelocity"),
        ("Viscosity",        "Viscosity",        "HVACViscosity"),
        ("Number",           "Number",           "Number"),
        ("Length",           "Length",           "Length"),
        ("Area",             "Area",             "Area"),
        ("Angle",            "Angle",            "Angle"),
    ],
    "Electrical": [
        ("Apparent Power",   "ElectricalApparentPower",    "ElectricalApparentPower"),
        ("Color Temperature","ColorTemperature",           "ElectricalColorTemperature"),
        ("Current",          "ElectricalCurrent",          "ElectricalCurrent"),
        ("Efficacy",         "ElectricalEfficacy",         "ElectricalEfficacy"),
        ("Frequency",        "ElectricalFrequency",        "ElectricalFrequency"),
        ("Illuminance",      "Illuminance",                "ElectricalIlluminance"),
        ("Luminance",        "Luminance",                  "ElectricalLuminance"),
        ("Luminous Flux",    "LuminousFlux",               "ElectricalLuminousFlux"),
        ("Luminous Intensity","LuminousIntensity",         "ElectricalLuminousIntensity"),
        ("Power",            "ElectricalPower",            "ElectricalPower"),
        ("Power Density",    "ElectricalPowerDensity",     "ElectricalPowerDensity"),
        ("Voltage",          "ElectricalVoltage",          "ElectricalVoltage"),
        ("Wattage",          "ElectricalPower",            "ElectricalPower"),
        ("Number",           "Number",                     "Number"),
        ("Length",           "Length",                     "Length"),
    ],
    "Piping": [
        ("Flow",             "PipingFlow",                 "PipingFlow"),
        ("Friction",         "PipingFriction",             "PipingFriction"),
        ("Pipe Size",        "PipeSize",                   "PipingSize"),
        ("Pressure",         "PipingPressure",             "PipingPressure"),
        ("Temperature",      "PipingTemperature",          "PipingTemperature"),
        ("Velocity",         "PipingVelocity",             "PipingVelocity"),
        ("Viscosity",        "PipingViscosity",            "PipingViscosity"),
        ("Volume",           "Volume",                     "Volume"),
        ("Density",          "PipingDensity",              "PipingDensity"),
        ("Number",           "Number",                     "Number"),
        ("Length",           "Length",                     "Length"),
        ("Area",             "Area",                       "Area"),
    ],
    "Energy": [
        ("Area",             "Area",                       "Area"),
        ("Cost Rate",        "CostRate",                   "CostRate"),
        ("Electrical Power", "ElectricalPower",            "ElectricalPower"),
        ("Heat Transfer",    "CoefficientOfHeatTransfer",  "HVACCoefficientOfHeatTransfer"),
        ("Length",           "Length",                     "Length"),
        ("Number",           "Number",                     "Number"),
        ("Temperature",      "HVACTemperature",            "HVACTemperature"),
        ("Thermal Resistance","ThermalResistance",         "HVACThermalResistance"),
    ],
}

def get_data_types_for_discipline(discipline):
    """Returns list of (display_label, spec_attr, param_type_attr) for a discipline."""
    return PARAM_DATA_TYPES.get(discipline, PARAM_DATA_TYPES["Common"])

# Flat token map kept for backward compat with add_family_parameter
_PARAM_TYPE_MAP = {
    # token                  SpecTypeId attr (2024+)            ParameterType (2022/23)
    # --- Geometry ---
    "Length":                ("Length",                          "Length"),
    "Angle":                 ("Angle",                           "Angle"),
    "Area":                  ("Area",                           "Area"),
    "Volume":                ("Volume",                          "Volume"),
    "Slope":                 ("Slope",                           "Slope"),
    # --- General numeric ---
    "Number":                ("Number",                          "Number"),
    "Integer":               ("Int64",                           "Integer"),
    "YesNo":                 ("Boolean",                         "YesNo"),
    "Currency":              ("Currency",                        "Currency"),
    # --- Text / reference ---
    "Text":                  ("String",                          "Text"),
    "URL":                   ("URL",                             "URL"),
    "Material":              ("Material",                        "Material"),
    "Image":                 ("Image",                           "Image"),
    "FamilyType":            ("FamilyType",                      "FamilyType"),
    # --- Mass / Force / Structural ---
    "Mass":                  ("Mass",                            "Mass"),
    "MassPerUnitLength":     ("MassPerUnitLength",               "MassPerUnitLength"),
    "MassDensity":           ("MassDensity",                     "MassDensity"),
    "Force":                 ("Force",                           "Force"),
    "LinearForce":           ("LinearForce",                     "LinearForce"),
    "AreaForce":             ("AreaForce",                       "AreaForce"),
    "Moment":                ("Moment",                          "Moment"),
    "Stress":                ("Stress",                          "Stress"),
    "UnitWeight":            ("SpecificWeight",                  "SpecificWeight"),
    "Weight":                ("Weight",                          "Weight"),
    "WeightPerUnitLength":   ("LinearForce",                     "LinearForce"),
    # --- Speed / Flow ---
    "Speed":                 ("Speed",                           "Speed"),
    "Flow":                  ("Flow",                            "HVACFlow"),
    # --- HVAC / Mechanical ---
    "Pressure":              ("Pressure",                        "HVACPressure"),
    "Temperature":           ("HVACTemperature",                 "HVACTemperature"),
    "DuctSize":              ("DuctSize",                        "HVACDuctSize"),
    "CrossSection":          ("CrossSection",                    "HVACCrossSection"),
    "HeatGain":              ("HeatGain",                        "HVACHeatGain"),
    "Roughness":             ("Roughness",                       "HVACRoughness"),
    "Factor":                ("Factor",                          "HVACFactor"),
    "AirFlow":               ("AirFlow",                         "HVACAirflow"),
    "CoefficientOfHeatTransfer": ("CoefficientOfHeatTransfer",   "HVACCoefficientOfHeatTransfer"),
    "ThermalMass":           ("ThermalMass",                     "HVACThermalMass"),
    "ThermalResistance":     ("ThermalResistance",               "HVACThermalResistance"),
    # --- Electrical ---
    "Power":                 ("ElectricalPower",                 "ElectricalPower"),
    "Current":               ("ElectricalCurrent",               "ElectricalCurrent"),
    "Voltage":               ("ElectricalVoltage",               "ElectricalVoltage"),
    "ApparentPower":         ("ElectricalApparentPower",         "ElectricalApparentPower"),
    "PowerDensity":          ("ElectricalPowerDensity",          "ElectricalPowerDensity"),
    "Efficacy":              ("ElectricalEfficacy",              "ElectricalEfficacy"),
    "Luminance":             ("Luminance",                       "ElectricalLuminance"),
    "Illuminance":           ("Illuminance",                     "ElectricalIlluminance"),
    "LuminousFlux":          ("LuminousFlux",                    "ElectricalLuminousFlux"),
    "LuminousIntensity":     ("LuminousIntensity",               "ElectricalLuminousIntensity"),
    "ColorTemperature":      ("ColorTemperature",                "ElectricalColorTemperature"),
    # --- Piping ---
    "PipingFlow":            ("PipingFlow",                      "PipingFlow"),
    "PipingPressure":        ("PipingPressure",                  "PipingPressure"),
    "PipingTemperature":     ("PipingTemperature",               "PipingTemperature"),
    "PipingVelocity":        ("PipingVelocity",                  "PipingVelocity"),
    "PipingViscosity":       ("PipingViscosity",                 "PipingViscosity"),
    "PipingDensity":         ("PipingDensity",                   "PipingDensity"),
    # --- Misc ---
    "LoadClassification":    ("LoadClassification",              "LoadClassification"),
    "BarDiameter":           ("ReinforcementBarDiameter",        "RebarBarDiameter"),
    "ReinforcementArea":     ("ReinforcementArea",               "RebarArea"),
    "ReinforcementVolume":   ("ReinforcementVolume",             "RebarVolume"),
    "ReinforcementCover":    ("ReinforcementCover",              "RebarCover"),
    "ReinforcementSpacing":  ("ReinforcementSpacing",            "RebarSpacing"),
    "RotationalFrequency":   ("RotationalFrequency",             "StructuralFrequency"),
    "Period":                ("Period",                          "StructuralPeriod"),
    "Energy":                ("Energy",                          "Energy"),
}

def _spec_type_id_from_token(token):
    """Revit 2024+: resolve a param type token to a SpecTypeId."""
    entry = _PARAM_TYPE_MAP.get(token, ("Number", "Number"))
    attr  = entry[0]
    try:
        return getattr(DB.SpecTypeId, attr)
    except:
        return DB.SpecTypeId.Number

def _param_type_from_token_2022(token):
    """Revit 2022/23: resolve a param type token to a ParameterType enum value."""
    entry = _PARAM_TYPE_MAP.get(token, ("Number", "Number"))
    attr  = entry[1]
    try:
        return getattr(DB.ParameterType, attr)
    except:
        return DB.ParameterType.Number

def list_param_type_tokens():
    """Returns the sorted list of param type token strings for the UI dropdown."""
    return sorted(_PARAM_TYPE_MAP.keys())


# =============================================================================
# EXCEL BRIDGE  (pure IronPython — no subprocess needed)
# =============================================================================

def export_family_to_xlsx(doc, xlsx_path, unit_mode="inches", selected_params=None):
    """
    Convenience wrapper: snapshots the family and writes directly to xlsx.
    Imports excel_bridge inline to keep type_core free of that dependency
    until it is actually needed.
    """
    import excel_bridge as eb
    snap = snapshot_family(doc)
    eb.export_to_xlsx(snap, xlsx_path, unit_mode=unit_mode, selected_params=selected_params)
    return snap

def import_xlsx_to_snapshot(xlsx_path):
    """
    Convenience wrapper: reads an xlsx and returns a TypeSnapshot dict
    ready to pass to apply_snapshot().
    """
    import excel_bridge as eb
    return eb.import_from_xlsx(xlsx_path)

# =============================================================================
# TEMP FILE HELPERS
# =============================================================================

def make_temp_json():
    """Returns a path to a new temp .json file in the system temp dir."""
    fd, path = tempfile.mkstemp(suffix=".json", prefix="nw_rfa_")
    os.close(fd)
    return path

def write_temp_json(data, path=None):
    """Serializes data to a temp JSON file. Returns the file path."""
    if path is None:
        path = make_temp_json()
    with open(path, "w") as f:
        json.dump(data, f, indent=2)
    return path

def read_temp_json(path):
    """Reads and returns a JSON file as a Python object."""
    with open(path, "r") as f:
        return json.load(f)

def cleanup_temp(path):
    """Silently deletes a temp file."""
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except:
        pass