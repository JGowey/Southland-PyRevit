# -*- coding: utf-8 -*-
"""
RFA Type Manager - Core Backend (type_core.py)
===============================================
Revit API interactions for the RFA Type Manager tool.

Pipeline contract (all values in Revit INTERNAL units throughout):
    snapshot_family()  -> TypeSnapshot  (internal units, no conversion)
    apply_snapshot()   -> Revit         (expects internal units)

Unit conversion happens ONLY at the display/IO boundary:
    script.py _populate_grid()     reads snapshot, converts to display
    script.py _grid_to_snapshot()  reads display, converts back to internal
    excel_bridge.py export/import  same boundary

Each param def carries:
    "storage_type"  : "Double" | "String" | "Integer" | "ElementId"
    "is_length"     : bool   -- True only for true length params (unit: feet)
    "unit_label"    : str    -- display unit label e.g. "ft", "kg", "deg", ""
    "to_display"    : float  -- internal * to_display = display value
    "from_display"  : float  -- display * from_display = internal value

For length params, to_display=1.0 and the grid/Excel layer applies the
user's inches/feet preference on top. For everything else, to_display
is the fixed physical unit scale.
"""

import os
import json

import clr
clr.AddReference("RevitAPI")
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB import (
    Transaction, TransactionGroup, StorageType, ElementId
)


# =============================================================================
# REVIT VERSION
# =============================================================================

def _rev_year():
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

_ST_TO_STR = {
    StorageType.String:              "String",
    StorageType.Integer:             "Integer",
    StorageType.Double:              "Double",
    StorageType.ElementId:           "ElementId",
    getattr(StorageType, "None"):    "None",
}
_STR_TO_ST = {v: k for k, v in _ST_TO_STR.items()}

def _st_str(st):
    return _ST_TO_STR.get(st, "None")


# =============================================================================
# UNIT RESOLUTION  (delegates to unit_prefs.py via spec_type_id)
# =============================================================================

def _get_spec_type_id(fp):
    """
    Returns the ForgeTypeId TypeId string for a FamilyParameter.
    Used as the key into unit_prefs.SPEC_TO_CATEGORY for conversion.
    Falls back to ParameterType string for older Revit versions.
    """
    try:
        return fp.Definition.GetDataType().TypeId or ""
    except:
        pass
    try:
        return str(fp.Definition.ParameterType)
    except:
        return ""


# =============================================================================
# PARAMETER VALUE WRITE
# =============================================================================

def _write_param_value(fm, fp, value):
    """Sets a parameter value via FamilyManager.Set(). Returns True on success."""
    if value is None:
        return True
    try:
        st = fp.StorageType
        if st == StorageType.String:
            fm.Set(fp, str(value))
        elif st == StorageType.Integer:
            # Handle Yes/No string input
            sv = str(value).strip().lower()
            if sv in ("yes", "true", "1"):
                int_val = 1
            elif sv in ("no", "false", "0"):
                int_val = 0
            else:
                int_val = int(round(float(value)))
            fm.Set(fp, int_val)
        elif st == StorageType.Double:
            fm.Set(fp, float(value))
        elif st == StorageType.ElementId:
            fm.Set(fp, ElementId(int(value)))
        else:
            return False
        return True
    except:
        return False


# =============================================================================
# SNAPSHOT  (Revit -> Python dict, all values in internal units)
# =============================================================================

def snapshot_family(doc):
    """
    Reads all family types and parameter values.
    Returns a TypeSnapshot dict with all Double values in Revit internal units.

    Each param def includes unit conversion info:
        is_length    : bool  -- length params get user inch/feet preference
        unit_label   : str   -- display label e.g. "kg", "deg", "ft"
        to_display   : float -- internal * to_display = display value
        from_display : float -- display * from_display = internal value
    """
    if not doc or not doc.IsFamilyDocument:
        raise ValueError("No open family document.")

    fm = doc.FamilyManager

    try:
        family_name = doc.Title.replace(".rfa", "")
    except:
        family_name = ""

    # ---- Parameter definitions ----
    param_defs = []
    fp_by_name = {}

    for fp in fm.Parameters:
        try:
            name      = fp.Definition.Name
            st        = fp.StorageType
            is_shared = fp.IsShared
            is_inst   = fp.IsInstance
            is_formula = fp.IsDeterminedByFormula

            guid = None
            if is_shared:
                try:
                    guid = str(fp.GUID)
                except:
                    pass

            # Formula expression
            formula_expr = ""
            try:
                expr = fp.Formula
                formula_expr = expr if expr is not None else ""
            except:
                pass

            # A param is read-only if it has a formula OR is formula-driven
            read_only = bool(is_formula) or bool(formula_expr)

            # spec_type_id: ForgeTypeId string used by unit_prefs for conversion
            spec_type_id = _get_spec_type_id(fp) if st == StorageType.Double else ""

            # Detect Yes/No params (Integer storage, Boolean/YesNo param type)
            is_yes_no = False
            if st == StorageType.Integer:
                try:
                    pt = str(fp.Definition.ParameterType)
                    is_yes_no = "yesno" in pt.lower() or "boolean" in pt.lower()
                except:
                    pass
                if not is_yes_no:
                    try:
                        tid = fp.Definition.GetDataType().TypeId or ""
                        is_yes_no = "boolean" in tid.lower() or "yesno" in tid.lower()
                    except:
                        pass

            param_defs.append({
                "name":         name,
                "storage_type": _st_str(st),
                "is_shared":    bool(is_shared),
                "is_instance":  bool(is_inst),
                "has_formula":  read_only,
                "formula_expr": formula_expr,
                "read_only":    read_only,
                "guid":         guid,
                "spec_type_id": spec_type_id,  # -> unit_prefs.to_display/to_internal/get_label
                "is_yes_no":    is_yes_no,      # True = display as Yes/No, write as 0/1
            })
            fp_by_name[name] = fp

        except:
            continue

    # Sort: writable first (alpha), then read-only (alpha)
    param_defs.sort(key=lambda p: (1 if p["read_only"] else 0, p["name"].lower()))

    writable_names = [p["name"] for p in param_defs if not p["read_only"]]

    # ---- Per-type values (no transaction needed for reads) ----
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
                    values[name] = ft.AsDouble(fp)  # internal units (feet for length)
                elif st == StorageType.ElementId:
                    eid = ft.AsElementId(fp)
                    values[name] = eid.IntegerValue if eid else -1
                else:
                    values[name] = None
            except:
                values[name] = None

        type_rows.append({"name": type_name, "values": values})

    type_rows.sort(key=lambda t: t["name"].lower())

    return {
        "family_name": family_name,
        "revit_year":  _rev_year(),
        "parameters":  param_defs,
        "types":       type_rows,
    }


# =============================================================================
# APPLY SNAPSHOT  (Python dict -> Revit, expects internal units)
# =============================================================================

def apply_snapshot(doc, snapshot, options=None):
    """
    Applies a TypeSnapshot to the open family document.
    All values in snapshot["types"][i]["values"] must be in Revit internal units.

    options:
        update_existing  (True)   overwrite existing type values
        create_new       (True)   create types not yet in family
        delete_missing   (False)  delete types not in snapshot
    """
    if options is None:
        options = {}
    do_update = options.get("update_existing", True)
    do_create = options.get("create_new",      True)
    do_delete = options.get("delete_missing",  False)

    fm = doc.FamilyManager

    # Park on a temp type for the entire batch so we are NEVER writing
    # to the currently active type. This is the same pattern the native
    # family editor uses internally — it keeps a scratch type selected
    # while it applies all changes, then restores the original selection.
    active_type_name = None
    active_type_obj  = None
    _global_temp     = None
    try:
        ct = fm.CurrentType
        if ct is not None:
            active_type_name = ct.Name
            active_type_obj  = ct
    except:
        pass

    # Index existing types and writable params
    existing_types = {}
    for ft in fm.Types:
        try:
            existing_types[ft.Name] = ft
        except:
            pass

    fp_by_name = {}
    for fp in fm.Parameters:
        try:
            if not fp.IsDeterminedByFormula:
                fp_by_name[fp.Definition.Name] = fp
        except:
            pass

    incoming_names = set(t["name"] for t in snapshot.get("types", []))

    added, updated, deleted, skipped, errors = [], [], [], [], []

    # Single transaction — simplest and most reliable
    t = Transaction(doc, "RFA Type Manager - Apply")
    try:
        from Autodesk.Revit.DB import (
            IFailuresPreprocessor, FailureProcessingResult, FailureSeverity
        )

        class _SilentPreprocessor(IFailuresPreprocessor):
            def PreprocessFailures(self, fa):
                for msg in list(fa.GetFailureMessages()):
                    try:
                        if msg.GetSeverity() == FailureSeverity.Warning:
                            fa.DeleteWarning(msg)
                    except:
                        pass
                return FailureProcessingResult.Continue

        fho = t.GetFailureHandlingOptions()
        fho.SetFailuresPreprocessor(_SilentPreprocessor())
        fho.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(fho)
    except:
        pass  # older Revit or import failure — proceed without preprocessor
    t.Start()
    try:
        # ----------------------------------------------------------------
        # Global temp pattern:
        # 1. Copy the active type -> __RFA_APPLY_TEMP__
        #    (copy preserves valid param values as base)
        # 2. Switch to temp — now parked there for the ENTIRE loop
        # 3. Write all N types freely (none of them are "current")
        # 4. After loop: switch to a known-good written type (same name
        #    as original active if it exists, else first written type)
        # 5. Delete temp
        # Zero mid-loop switches = zero intermediate constraint checks.
        # ----------------------------------------------------------------
        _global_temp      = None
        _restore_type_name = active_type_name  # prefer to restore to original

        try:
            _global_temp = fm.NewType("__RFA_APPLY_TEMP__")
            fm.CurrentType = _global_temp
        except:
            _global_temp = None  # non-fatal — fall back to per-type switching
        # ---- Delete ----
        if do_delete:
            for name in list(existing_types.keys()):
                if name in incoming_names:
                    continue
                ft = existing_types[name]
                # Can't delete active type — use placeholder
                if name == active_type_name:
                    try:
                        tmp = fm.NewType("__DEL_TEMP__")
                        fm.CurrentType = tmp
                        fm.CurrentType = ft
                        fm.DeleteCurrentType()
                        fm.CurrentType = tmp
                        fm.DeleteCurrentType()
                        active_type_name = None
                        active_type_obj  = None
                    except Exception as ex:
                        skipped.append((name, "delete active type failed: " + str(ex)))
                        continue
                else:
                    try:
                        fm.CurrentType = ft
                        fm.DeleteCurrentType()
                    except Exception as ex:
                        skipped.append((name, "delete failed: " + str(ex)))
                        continue
                deleted.append(name)
                del existing_types[name]

        # ---- Create new types ----
        # Strategy: before creating, switch CurrentType to an existing type
        # that has the most similar name (or just first existing type).
        # This controls what Revit copies as the base for NewType().
        # Then write ALL provided values to the new type.
        newly_created = {}
        if do_create:
            # Pick a stable base type to copy from (first existing type)
            base_ft = None
            if existing_types:
                base_ft = list(existing_types.values())[0]

            for row in snapshot.get("types", []):
                name = row["name"]
                if name not in existing_types:
                    try:
                        # Set CurrentType to base before NewType so Revit
                        # copies from a known good state, not a random type
                        if base_ft is not None:
                            try:
                                fm.CurrentType = base_ft
                            except:
                                pass
                        ft = fm.NewType(name)
                        newly_created[name] = ft
                        existing_types[name] = ft
                        added.append(name)
                    except Exception as ex:
                        skipped.append((name, "create failed: " + str(ex)))

        # ---- Set values ----
        for row in snapshot.get("types", []):
            name   = row["name"]
            values = row.get("values", {})

            ft = existing_types.get(name)
            if ft is None:
                continue

            is_new = name in newly_created
            if not is_new and not do_update:
                continue

            # Switch to this type.
            # If we're parked on _global_temp, ft is never the "current" type
            # so Revit skips per-Set() constraint re-evaluation entirely.
            fm.CurrentType = ft

            # Skip-unchanged optimization:
            # ONLY safe when NOT using the global temp pattern.
            # When parked on _global_temp, ft.AsDouble() mid-transaction
            # can return the temp's inherited values instead of the type's
            # actual values, causing subsequent types to be skipped wrongly.
            if not is_new and values and _global_temp is None:
                current_vals = {}
                for param_name in values:
                    fp_chk = fp_by_name.get(param_name)
                    if fp_chk is None:
                        continue
                    try:
                        st = fp_chk.StorageType
                        if st == StorageType.Double:
                            current_vals[param_name] = ft.AsDouble(fp_chk)
                        elif st == StorageType.String:
                            current_vals[param_name] = ft.AsString(fp_chk) or ""
                        elif st == StorageType.Integer:
                            current_vals[param_name] = ft.AsInteger(fp_chk)
                    except:
                        pass
            else:
                current_vals = {}  # write everything — no skip

            wrote_any = False

            # For new types: write in safe order to avoid constraint violations.
            # 1. Strings and integers first (no unit constraints)
            # 2. Doubles sorted largest-first (set big dimensions before small ones
            #    that may be constrained by them)
            # For existing types: skip-unchanged optimization still applies.
            if is_new:
                def _write_order(item):
                    pname, val = item
                    fp2 = fp_by_name.get(pname)
                    if fp2 is None:
                        return (99, 0)
                    st2 = fp2.StorageType
                    if st2 == StorageType.String:
                        return (0, 0)
                    if st2 == StorageType.Integer:
                        return (1, 0)
                    if st2 == StorageType.Double:
                        try:
                            return (2, -abs(float(val)))  # largest first
                        except:
                            return (2, 0)
                    return (3, 0)
                sorted_values = sorted(values.items(), key=_write_order)
            else:
                sorted_values = list(values.items())

            for param_name, value in sorted_values:
                if value is None:
                    continue
                fp = fp_by_name.get(param_name)
                if fp is None or fp.IsDeterminedByFormula:
                    continue
                # Skip-unchanged for existing types
                if not is_new:
                    cur = current_vals.get(param_name)
                    if cur is not None:
                        try:
                            if fp.StorageType == StorageType.Double:
                                if abs(float(cur) - float(value)) < 1e-9:
                                    continue
                            elif str(cur) == str(value):
                                continue
                        except:
                            pass
                if not _write_param_value(fm, fp, value):
                    errors.append((name, param_name, "write failed"))
                else:
                    wrote_any = True

            if not is_new:
                if wrote_any:
                    updated.append(name)
                # else: no changes needed, silently skip

            # (no per-type cleanup needed — global temp handles active type)

        # ---- Restore and cleanup ----
        if _global_temp is not None:
            try:
                # Find the best type to land on after cleanup:
                # prefer the original active type by name (now freshly written),
                # fall back to first type in the written batch,
                # fall back to any existing type.
                restore_ft = None

                # Try original active type name first
                if _restore_type_name:
                    for ft_r in fm.Types:
                        try:
                            if ft_r.Name == _restore_type_name:
                                restore_ft = ft_r
                                break
                        except:
                            pass

                # Fall back to first type in the snapshot
                if restore_ft is None and snapshot.get("types"):
                    first_name = snapshot["types"][0]["name"]
                    for ft_r in fm.Types:
                        try:
                            if ft_r.Name == first_name:
                                restore_ft = ft_r
                                break
                        except:
                            pass

                # Final fallback: any type that isn't the temp
                if restore_ft is None:
                    temp_name = _global_temp.Name
                    for ft_r in fm.Types:
                        try:
                            if ft_r.Name != temp_name:
                                restore_ft = ft_r
                                break
                        except:
                            pass

                # Switch to restore target, then delete temp
                if restore_ft is not None:
                    fm.CurrentType = restore_ft
                fm.CurrentType = _global_temp
                fm.DeleteCurrentType()
                if restore_ft is not None:
                    fm.CurrentType = restore_ft

            except:
                pass  # cleanup failure is non-fatal

        t.Commit()

    except Exception as e:
        # Best-effort cleanup on failure
        if _global_temp is not None:
            try:
                for ft_r in fm.Types:
                    try:
                        if ft_r.Name != "__RFA_APPLY_TEMP__":
                            fm.CurrentType = ft_r
                            break
                    except:
                        pass
                fm.CurrentType = _global_temp
                fm.DeleteCurrentType()
            except:
                pass
        try:
            t.RollBack()
        except:
            pass
        raise RuntimeError("Apply failed: " + str(e))

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

def add_family_parameter(doc, name, param_type_token, group_key,
                         is_instance, formula=None, discipline=None):
    """Adds a family-owned (non-shared) parameter."""
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
    """Removes a parameter from the family."""
    fm = doc.FamilyManager
    fp = next((p for p in fm.Parameters
               if p.Definition.Name == param_name), None)
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
    """Sets or clears the formula on a named parameter."""
    fm = doc.FamilyManager
    fp = next((p for p in fm.Parameters
               if p.Definition.Name == param_name), None)
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
# PARAM TYPE RESOLUTION  (for Add Parameter dialog)
# =============================================================================

PARAM_DISCIPLINES = ["Common","Structural","HVAC","Electrical","Piping","Energy"]

PARAM_DATA_TYPES = {
    "Common": [
        ("Length","Length","Length"),("Area","Area","Area"),
        ("Volume","Volume","Volume"),("Angle","Angle","Angle"),
        ("Slope","Slope","Slope"),("Number","Number","Number"),
        ("Integer","Int64","Integer"),("Yes/No","Boolean","YesNo"),
        ("Text","String","Text"),("URL","URL","URL"),
        ("Material","Material","Material"),("Image","Image","Image"),
        ("Family Type","FamilyType","FamilyType"),
        ("Currency","Currency","Currency"),("Mass","Mass","Mass"),
        ("Speed","Speed","Speed"),
    ],
    "Structural": [
        ("Length","Length","Length"),("Area","Area","Area"),
        ("Volume","Volume","Volume"),("Angle","Angle","Angle"),
        ("Number","Number","Number"),("Force","Force","Force"),
        ("Linear Force","LinearForce","LinearForce"),
        ("Area Force","AreaForce","AreaForce"),
        ("Moment","Moment","Moment"),("Stress","Stress","Stress"),
        ("Mass","Mass","Mass"),
        ("Mass per Unit Length","MassPerUnitLength","MassPerUnitLength"),
        ("Weight","Weight","Weight"),
        ("Weight per Unit Length","LinearForce","LinearForce"),
        ("Unit Weight","SpecificWeight","SpecificWeight"),
        ("Bar Diameter","ReinforcementBarDiameter","RebarBarDiameter"),
        ("Rebar Area","ReinforcementArea","RebarArea"),
        ("Rebar Volume","ReinforcementVolume","RebarVolume"),
        ("Rebar Cover","ReinforcementCover","RebarCover"),
        ("Rebar Spacing","ReinforcementSpacing","RebarSpacing"),
        ("Period","Period","StructuralPeriod"),
        ("Energy","Energy","Energy"),
    ],
    "HVAC": [
        ("Air Flow","AirFlow","HVACAirflow"),
        ("Duct Size","DuctSize","HVACDuctSize"),
        ("Factor","Factor","HVACFactor"),
        ("Flow","Flow","HVACFlow"),
        ("Pressure","Pressure","HVACPressure"),
        ("Temperature","HVACTemperature","HVACTemperature"),
        ("Thermal Resistance","ThermalResistance","HVACThermalResistance"),
        ("Velocity","Velocity","HVACVelocity"),
        ("Number","Number","Number"),("Length","Length","Length"),
        ("Area","Area","Area"),("Angle","Angle","Angle"),
    ],
    "Electrical": [
        ("Apparent Power","ElectricalApparentPower","ElectricalApparentPower"),
        ("Color Temperature","ColorTemperature","ElectricalColorTemperature"),
        ("Current","ElectricalCurrent","ElectricalCurrent"),
        ("Efficacy","ElectricalEfficacy","ElectricalEfficacy"),
        ("Illuminance","Illuminance","ElectricalIlluminance"),
        ("Luminance","Luminance","ElectricalLuminance"),
        ("Luminous Flux","LuminousFlux","ElectricalLuminousFlux"),
        ("Power","ElectricalPower","ElectricalPower"),
        ("Voltage","ElectricalVoltage","ElectricalVoltage"),
        ("Number","Number","Number"),("Length","Length","Length"),
    ],
    "Piping": [
        ("Flow","PipingFlow","PipingFlow"),
        ("Pressure","PipingPressure","PipingPressure"),
        ("Temperature","PipingTemperature","PipingTemperature"),
        ("Velocity","PipingVelocity","PipingVelocity"),
        ("Number","Number","Number"),("Length","Length","Length"),
        ("Area","Area","Area"),
    ],
    "Energy": [
        ("Area","Area","Area"),
        ("Electrical Power","ElectricalPower","ElectricalPower"),
        ("Heat Transfer","CoefficientOfHeatTransfer","HVACCoefficientOfHeatTransfer"),
        ("Length","Length","Length"),("Number","Number","Number"),
        ("Temperature","HVACTemperature","HVACTemperature"),
        ("Thermal Resistance","ThermalResistance","HVACThermalResistance"),
    ],
}

def get_data_types_for_discipline(discipline):
    return PARAM_DATA_TYPES.get(discipline, PARAM_DATA_TYPES["Common"])

_PARAM_TYPE_MAP = {
    "Length":("Length","Length"),"Angle":("Angle","Angle"),
    "Area":("Area","Area"),"Volume":("Volume","Volume"),
    "Slope":("Slope","Slope"),"Number":("Number","Number"),
    "Integer":("Int64","Integer"),"YesNo":("Boolean","YesNo"),
    "Text":("String","Text"),"URL":("URL","URL"),
    "Material":("Material","Material"),"Image":("Image","Image"),
    "FamilyType":("FamilyType","FamilyType"),"Currency":("Currency","Currency"),
    "Mass":("Mass","Mass"),"MassPerUnitLength":("MassPerUnitLength","MassPerUnitLength"),
    "MassDensity":("MassDensity","MassDensity"),"Force":("Force","Force"),
    "LinearForce":("LinearForce","LinearForce"),"AreaForce":("AreaForce","AreaForce"),
    "Moment":("Moment","Moment"),"Stress":("Stress","Stress"),
    "SpecificWeight":("SpecificWeight","SpecificWeight"),
    "Weight":("Weight","Weight"),"Speed":("Speed","Speed"),
    "Flow":("Flow","HVACFlow"),"Pressure":("Pressure","HVACPressure"),
    "HVACTemperature":("HVACTemperature","HVACTemperature"),
    "DuctSize":("DuctSize","HVACDuctSize"),
    "AirFlow":("AirFlow","HVACAirflow"),
    "ThermalResistance":("ThermalResistance","HVACThermalResistance"),
    "Velocity":("Velocity","HVACVelocity"),
    "ElectricalPower":("ElectricalPower","ElectricalPower"),
    "ElectricalCurrent":("ElectricalCurrent","ElectricalCurrent"),
    "ElectricalVoltage":("ElectricalVoltage","ElectricalVoltage"),
    "ElectricalApparentPower":("ElectricalApparentPower","ElectricalApparentPower"),
    "Illuminance":("Illuminance","ElectricalIlluminance"),
    "Luminance":("Luminance","ElectricalLuminance"),
    "LuminousFlux":("LuminousFlux","ElectricalLuminousFlux"),
    "ColorTemperature":("ColorTemperature","ElectricalColorTemperature"),
    "ElectricalEfficacy":("ElectricalEfficacy","ElectricalEfficacy"),
    "PipingFlow":("PipingFlow","PipingFlow"),
    "PipingPressure":("PipingPressure","PipingPressure"),
    "PipingTemperature":("PipingTemperature","PipingTemperature"),
    "PipingVelocity":("PipingVelocity","PipingVelocity"),
    "ReinforcementBarDiameter":("ReinforcementBarDiameter","RebarBarDiameter"),
    "ReinforcementArea":("ReinforcementArea","RebarArea"),
    "ReinforcementVolume":("ReinforcementVolume","RebarVolume"),
    "ReinforcementCover":("ReinforcementCover","RebarCover"),
    "ReinforcementSpacing":("ReinforcementSpacing","RebarSpacing"),
    "Period":("Period","StructuralPeriod"),
    "Energy":("Energy","Energy"),
    "CoefficientOfHeatTransfer":("CoefficientOfHeatTransfer","HVACCoefficientOfHeatTransfer"),
}

def _spec_type_id_from_token(token):
    entry = _PARAM_TYPE_MAP.get(token, ("Number","Number"))
    try:
        return getattr(DB.SpecTypeId, entry[0])
    except:
        return DB.SpecTypeId.Number

def _param_type_from_token_2022(token):
    entry = _PARAM_TYPE_MAP.get(token, ("Number","Number"))
    try:
        return getattr(DB.ParameterType, entry[1])
    except:
        return DB.ParameterType.Number

def list_param_type_tokens():
    return sorted(_PARAM_TYPE_MAP.keys())


# =============================================================================
# DIAGNOSTICS
# =============================================================================

def diagnose_snapshot(doc):
    """Quick API sanity check. Returns a formatted string."""
    lines = []
    try:
        fm   = doc.FamilyManager
        types  = list(fm.Types)
        params = list(fm.Parameters)
        lines.append("Types: " + str(len(types)))
        lines.append("Params: " + str(len(params)))
        if not types:
            return "\n".join(lines)
        ft = types[0]
        lines.append("Reading type: " + str(ft.Name))
        for fp in params[:8]:
            try:
                name = fp.Definition.Name
                st   = fp.StorageType
                is_len, lbl, to_d, _ = get_param_unit_info(fp)
                formula = ""
                try:
                    formula = fp.Formula or ""
                except:
                    pass
                if st == StorageType.Double:
                    raw = ft.AsDouble(fp)
                    disp = raw * to_d if not is_len else raw * 12.0
                    lines.append("  {} = {} internal  ({:.4f} {}){}".format(
                        name, raw, disp, lbl or "in",
                        "  [=" + formula + "]" if formula else ""))
                elif st == StorageType.String:
                    lines.append("  {} = {}{}".format(
                        name, repr(ft.AsString(fp)),
                        "  [=" + formula + "]" if formula else ""))
                elif st == StorageType.Integer:
                    lines.append("  {} = {}".format(name, ft.AsInteger(fp)))
            except Exception as ex:
                lines.append("  ERR " + fp.Definition.Name + ": " + str(ex))
    except Exception as ex:
        import traceback
        lines.append("ERROR: " + str(ex))
        lines.append(traceback.format_exc())
    return "\n".join(lines)