# -*- coding: utf-8 -*-
# RULE must be first; ASCII-only so the loader can detect it.

RULE = {
    "name": "CP_Weight - ITM and RFA dry weight",
    "description": "Computes dry weight for fabrication parts and RFAs, writing kg to CP_Weight with fallbacks.",
    "priority": 80,
    "filter": {
        "categories": [
            "Air Terminals",
            "Cable Trays",
            "Cable Tray Fittings",
            "Columns",
            "Communication Devices",
            "Conduits",
            "Conduit Fittings",
            "Duct Accessories",
            "Duct Fittings",
            "Ducts",
            "Electrical Equipment",
            "Electrical Fixtures",
            "Fire Alarm Devices",
            "Flex Ducts",
            "Flex Pipes",
            "Food Service Equipment",
            "Generic Models",
            "Lighting Devices",
            "Lighting Fixtures",
            "Mechanical Equipment",
            "Medical Equipment",
            "MEP Fabrication Containment",
            "MEP Fabrication Ductwork",
            "MEP Fabrication Hangers",
            "MEP Fabrication Pipework",
            "Pipe Accessories",
            "Pipe Fittings",
            "Pipes",
            "Plumbing Equipment",
            "Plumbing Fixtures",
            "Security Devices",
            "Specialty Equipment",
            "Sprinklers",
            "Structural Beam Systems",
            "Structural Columns",
            "Structural Connections",
            "Structural Foundations",
            "Structural Framing",
            "Supports",
            "Telecommunication Devices",
            "Vibration Isolators"
        ]
    },
    "target": {
        "name": "CP_Weight"
    },
    "combine": "last_wins"
}

# NOTE:
#  - No imports at top level so fast rule discovery (with __builtins__ = {}) can exec this file.
#  - All imports are inside helper functions and only run at evaluation time.

_WEIGHT_PER_FT_STD = {
    0.125: 0.24, 0.25: 0.42, 0.375: 0.57, 0.5: 0.85, 0.75: 1.13, 1.0: 1.68, 1.25: 2.28, 1.5: 2.72,
    2.0: 3.66, 2.5: 5.10, 3.0: 6.60, 3.5: 8.17, 4.0: 10.79, 5.0: 14.62, 6.0: 18.97, 8.0: 28.55,
    10.0: 40.48, 12.0: 53.52, 14.0: 63.52, 16.0: 74.75, 18.0: 92.00, 20.0: 105.00, 24.0: 140.00,
    30.0: 182.00, 36.0: 216.00, 42.0: 254.00, 48.0: 280.00,
}

_WEIGHT_PER_FT_SCH10 = {
    0.5: 0.65, 0.75: 0.72, 1.0: 1.27, 1.25: 1.63, 1.5: 2.03, 2.0: 2.64, 2.5: 3.33, 3.0: 4.30, 4.0: 5.62,
    5.0: 7.77, 6.0: 9.29, 8.0: 14.97, 10.0: 20.55, 12.0: 26.40, 14.0: 30.63, 16.0: 35.43, 18.0: 40.56,
    20.0: 45.62, 24.0: 56.84, 30.0: 77.45, 36.0: 92.69, 42.0: 109.56, 48.0: 123.93,
}

_WEIGHT_PER_FT_SCH40 = {
    0.125: 0.24, 0.25: 0.42, 0.375: 0.57, 0.5: 0.85, 0.75: 1.13, 1.0: 1.68, 1.25: 2.28, 1.5: 2.72,
    2.0: 3.66, 2.5: 5.80, 3.0: 7.58, 3.5: 9.11, 4.0: 10.79, 5.0: 14.62, 6.0: 18.97, 8.0: 28.55,
    10.0: 40.48, 12.0: 53.52, 14.0: 63.52, 16.0: 74.75, 18.0: 92.00, 20.0: 105.00, 24.0: 140.00,
    30.0: 182.00, 36.0: 216.00, 42.0: 254.00, 48.0: 280.00,
}

_LBS_TO_KG = 0.45359237


def _get_revit_types():
    try:
        from Autodesk.Revit.DB import FabricationPart, FamilyInstance
        return FabricationPart, FamilyInstance
    except Exception:
        return None, None


def _safe_as_double(param):
    try:
        if param is None or not param.HasValue:
            return None
        return float(param.AsDouble())
    except Exception:
        return None


def _get_cp_weight_existing(element):
    try:
        p = element.LookupParameter("CP_Weight")
    except Exception:
        p = None
    return _safe_as_double(p)


def _compute_fishmouth_itm_kg(itm):
    import re as _re

    FabricationPart, _ = _get_revit_types()
    if FabricationPart is None:
        return None

    cid = None
    try:
        item = getattr(itm, "Item", None)
        if item is not None:
            cid = getattr(item, "CID", None)
    except Exception:
        pass
    if cid is None:
        try:
            ft = getattr(itm, "FittingType", None)
            if ft is not None:
                cid = getattr(ft, "CID", None)
        except Exception:
            pass
    if cid is None:
        try:
            cid = getattr(itm, "ItemCustomId", None)
        except Exception:
            pass
    if cid != 2875:
        return None

    length = getattr(itm, "CenterlineLength", None)
    if not length or length <= 0:
        for alt in ("Tap Length", "Cut Length", "Height"):
            try:
                p = itm.LookupParameter(alt)
            except Exception:
                p = None
            v = _safe_as_double(p)
            if v and v > 0:
                length = v
                break
    if not length or length <= 0:
        return None

    entry = getattr(itm, "ProductEntry", None)
    if not entry:
        try:
            pe = itm.LookupParameter("Product Entry")
        except Exception:
            pe = None
        if pe is not None and pe.HasValue:
            try:
                entry = pe.AsString()
            except Exception:
                entry = None
    if not entry:
        return None

    nums = []
    for m in _re.finditer(r"(\d+(?:\.\d+)?|\d+\s*[-/]\s*\d+)", entry):
        v = m.group(1).replace(" ", "")
        if "-" in v or "/" in v:
            try:
                num, den = v.replace("-", "/").split("/")
                nums.append(float(num) / float(den))
            except Exception:
                continue
        else:
            try:
                nums.append(float(v))
            except Exception:
                continue
    if not nums:
        return None
    tap_size = sorted(nums)[0]

    schedule = "STD"
    spec_str = None
    for attr in ("ProductSpecificationDescription", "Specification"):
        try:
            spec_str = getattr(itm, attr, None)
        except Exception:
            spec_str = None
        if spec_str:
            break
    if spec_str:
        s = spec_str.upper()
        if "SCH 10" in s or "SCH10" in s:
            schedule = "SCH10"
        elif "SCH 40" in s or "SCH40" in s:
            schedule = "SCH40"
        elif "STD" in s:
            schedule = "STD"

    if schedule == "SCH10":
        table = _WEIGHT_PER_FT_SCH10
    elif schedule == "SCH40":
        table = _WEIGHT_PER_FT_SCH40
    else:
        table = _WEIGHT_PER_FT_STD

    keys = list(table.keys())
    nearest_key = min(keys, key=lambda k: abs(k - tap_size))
    lbs_per_ft = table.get(nearest_key)
    if not lbs_per_ft:
        return None

    raw_lbs = lbs_per_ft * float(length)
    kg = raw_lbs * _LBS_TO_KG
    return kg


def _compute_itm_weight_kg(itm):
    FabricationPart, _ = _get_revit_types()
    if FabricationPart is None:
        return None

    kg = _compute_fishmouth_itm_kg(itm)
    if kg is not None and kg > 0:
        return kg

    try:
        get_w = getattr(itm, "GetWeight", None)
    except Exception:
        get_w = None
    if callable(get_w):
        try:
            kg_val = float(get_w())
            if kg_val > 0:
                return kg_val
        except Exception:
            pass

    try:
        w_prop = getattr(itm, "Weight", None)
        if isinstance(w_prop, (int, float)) and w_prop > 0:
            return float(w_prop)
    except Exception:
        pass

    try:
        p = itm.LookupParameter("Weight")
    except Exception:
        p = None
    kg_val = _safe_as_double(p)
    if kg_val is not None and kg_val > 0:
        return kg_val

    kg_val = _get_cp_weight_existing(itm)
    if kg_val is not None and kg_val > 0:
        return kg_val

    return None


def _compute_rfa_weight_kg(elem):
    _, FamilyInstance = _get_revit_types()

    kg_val = _get_cp_weight_existing(elem)
    if kg_val is not None and kg_val > 0:
        return kg_val

    try:
        if FamilyInstance is not None and isinstance(elem, FamilyInstance):
            symbol = elem.Symbol
        else:
            symbol = None
    except Exception:
        symbol = None

    if symbol is not None:
        try:
            p = symbol.LookupParameter("CP_Weight")
        except Exception:
            p = None
        kg_val = _safe_as_double(p)
        if kg_val is not None and kg_val > 0:
            return kg_val

    return None


def predicate(element, context):
    return True


def compute(element, context):
    FabricationPart, _ = _get_revit_types()

    kg_val = None
    try:
        if FabricationPart is not None and isinstance(element, FabricationPart):
            kg_val = _compute_itm_weight_kg(element)
    except Exception:
        kg_val = None

    if kg_val is None:
        kg_val = _compute_rfa_weight_kg(element)

    if kg_val is None or not (kg_val == kg_val):
        existing = _get_cp_weight_existing(element)
        if existing is not None:
            return float(existing)
        return 0.0

    return float(kg_val)
