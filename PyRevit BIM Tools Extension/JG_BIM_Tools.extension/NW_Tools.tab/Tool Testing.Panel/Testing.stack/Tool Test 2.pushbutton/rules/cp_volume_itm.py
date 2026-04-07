# -*- coding: utf-8 -*-
# RULE must be first; ASCII-only so the loader can detect it.

RULE = {
    "name": "CP_Volume - ITM internal volume",
    "description": "Computes internal volume for fabrication pipe, caps, fishmouths, and couplers and writes ft^3 to CP_Volume.",
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
        "name": "CP_Volume"
    },
    "combine": "last_wins"
}

# No imports at top level; everything imported inside helpers.

_GPF_STD = {
    0.125: 0.002, 0.25: 0.005, 0.375: 0.011, 0.5: 0.018, 0.75: 0.041, 1.0: 0.065,
    1.25: 0.100, 1.5: 0.140, 2.0: 0.225, 3.0: 0.500, 4.0: 0.900, 6.0: 2.000,
    8.0: 3.500, 10.0: 5.500, 12.0: 8.000, 14.0: 10.500, 16.0: 13.500,
    18.0: 17.000, 20.0: 21.000, 24.0: 30.000, 30.0: 46.000, 36.0: 66.000,
    42.0: 90.000, 48.0: 117.000,
}

_GPF_SCH10 = {
    0.125: 0.003, 0.25: 0.006, 0.375: 0.012, 0.5: 0.020, 0.75: 0.045, 1.0: 0.070,
    1.25: 0.110, 1.5: 0.150, 2.0: 0.250, 3.0: 0.540, 4.0: 1.000, 6.0: 2.200,
    8.0: 3.900, 10.0: 6.100, 12.0: 8.800, 14.0: 11.600, 16.0: 14.800,
    18.0: 18.600, 20.0: 22.800, 24.0: 32.800, 30.0: 50.000, 36.0: 72.000,
    42.0: 98.000, 48.0: 128.000,
}

_GPF_SCH40 = {
    0.125: 0.002, 0.25: 0.005, 0.375: 0.010, 0.5: 0.017, 0.75: 0.038, 1.0: 0.060,
    1.25: 0.092, 1.5: 0.130, 2.0: 0.210, 3.0: 0.475, 4.0: 0.850, 6.0: 1.900,
    8.0: 3.300, 10.0: 5.200, 12.0: 7.600, 14.0: 10.000, 16.0: 12.900,
    18.0: 16.200, 20.0: 20.000, 24.0: 28.800, 30.0: 44.000, 36.0: 63.000,
    42.0: 86.000, 48.0: 112.000,
}

_GAL_TO_FT3 = 1.0 / 7.48051948


def _get_revit_type():
    try:
        from Autodesk.Revit.DB import FabricationPart
        return FabricationPart
    except Exception:
        return None


def _safe_as_double(param):
    try:
        if param is None or not param.HasValue:
            return None
        return float(param.AsDouble())
    except Exception:
        return None


def _get_cp_volume_existing(element):
    try:
        p = element.LookupParameter("CP_Volume")
    except Exception:
        p = None
    return _safe_as_double(p)


def _parse_inches(value_string):
    import re as _re
    if not value_string:
        return None
    s = value_string.strip()
    m = _re.match(r"^(\d+)\s+(\d+)/(\d+)$", s)
    if m:
        whole = float(m.group(1))
        num = float(m.group(2))
        den = float(m.group(3))
        return whole + num / den
    m = _re.match(r"^(\d+)/(\d+)$", s)
    if m:
        num = float(m.group(1))
        den = float(m.group(2))
        return num / den
    m = _re.match(r"^(\d+(?:\.\d*)?)", s)
    if m:
        return float(m.group(1))
    return None


def _get_itm_cid(itm):
    cid = None
    try:
        item = getattr(itm, "Item", None)
        if item is not None:
            cid = getattr(item, "CID", None)
    except Exception:
        cid = None
    if cid is None:
        try:
            ft = getattr(itm, "FittingType", None)
            if ft is not None:
                cid = getattr(ft, "CID", None)
        except Exception:
            cid = None
    if cid is None:
        try:
            cid = getattr(itm, "ItemCustomId", None)
        except Exception:
            cid = None
    return cid


def _compute_cap_volume_ft3(itm, id_inches):
    import math as _m
    if id_inches is None or id_inches <= 0:
        return None
    try:
        dims = list(itm.GetDimensions())
    except Exception:
        dims = []
    collar = None
    height = None
    for d in dims:
        try:
            name = getattr(d, "Name", None)
        except Exception:
            name = None
        if name == "Collar":
            try:
                collar = float(itm.GetDimensionValue(d))
            except Exception:
                pass
        elif name == "Height":
            try:
                height = float(itm.GetDimensionValue(d))
            except Exception:
                pass
    if collar is None or height is None:
        return None
    id_feet = id_inches / 12.0
    r = id_feet / 2.0
    dome = height - collar
    if dome < 0:
        dome = 0.0
    vol_cylinder = _m.pi * r * r * collar
    vol_dome = (2.0 / 3.0) * _m.pi * (r ** 3)
    return vol_cylinder + vol_dome


def _compute_fishmouth_volume_ft3(itm, length_ft):
    import re as _re
    cid = _get_itm_cid(itm)
    if cid != 2875:
        return None
    if not length_ft or length_ft <= 0:
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
        table = _GPF_SCH10
    elif schedule == "SCH40":
        table = _GPF_SCH40
    else:
        table = _GPF_STD

    keys = list(table.keys())
    nearest_key = min(keys, key=lambda k: abs(k - tap_size))
    gpf = table.get(nearest_key)
    if not gpf:
        return None

    gallons = float(length_ft) * float(gpf)
    return gallons * _GAL_TO_FT3


def _compute_coupler_volume_ft3(itm, id_inches, centerline_length):
    import math as _m
    cid = _get_itm_cid(itm)
    if cid != 2522:
        return None
    if id_inches is None or id_inches <= 0:
        return None

    gap_ft = None
    try:
        cm_prop = itm.GetType().GetProperty("ConnectorManager")
        if cm_prop is not None:
            cm = cm_prop.GetValue(itm, None)
            if cm is not None:
                conns_prop = cm.GetType().GetProperty("Connectors")
                if conns_prop is not None:
                    conns = conns_prop.GetValue(cm, None)
                    if conns is not None:
                        p1 = None
                        p2 = None
                        for c in conns:
                            p = None
                            try:
                                p = c.GetType().GetProperty("Origin").GetValue(c, None)
                            except Exception:
                                p = None
                            if p is None:
                                try:
                                    tf = c.GetType().GetProperty("CoordinateSystem").GetValue(c, None)
                                    if tf is not None:
                                        p = tf.Origin
                                except Exception:
                                    p = None
                            if p is not None:
                                if p1 is None:
                                    p1 = p
                                else:
                                    p2 = p
                                    break
                        if p1 is not None and p2 is not None:
                            dx = p2.X - p1.X
                            dy = p2.Y - p1.Y
                            dz = p2.Z - p1.Z
                            gap_ft = _m.sqrt(dx * dx + dy * dy + dz * dz)
    except Exception:
        gap_ft = None

    if gap_ft is None or gap_ft <= 0:
        gap_ft = centerline_length or 0.0
    if gap_ft <= 0:
        return None

    id_feet = id_inches / 12.0
    r = id_feet / 2.0
    return _m.pi * r * r * gap_ft


def _compute_pipe_volume_ft3(id_inches, length_ft):
    import math as _m
    if id_inches is None or id_inches <= 0 or not length_ft or length_ft <= 0:
        return None
    id_feet = id_inches / 12.0
    r = id_feet / 2.0
    return _m.pi * r * r * length_ft


def predicate(element, context):
    return True


def compute(element, context):
    FabricationPart = _get_revit_type()

    vol_ft3 = None

    try:
        if FabricationPart is not None and isinstance(element, FabricationPart):
            itm = element
            length = getattr(itm, "CenterlineLength", None) or 0.0

            id_inches = None
            try:
                p_id = itm.LookupParameter("Inside Diameter")
            except Exception:
                p_id = None
            if p_id is not None and p_id.HasValue:
                try:
                    id_inches = _parse_inches(p_id.AsValueString())
                except Exception:
                    id_inches = None

            cid = _get_itm_cid(itm)
            desc = getattr(itm, "ProductLongDescription", None) or ""

            if cid == 2060 and "CAP" in desc.upper():
                vol_ft3 = _compute_cap_volume_ft3(itm, id_inches)

            if vol_ft3 is None and cid == 2875:
                vol_ft3 = _compute_fishmouth_volume_ft3(itm, length)

            if vol_ft3 is None and cid == 2522:
                vol_ft3 = _compute_coupler_volume_ft3(itm, id_inches, length)

            if vol_ft3 is None:
                if id_inches is None or id_inches <= 0:
                    for name in ("Inside Diameter 1", "Inside Diameter 2"):
                        try:
                            p_alt = itm.LookupParameter(name)
                        except Exception:
                            p_alt = None
                        v = None
                        if p_alt is not None and p_alt.HasValue:
                            try:
                                v = _parse_inches(p_alt.AsValueString())
                            except Exception:
                                v = None
                        if v is not None and v > 0 and (id_inches is None or v < id_inches):
                            id_inches = v
                vol_ft3 = _compute_pipe_volume_ft3(id_inches, length)
    except Exception:
        vol_ft3 = None

    if vol_ft3 is None or not (vol_ft3 == vol_ft3):
        existing = _get_cp_volume_existing(element)
        if existing is not None:
            return float(existing)
        return 0.0

    return float(vol_ft3)
