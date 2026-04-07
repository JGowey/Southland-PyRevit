# -*- coding: utf-8 -*-
"""
pyRevit - Batch Writer for:
  - CP_Weight        (internal kg)
  - CP_Volume        (internal ft^3)
  - CP_BOM_Category  (string)

UI:
  1) Checkbox popup: pick 1–3 operations
  2) Scope popup (same UI style): Entire Model / Current View / Current Selection
     Default: Current View

Performance:
  - Category-name based filter (matches your checkbox list screenshot)
  - Chunked Transactions for large models
  - No project-unit detection and no display-unit conversions
"""

from __future__ import division

import re
import math

from pyrevit import forms, script

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementMulticategoryFilter,
    TransactionGroup,
    Transaction,
    ElementId,
)

from System.Collections.Generic import List  # typed .NET list

# Fabrication API class (safe import pattern)
try:
    from Autodesk.Revit.DB import FabricationPart
except:
    FabricationPart = None

logger = script.get_logger()
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# -----------------------------------------------------------------------------
# CATEGORY FILTER: Match screenshot labels by Category.Name
# -----------------------------------------------------------------------------
CATEGORY_NAMES = set([
    "Air Terminals",
    "Cable Tray Fittings",
    "Cable Trays",
    "Columns",
    "Communication Devices",
    "Conduit Fittings",
    "Conduits",
    "Data Devices",
    "Doors",
    "Duct Accessories",
    "Duct Fittings",
    "Ducts",
    "Electrical Equipment",
    "Electrical Fixtures",
    "Fire Alarm Devices",
    "Fire Protection",
    "Flex Ducts",
    "Flex Pipes",
    "Food Service Equipment",
    "Generic Models",
    "Lighting Devices",
    "Lighting Fixtures",
    "Mechanical Control Devices",
    "Mechanical Equipment",
    "Medical Equipment",
    "MEP Fabrication Ductwork",
    "MEP Fabrication Hangers",
    "MEP Fabrication Pipework",
    "Parts",
    "Pipe Accessories",
    "Pipe Fittings",
    "Pipes",
    "Plumbing Equipment",
    "Plumbing Fixtures",
    "Railings",
    "Ramps",
    "Security Devices",
    "Specialty Equipment",
    "Sprinklers",
    "Structural Columns",
    "Structural Connections",
    "Structural Foundations",
    "Structural Framing",
    "Structural Rebar",
    "Structural Stiffeners",
    "Structural Trusses",
    "Telephone Devices",
    "Walls",
    "Windows",
    "Wires",
])

# -----------------------------------------------------------------------------
# Fishmouth lookup tables
# -----------------------------------------------------------------------------
WEIGHT_PER_FOOT_STD = {
    "0.125":0.24,"0.25":0.42,"0.375":0.57,"0.5":0.85,"0.75":1.13,"1":1.68,"1.25":2.28,"1.5":2.72,
    "2":3.66,"2.5":5.1,"3":6.6,"3.5":8.17,"4":10.79,"5":14.62,"6":18.97,"8":28.55,"10":40.48,
    "12":53.52,"14":63.52,"16":74.75,"18":92,"20":105,"24":140,"30":182,"36":216,"42":254,"48":280
}
WEIGHT_PER_FOOT_SCH10 = {
    "0.5":0.65,"0.75":0.72,"1":1.27,"1.25":1.63,"1.5":2.03,"2":2.64,"2.5":3.33,"3":4.3,"4":5.62,
    "5":7.77,"6":9.29,"8":14.97,"10":20.55,"12":26.4,"14":30.63,"16":35.43,"18":40.56,"20":45.62,
    "24":56.84,"30":77.45,"36":92.69,"42":109.56,"48":123.93
}
WEIGHT_PER_FOOT_SCH40 = {
    "0.125":0.24,"0.25":0.42,"0.375":0.57,"0.5":0.85,"0.75":1.13,"1":1.68,"1.25":2.28,"1.5":2.72,
    "2":3.66,"2.5":5.8,"3":7.58,"3.5":9.11,"4":10.79,"5":14.62,"6":18.97,"8":28.55,"10":40.48,
    "12":53.52,"14":63.52,"16":74.75,"18":92,"20":105,"24":140,"30":182,"36":216,"42":254,"48":280
}

GPF_STD = {
    0.125:0.002, 0.25:0.005, 0.375:0.011, 0.5:0.018, 0.75:0.041, 1.0:0.065,
    1.25:0.100, 1.5:0.140, 2.0:0.225, 3.0:0.500, 4.0:0.900, 6.0:2.000,
    8.0:3.500, 10.0:5.500, 12.0:8.000, 14.0:10.500, 16.0:13.500,
    18.0:17.000, 20.0:21.000, 24.0:30.000, 30.0:46.000, 36.0:66.000,
    42.0:90.000, 48.0:117.000
}
GPF_SCH10 = {
    0.125:0.003, 0.25:0.006, 0.375:0.012, 0.5:0.020, 0.75:0.045, 1.0:0.070,
    1.25:0.110, 1.5:0.150, 2.0:0.250, 3.0:0.540, 4.0:1.000, 6.0:2.200,
    8.0:3.900, 10.0:6.100, 12.0:8.800, 14.0:11.600, 16.0:14.800,
    18.0:18.600, 20.0:22.800, 24.0:32.800, 30.0:50.000, 36.0:72.000,
    42.0:98.000, 48.0:128.000
}
GPF_SCH40 = {
    0.125:0.002, 0.25:0.005, 0.375:0.010, 0.5:0.017, 0.75:0.038, 1.0:0.060,
    1.25:0.092, 1.5:0.130, 2.0:0.210, 3.0:0.475, 4.0:0.850, 6.0:1.900,
    8.0:3.300, 10.0:5.200, 12.0:7.600, 14.0:10.000, 16.0:12.900,
    18.0:16.200, 20.0:20.000, 24.0:28.800, 30.0:44.000, 36.0:63.000,
    42.0:86.000, 48.0:112.000
}

KG_PER_LB = 0.45359237
GAL_TO_FT3 = 1.0 / 7.48051948  # gal -> ft^3

# -----------------------------------------------------------------------------
# UI (same style for both popups)
# -----------------------------------------------------------------------------
ops = forms.SelectFromList.show(
    ["Write CP_Weight", "Write CP_Volume", "Write CP_BOM_Category"],
    multiselect=True,
    button_name="Run",
    title="Choose outputs to write",
    width=450,
    height=300
)
if not ops:
    script.exit()

scope = forms.SelectFromList.show(
    ["Current View", "Entire Model", "Current Selection"],
    multiselect=False,
    button_name="Run",
    title="Choose scope",
    width=450,
    height=230,
    default=["Current View"]
)
if not scope:
    script.exit()

DO_WEIGHT = "Write CP_Weight" in ops
DO_VOLUME = "Write CP_Volume" in ops
DO_BOMCAT = "Write CP_BOM_Category" in ops

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def is_fab_part(el):
    return (FabricationPart is not None) and isinstance(el, FabricationPart)

def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None

def set_param_double(el, name, val):
    p = get_param(el, name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(float(val))
        return True
    except:
        return False

def set_param_string(el, name, val):
    p = get_param(el, name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(val if val is not None else "")
        return True
    except:
        return False

def get_type_elem(el):
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue != -1:
            return doc.GetElement(tid)
    except:
        pass
    return None

def get_category_ids_by_name(doc, names_set):
    ids = []
    try:
        cats = doc.Settings.Categories
        for c in cats:
            try:
                if c and c.Name in names_set:
                    ids.append(c.Id)
            except:
                pass
    except:
        pass
    return ids

# -----------------------------------------------------------------------------
# Fabrication helpers (CID, spec, etc.)
# -----------------------------------------------------------------------------
def get_cid_fab(itm):
    cid = None
    try:
        cid = itm.ItemCustomId
    except:
        cid = None
    if cid:
        return cid
    p = get_param(itm, "ItemCustomId")
    if p and p.HasValue:
        try:
            return int(p.AsString())
        except:
            pass
    return None

def get_spec_string_fab(itm):
    for attr in ("ProductSpecificationDescription", "Specification"):
        try:
            s = getattr(itm, attr)
            if s:
                return s
        except:
            pass
    return ""

def schedule_key(spec_str):
    s = (spec_str or "").upper()
    if "SCH 10" in s:
        return "SCH10"
    if "SCH 40" in s:
        return "SCH40"
    if "STD" in s:
        return "STD"
    return "STD"

def closest_key_numeric(keys, target):
    best = None
    bestdiff = None
    for k in keys:
        try:
            diff = abs(float(k) - float(target))
        except:
            continue
        if best is None or diff < bestdiff:
            best = k
            bestdiff = diff
    return best

# -----------------------------------------------------------------------------
# CP_BOM_Category calculator
# -----------------------------------------------------------------------------
def calc_bom_category(el):
    if is_fab_part(el):
        try:
            if el.Category:
                return el.Category.Name
        except:
            pass
        return "Unknown ITM Category"

    p = get_param(el, "CPL_Family Category")
    if p and p.HasValue:
        try:
            v = (p.AsString() or "").strip()
            if v:
                return v
        except:
            pass

    te = get_type_elem(el)
    if te:
        tp = get_param(te, "CPL_Family Category")
        if tp and tp.HasValue:
            try:
                v = (tp.AsString() or "").strip()
                if v:
                    return v
            except:
                pass

    try:
        if el.Category:
            return el.Category.Name
    except:
        pass
    return "Unknown Category"

# -----------------------------------------------------------------------------
# CP_Weight calculator (returns kg, internal)
# -----------------------------------------------------------------------------
def parse_fishmouth_tap_size(entry_str):
    if not entry_str:
        return None
    matches = re.findall(r"(\d+(?:\.\d+)?|\d+\s*[-/]\s*\d+)", entry_str)
    nums = []
    for m in matches:
        v = re.sub(r"\s", "", m)
        if "-" in v:
            v = v.replace("-", "/")
        if "/" in v:
            try:
                parts = v.split("/")
                if len(parts) == 2:
                    nums.append(float(parts[0]) / float(parts[1]))
            except:
                pass
        else:
            try:
                nums.append(float(v))
            except:
                pass
    if not nums:
        return None
    return sorted(nums)[0]

def fishmouth_weight_kg(itm):
    entry = None
    for attr in ("ProductEntry", "ProductSizeDescription"):
        try:
            entry = getattr(itm, attr)
            if entry:
                break
        except:
            pass
    if not entry:
        pe = get_param(itm, "Product Entry")
        if pe and pe.HasValue:
            try:
                entry = pe.AsString()
            except:
                entry = None

    tap = parse_fishmouth_tap_size(entry)
    if tap is None:
        return 0.0

    spec = get_spec_string_fab(itm)
    sk = schedule_key(spec)

    table = WEIGHT_PER_FOOT_STD
    if sk == "SCH10":
        table = WEIGHT_PER_FOOT_SCH10
    elif sk == "SCH40":
        table = WEIGHT_PER_FOOT_SCH40

    tap_key = "{0}".format(tap)
    wpf = table.get(tap_key, None)
    if wpf is None:
        ck = closest_key_numeric(table.keys(), tap)
        if ck is not None:
            wpf = table.get(ck, None)
    if wpf is None:
        return 0.0

    length = 0.0
    try:
        length = float(itm.CenterlineLength or 0.0)
    except:
        length = 0.0
    if length <= 0:
        for alt in ("Tap Length", "Cut Length", "Height"):
            p = get_param(itm, alt)
            if p and p.HasValue:
                try:
                    length = float(p.AsDouble())
                    if length > 0:
                        break
                except:
                    pass
    if length <= 0:
        return 0.0

    raw_lbs = float(wpf) * float(length)
    return raw_lbs * KG_PER_LB

def calc_weight_kg(el):
    if is_fab_part(el):
        cid = get_cid_fab(el)
        if cid == 2875:
            return fishmouth_weight_kg(el)

        try:
            w = getattr(el, "Weight", None)
            if w and w > 0:
                return float(w)
        except:
            pass
        try:
            w = el.Get_Weight()
            if w and w > 0:
                return float(w)
        except:
            pass

        p = get_param(el, "Weight")
        if p and p.HasValue:
            try:
                v = float(p.AsDouble())
                if v > 0:
                    return v
            except:
                pass

        p = get_param(el, "CP_Weight")
        if p and p.HasValue:
            try:
                return float(p.AsDouble())
            except:
                pass

        return 0.0

    p = get_param(el, "CP_Weight")
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            pass

    te = get_type_elem(el)
    if te:
        tp = get_param(te, "CP_Weight")
        if tp and tp.HasValue:
            try:
                return float(tp.AsDouble())
            except:
                pass

    return 0.0

# -----------------------------------------------------------------------------
# CP_Volume calculator (returns ft^3, internal)
# -----------------------------------------------------------------------------
def fishmouth_volume_ft3(itm):
    spec = get_spec_string_fab(itm)
    sk = schedule_key(spec)

    lookup = GPF_STD
    if sk == "SCH10":
        lookup = GPF_SCH10
    elif sk == "SCH40":
        lookup = GPF_SCH40

    entry = None
    p = get_param(itm, "Product Entry")
    if p and p.HasValue:
        try:
            entry = p.AsString()
        except:
            entry = None
    if not entry:
        return 0.0

    m = re.search(r"\d+[xX](\d+\.?\d*)", entry)
    if not m:
        return 0.0
    try:
        tap_size = float(m.group(1))
    except:
        return 0.0

    ck = None
    bestd = None
    for k in lookup.keys():
        d = abs(float(k) - tap_size)
        if ck is None or d < bestd:
            ck = k
            bestd = d
    if ck is None:
        return 0.0

    gpf = float(lookup[ck])
    try:
        length = float(itm.CenterlineLength or 0.0)
    except:
        length = 0.0
    if length <= 0:
        return 0.0

    gallons = length * gpf
    return gallons * GAL_TO_FT3

def get_double_param(el, name):
    p = get_param(el, name)
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            return None
    return None

def coupler_length_ft(itm):
    try:
        cm = getattr(itm, "ConnectorManager", None)
        if cm:
            conns = cm.Connectors
            pts = []
            for c in conns:
                try:
                    pts.append(c.Origin)
                except:
                    try:
                        pts.append(c.CoordinateSystem.Origin)
                    except:
                        pass
            if len(pts) >= 2:
                p1, p2 = pts[0], pts[1]
                dx = p2.X - p1.X
                dy = p2.Y - p1.Y
                dz = p2.Z - p1.Z
                d = math.sqrt(dx*dx + dy*dy + dz*dz)
                if d > 0:
                    return float(d)
    except:
        pass

    try:
        return float(itm.CenterlineLength or 0.0)
    except:
        return 0.0

def calc_volume_ft3(el):
    if is_fab_part(el):
        cid = get_cid_fab(el)
        desc = ""
        try:
            desc = el.ProductLongDescription or ""
        except:
            desc = ""

        id_ft = get_double_param(el, "Inside Diameter")

        # CAP
        if cid == 2060 and ("Cap" in desc):
            try:
                dims = el.GetDimensions()
                collar = None
                height = None
                for dim in dims:
                    if dim.Name == "Collar":
                        collar = float(el.GetDimensionValue(dim))
                    elif dim.Name == "Height":
                        height = float(el.GetDimensionValue(dim))
                if id_ft and id_ft > 0 and collar and height and collar > 0 and height > 0:
                    r = id_ft / 2.0
                    vol_cyl = math.pi * r * r * collar
                    vol_dome = (2.0 / 3.0) * math.pi * (r ** 3)
                    return float(vol_cyl + vol_dome)
            except:
                pass

        # FISHMOUTH
        if cid == 2875:
            v = fishmouth_volume_ft3(el)
            if v > 0:
                return v

        # COUPLER
        if cid == 2522:
            if not id_ft or id_ft <= 0:
                cand = []
                for nm in ("Inside Diameter 1", "Inside Diameter 2"):
                    vv = get_double_param(el, nm)
                    if vv and vv > 0:
                        cand.append(vv)
                if cand:
                    id_ft = min(cand)

            gap = coupler_length_ft(el)
            if id_ft and id_ft > 0 and gap and gap > 0:
                r = id_ft / 2.0
                return float(math.pi * r * r * gap)

        # PIPE default
        try:
            length = float(el.CenterlineLength or 0.0)
        except:
            length = 0.0

        if id_ft and id_ft > 0 and length > 0:
            r = id_ft / 2.0
            return float(math.pi * r * r * length)

        # fallback CP_Volume
        p = get_param(el, "CP_Volume")
        if p and p.HasValue:
            try:
                return float(p.AsDouble())
            except:
                pass
        return 0.0

    # Non-fab fallback only
    p = get_param(el, "CP_Volume")
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            pass
    return 0.0

# -----------------------------------------------------------------------------
# Collect elements
# -----------------------------------------------------------------------------
cat_ids_py = get_category_ids_by_name(doc, CATEGORY_NAMES)

cat_ids = List[ElementId]()
for cid in cat_ids_py:
    cat_ids.Add(cid)

cat_filter = ElementMulticategoryFilter(cat_ids) if cat_ids.Count > 0 else None

def collect_elements():
    if scope == "Current Selection":
        ids = list(uidoc.Selection.GetElementIds())
        els = [doc.GetElement(i) for i in ids]
        els = [e for e in els if e is not None]
        filtered = []
        for e in els:
            try:
                if is_fab_part(e):
                    filtered.append(e)
                elif e.Category and e.Category.Name in CATEGORY_NAMES:
                    filtered.append(e)
            except:
                pass
        return filtered

    if scope == "Current View":
        collector = FilteredElementCollector(doc, doc.ActiveView.Id)
    else:
        collector = FilteredElementCollector(doc)

    collector = collector.WhereElementIsNotElementType()
    if cat_filter:
        collector = collector.WherePasses(cat_filter)

    filtered = []
    for e in collector:
        try:
            if is_fab_part(e):
                filtered.append(e)
            elif e.Category and e.Category.Name in CATEGORY_NAMES:
                filtered.append(e)
        except:
            pass
    return filtered

elements = collect_elements()
if not elements:
    forms.alert("No elements found for the chosen scope + category filter.", title="Nothing to do")
    script.exit()

# -----------------------------------------------------------------------------
# Write loop (chunked Transactions inside TransactionGroup)
# -----------------------------------------------------------------------------
CHUNK_SIZE = 10000
total = len(elements)

w_ok = w_skip = 0
v_ok = v_skip = 0
b_ok = b_skip = 0

with forms.ProgressBar(title="Writing parameters...", cancellable=True, step=1) as pb:
    tg = TransactionGroup(doc, "Write CP_Weight / CP_Volume / CP_BOM_Category")
    tg.Start()
    try:
        for start in range(0, total, CHUNK_SIZE):
            if pb.cancelled:
                break

            t = Transaction(doc, "Write params ({}-{})".format(start + 1, min(total, start + CHUNK_SIZE)))
            t.Start()
            try:
                chunk = elements[start:start + CHUNK_SIZE]
                for idx, el in enumerate(chunk):
                    if pb.cancelled:
                        break
                    pb.update_progress(start + idx + 1, total)

                    if DO_WEIGHT:
                        val = calc_weight_kg(el)
                        if set_param_double(el, "CP_Weight", val):
                            w_ok += 1
                        else:
                            w_skip += 1

                    if DO_VOLUME:
                        val = calc_volume_ft3(el)
                        if set_param_double(el, "CP_Volume", val):
                            v_ok += 1
                        else:
                            v_skip += 1

                    if DO_BOMCAT:
                        val = calc_bom_category(el)
                        if set_param_string(el, "CP_BOM_Category", val):
                            b_ok += 1
                        else:
                            b_skip += 1

                t.Commit()
            except:
                logger.exception("Chunk failed; rolling back transaction.")
                t.RollBack()

        tg.Assimilate()
    except:
        logger.exception("Failed overall; rolling back transaction group.")
        tg.RollBack()
        raise

msg_lines = [
    "Scope: {0}".format(scope),
    "Elements processed: {0}".format(total),
]
if DO_WEIGHT:
    msg_lines.append("CP_Weight: wrote {0}, skipped {1}".format(w_ok, w_skip))
if DO_VOLUME:
    msg_lines.append("CP_Volume: wrote {0}, skipped {1}".format(v_ok, v_skip))
if DO_BOMCAT:
    msg_lines.append("CP_BOM_Category: wrote {0}, skipped {1}".format(b_ok, b_skip))

forms.alert("\n".join(msg_lines), title="Done")
