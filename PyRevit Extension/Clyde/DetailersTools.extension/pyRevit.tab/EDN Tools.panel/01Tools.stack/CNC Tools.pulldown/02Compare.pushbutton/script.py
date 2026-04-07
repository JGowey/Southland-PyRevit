# -*- coding: utf-8 -*-
"""Compare - Unified comparison tool.

Consolidates Compare Elements, Families, Schedules, and Views into a single
branded WPF window with DataGrid results, filter modes, and CSV/XLS export.

Author: Jeremiah Griffith
Version: 1.0.0
"""

from pyrevit import revit, DB, forms, script
from System import Enum
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import (DataGridTextColumn, DataGridLength,
                                     DataGridLengthUnitType)
from System.Windows.Data import Binding
from Autodesk.Revit.UI.Selection import ObjectType
import System.Windows
import System.Windows.Input
import csv
import codecs
import datetime
import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

doc   = revit.doc
uidoc = revit.uidoc
log   = script.get_logger()

XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")


# ================================================================
#  DATA MODEL
# ================================================================
class ComparisonRow(object):
    """One row in the results DataGrid.

    Properties are set dynamically: Property, IsDifferent, Section,
    plus Item0..ItemN for each compared item.
    """
    def __init__(self, prop_name, values, is_different, section=""):
        self.Property    = prop_name
        self.Values      = values
        self.IsDifferent = is_different
        self.Section     = section
        for i, val in enumerate(values):
            setattr(self, "Item{}".format(i), val)


# ================================================================
#  FORMATTING HELPERS (from Compare Elements)
# ================================================================
def _trim_zeros(numstr):
    if "." not in numstr:
        return numstr
    s = numstr.rstrip("0").rstrip(".")
    return s if s else "0"


def _double_to_proj(val):
    try:
        fo = doc.GetUnits().GetFormatOptions(DB.SpecTypeId.Length)
        unit_id = fo.GetUnitTypeId()
        x = float(val)
        if abs(x) < 1e-9:
            return "0"
        s = "{0:.3f}".format(DB.UnitUtils.ConvertFromInternalUnits(x, unit_id))
        return _trim_zeros(s)
    except:
        try:
            s = "{0:.3f}".format(float(val))
            return _trim_zeros(s)
        except:
            return str(val)


def _string_or_value(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p and p.HasValue:
            try:
                v = p.AsValueString()
                if v:
                    return str(v)
            except:
                pass
            try:
                st = p.StorageType
                if st == DB.StorageType.String:
                    return p.AsString() or ""
                if st == DB.StorageType.Double:
                    return _double_to_proj(p.AsDouble())
                if st == DB.StorageType.Integer:
                    return str(p.AsInteger())
            except:
                pass
    except:
        pass
    return ""


def _get_param_value(param):
    """Safely get a parameter's value as a displayable string."""
    if not param:
        return "N/A"
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString() or ""
        elif st == DB.StorageType.Double:
            return str(round(param.AsDouble(), 4))
        elif st == DB.StorageType.Integer:
            return str(param.AsInteger())
        elif st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                elem = doc.GetElement(eid)
                return getattr(elem, "Name", str(eid.IntegerValue)) if elem else str(eid.IntegerValue)
            return "None"
        else:
            return "Unhandled Type"
    except:
        return "Error"


def _safe_param_dict(paramset):
    result = {}
    for p in paramset:
        try:
            d = p.Definition
            if d and d.Name:
                result[d.Name] = p
        except:
            continue
    return result


# ================================================================
#  ELEMENTS ENGINE
# ================================================================
def _is_fab(e):
    return isinstance(e, DB.FabricationPart)


def _cfg():
    try:
        return DB.FabricationConfiguration.GetFabricationConfiguration(doc)
    except:
        return None


def _resolve_id_to_name(prop_name, id_value):
    cfg = _cfg()
    if cfg is None:
        return str(id_value)
    try:
        if prop_name == "ServiceId":
            return cfg.GetServiceName(id_value) or str(id_value)
        if prop_name == "GaugeId":
            return cfg.GetGaugeName(id_value) or str(id_value)
        if prop_name in ("MaterialId", "ItemMaterialId"):
            return cfg.GetMaterialName(id_value) or str(id_value)
        if prop_name == "SpecificationId":
            return cfg.GetSpecificationName(id_value) or str(id_value)
        if prop_name == "InsulationSpecificationId":
            return cfg.GetInsulationSpecificationName(id_value) or str(id_value)
        if prop_name == "LiningSpecificationId":
            return cfg.GetLiningSpecificationName(id_value) or str(id_value)
    except:
        pass
    return str(id_value)


def _get_connector_name_by_id(elem, end_id_int):
    try:
        cfg = _cfg()
        cm = getattr(elem, "ConnectorManager", None)
        if not (cfg and cm):
            return "-"
        selected = None
        for c in cm.Connectors:
            try:
                if int(getattr(c, "Id", -1)) == end_id_int:
                    selected = c
                    break
            except:
                pass
        if selected is None:
            conns = list(cm.Connectors)
            if len(conns) > end_id_int:
                selected = conns[end_id_int]
        if selected is None:
            return "-"
        fci = selected.GetFabricationConnectorInfo()
        if fci:
            try:
                return cfg.GetFabricationConnectorName(fci.BodyConnectorId) or "-"
            except:
                return "-"
    except:
        pass
    return "-"


def _end_allowance(elem, endnum):
    names = [
        "End {} Connector Allowance", "Connector {} Allowance",
        "End {} Allowance", "Joint {} Allowance",
        "End {} Joint Allowance", "Allowance End {}"
    ]
    for pat in names:
        p = elem.LookupParameter(pat.format(endnum))
        if p and p.HasValue:
            try:
                st = p.StorageType
                if st == DB.StorageType.Double:
                    return _double_to_proj(p.AsDouble())
                if st == DB.StorageType.Integer:
                    return str(p.AsInteger())
                if st == DB.StorageType.String:
                    return p.AsString()
            except:
                pass
    return "-"


def _get_assembly_name(elem):
    try:
        aid = getattr(elem, "AssemblyInstanceId", None)
        if isinstance(aid, DB.ElementId) and aid.IntegerValue > 0:
            inst = doc.GetElement(aid)
            if inst and hasattr(inst, "Name") and inst.Name:
                return str(inst.Name)
    except:
        pass
    return "-"


SKIP_PROPS = set([
    "ConnectorManager", "GeometryChecksum", "HangerRodKit", "LiningArea",
    "Origin", "Oversize", "PartGuid", "PartStatus", "SpoolName",
    "ValidationStatus", "DoubleWallMaterialArea", "InsulationArea",
    "ItemCustomId", "ServiceName", "ServiceType", "Specification",
    "ServiceAbbreviation"
])
SKIP_PREFIXES = ("Product",)
HIDE_PROPS = set(["DisplayName", "DisplayNameSource", "PartType", "ItemName"])

DESC_PARAM_NAMES = [
    "Description", "Part Description", "Item Description",
    "MAP Description", "Fabrication Description", "Button Description"
]

PARAM_ORDER = [
    "DomainType", "Family", "Alias", "ItemNumber", "Assembly Name",
    "EndConnector1", "EndConnector2", "End1Allowance", "End2Allowance",
    "CutType", "CenterlineLength", "Size", "FreeSize", "OverallSize",
    "ServiceId", "Material", "MaterialGauge", "MaterialThickness",
    "Weight", "SheetMetalArea", "HasDoubleWall", "DoubleWallMaterial",
    "DoubleWallMaterialThickness", "HasInsulation", "InsulationSpecification",
    "InsulationThickness", "InsulationType", "HasLining", "LiningThickness",
    "LiningType", "LevelOffset", "BottomOfPartElevation",
    "TopOfPartElevation", "Slope", "IsBoughtOut", "Notes"
]


def _get_param_string(elem, pname):
    try:
        p = elem.LookupParameter(pname)
        if p and p.HasValue:
            if p.StorageType == DB.StorageType.String:
                return p.AsString() or ""
            elif p.StorageType == DB.StorageType.Integer:
                return str(p.AsInteger())
            elif p.StorageType == DB.StorageType.Double:
                return _double_to_proj(p.AsDouble())
    except:
        pass
    return ""


def _get_fab_display_name(elem):
    for nm in DESC_PARAM_NAMES:
        val = _get_param_string(elem, nm)
        if val and val.strip():
            return val.strip()
    try:
        if hasattr(elem, "ItemName") and elem.ItemName:
            return str(elem.ItemName)
    except:
        pass
    try:
        if hasattr(elem, "ItemSourceFileName") and elem.ItemSourceFileName:
            base = os.path.splitext(
                os.path.basename(elem.ItemSourceFileName))[0]
            if base:
                return str(base)
    except:
        pass
    return "-"


def harvest_fp_api(elem):
    """Harvest all comparison properties from a FabricationPart."""
    d = {}
    try:
        t = elem.GetType()
        props = [p for p in t.GetProperties()
                 if p.DeclaringType == t and p.CanRead]
        for p in props:
            pname = str(p.Name)
            if pname in SKIP_PROPS:
                continue
            skip = False
            for pref in SKIP_PREFIXES:
                if pname.startswith(pref):
                    skip = True
                    break
            if skip:
                continue
            try:
                raw = p.GetValue(elem, None)
            except:
                raw = None
            if raw is None:
                val = "-"
            else:
                try:
                    if isinstance(raw, (int, long)):
                        if pname.endswith("Id"):
                            val = _resolve_id_to_name(pname, raw)
                        else:
                            val = str(raw)
                    elif isinstance(raw, float):
                        lower = pname.lower()
                        if lower in ("width", "height", "diameter",
                                     "insulationthickness",
                                     "liningthickness", "length",
                                     "centerlinelength",
                                     "materialthickness"):
                            val = _double_to_proj(raw)
                        else:
                            val = str(raw)
                    elif isinstance(raw, bool):
                        val = "True" if raw else "False"
                    elif isinstance(raw, DB.ElementId):
                        val = str(raw.IntegerValue)
                    else:
                        val = str(raw)
                except:
                    val = str(raw) if raw is not None else "-"
            d[pname] = val
    except:
        pass

    for nm in ("Size", "FreeSize", "OverallSize", "CenterlineLength",
               "MaterialThickness", "InsulationThickness",
               "LiningThickness"):
        val = _string_or_value(elem, nm)
        if val:
            d[nm] = val

    d["EndConnector1"] = _get_connector_name_by_id(elem, 0)
    d["EndConnector2"] = _get_connector_name_by_id(elem, 1)
    d["End1Allowance"] = _end_allowance(elem, 1)
    d["End2Allowance"] = _end_allowance(elem, 2)
    d["Assembly Name"] = _get_assembly_name(elem)

    try:
        disp_val = _get_fab_display_name(elem)
        d["ItemName"] = str(getattr(elem, "ItemName", "") or "-")
        try:
            d["PartType"] = str(
                Enum.GetName(DB.FabricationPartType, elem.PartType)
                or elem.PartType.ToString())
        except:
            d["PartType"] = "-"
        d["DisplayName"] = str(disp_val if disp_val else "-")
        fam_only = d["DisplayName"]
        if fam_only and ":" in fam_only:
            fam_only = fam_only.split(":", 1)[0].strip()
        d["Family"] = fam_only or "-"
    except:
        d["ItemName"] = d["PartType"] = d["DisplayName"] = d["Family"] = "-"

    return d


def compare_elements(elements):
    """Compare 2+ fabrication elements. Returns (rows, headers)."""
    data = [harvest_fp_api(el) for el in elements]
    prop_set = set()
    for d in data:
        prop_set.update(d.keys())
    prop_set = prop_set.difference(HIDE_PROPS)

    all_props = [p for p in PARAM_ORDER if p in prop_set]
    remaining = [p for p in prop_set if p not in PARAM_ORDER]
    all_props.extend(sorted(remaining, key=lambda s: s.lower()))

    rows = []
    for prop in all_props:
        vals = [data[i].get(prop, "-") for i in range(len(elements))]
        is_diff = len(set(vals)) > 1
        rows.append(ComparisonRow(prop, vals, is_diff, "Element Properties"))

    headers = ["E{} (Id: {})".format(i + 1, e.Id.IntegerValue)
               for i, e in enumerate(elements)]
    return rows, headers


# ================================================================
#  FAMILIES ENGINE
# ================================================================
def _get_families_by_category():
    families_by_cat = {}
    all_families = (DB.FilteredElementCollector(doc)
                    .OfClass(DB.Family).ToElements())
    for f in all_families:
        if f.IsUserCreated and f.FamilyCategory:
            cat_name = f.FamilyCategory.Name
            if cat_name not in families_by_cat:
                families_by_cat[cat_name] = []
            families_by_cat[cat_name].append(f)
    return families_by_cat


def pick_two_families():
    """Two-step picker: categories then families. Returns 2 Family objects."""
    families_by_cat = _get_families_by_category()
    if not families_by_cat:
        forms.alert("No loadable families found.", exitscript=True)

    sorted_cats = sorted(families_by_cat.keys())
    selected_cats = forms.SelectFromList.show(
        sorted_cats, title="Step 1: Select Categories",
        multiselect=True, button_name="Next",
        dark_theme=True, height=600)
    if not selected_cats:
        return None

    candidates = []
    for cat in selected_cats:
        candidates.extend(families_by_cat.get(cat, []))
    if len(candidates) < 2:
        forms.alert("Fewer than 2 families in selected categories.")
        return None

    sorted_cands = sorted(candidates, key=lambda f: f.Name)
    picks = forms.SelectFromList.show(
        [f.Name for f in sorted_cands],
        title="Step 2: Pick Exactly Two Families",
        multiselect=True, button_name="Compare",
        dark_theme=True, height=600)
    if not picks or len(picks) != 2:
        forms.alert("Please pick exactly two families.")
        return None

    return [f for f in sorted_cands if f.Name in picks]


def compare_families(fam_a, fam_b, aspects):
    """Compare two families. Returns (rows, headers)."""
    rows = []
    headers = [fam_a.Name, fam_b.Name]

    if "Family Properties" in aspects:
        props_to_check = {
            "Category": lambda f: f.FamilyCategory.Name if f.FamilyCategory else "None",
            "Is Shared": lambda f: "Yes" if (
                f.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED) and
                f.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED).AsInteger() == 1
            ) else "No",
            "Is In-Place": lambda f: str(f.IsInPlace),
        }
        for prop_name, getter in props_to_check.items():
            va, vb = getter(fam_a), getter(fam_b)
            rows.append(ComparisonRow(prop_name, [va, vb],
                                      va != vb, "Family Properties"))

    if "Family Parameters" in aspects:
        params_a = _safe_param_dict(fam_a.Parameters)
        params_b = _safe_param_dict(fam_b.Parameters)
        all_names = sorted(set(params_a.keys()) | set(params_b.keys()))
        for name in all_names:
            va = _get_param_value(params_a.get(name))
            vb = _get_param_value(params_b.get(name))
            rows.append(ComparisonRow(name, [va, vb],
                                      va != vb, "Family Parameters"))

    if "Family Types" in aspects:
        types_a, types_b = {}, {}
        try:
            for tid in fam_a.GetFamilySymbolIds():
                sym = doc.GetElement(tid)
                if sym:
                    types_a[sym.Name] = sym
        except:
            pass
        try:
            for tid in fam_b.GetFamilySymbolIds():
                sym = doc.GetElement(tid)
                if sym:
                    types_b[sym.Name] = sym
        except:
            pass

        all_types = sorted(set(types_a.keys()) | set(types_b.keys()))
        for type_name in all_types:
            ta, tb = types_a.get(type_name), types_b.get(type_name)
            if ta and tb:
                pa = _safe_param_dict(ta.Parameters)
                pb = _safe_param_dict(tb.Parameters)
                all_pn = sorted(set(pa.keys()) | set(pb.keys()))
                for pn in all_pn:
                    va = _get_param_value(pa.get(pn))
                    vb = _get_param_value(pb.get(pn))
                    rows.append(ComparisonRow(
                        pn, [va, vb], va != vb,
                        "Type: {}".format(type_name)))
            else:
                status = "Only in A" if ta else "Only in B"
                rows.append(ComparisonRow(
                    type_name, [status, ""],
                    True, "Family Types"))

    return rows, headers


# ================================================================
#  SCHEDULES ENGINE
# ================================================================
def pick_schedules():
    """Pick 2+ schedules from project. Returns list or None."""
    cands = sorted(
        [s for s in DB.FilteredElementCollector(doc)
         .OfClass(DB.ViewSchedule)],
        key=lambda s: s.Name.lower())
    if len(cands) < 2:
        forms.alert("Need at least two schedules in this model.")
        return None
    labels = [s.Name for s in cands]
    mapby = dict(zip(labels, cands))
    picks = forms.SelectFromList.show(
        labels, "Pick 2+ Schedules", multiselect=True,
        button_name="Compare", dark_theme=True, height=600)
    if not picks or len(picks) < 2:
        return None
    return [mapby[p] for p in picks]


def _defn(s):
    try:
        return s.Definition
    except:
        return None


def _harvest_fields(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            f = d.GetField(i)
            nm = f.GetName() if hasattr(f, "GetName") else "<field>"
            hdr = f.ColumnHeading if hasattr(f, "ColumnHeading") else ""
            hidden = bool(getattr(f, "IsHidden", False))
            calc = bool(getattr(f, "IsCalculatedField", False))
            outl.append((nm, hdr, hidden, calc))
        except:
            pass
    return outl


def _harvest_filters(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFilterCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            flt = d.GetFilter(i)
            fname = "<field>"
            try:
                fid = flt.FieldId
                fld = d.GetField(fid)
                fname = fld.GetName() if hasattr(fld, "GetName") else fname
            except:
                pass
            ftype = str(getattr(flt, "FilterType", ""))
            vstr = ""
            try:
                vstr = flt.GetStringValue()
            except:
                try:
                    vstr = str(flt.GetDoubleValue())
                except:
                    try:
                        vstr = str(flt.GetIntegerValue())
                    except:
                        vstr = ""
            outl.append((fname, ftype, vstr))
        except:
            pass
    return outl


def _harvest_sorting(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetSortGroupFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            sg = d.GetSortGroupField(i)
            fname = "<field>"
            try:
                fld = d.GetField(sg.FieldId)
                fname = fld.GetName() if hasattr(fld, "GetName") else fname
            except:
                pass
            order = ("Ascending" if getattr(sg, "SortOrder", 0) ==
                     DB.SortOrder.Ascending else "Descending")
            header = bool(getattr(sg, "ShowHeader", False))
            footer = bool(getattr(sg, "ShowFooter", False))
            outl.append((fname, order, header, footer))
        except:
            pass
    return outl


def _harvest_formatting(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            f = d.GetField(i)
            nm = f.GetName() if hasattr(f, "GetName") else "<field>"
            width = getattr(f, "GridColumnWidth", None)
            align = str(getattr(f, "HorizontalAlignment", None))
            custom = bool(getattr(f, "HasCustomFormatting", False))
            outl.append((nm, width, align, custom))
        except:
            pass
    return outl


def _harvest_appearance(s):
    m = {}
    try:
        m["ShowTitle"] = str(bool(getattr(s, "ShowTitle", False)))
    except:
        pass
    try:
        m["Title"] = getattr(s, "Title", "")
    except:
        pass
    try:
        m["ShowHeaders"] = str(bool(getattr(s, "ShowHeaders", True)))
    except:
        pass
    try:
        m["ShowGrandTotals"] = str(bool(getattr(s, "ShowGrandTotals", False)))
    except:
        pass
    return m


def compare_schedules(schedules, aspects):
    """Compare 2+ schedules. Returns (rows, headers)."""
    rows = []
    headers = [s.Name for s in schedules]

    if "Fields" in aspects:
        per = [_harvest_fields(s) for s in schedules]
        names = set()
        for lst in per:
            names |= set([x[0] for x in lst])
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                if hit:
                    _, hdr, hidden, calc = hit[0]
                    vals.append("hdr='{}'; hidden={}; calc={}".format(
                        hdr, hidden, calc))
                else:
                    vals.append("Missing")
            is_diff = len(set(vals)) > 1
            rows.append(ComparisonRow(nm, vals, is_diff, "Fields"))

    if "Filters" in aspects:
        per = [_harvest_filters(s) for s in schedules]
        names = set()
        for lst in per:
            names |= set([x[0] for x in lst])
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                same_field = [x for x in lst if x[0] == nm]
                vals.append(
                    "; ".join(["{} {}".format(x[1], x[2])
                               for x in same_field])
                    if same_field else "None")
            is_diff = len(set(vals)) > 1
            rows.append(ComparisonRow(nm, vals, is_diff, "Filters"))

    if "Sorting/Grouping" in aspects:
        per = [_harvest_sorting(s) for s in schedules]
        names = set()
        for lst in per:
            names |= set([x[0] for x in lst])
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                vals.append(
                    "{}; hdr={}; ftr={}".format(
                        hit[0][1], hit[0][2], hit[0][3])
                    if hit else "None")
            is_diff = len(set(vals)) > 1
            rows.append(ComparisonRow(nm, vals, is_diff, "Sorting/Grouping"))

    if "Field Formatting" in aspects:
        per = [_harvest_formatting(s) for s in schedules]
        names = set()
        for lst in per:
            names |= set([x[0] for x in lst])
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                vals.append(
                    "w={}; align={}; custom={}".format(
                        hit[0][1], hit[0][2], hit[0][3])
                    if hit else "None")
            is_diff = len(set(vals)) > 1
            rows.append(ComparisonRow(nm, vals, is_diff, "Field Formatting"))

    if "Appearance" in aspects:
        per = [_harvest_appearance(s) for s in schedules]
        keys = set()
        for m in per:
            keys |= set(m.keys())
        for k in sorted(keys, key=lambda s: s.lower()):
            vals = [str(m.get(k, "")) for m in per]
            is_diff = len(set(vals)) > 1
            rows.append(ComparisonRow(k, vals, is_diff, "Appearance"))

    return rows, headers


# ================================================================
#  VIEWS ENGINE
# ================================================================
def _is_schedule(v):
    return isinstance(v, DB.ViewSchedule)


def _nice_label(v):
    vt = getattr(v, "ViewType", None)
    vtype = str(vt) if vt is not None else "View"
    tag = "TEMPLATE: " if getattr(v, "IsTemplate", False) else ""
    return u"{}{}  ({})".format(tag, v.Name, vtype)


def pick_views_or_templates():
    """Pick 2 views or 2 view templates. Returns list or None."""
    source = forms.CommandSwitchWindow.show(
        ["Project Views", "View Templates"],
        message="What would you like to compare?",
        title="Compare", dark_theme=True, width=420, height=200)
    if not source:
        return None

    if source == "Project Views":
        cands = sorted(
            [v for v in DB.FilteredElementCollector(doc)
             .OfClass(DB.View).ToElements()
             if not v.IsTemplate and not _is_schedule(v)],
            key=lambda x: x.Name.lower())
    else:
        cands = sorted(
            [v for v in DB.FilteredElementCollector(doc)
             .OfClass(DB.View).ToElements()
             if getattr(v, "IsTemplate", False) and not _is_schedule(v)],
            key=lambda x: x.Name.lower())

    if len(cands) < 2:
        forms.alert("Need at least two items to compare.")
        return None

    label_map = {_nice_label(v): v for v in cands}
    labels_sorted = sorted(label_map.keys(), key=lambda s: s.lower())
    picks = forms.SelectFromList.show(
        labels_sorted, title="Pick exactly two",
        multiselect=True, button_name="Compare",
        dark_theme=True, height=600)
    if not picks or len(picks) != 2:
        return None
    return [label_map[p] for p in picks]


def _cat_groups():
    groups = {"Model": [], "Annotation": [], "Analytical": [],
              "Import/Other": []}
    for cat in doc.Settings.Categories:
        try:
            if not isinstance(cat, DB.Category) or cat.Parent is not None:
                continue
            if cat.CategoryType == DB.CategoryType.Model:
                groups["Model"].append(cat)
            elif cat.CategoryType == DB.CategoryType.Annotation:
                groups["Annotation"].append(cat)
            elif cat.CategoryType == DB.CategoryType.AnalyticalModel:
                groups["Analytical"].append(cat)
            else:
                groups["Import/Other"].append(cat)
        except:
            pass
    for k in groups:
        groups[k].sort(key=lambda c: c.Name.lower())
    return groups


def _cat_hidden(vw, cat):
    try:
        return not vw.IsCategoryVisible(cat.Id)
    except:
        try:
            return vw.GetCategoryHidden(cat.Id)
        except:
            return False


def _ogs_as_tuple(ogs):
    try:
        def _get(getter, default=None):
            try:
                return getter()
            except:
                return default
        def _col(c):
            return (c.Red, c.Green, c.Blue) if isinstance(c, DB.Color) else None
        def _eid(eid):
            return eid.IntegerValue if isinstance(eid, DB.ElementId) else None
        return (
            _col(_get(ogs.GetCutLineColor)),
            _col(_get(ogs.GetProjectionLineColor)),
            _get(ogs.GetCutLineWeight),
            _get(ogs.GetProjectionLineWeight),
            _get(ogs.IsHalftone),
            _get(ogs.GetSurfaceTransparency),
            _eid(_get(ogs.GetSurfaceForegroundPatternId)),
            _eid(_get(ogs.GetSurfaceBackgroundPatternId)),
            _eid(_get(ogs.GetCutForegroundPatternId))
        )
    except:
        return None


def compare_views(view_a, view_b, aspects):
    """Compare two views/templates. Returns (rows, headers)."""
    rows = []
    headers = [view_a.Name, view_b.Name]

    aspect_map = {
        "Model": "Model Categories",
        "Annotation": "Annotation Categories",
        "Analytical": "Analytical Model Categories",
        "Import/Other": "Imported Categories"
    }
    all_cats = _cat_groups()

    for group_key, aspect_name in aspect_map.items():
        if aspect_name not in aspects:
            continue
        cats = all_cats.get(group_key, [])

        for c in cats:
            ha = _cat_hidden(view_a, c)
            hb = _cat_hidden(view_b, c)
            sa = "Hidden" if ha else "Shown"
            sb = "Hidden" if hb else "Shown"
            rows.append(ComparisonRow(
                c.Name, [sa, sb], sa != sb,
                "{} Visibility".format(aspect_name)))

        for c in cats:
            ta = _ogs_as_tuple(view_a.GetCategoryOverrides(c.Id))
            tb = _ogs_as_tuple(view_b.GetCategoryOverrides(c.Id))
            is_diff = ta != tb
            va = "Custom" if ta else "Default"
            vb = "Custom" if tb else "Default"
            if is_diff:
                va = "Different"
                vb = "Different"
            else:
                va = "Same"
                vb = "Same"
            rows.append(ComparisonRow(
                c.Name, [va, vb], is_diff,
                "{} Overrides".format(aspect_name)))

    if "Filters" in aspects:
        fa_ids = {eid.IntegerValue
                  for eid in (view_a.GetFilters() or [])}
        fb_ids = {eid.IntegerValue
                  for eid in (view_b.GetFilters() or [])}
        all_ids = sorted(fa_ids | fb_ids)
        for i in all_ids:
            nm = "-"
            e = doc.GetElement(DB.ElementId(i))
            if e:
                nm = e.Name
            in_a = i in fa_ids
            in_b = i in fb_ids
            if in_a and not in_b:
                va, vb = "Present", "Missing"
            elif not in_a and in_b:
                va, vb = "Missing", "Present"
            else:
                oa = view_a.GetFilterOverrides(DB.ElementId(i))
                ob = view_b.GetFilterOverrides(DB.ElementId(i))
                if _ogs_as_tuple(oa) != _ogs_as_tuple(ob):
                    va, vb = "Different overrides", "Different overrides"
                else:
                    va, vb = "Same", "Same"
            rows.append(ComparisonRow(
                nm, [va, vb], va != vb, "Filters"))

    if "Worksets" in aspects:
        try:
            wss = list(DB.FilteredWorksetCollector(doc)
                       .OfKind(DB.WorksetKind.UserWorkset)) \
                if DB.WorksharingUtils.IsWorkshared(doc) else []
        except:
            wss = []
        for ws in sorted(wss, key=lambda w: w.Name.lower()):
            try:
                va = str(view_a.GetWorksetVisibility(ws.Id))
            except:
                va = "Default"
            try:
                vb = str(view_b.GetWorksetVisibility(ws.Id))
            except:
                vb = "Default"
            rows.append(ComparisonRow(
                ws.Name, [va, vb], va != vb, "Worksets"))

    if "Revit Links" in aspects:
        try:
            links = list(DB.FilteredElementCollector(doc)
                         .OfClass(DB.RevitLinkInstance))
        except:
            links = []
        for inst in sorted(links, key=lambda x: x.Name.lower()):
            try:
                ha = inst.IsHidden(view_a)
            except:
                ha = False
            try:
                hb = inst.IsHidden(view_b)
            except:
                hb = False
            sa = "Hidden" if ha else "Shown"
            sb = "Hidden" if hb else "Shown"
            rows.append(ComparisonRow(
                inst.Name, [sa, sb], sa != sb, "Revit Links"))

    return rows, headers


# ================================================================
#  ASPECT LISTS (for the aspect picker per mode)
# ================================================================
ASPECTS = {
    "Families": [
        "Family Properties",
        "Family Parameters",
        "Family Types"
    ],
    "Schedules": [
        "Fields",
        "Filters",
        "Sorting/Grouping",
        "Field Formatting",
        "Appearance"
    ],
    "Views": [
        "Model Categories",
        "Annotation Categories",
        "Analytical Model Categories",
        "Imported Categories",
        "Filters",
        "Worksets",
        "Revit Links"
    ],
}


# ================================================================
#  EXPORT HELPERS
# ================================================================
def export_csv(rows, headers, filepath):
    """Write comparison results to CSV."""
    with codecs.open(filepath, "w", "utf-8") as f:
        w = csv.writer(f, lineterminator="\n")
        w.writerow(["Section", "Property"] + headers + ["Status"])
        for r in rows:
            csv_row = [r.Section, r.Property] + r.Values
            csv_row.append("Different" if r.IsDifferent else "Same")
            w.writerow(csv_row)


def export_xls(rows, headers, filepath):
    """Write comparison results to XLS-compatible HTML table."""
    with codecs.open(filepath, "w", "utf-8") as f:
        f.write('<html xmlns:x="urn:schemas-microsoft-com:office:excel">\n')
        f.write('<head><meta charset="utf-8"/>\n')
        f.write('<style>\n')
        f.write('  table { border-collapse: collapse; font-family: Segoe UI, Arial; font-size: 11px; }\n')
        f.write('  th { background: #2563EB; color: white; padding: 6px 12px; text-align: left; }\n')
        f.write('  td { border: 1px solid #E5E8F0; padding: 4px 12px; }\n')
        f.write('  tr:nth-child(even) { background: #FAFBFF; }\n')
        f.write('  .diff { background: #FFF9E6; }\n')
        f.write('  .same { }\n')
        f.write('</style>\n</head>\n<body>\n')
        f.write('<table>\n')
        f.write('<tr><th>Section</th><th>Property</th>')
        for h in headers:
            f.write('<th>{}</th>'.format(h))
        f.write('<th>Status</th></tr>\n')
        for r in rows:
            cls = "diff" if r.IsDifferent else "same"
            f.write('<tr class="{}"><td>{}</td><td>{}</td>'.format(
                cls, r.Section, r.Property))
            for v in r.Values:
                f.write('<td>{}</td>'.format(v))
            f.write('<td>{}</td></tr>\n'.format(
                "Different" if r.IsDifferent else "Same"))
        f.write('</table>\n</body></html>\n')


# ================================================================
#  WINDOW CLASS (XamlReader + pick-loop)
# ================================================================
class CompareWindow(object):
    """Main window for the unified Compare tool."""

    def __init__(self, xaml_path):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        self.all_rows     = []
        self.headers      = []
        self._pending_pick = None

        self._get_controls()
        self._wire_events()
        self._update_state()

    def _get_controls(self):
        find = self.window.FindName
        self.rb_elements   = find("RbElements")
        self.rb_families   = find("RbFamilies")
        self.rb_schedules  = find("RbSchedules")
        self.rb_views      = find("RbViews")
        self.btn_select    = find("BtnSelectItems")
        self.txt_selection = find("TxtSelectionStatus")
        self.btn_everything  = find("BtnEverything")
        self.btn_differences = find("BtnDifferences")
        self.btn_same        = find("BtnSame")
        self.txt_filter      = find("TxtFilterStatus")
        self.results_card  = find("ResultsCard")
        self.dg_results    = find("DgResults")
        self.txt_results   = find("TxtResultsSummary")
        self.btn_csv       = find("BtnExportCsv")
        self.btn_xls       = find("BtnExportXls")
        self.btn_close     = find("BtnClose")
        self.txt_status    = find("TxtStatus")

    def _wire_events(self):
        self.btn_select.Click      += self._on_select_items
        self.btn_everything.Click  += self._on_everything
        self.btn_differences.Click += self._on_differences
        self.btn_same.Click        += self._on_same
        self.btn_csv.Click         += self._on_export_csv
        self.btn_xls.Click         += self._on_export_xls
        self.btn_close.Click       += self._on_close
        self.window.KeyDown        += self._on_key

    # ---- mode helpers ----
    def _selected_mode(self):
        if self.rb_elements.IsChecked:
            return "Elements"
        if self.rb_families.IsChecked:
            return "Families"
        if self.rb_schedules.IsChecked:
            return "Schedules"
        if self.rb_views.IsChecked:
            return "Views"
        return "Elements"

    def _update_state(self):
        has_rows = len(self.all_rows) > 0
        self.results_card.Visibility = (
            Visibility.Visible if has_rows else Visibility.Collapsed)
        self.btn_everything.IsEnabled  = has_rows
        self.btn_differences.IsEnabled = has_rows
        self.btn_same.IsEnabled        = has_rows

    # ---- select items ----
    def _on_select_items(self, sender, args):
        mode = self._selected_mode()
        if mode == "Elements":
            self._pending_pick = "elements"
            self.window.Hide()
        elif mode == "Families":
            self._run_families()
        elif mode == "Schedules":
            self._run_schedules()
        elif mode == "Views":
            self._run_views()

    def _run_elements_pick(self):
        """Called outside ShowDialog for element pick-loop."""
        cur = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
        fab = [e for e in cur if _is_fab(e)]
        if len(fab) >= 2:
            elements = fab
        else:
            try:
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    "Pick 2-100 Fabrication Parts, then Finish")
                elements = [doc.GetElement(r.ElementId) for r in refs]
                elements = [e for e in elements if _is_fab(e)]
            except:
                elements = []

        if len(elements) < 2:
            self.txt_status.Text = "Need at least 2 fabrication parts."
            return

        self.txt_status.Text = "Comparing {} elements...".format(
            len(elements))
        rows, headers = compare_elements(elements)
        self.all_rows = rows
        self.headers  = headers
        self.txt_selection.Text = "{} elements selected".format(
            len(elements))
        self._bind_results("Everything")
        self._update_state()

    def _run_families(self):
        fams = pick_two_families()
        if not fams or len(fams) != 2:
            self.txt_status.Text = "Family selection cancelled."
            return
        aspects = forms.SelectFromList.show(
            ASPECTS["Families"],
            title="Choose Aspects to Compare",
            multiselect=True, button_name="Compare",
            dark_theme=True, default_selected=True)
        if not aspects:
            return
        rows, headers = compare_families(fams[0], fams[1], aspects)
        self.all_rows = rows
        self.headers  = headers
        self.txt_selection.Text = "{} vs {}".format(
            fams[0].Name, fams[1].Name)
        self._bind_results("Everything")
        self._update_state()

    def _run_schedules(self):
        scheds = pick_schedules()
        if not scheds:
            self.txt_status.Text = "Schedule selection cancelled."
            return
        aspects = forms.SelectFromList.show(
            ASPECTS["Schedules"],
            title="Choose Aspects to Compare",
            multiselect=True, button_name="Compare",
            dark_theme=True, default_selected=True)
        if not aspects:
            return
        rows, headers = compare_schedules(scheds, aspects)
        self.all_rows = rows
        self.headers  = headers
        self.txt_selection.Text = "{} schedules".format(len(scheds))
        self._bind_results("Everything")
        self._update_state()

    def _run_views(self):
        views = pick_views_or_templates()
        if not views or len(views) != 2:
            self.txt_status.Text = "View selection cancelled."
            return
        aspects = forms.SelectFromList.show(
            ASPECTS["Views"],
            title="Choose Aspects to Compare",
            multiselect=True, button_name="Compare",
            dark_theme=True, default_selected=True)
        if not aspects:
            return
        rows, headers = compare_views(views[0], views[1], aspects)
        self.all_rows = rows
        self.headers  = headers
        self.txt_selection.Text = "{} vs {}".format(
            views[0].Name, views[1].Name)
        self._bind_results("Everything")
        self._update_state()

    # ---- filter modes ----
    def _on_everything(self, sender, args):
        self._bind_results("Everything")

    def _on_differences(self, sender, args):
        self._bind_results("Differences")

    def _on_same(self, sender, args):
        self._bind_results("Same")

    def _bind_results(self, mode):
        if mode == "Everything":
            filtered = self.all_rows
        elif mode == "Differences":
            filtered = [r for r in self.all_rows if r.IsDifferent]
        else:
            filtered = [r for r in self.all_rows if not r.IsDifferent]

        diff_count = sum(1 for r in self.all_rows if r.IsDifferent)
        same_count = len(self.all_rows) - diff_count

        if mode == "Everything":
            self.txt_filter.Text = "{} properties ({} diff, {} same)".format(
                len(filtered), diff_count, same_count)
        elif mode == "Differences":
            self.txt_filter.Text = "{} different properties".format(
                len(filtered))
        else:
            self.txt_filter.Text = "{} identical properties".format(
                len(filtered))

        # Rebuild columns
        self.dg_results.Columns.Clear()

        # Section column
        sec_col = DataGridTextColumn()
        sec_col.Header = "Section"
        sec_col.Binding = Binding("Section")
        sec_col.Width = DataGridLength(140)
        self.dg_results.Columns.Add(sec_col)

        # Property column
        prop_col = DataGridTextColumn()
        prop_col.Header = "Property"
        prop_col.Binding = Binding("Property")
        prop_col.Width = DataGridLength(180)
        self.dg_results.Columns.Add(prop_col)

        # Dynamic item columns
        for i, hdr in enumerate(self.headers):
            col = DataGridTextColumn()
            col.Header = hdr
            col.Binding = Binding("Item{}".format(i))
            col.Width = DataGridLength(1, DataGridLengthUnitType.Star)
            self.dg_results.Columns.Add(col)

        # Bind data
        oc = ObservableCollection[ComparisonRow]()
        for row in filtered:
            oc.Add(row)
        self.dg_results.ItemsSource = oc

        self.txt_results.Text = "{} of {} shown".format(
            len(filtered), len(self.all_rows))
        self.txt_status.Text = "Showing {} ({} diff, {} same)".format(
            mode.lower(), diff_count, same_count)

    # ---- export ----
    def _on_export_csv(self, sender, args):
        if not self.all_rows:
            self.txt_status.Text = "No results to export."
            return
        filepath = forms.save_file(
            file_ext="csv", default_name="compare_results.csv")
        if filepath:
            try:
                export_csv(self.all_rows, self.headers, filepath)
                self.txt_status.Text = "CSV exported: {}".format(
                    os.path.basename(filepath))
            except Exception as ex:
                self.txt_status.Text = "Export failed: {}".format(ex)

    def _on_export_xls(self, sender, args):
        if not self.all_rows:
            self.txt_status.Text = "No results to export."
            return
        filepath = forms.save_file(
            file_ext="xls", default_name="compare_results.xls")
        if filepath:
            try:
                export_xls(self.all_rows, self.headers, filepath)
                self.txt_status.Text = "XLS exported: {}".format(
                    os.path.basename(filepath))
            except Exception as ex:
                self.txt_status.Text = "Export failed: {}".format(ex)

    # ---- close / key ----
    def _on_close(self, sender, args):
        self.window.Close()

    def _on_key(self, sender, args):
        if args.Key == System.Windows.Input.Key.Escape:
            self.window.Close()

    # ---- show loop (pick-loop for elements) ----
    def show(self):
        while True:
            self._pending_pick = None
            self.window.ShowDialog()
            if self._pending_pick == "elements":
                self._run_elements_pick()
                continue
            else:
                break


# ================================================================
#  MAIN
# ================================================================
def main():
    win = CompareWindow(XAML_PATH)
    win.show()


if __name__ == "__main__":
    main()
