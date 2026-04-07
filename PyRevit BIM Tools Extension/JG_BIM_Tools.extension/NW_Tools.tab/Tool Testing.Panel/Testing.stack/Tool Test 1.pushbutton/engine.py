# -*- coding: utf-8 -*-
# Filter Manager v2.5 — Revit 2022–2025 (IronPython 2.7 / pyRevit)
# - Import always opens an editable Preview (colors, line-weights, transparency, halftone, fill patterns)
# - Resizable preview window + horizontal scrolling; columns are user-resizable/orderable
# - Guard against header/row-selector clicks (fixes IndexOutOfRange in grid cell handler)
# - CSV schema unchanged from v2.1; functionality identical to v2.4 aside from UI/robustness

import os, csv, re, io
import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System import Byte, Environment, Enum
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Collections.Generic import List
import Autodesk.Revit.DB as DB
from pyrevit import revit, forms

doc = revit.doc

HEADERS = [
    "FilterName","NewFilterName","Categories","SetIndex","Combine","Rules",
    "ProjLineColorRGB","ProjLineWeight",
    "ProjFgPattern","ProjFgColorRGB","ProjBgPattern","ProjBgColorRGB",
    "CutLineColorRGB","CutLineWeight",
    "CutFgPattern","CutFgColorRGB","CutBgPattern","CutBgColorRGB",
    "Halftone","Transparency","ApplyTo","ImportMode"
]

EXAMPLE_ROWS = [
    ["Example Copper","", "Pipes;Pipe Fittings","1","AND",
     "Type Name|String|Contains|Copper|False; Diameter|Number|GreaterOrEqual|2 in|False",
     "0,112,192","4",
     "Solid Fill","0,112,192","","",
     "","","","","",
     "False","25","View: * (current)","UpdateOrCreate"]
]

_LAST_DIR = [Environment.GetFolderPath(Environment.SpecialFolder.Desktop)]

# ------------------------- helpers -------------------------

def _alert(msg, title="Filter Manager", warn=False):
    try:
        forms.alert(msg, title=title, warn_icon=True if warn else False)
    except:
        WinForms.MessageBox.Show(msg, title)

def _bool(s, default=False):
    if s is None or s == "": return default
    return str(s).strip().lower() in ("true","1","yes","y","on")

def _clamp(v, lo, hi):
    try: return max(lo, min(hi, v))
    except: return v

def _parse_rgb(s):
    if not s: return None
    parts = [p.strip() for p in str(s).split(",")]
    if len(parts) != 3: return None
    try:
        r = _clamp(int(parts[0]),0,255)
        g = _clamp(int(parts[1]),0,255)
        b = _clamp(int(parts[2]),0,255)
        return DB.Color(Byte(r), Byte(g), Byte(b))
    except: return None

def _rgb_to_text(r, g, b):
    try: return "%d,%d,%d" % (int(r), int(g), int(b))
    except: return ""

def _color_to_text(c):
    try:
        if c is not None and hasattr(c, "Red"):
            return "%d,%d,%d" % (c.Red, c.Green, c.Blue)
    except: pass
    return ""

def _parse_num_with_units(sval):
    if sval is None: return None, None
    txt = str(sval).strip().lower()
    numrx = r'[^0-9\.\-]+'
    if txt.endswith('"') or txt.endswith(" in") or txt.endswith(" in."):
        try: return float(re.sub(numrx,'',txt))/12.0, "LENGTH"
        except: return None, None
    if txt.endswith("'") or txt.endswith(" ft") or txt.endswith(" ft."):
        try: return float(re.sub(numrx,'',txt)), "LENGTH"
        except: return None, None
    if txt.endswith(" mm"):
        try: return float(re.sub(numrx,'',txt))/304.8, "LENGTH"
        except: return None, None
    if txt.endswith(" cm"):
        try: return (float(re.sub(numrx,'',txt))*10.0)/304.8, "LENGTH"
        except: return None, None
    try: return float(re.sub(numrx,'',txt)), "NUMBER"
    except: return None, None

def _is_type(obj, typename):
    try: return obj is not None and obj.GetType().Name == typename
    except: return False

def _get_attr(obj, *names):
    for n in names:
        try:
            v = getattr(obj, n)
            if callable(v):
                try: return v()
                except: pass
            else:
                return v
        except: pass
        try:
            m = getattr(obj, "Get" + n[0].upper() + n[1:])
            return m()
        except: pass
    return None

def _param_elements_by_name():
    m = {}
    for pe in DB.FilteredElementCollector(doc).OfClass(DB.ParameterElement):
        try:
            if pe.Name and pe.Name not in m: m[pe.Name] = pe
        except: pass
    return m

def _resolve_param_id(param_name):
    if not param_name: return None
    pe = _param_elements_by_name().get(param_name)
    if pe: return pe.Id
    for bip in Enum.GetValues(DB.BuiltInParameter):
        try:
            lbl = DB.LabelUtils.GetLabelFor(bip)
            if lbl and lbl.strip().lower() == param_name.strip().lower():
                return DB.ElementId(bip)
        except: pass
    return None

def _resolve_param_name_from_id(eid):
    try:
        if eid.IntegerValue < 0:
            return DB.LabelUtils.GetLabelFor(DB.BuiltInParameter(eid.IntegerValue))
        pe = doc.GetElement(eid)
        if isinstance(pe, DB.ParameterElement): return pe.Name
    except: pass
    return "UNKNOWN"

def _catnames_from_ids(id_iterable):
    names = []
    for cid in id_iterable:
        name = None
        try:
            if cid.IntegerValue < 0:
                name = DB.LabelUtils.GetLabelFor(DB.BuiltInCategory(cid.IntegerValue))
        except: pass
        if not name:
            try:
                for cat in doc.Settings.Categories:
                    if cat.Id.IntegerValue == cid.IntegerValue:
                        name = cat.Name; break
            except: pass
        if name: names.append(name)
    seen, out = set(), []
    for n in names:
        if n not in seen:
            out.append(n); seen.add(n)
    return ";".join(out)

def _resolve_categories(catnames):
    ids = List[DB.ElementId]()
    if not catnames: return ids
    wanted = [c.strip() for c in catnames.split(";") if c.strip()]
    label_to_bic = {}
    for bic in Enum.GetValues(DB.BuiltInCategory):
        try: label_to_bic[DB.LabelUtils.GetLabelFor(bic).strip().lower()] = bic
        except: pass
    for nm in wanted:
        key = nm.lower()
        bic = label_to_bic.get(key)
        if bic:
            ids.Add(DB.ElementId(bic)); continue
        try:
            for cat in doc.Settings.Categories:
                if cat.Name.strip().lower() == key:
                    ids.Add(cat.Id); break
        except: pass
    return ids

# ---------------- rules extraction ----------------

def _get_case_sensitive(rule):
    try: return bool(rule.CaseSensitive)
    except:
        try: return bool(rule.IsCaseSensitive)
        except:
            try: return bool(rule.GetCaseSensitive())
            except: return False

def _extract_rules_from_elementfilter(ef):
    results = []
    if _is_type(ef, "LogicalAndFilter") or _is_type(ef, "LogicalOrFilter"):
        try:
            for child in ef.GetFilters():
                results.extend(_extract_rules_from_elementfilter(child))
        except: pass
        return results
    if _is_type(ef, "LogicalNotFilter"):
        try:
            inner = _extract_rules_from_elementfilter(ef.NotFilter)
            for r in inner:
                if r["RuleOp"] == "Equals": r["RuleOp"] = "NotEquals"
            results.extend(inner)
        except: pass
        return results
    if isinstance(ef, DB.ElementParameterFilter):
        try: rules = list(ef.GetRules())
        except: rules = []
        try: pid_list = list(ef.GetRuleParameters())
        except: pid_list = []
        for idx, r in enumerate(rules):
            pid = None
            try: pid = r.GetRuleParameter()
            except:
                try: pid = pid_list[idx] if idx < len(pid_list) else None
                except: pid = None
            pname = _resolve_param_name_from_id(pid) if pid else "UNKNOWN"
            if isinstance(r, DB.FilterStringRule):
                ev = r.GetEvaluator()
                if   isinstance(ev, DB.FilterStringEquals):     op = "Equals"
                elif isinstance(ev, DB.FilterStringContains):   op = "Contains"
                elif isinstance(ev, DB.FilterStringBeginsWith): op = "BeginsWith"
                elif isinstance(ev, DB.FilterStringEndsWith):   op = "EndsWith"
                else: op = "Equals"
                try: sval = r.RuleString
                except:
                    try: sval = r.GetRuleString()
                    except: sval = ""
                results.append(dict(ParamName=pname, RuleType="String", RuleOp=op,
                                    RuleValue=sval, CaseSensitive=str(_get_case_sensitive(r))))
            elif isinstance(r, DB.FilterDoubleRule) or isinstance(r, DB.FilterIntegerRule):
                ev = r.GetEvaluator()
                if   isinstance(ev, DB.FilterNumericEquals):         op = "Equals"
                elif isinstance(ev, DB.FilterNumericNotEquals):      op = "NotEquals"
                elif isinstance(ev, DB.FilterNumericGreater):        op = "Greater"
                elif isinstance(ev, DB.FilterNumericGreaterOrEqual): op = "GreaterOrEqual"
                elif isinstance(ev, DB.FilterNumericLess):           op = "Less"
                elif isinstance(ev, DB.FilterNumericLessOrEqual):    op = "LessOrEqual"
                else: op = "Equals"
                try: val = r.RuleValue
                except: val = ""
                results.append(dict(ParamName=pname, RuleType="Number", RuleOp=op,
                                    RuleValue=str(val), CaseSensitive="False"))
            else:
                results.append(dict(ParamName=pname, RuleType="String", RuleOp="Equals",
                                    RuleValue="", CaseSensitive="False"))
        return results
    return results

def _flatten_top(ef):
    if _is_type(ef, "LogicalAndFilter"): return "AND", list(ef.GetFilters())
    if _is_type(ef, "LogicalOrFilter"):  return "OR", list(ef.GetFilters())
    return None, [ef]

# --------------- patterns & overrides ----------------

def _pattern_name_from_id(eid):
    try:
        fpe = doc.GetElement(eid)
        if isinstance(fpe, DB.FillPatternElement):
            return fpe.GetFillPattern().Name
    except: pass
    return ""

def _pattern_id_from_name(name):
    if not name: return DB.ElementId.InvalidElementId
    try:
        coll = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
        for fpe in coll:
            fp = fpe.GetFillPattern()
            if fp and fp.Name.strip().lower() == name.strip().lower():
                return fpe.Id
    except: pass
    return DB.ElementId.InvalidElementId

def _all_pattern_names():
    names = [""]
    try:
        coll = DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement)
        for fpe in coll:
            fp = fpe.GetFillPattern()
            if fp and fp.Name: names.append(fp.Name)
    except: pass
    return sorted(list(set(names)))

def _get_ogs_fields(vw, fid):
    out = {
        "ProjLineColorRGB":"", "ProjLineWeight":"",
        "ProjFgPattern":"", "ProjFgColorRGB":"", "ProjBgPattern":"", "ProjBgColorRGB":"",
        "CutLineColorRGB":"",  "CutLineWeight":"",
        "CutFgPattern":"",  "CutFgColorRGB":"",  "CutBgPattern":"",  "CutBgColorRGB":"",
        "Halftone":"", "Transparency":""
    }
    if fid not in list(vw.GetFilters()): return out
    try:
        ogs = vw.GetFilterOverrides(fid)
        out["ProjLineColorRGB"] = _color_to_text(_get_attr(ogs, "ProjectionLineColor"))
        lw = _get_attr(ogs, "ProjectionLineWeight")
        out["ProjLineWeight"] = "" if lw in (None, 0) else str(lw)

        pid_fg = _get_attr(ogs, "SurfaceForegroundPatternId")
        if pid_fg is None or pid_fg == DB.ElementId.InvalidElementId:
            pid_fg = _get_attr(ogs, "ProjectionFillPatternId")  # 2022
        if pid_fg: out["ProjFgPattern"] = _pattern_name_from_id(pid_fg)
        c_fg = _get_attr(ogs, "SurfaceForegroundPatternColor")
        if c_fg is None: c_fg = _get_attr(ogs, "ProjectionFillColor")  # 2022
        out["ProjFgColorRGB"] = _color_to_text(c_fg)

        pid_bg = _get_attr(ogs, "SurfaceBackgroundPatternId")
        if pid_bg: out["ProjBgPattern"] = _pattern_name_from_id(pid_bg)
        c_bg = _get_attr(ogs, "SurfaceBackgroundPatternColor")
        out["ProjBgColorRGB"] = _color_to_text(c_bg)

        out["CutLineColorRGB"] = _color_to_text(_get_attr(ogs, "CutLineColor"))
        lwc = _get_attr(ogs, "CutLineWeight")
        out["CutLineWeight"] = "" if lwc in (None, 0) else str(lwc)

        cid_fg = _get_attr(ogs, "CutForegroundPatternId")
        if cid_fg is None or cid_fg == DB.ElementId.InvalidElementId:
            cid_fg = _get_attr(ogs, "CutFillPatternId")         # 2022
        if cid_fg: out["CutFgPattern"] = _pattern_name_from_id(cid_fg)
        c_cut_fg = _get_attr(ogs, "CutForegroundPatternColor")
        if c_cut_fg is None: c_cut_fg = _get_attr(ogs, "CutFillColor")  # 2022
        out["CutFgColorRGB"] = _color_to_text(c_cut_fg)

        cid_bg = _get_attr(ogs, "CutBackgroundPatternId")
        if cid_bg: out["CutBgPattern"] = _pattern_name_from_id(cid_bg)
        c_cut_bg = _get_attr(ogs, "CutBackgroundPatternColor")
        out["CutBgColorRGB"] = _color_to_text(c_cut_bg)

        hf = _get_attr(ogs, "Halftone")
        out["Halftone"] = "True" if bool(hf) else "False"
        tr = _get_attr(ogs, "SurfaceTransparency")
        out["Transparency"] = "" if (tr is None or tr < 0) else str(tr)
    except: pass
    return out

def _apply_ogs_fields(vw, fid, d):
    ogs = DB.OverrideGraphicSettings()

    c = _parse_rgb(d.get("ProjLineColorRGB"));     lw = d.get("ProjLineWeight")
    if c: ogs.SetProjectionLineColor(c)
    if lw:
        try: ogs.SetProjectionLineWeight(int(lw))
        except: pass

    pname_fg = d.get("ProjFgPattern"); pid_fg = _pattern_id_from_name(pname_fg)
    if pid_fg and pid_fg != DB.ElementId.InvalidElementId:
        try: ogs.SetSurfaceForegroundPatternId(pid_fg)
        except:
            try: ogs.SetProjectionFillPatternId(pid_fg)  # 2022
            except: pass
    c = _parse_rgb(d.get("ProjFgColorRGB"))
    if c:
        try: ogs.SetSurfaceForegroundPatternColor(c)
        except:
            try: ogs.SetProjectionFillColor(c)           # 2022
            except: pass

    pname_bg = d.get("ProjBgPattern"); pid_bg = _pattern_id_from_name(pname_bg)
    if pid_bg and pid_bg != DB.ElementId.InvalidElementId:
        try: ogs.SetSurfaceBackgroundPatternId(pid_bg)
        except: pass
    c = _parse_rgb(d.get("ProjBgColorRGB"))
    if c:
        try: ogs.SetSurfaceBackgroundPatternColor(c)
        except: pass

    c = _parse_rgb(d.get("CutLineColorRGB"));       lw = d.get("CutLineWeight")
    if c: ogs.SetCutLineColor(c)
    if lw:
        try: ogs.SetCutLineWeight(int(lw))
        except: pass

    cname_fg = d.get("CutFgPattern"); cid_fg = _pattern_id_from_name(cname_fg)
    if cid_fg and cid_fg != DB.ElementId.InvalidElementId:
        try: ogs.SetCutForegroundPatternId(cid_fg)
        except:
            try: ogs.SetCutFillPatternId(cid_fg)          # 2022
            except: pass
    c = _parse_rgb(d.get("CutFgColorRGB"))
    if c:
        try: ogs.SetCutForegroundPatternColor(c)
        except:
            try: ogs.SetCutFillColor(c)                   # 2022
            except: pass

    cname_bg = d.get("CutBgPattern"); cid_bg = _pattern_id_from_name(cname_bg)
    if cid_bg and cid_bg != DB.ElementId.InvalidElementId:
        try: ogs.SetCutBackgroundPatternId(cid_bg)
        except: pass
    c = _parse_rgb(d.get("CutBgColorRGB"))
    if c:
        try: ogs.SetCutBackgroundPatternColor(c)
        except: pass

    h = d.get("Halftone")
    if h not in (None, ""): ogs.SetHalftone(_bool(h))
    t = d.get("Transparency")
    if t not in (None, ""):
        try: ogs.SetSurfaceTransparency(_clamp(int(t), 0, 100))
        except: pass

    if fid not in list(vw.GetFilters()): vw.AddFilter(fid)
    vw.SetFilterOverrides(fid, ogs)

# ---------------- element-filter builders ----------------

def _make_string_rule(pvp, ev, value, cs):
    try:
        return DB.FilterStringRule(pvp, ev, value or "", bool(cs))  # 2023+
    except TypeError:
        return DB.FilterStringRule(pvp, ev, value or "")            # 2022

def _string_filter(param_id, op, value, cs):
    pvp = DB.ParameterValueProvider(param_id)
    op_l = (op or "Equals").strip().lower()
    if op_l == "notequals":
        ev = DB.FilterStringEquals()
        r = _make_string_rule(pvp, ev, value, cs)
        return DB.LogicalNotFilter(DB.ElementParameterFilter(r))
    if   op_l == "contains":   ev = DB.FilterStringContains()
    elif op_l == "beginswith": ev = DB.FilterStringBeginsWith()
    elif op_l == "endswith":   ev = DB.FilterStringEndsWith()
    else:                       ev = DB.FilterStringEquals()
    r = _make_string_rule(pvp, ev, value, cs)
    return DB.ElementParameterFilter(r)

def _number_filter(param_id, op, val):
    pvp = DB.ParameterValueProvider(param_id)
    op_l = (op or "Equals").strip().lower()
    if   op_l == "equals":         ev = DB.FilterNumericEquals()
    elif op_l == "notequals":      ev = DB.FilterNumericNotEquals()
    elif op_l == "greater":        ev = DB.FilterNumericGreater()
    elif op_l == "greaterorequal": ev = DB.FilterNumericGreaterOrEqual()
    elif op_l == "less":           ev = DB.FilterNumericLess()
    elif op_l == "lessorequal":    ev = DB.FilterNumericLessOrEqual()
    else:                           ev = DB.FilterNumericEquals()
    r = DB.FilterDoubleRule(pvp, ev, val, 1e-06)
    return DB.ElementParameterFilter(r)

def _combine(filters, use_and):
    if not filters: return None
    if len(filters) == 1: return filters[0]
    lst = List[DB.ElementFilter]()
    for f in filters: lst.Add(f)
    return DB.LogicalAndFilter(lst) if use_and else DB.LogicalOrFilter(lst)

# ---------------- rules text ----------------

def _rules_to_text(rules_list):
    toks = []
    for r in rules_list:
        toks.append("{0}|{1}|{2}|{3}|{4}".format(
            r.get("ParamName",""),
            r.get("RuleType",""),
            r.get("RuleOp",""),
            r.get("RuleValue",""),
            r.get("CaseSensitive","False")
        ))
    return "; ".join(toks)

def _rules_from_text(txt):
    items = []
    if not txt: return items
    parts = [p.strip() for p in txt.split(";") if p.strip()]
    for p in parts:
        fields = [x.strip() for x in p.split("|")]
        if len(fields) < 5: continue
        items.append(dict(
            ParamName=fields[0],
            RuleType=fields[1],
            RuleOp=fields[2],
            RuleValue=fields[3],
            CaseSensitive=fields[4]
        ))
    return items

# ---------------- exports ----------------

def _safe_save_dialog(default_name):
    sfd = WinForms.SaveFileDialog()
    sfd.Filter = "CSV (*.csv)|*.csv"
    sfd.FileName = default_name
    try: sfd.InitialDirectory = _LAST_DIR[0]
    except: pass
    if sfd.ShowDialog() != WinForms.DialogResult.OK: return None
    path = sfd.FileName
    try: _LAST_DIR[0] = os.path.dirname(path)
    except: pass
    base, ext = os.path.splitext(path)
    n = 1
    while True:
        try:
            with io.open(path, 'w', encoding='utf-8') as _:
                pass
            break
        except:
            n += 1
            path = u"%s (%d)%s" % (base, n, ext)
    return path

def export_blank_template():
    out = _safe_save_dialog("FilterPack_Template.csv")
    if not out: return
    with io.open(out, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f); w.writerow(HEADERS)
        for row in EXAMPLE_ROWS: w.writerow(row)
    _alert("Template saved:\n" + out)

def _rows_for_pfe(pfe, include_overrides_from=None):
    fname = pfe.Name
    cats  = _catnames_from_ids(pfe.GetCategories())
    ef    = pfe.GetElementFilter()
    topc, children = _flatten_top(ef)
    sets = children if topc in ("AND","OR") else [ef]
    rows = []
    for idx, node in enumerate(sets, 1):
        rules = _extract_rules_from_elementfilter(node)
        combine_inside, _ = _flatten_top(node)
        combine_txt = combine_inside or ("AND" if topc != "OR" else "OR")
        if include_overrides_from:
            ogs_map = _get_ogs_fields(include_overrides_from, pfe.Id)
            apply_to_txt = ("Template: " if include_overrides_from.IsTemplate else "View: ") + include_overrides_from.Name
        else:
            ogs_map = {}; apply_to_txt = ""
        rules_text = _rules_to_text(rules)
        rows.append([
            fname,"",cats,str(idx),combine_txt,rules_text,
            ogs_map.get("ProjLineColorRGB",""), ogs_map.get("ProjLineWeight",""),
            ogs_map.get("ProjFgPattern",""),    ogs_map.get("ProjFgColorRGB",""),
            ogs_map.get("ProjBgPattern",""),    ogs_map.get("ProjBgColorRGB",""),
            ogs_map.get("CutLineColorRGB",""),  ogs_map.get("CutLineWeight",""),
            ogs_map.get("CutFgPattern",""),     ogs_map.get("CutFgColorRGB",""),
            ogs_map.get("CutBgPattern",""),     ogs_map.get("CutBgColorRGB",""),
            ogs_map.get("Halftone",""),         ogs_map.get("Transparency",""),
            apply_to_txt, "UpdateOrCreate"
        ])
    return rows

def export_base_filters():
    pfes = list(DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement))
    if not pfes:
        _alert("No native rule filters found.", warn=True); return
    labels = []
    for p in sorted(pfes, key=lambda x: x.Name or ""):
        try: labels.append("%s   {%s}" % (p.Name, _catnames_from_ids(p.GetCategories())))
        except: labels.append(p.Name)
    picked = forms.SelectFromList.show(labels, title="Pick base filters (no overrides)", multiselect=True)
    if not picked: return
    label_to_pfe = {("%s   {%s}" % (p.Name, _catnames_from_ids(p.GetCategories()))): p for p in pfes}
    out = _safe_save_dialog("FilterPack_Base_Selected.csv")
    if not out: return
    rows = []
    for lab in picked:
        pfe = label_to_pfe.get(lab)
        if not pfe: pfe = next((x for x in pfes if lab.startswith(x.Name)), None)
        if pfe: rows.extend(_rows_for_pfe(pfe, include_overrides_from=None))
    with io.open(out, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f); w.writerow(HEADERS); w.writerows(rows)
    _alert("Exported %d filter(s) to:\n%s" % (len(picked), out))

def export_from_view_with_overrides():
    all_views = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
    names = [v.Name for v in all_views]
    pickname = forms.SelectFromList.show(names, title="Pick a view/template to pull overrides from", multiselect=False)
    if not pickname: return
    src = next((v for v in all_views if v.Name == pickname), None)
    if not src: _alert("Could not resolve selected view/template.", warn=True); return

    src_filters = []
    for fid in src.GetFilters():
        el = doc.GetElement(fid)
        if isinstance(el, DB.ParameterFilterElement): src_filters.append(el)
    if not src_filters:
        _alert("Selected view/template has no filters.", warn=True); return

    choices = sorted([p.Name for p in src_filters])
    picked_names = forms.SelectFromList.show(
        choices,
        title="Pick filters from '%s' to export (includes overrides)" % pickname,
        multiselect=True
    )
    if not picked_names: return

    clean = re.sub(r'[\\/:*?\"<>|]+','_', pickname)
    out = _safe_save_dialog("FilterPack_From_%s.csv" % clean)
    if not out: return

    rows = []
    name_to_pfe = {p.Name:p for p in src_filters}
    for nm in picked_names:
        pfe = name_to_pfe.get(nm)
        if pfe: rows.extend(_rows_for_pfe(pfe, include_overrides_from=src))
    with io.open(out, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f); w.writerow(HEADERS); w.writerows(rows)
    _alert("Exported %d filter(s) with overrides from '%s' to:\n%s" % (len(picked_names), pickname, out))

# ---------------- import planning ----------------

def _group_rows(rows):
    groups = {}
    for r in rows:
        base = (r.get("FilterName") or "").strip()
        newn = (r.get("NewFilterName") or "").strip()
        tgt  = newn if newn else base
        if not tgt: continue
        si = r.get("SetIndex") or "1"
        try: si = str(int(si))
        except: si = "1"
        groups.setdefault(tgt, {}).setdefault(si, []).append(r)
    return groups

def _read_csv_rows(path):
    with io.open(path, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f); return [r for r in reader]

def _find_targets(apply_to):
    views, templates = [], []
    if not apply_to: return views, templates
    items = [s.strip() for s in apply_to.split(";") if s.strip()]
    allv = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
    byname = {}
    for v in allv:
        try: byname.setdefault(v.Name, []).append(v)
        except: pass
    for it in items:
        low = it.lower()
        if low in ("view: * (current)","view:* (current)","view:*"):
            if doc.ActiveView: views.append(doc.ActiveView)
            continue
        m = re.match(r'^(View|Template)\s*:\s*(.+)$', it, re.I)
        if not m: continue
        kind = m.group(1).lower(); nm = m.group(2).strip()
        for v in byname.get(nm, []):
            if kind == "template" and v.IsTemplate: templates.append(v)
            if kind == "view" and not v.IsTemplate: views.append(v)
    return views, templates

def _plan_import(rows, create_only, apply_to_picked, picked_targets):
    errors = []
    groups = _group_rows(rows)
    plans = []

    for tgt_name, setmap in groups.items():
        cat_ids = List[DB.ElementId](); seen = set()
        for rows_in_set in setmap.values():
            for it in rows_in_set:
                ids = _resolve_categories(it.get("Categories") or "")
                for cid in ids:
                    if cid.IntegerValue not in seen:
                        cat_ids.Add(cid); seen.add(cid.IntegerValue)
        if cat_ids.Count == 0:
            errors.append("Skip [%s]: no valid Categories." % tgt_name); continue

        set_filters = []
        for si, items in sorted(setmap.items(), key=lambda kv: int(kv[0])):
            combine = (items[0].get("Combine") or "AND").strip().upper()
            use_and = (combine == "AND")
            inner = []
            for it in items:
                rules_blob = (it.get("Rules") or "").strip()
                rule_objs = _rules_from_text(rules_blob) if rules_blob else []
                if not rule_objs and (it.get("ParamName") or "").strip():
                    rule_objs = [dict(
                        ParamName=(it.get("ParamName") or "").strip(),
                        RuleType=(it.get("RuleType") or "String").strip(),
                        RuleOp=(it.get("RuleOp") or "Equals").strip(),
                        RuleValue=(it.get("RuleValue") or "").strip(),
                        CaseSensitive=(it.get("CaseSensitive") or "False").strip()
                    )]
                for r in rule_objs:
                    pname = r.get("ParamName","")
                    if not pname: continue
                    pid = _resolve_param_id(pname)
                    if not pid:
                        errors.append("[%s]: Param '%s' not found." % (tgt_name, pname)); continue
                    rtype = (r.get("RuleType") or "String").strip().lower()
                    rop   = (r.get("RuleOp") or "Equals").strip()
                    rval  = r.get("RuleValue") or ""
                    if rtype == "string":
                        ef = _string_filter(pid, rop, rval, _bool(r.get("CaseSensitive"), False))
                    else:
                        num, _ = _parse_num_with_units(rval)
                        if num is None:
                            errors.append("[%s]: Bad numeric value '%s'." % (tgt_name, rval)); continue
                        ef = _number_filter(pid, rop, num)
                    inner.append(ef)
            if inner: set_filters.append(_combine(inner, use_and))
        if not set_filters:
            errors.append("Skip [%s]: no valid rules." % tgt_name); continue

        combined = _combine(set_filters, False)

        def _find_existing(name):
            for p in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
                try:
                    if p.Name == name: return p
                except: pass
            return None

        sample = list(setmap.values())[0][0]
        mode = (sample.get("ImportMode") or "UpdateOrCreate").strip().lower()
        exists = _find_existing(tgt_name)

        action = "create"; final_name = tgt_name
        if mode == "createnew":
            if exists:
                existing_names = set()
                for p in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
                    try: existing_names.add(p.Name)
                    except: pass
                base = tgt_name; n = 2; name = base
                while name in existing_names: name = "%s (%d)" % (base, n); n += 1
                final_name = name
            action = "create"
        elif mode == "update":
            if not exists:
                errors.append("Skip [%s]: Update mode but filter not found." % tgt_name); continue
            action = "update"
        else:
            action = "update" if exists else "create"

        ogs_dict, apply_parts = {}, []
        if not create_only:
            for rows_in_set in setmap.values():
                for it in rows_in_set:
                    if (it.get("ApplyTo") or "").strip():
                        apply_parts.append(it.get("ApplyTo").strip())
                    for k in ("ProjLineColorRGB","ProjLineWeight",
                              "ProjFgPattern","ProjFgColorRGB","ProjBgPattern","ProjBgColorRGB",
                              "CutLineColorRGB","CutLineWeight",
                              "CutFgPattern","CutFgColorRGB","CutBgPattern","CutBgColorRGB",
                              "Halftone","Transparency"):
                        val = it.get(k)
                        if val is not None and val != "": ogs_dict[k] = val
            if rows:
                r0 = rows[0]
                if "ProjFillPattern" in r0 and "ProjFgPattern" not in ogs_dict:
                    ogs_dict["ProjFgPattern"] = r0.get("ProjFillPattern") or r0.get("FillPattern") or ""
                if "ProjFillColorRGB" in r0 and "ProjFgColorRGB" not in ogs_dict:
                    ogs_dict["ProjFgColorRGB"] = r0.get("ProjFillColorRGB") or r0.get("FillColorRGB") or ""
                if "CutFillPattern" in r0 and "CutFgPattern" not in ogs_dict:
                    ogs_dict["CutFgPattern"] = r0.get("CutFillPattern") or ""
                if "CutFillColorRGB" in r0 and "CutFgColorRGB" not in ogs_dict:
                    ogs_dict["CutFgColorRGB"] = r0.get("CutFillColorRGB") or ""
                if "LineColorRGB" in r0 and "ProjLineColorRGB" not in ogs_dict:
                    ogs_dict["ProjLineColorRGB"] = r0.get("LineColorRGB") or ""
                if "LineWeight" in r0 and "ProjLineWeight" not in ogs_dict:
                    ogs_dict["ProjLineWeight"] = r0.get("LineWeight") or ""

        if apply_to_picked and not create_only:
            target_txt = "; ".join(picked_targets)
        else:
            target_txt = "; ".join(apply_parts)
        views, templates = _find_targets(target_txt)

        plans.append(dict(
            name=tgt_name, final_name=final_name, action=action, mode=mode,
            cat_ids=cat_ids, combined=combined, ogs=ogs_dict,
            views=views, templates=templates, will_apply=(not create_only)
        ))
    return plans, errors

# ---------------- Preview dialog (editable) ----------------

_COLOR_KEYS = [
    ("ProjLineColorRGB","Proj Line Color"),
    ("ProjFgColorRGB","Proj FG Color"),
    ("ProjBgColorRGB","Proj BG Color"),
    ("CutLineColorRGB","Cut Line Color"),
    ("CutFgColorRGB","Cut FG Color"),
    ("CutBgColorRGB","Cut BG Color"),
]
_NUM_KEYS = [
    ("ProjLineWeight","Proj LW"),
    ("CutLineWeight","Cut LW"),
    ("Transparency","Transparency (0–100)"),
]
_CHECK_KEYS = [("Halftone","Halftone")]
_PATTERN_KEYS = [
    ("ProjFgPattern","Proj FG Pattern"),
    ("ProjBgPattern","Proj BG Pattern"),
    ("CutFgPattern","Cut FG Pattern"),
    ("CutBgPattern","Cut BG Pattern"),
]

def _draw_col_to_cell(cell, rgb_txt):
    try:
        if rgb_txt and rgb_txt.strip():
            parts = [int(x.strip()) for x in rgb_txt.split(",")]
            if len(parts) == 3:
                cell.Style.BackColor = Drawing.Color.FromArgb(parts[0], parts[1], parts[2])
                cell.Value = " "
            else:
                cell.Style.BackColor = Drawing.Color.Empty; cell.Value = ""
        else:
            cell.Style.BackColor = Drawing.Color.Empty; cell.Value = ""
    except:
        cell.Style.BackColor = Drawing.Color.Empty; cell.Value = ""

class PreviewDialog(WinForms.Form):
    def __init__(self, plans, errors, title="Import Preview"):
        WinForms.Form.__init__(self)
        self.Text = title
        self.StartPosition = WinForms.FormStartPosition.CenterScreen
        self.FormBorderStyle = WinForms.FormBorderStyle.Sizable
        self.MaximizeBox = True; self.MinimizeBox = True
        self.AutoScaleMode = WinForms.AutoScaleMode.Dpi
        self.Width = 1200; self.Height = 640; self.MinimumSize = Drawing.Size(900, 480)
        self.Padding = WinForms.Padding(10)

        layout = WinForms.TableLayoutPanel()
        layout.Dock = WinForms.DockStyle.Fill
        layout.ColumnCount = 1; layout.RowCount = 3
        layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))
        layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.Percent, 100))
        layout.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))
        self.Controls.Add(layout)

        hdr = WinForms.Label()
        hdr.Text = "Preview — edit settings, then Apply. Tip: resize columns; horizontal scroll is enabled."
        hdr.AutoSize = True; hdr.Margin = WinForms.Padding(0,0,0,8)
        layout.Controls.Add(hdr, 0, 0)

        self.grid = WinForms.DataGridView()
        self.grid.AllowUserToAddRows = False; self.grid.AllowUserToDeleteRows = False
        self.grid.SelectionMode = WinForms.DataGridViewSelectionMode.FullRowSelect
        self.grid.AutoSizeColumnsMode = WinForms.DataGridViewAutoSizeColumnsMode.None
        self.grid.AutoSizeRowsMode = WinForms.DataGridViewAutoSizeRowsMode.None
        self.grid.Dock = WinForms.DockStyle.Fill
        self.grid.ScrollBars = WinForms.ScrollBars.Both
        self.grid.AllowUserToResizeColumns = True
        self.grid.AllowUserToOrderColumns = True
        self.grid.ColumnHeadersHeightSizeMode = WinForms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        layout.Controls.Add(self.grid, 0, 1)

        # basic columns
        def _add_txt(name, header, ro=True, width=90):
            c = WinForms.DataGridViewTextBoxColumn()
            c.HeaderText = header; c.Name = name; c.ReadOnly = ro; c.Width = width; c.MinimumWidth = 60
            self.grid.Columns.Add(c)
        _add_txt("Action","Action", ro=True, width=80)
        _add_txt("Name","Name", ro=True, width=240)
        _add_txt("Views","Views", ro=True, width=70)
        _add_txt("Templates","Templates", ro=True, width=90)

        # pattern dropdowns
        pattern_names = _all_pattern_names()
        for key, title in _PATTERN_KEYS:
            c = WinForms.DataGridViewComboBoxColumn()
            c.HeaderText = title; c.Name = key; c.Width = 140; c.MinimumWidth = 120
            for nm in pattern_names: c.Items.Add(nm)
            self.grid.Columns.Add(c)

        # numeric / checkbox
        for key, title in _NUM_KEYS:
            _add_txt(key, title, ro=False, width=140)

        cb = WinForms.DataGridViewCheckBoxColumn()
        cb.HeaderText = "Halftone"; cb.Name = "Halftone"; cb.Width = 90
        self.grid.Columns.Add(cb)

        # color columns (click to change)
        for key, title in _COLOR_KEYS:
            c = WinForms.DataGridViewTextBoxColumn()
            c.HeaderText = title; c.Name = key; c.Width = 120; c.MinimumWidth = 110
            self.grid.Columns.Add(c)

        # load rows
        self.plans = plans
        for p in plans:
            r = self.grid.Rows[self.grid.Rows.Add()]
            r.Cells["Action"].Value = ("Create" if p["action"]=="create" else "Update")
            r.Cells["Name"].Value = p["final_name"]
            r.Cells["Views"].Value = len(p["views"]) if p["will_apply"] else 0
            r.Cells["Templates"].Value = len(p["templates"]) if p["will_apply"] else 0
            ogs = p["ogs"]

            for key, _title in _PATTERN_KEYS:
                try: r.Cells[key].Value = ogs.get(key, "")
                except: pass

            r.Cells["ProjLineWeight"].Value = ogs.get("ProjLineWeight","")
            r.Cells["CutLineWeight"].Value  = ogs.get("CutLineWeight","")
            r.Cells["Transparency"].Value   = ogs.get("Transparency","")
            try: r.Cells["Halftone"].Value  = _bool(ogs.get("Halftone","False"))
            except: pass

            for key, _title in _COLOR_KEYS:
                _draw_col_to_cell(r.Cells[key], ogs.get(key,""))
            r.Tag = p

        # buttons
        btns = WinForms.FlowLayoutPanel()
        btns.FlowDirection = WinForms.FlowDirection.RightToLeft
        btns.Dock = WinForms.DockStyle.Fill
        layout.Controls.Add(btns, 0, 2)

        close = WinForms.Button(); close.Text = "Cancel"; close.Click += (lambda s,a: self.Close()); btns.Controls.Add(close)

        apply = WinForms.Button(); apply.Text = "Apply"
        def _on_apply(sender, args):
            for r in self.grid.Rows:
                p = r.Tag
                if p is None: continue
                ogs = p["ogs"]
                # patterns
                for key, _ in _PATTERN_KEYS:
                    try: ogs[key] = r.Cells[key].Value or ""
                    except: pass
                # numerics
                def _numcell(name, clamp=None):
                    v = r.Cells[name].Value
                    if v is None or str(v).strip()=="":
                        ogs[name] = ""; return
                    try:
                        n = int(float(str(v)))
                        if clamp is not None: n = _clamp(n, clamp[0], clamp[1])
                        ogs[name] = str(n)
                    except:
                        pass
                _numcell("ProjLineWeight")
                _numcell("CutLineWeight")
                _numcell("Transparency", (0,100))
                # halftone
                try: ogs["Halftone"] = "True" if bool(r.Cells["Halftone"].Value) else "False"
                except: pass
                # colors
                for key, _ in _COLOR_KEYS:
                    try:
                        col = r.Cells[key].Style.BackColor
                        if col.A != 0: ogs[key] = _rgb_to_text(col.R, col.G, col.B)
                        else: ogs[key] = ""
                    except:
                        pass
            self.DialogResult = WinForms.DialogResult.OK
            self.Close()
        apply.Click += _on_apply
        btns.Controls.Add(apply)

        # color pickers (robust against header/invalid clicks)
        def _cell_click(sender, e):
            try:
                if e is None: return
                if e.RowIndex is None or e.ColumnIndex is None: return
                if e.RowIndex < 0 or e.ColumnIndex < 0: return   # headers / selector
                if e.RowIndex >= self.grid.Rows.Count: return
                if e.ColumnIndex >= self.grid.Columns.Count: return
                colname = self.grid.Columns[e.ColumnIndex].Name
                names = [k for (k, _t) in _COLOR_KEYS]
                if colname not in names: return
                cell = self.grid.Rows[e.RowIndex].Cells[e.ColumnIndex]
                cd = WinForms.ColorDialog()
                try: cd.Color = cell.Style.BackColor
                except: pass
                if cd.ShowDialog() == WinForms.DialogResult.OK:
                    cell.Style.BackColor = cd.Color; cell.Value = " "
            except:
                # swallow any odd clicks from row headers, etc.
                return
        self.grid.CellClick += _cell_click

        if errors:
            warn = "\n".join(["⚠ " + e for e in errors])
            _alert(warn, title="Warnings")

# ---------------- apply ----------------

def _summarize_and_apply(plans):
    created = updated = skipped = 0
    applied_views = applied_templates = 0
    lines = []
    t = DB.Transaction(doc, "Filter Import v2.5")
    t.Start()
    try:
        for p in plans:
            action = p["action"]; tgt_name = p["name"]
            final_name = p["final_name"]; cat_ids = p["cat_ids"]; combined = p["combined"]

            def _find_existing(name):
                for x in DB.FilteredElementCollector(doc).OfClass(DB.ParameterFilterElement):
                    try:
                        if x.Name == name: return x
                    except: pass
                return None

            if action == "create":
                pfe = DB.ParameterFilterElement.Create(doc, final_name, cat_ids, combined)
                created += 1; act_text = "created"
            else:
                pfe = _find_existing(tgt_name)
                if not pfe:
                    lines.append("⚠ Skip [%s]: Update mode but not found." % tgt_name)
                    skipped += 1; continue
                pfe.SetCategories(cat_ids); pfe.SetElementFilter(combined)
                updated += 1; act_text = "updated"

            if p["will_apply"]:
                views = p["views"]; templates = p["templates"]
                if (not views) and (not templates) and doc.ActiveView:
                    views = [doc.ActiveView]
                for v in views: _apply_ogs_fields(v, pfe.Id, p["ogs"])
                for v in templates: _apply_ogs_fields(v, pfe.Id, p["ogs"])
                applied_views += len(views); applied_templates += len(templates)
                lines.append("OK [%s]: %s | views=%d templates=%d" %
                             (final_name, act_text, len(views), len(templates)))
            else:
                lines.append("OK [%s]: %s (no apply)" % (final_name, act_text))
        t.Commit()
    except Exception as e:
        t.RollBack()
        _alert("Import failed:\n" + str(e), warn=True); return
    summary = "Created %d • Updated %d • Skipped %d • Applied to %d view(s), %d template(s)" % (
        created, updated, skipped, applied_views, applied_templates)
    body = "\n".join(lines[:400]) + ("\n... (truncated)" if len(lines) > 400 else "")
    _alert(summary + "\n\n" + body, title="Import Results")

# ---------------- run flows ----------------

def _open_csv_path():
    ofd = WinForms.OpenFileDialog(); ofd.Filter = "CSV (*.csv)|*.csv"
    try: ofd.InitialDirectory = _LAST_DIR[0]
    except: pass
    if ofd.ShowDialog() != WinForms.DialogResult.OK: return None
    try: _LAST_DIR[0] = os.path.dirname(ofd.FileName)
    except: pass
    return ofd.FileName

def _run_import(create_only=False, apply_to_picked=False, force_active=False):
    path = _open_csv_path()
    if not path: return
    rows = _read_csv_rows(path)

    picked_targets = []
    if not create_only:
        if force_active:
            picked_targets = ["View: * (current)"]; apply_to_picked = True
        elif apply_to_picked:
            allv = list(DB.FilteredElementCollector(doc).OfClass(DB.View))
            labels = [("Template: " if v.IsTemplate else "View: ") + v.Name for v in sorted(allv, key=lambda x: x.Name or "")]
            picked = forms.SelectFromList.show(labels, title="Pick views/templates to apply filters", multiselect=True)
            if not picked:
                _alert("Nothing selected. Import cancelled.", warn=True); return
            picked_targets = picked

    plans, errors = _plan_import(rows, create_only, apply_to_picked, picked_targets)

    dlg = PreviewDialog(plans, errors, title="Import Preview")
    if dlg.ShowDialog() == WinForms.DialogResult.OK:
        _summarize_and_apply(plans)

# ---------------- UI ----------------

def _make_button(text, on_click):
    b = WinForms.Button()
    b.Text = text; b.AutoSize = True; b.Dock = WinForms.DockStyle.Fill
    b.Margin = WinForms.Padding(6, 4, 6, 4); b.Click += on_click
    return b

def show_ui():
    dlg = WinForms.Form()
    dlg.Text = "Filter Manager v2.5"
    dlg.StartPosition = WinForms.FormStartPosition.CenterScreen
    dlg.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog
    dlg.MaximizeBox = False; dlg.MinimizeBox = False
    dlg.AutoScaleMode = WinForms.AutoScaleMode.Dpi
    dlg.Padding = WinForms.Padding(10); dlg.AutoSize = True
    dlg.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink

    tlp = WinForms.TableLayoutPanel()
    tlp.ColumnCount = 2  # Exports | Imports
    tlp.RowCount = 6
    tlp.Dock = WinForms.DockStyle.Fill; tlp.AutoSize = True
    tlp.AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink
    tlp.Padding = WinForms.Padding(0); tlp.Margin = WinForms.Padding(0)
    tlp.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50.0))
    tlp.ColumnStyles.Add(WinForms.ColumnStyle(WinForms.SizeType.Percent, 50.0))
    for _ in range(6): tlp.RowStyles.Add(WinForms.RowStyle(WinForms.SizeType.AutoSize))

    lbl = WinForms.Label()
    lbl.Text = ("Workflow: Export or template → Edit CSV → Import (opens Preview) → Adjust colors/weights/patterns → Apply.")
    lbl.AutoSize = True; lbl.Margin = WinForms.Padding(0, 0, 0, 8)
    tlp.Controls.Add(lbl, 0, 0); tlp.SetColumnSpan(lbl, 2)

    sec1 = WinForms.Label(); sec1.Text = "Export"; sec1.AutoSize = True
    sec1.Font = Drawing.Font(sec1.Font, Drawing.FontStyle.Bold); tlp.Controls.Add(sec1, 0, 1)
    sec2 = WinForms.Label(); sec2.Text = "Import"; sec2.AutoSize = True
    sec2.Font = Drawing.Font(sec2.Font, Drawing.FontStyle.Bold); tlp.Controls.Add(sec2, 1, 1)

    # Exports (left)
    tlp.Controls.Add(_make_button("Blank Template", lambda s,a: export_blank_template()), 0, 2)
    tlp.Controls.Add(_make_button("Pick Base Filters", lambda s,a: export_base_filters()), 0, 3)
    tlp.Controls.Add(_make_button("From a View (includes overrides)", lambda s,a: export_from_view_with_overrides()), 0, 4)

    # Imports (right) — always open preview
    tlp.Controls.Add(_make_button("Import: Create filters only", lambda s,a: _run_import(create_only=True, apply_to_picked=False)), 1, 2)
    tlp.Controls.Add(_make_button("Import + Apply (pick views/templates)", lambda s,a: _run_import(create_only=False, apply_to_picked=True)), 1, 3)
    tlp.Controls.Add(_make_button("Import + Apply to ACTIVE VIEW", lambda s,a: _run_import(create_only=False, apply_to_picked=False, force_active=True)), 1, 4)

    close = _make_button("Close", lambda s,a: dlg.Close())
    tlp.Controls.Add(close, 0, 5); tlp.SetColumnSpan(close, 2)

    dlg.Controls.Add(tlp)
    dlg.MinimumSize = Drawing.Size(780, 0)
    dlg.ShowDialog()

# run
show_ui()
