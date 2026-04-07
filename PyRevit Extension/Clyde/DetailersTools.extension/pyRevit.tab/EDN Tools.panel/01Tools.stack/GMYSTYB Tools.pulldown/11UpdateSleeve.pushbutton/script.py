# -*- coding: utf-8 -*-
"""UpdateSleeve v3.4.0 -- Sleeve finder with auto-detected size params.

Author: Jeremiah Griffith
Version: 1.0.0
"""
from __future__ import division
from pyrevit import revit, DB, forms, script, UI
import clr
import os
import json
import datetime

clr.AddReference("System")
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Markup import XamlReader
from System.IO             import StreamReader
import System.Windows

doc   = revit.doc
uidoc = revit.uidoc

# ── Config ────────────────────────────────────────────────────────
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "sleeve_config.json")

DEFAULT_CONFIG = {
    "param_width":       "",
    "param_height":      "",
    "param_dia":         "",
    "oversize_rect_in":  2.0,
    "oversize_round_in": 2.0,
    "family_params":     {},   # keyed by family name -> {w, h, d}
    "rules": [
        {"match": "Fabrication Ductwork", "kw1": "sleeve", "kw2": "rect", "kw3": "rec",   "kw4": "",     "enabled": True},
        {"match": "Fabrication Ductwork", "kw1": "sleeve", "kw2": "round","kw3": "rd",    "kw4": "circ", "enabled": True},
        {"match": "Fabrication Pipework",  "kw1": "sleeve", "kw2": "pipe", "kw3": "",      "kw4": "",     "enabled": True},
    ],
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as fh:
                cfg = json.load(fh)
            for k, v in DEFAULT_CONFIG.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
        except Exception:
            pass
    import copy
    return copy.deepcopy(DEFAULT_CONFIG)

def save_config(cfg):
    try:
        with open(CONFIG_FILE, "w") as fh:
            json.dump(cfg, fh, indent=2)
    except Exception:
        pass

def get_family_params(config, family_name):
    """Return saved (w, h, d) for a specific family, falling back to globals."""
    fp = config.get("family_params", {})
    entry = fp.get(family_name, {})
    w = entry.get("w", config.get("param_width",  ""))
    h = entry.get("h", config.get("param_height", ""))
    d = entry.get("d", config.get("param_dia",    ""))
    return w, h, d

def set_family_params(config, family_name, w, h, d):
    """Save (w, h, d) for a specific family into config."""
    if "family_params" not in config:
        config["family_params"] = {}
    config["family_params"][family_name] = {"w": w, "h": h, "d": d}
    # Also update globals so Part+Sleeve fallback works
    config["param_width"]  = w
    config["param_height"] = h
    config["param_dia"]    = d

# ── Utilities ─────────────────────────────────────────────────────
def ft_to_ftin(v_ft):
    neg = v_ft < 0
    v   = abs(v_ft)
    ft  = int(v)
    inch = (v - ft) * 12
    s = "{0}'-{1:.1f}\"".format(ft, inch)
    return ("-" + s) if neg else s

def safe_float(text, default=0.0):
    try:
        return float(text)
    except Exception:
        return default

def _safe_family_name(elem):
    try:
        et = doc.GetElement(elem.GetTypeId())
        try:
            return et.FamilyName or "?"
        except Exception:
            pass
        try:
            p = et.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
            return p.AsString() if p else "?"
        except Exception:
            return "?"
    except Exception:
        try:
            return elem.Category.Name
        except Exception:
            return "?"

def _safe_type_name(elem):
    try:
        et = doc.GetElement(elem.GetTypeId())
        for bip in (DB.BuiltInParameter.SYMBOL_NAME_PARAM,
                    DB.BuiltInParameter.ALL_MODEL_TYPE_NAME):
            try:
                p = et.get_Parameter(bip)
                if p and p.HasValue:
                    v = p.AsString()
                    if v:
                        return v
            except Exception:
                pass
        return et.Name or "?"
    except Exception:
        return "?"

def _safe_category(elem):
    try:
        return elem.Category.Name
    except Exception:
        return "?"

# ── Source data extraction ────────────────────────────────────────
def extract_source_data(elem):
    d = {"Width": None, "Height": None, "Diameter": None,
         "shape": "unknown", "family": "--", "category": "--",
         "display_size": "--", "display_elev": "--"}
    try:
        d["category"] = elem.Category.Name
    except Exception:
        pass
    d["family"] = _safe_family_name(elem)

    def _lkd(name, bip=None):
        p = elem.LookupParameter(name)
        if p and p.HasValue and p.StorageType == DB.StorageType.Double:
            return p.AsDouble()
        if bip is not None:
            try:
                p2 = elem.get_Parameter(bip)
                if p2 and p2.HasValue and p2.StorageType == DB.StorageType.Double:
                    return p2.AsDouble()
            except Exception:
                pass
        return None

    d["Width"]    = _lkd("Main Primary Width",    DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    d["Height"]   = _lkd("Main Primary Depth",    DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    d["Diameter"] = _lkd("Main Primary Diameter", DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)

    if d["Width"] is not None and d["Height"] is not None:
        d["shape"] = "rect"
        d["display_size"] = "{0}x{1}\"".format(
            int(round(d["Width"]*12)), int(round(d["Height"]*12)))
    elif d["Diameter"] is not None:
        d["shape"] = "round"
        d["display_size"] = "{0}\" dia".format(int(round(d["Diameter"]*12)))

    # Build a descriptive shape label that names both shape and content type
    cat = d.get("category", "")
    shape = d.get("shape", "unknown")
    if "Ductwork" in cat or "Duct" in cat:
        if shape == "rect":
            d["shape_label"] = "Rect Duct"
        elif shape == "round":
            d["shape_label"] = "Round Duct"
        else:
            d["shape_label"] = "Duct"
    elif "Pipework" in cat or "Pipe" in cat:
        d["shape_label"] = "Round Pipe"
    else:
        d["shape_label"] = shape.capitalize() if shape != "unknown" else cat

    # Also grab product entry name for ductwork (Straight, Elbow, etc.)
    try:
        pp = elem.LookupParameter("Product Entry")
        if pp and pp.HasValue:
            d["product"] = pp.AsString() or ""
        else:
            d["product"] = ""
    except Exception:
        d["product"] = ""

    p_lo = elem.LookupParameter("Lower End Bottom Elevation")
    p_hi = elem.LookupParameter("Upper End Bottom Elevation")
    bot = None
    if p_lo and p_lo.HasValue:
        lo = p_lo.AsDouble()
        hi = p_hi.AsDouble() if (p_hi and p_hi.HasValue) else lo
        bot = (lo + hi) / 2.0
    else:
        try:
            loc = elem.Location
            if hasattr(loc, "Curve"):
                c = loc.Curve
                mid_z = (c.GetEndPoint(0).Z + c.GetEndPoint(1).Z) / 2.0
                half = (d["Height"]/2.0 if d["Height"] else
                        (d["Diameter"]/2.0 if d["Diameter"] else 0.0))
                bot = mid_z - half
        except Exception:
            pass

    if bot is not None:
        half = (d["Height"]/2.0 if d["Height"] else (d["Diameter"]/2.0 if d["Diameter"] else 0.0))
        full = d["Height"] or d["Diameter"] or 0.0
        d["BottomElev"] = bot
        d["CenterElev"] = bot + half
        d["TopElev"]    = bot + full
        d["display_elev"] = "Bot {0}".format(ft_to_ftin(bot))

    return d

# ── Sleeve scanner ────────────────────────────────────────────────
def _sleeve_elevation(elem):
    try:
        loc = elem.Location
        if hasattr(loc, "Point"):
            return ft_to_ftin(loc.Point.Z)
        if hasattr(loc, "Curve"):
            c = loc.Curve
            return ft_to_ftin((c.GetEndPoint(0).Z + c.GetEndPoint(1).Z) / 2.0)
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            return ft_to_ftin((bb.Min.Z + bb.Max.Z) / 2.0)
    except Exception:
        pass
    return "--"

def _sleeve_size(elem, p_width="", p_height="", p_dia=""):
    """Read size from element. Uses configured param names when provided,
    falls back to keyword scan if not."""
    def _read(name):
        if not name or name == CONFIG_NONE:
            return None
        try:
            p = elem.LookupParameter(name)
            if p is None:
                # case-insensitive fallback
                nl = name.lower()
                for cp in elem.Parameters:
                    try:
                        if cp.Definition.Name.lower() == nl and cp.HasValue:
                            p = cp
                            break
                    except Exception:
                        pass
            if p and p.HasValue and p.StorageType == DB.StorageType.Double:
                v = p.AsDouble()
                if v > 0:
                    return v
        except Exception:
            pass
        return None

    # Try configured params first
    width_val  = _read(p_width)
    height_val = _read(p_height)
    dia_val    = _read(p_dia)

    # Fall back to keyword scan for any that are missing
    if width_val is None and height_val is None and dia_val is None:
        try:
            for p in elem.Parameters:
                try:
                    if not p.HasValue:
                        continue
                    if p.StorageType != DB.StorageType.Double:
                        continue
                    v = p.AsDouble()
                    if v <= 0:
                        continue
                    nl = p.Definition.Name.lower()
                    if "width" in nl and width_val is None:
                        width_val = v
                    elif ("height" in nl or "depth" in nl) and height_val is None:
                        height_val = v
                    elif "diameter" in nl and dia_val is None:
                        dia_val = v
                except Exception:
                    pass
        except Exception:
            pass

    if width_val is not None and height_val is not None:
        return "{0:.0f}x{1:.0f}\"".format(width_val*12, height_val*12)
    if dia_val is not None:
        return "{0:.0f}\" dia".format(dia_val*12)
    if width_val is not None:
        return "{0:.0f}\"".format(width_val*12)
    try:
        bb = elem.get_BoundingBox(None)
        if bb:
            dims = sorted([abs(bb.Max.X-bb.Min.X)*12,
                           abs(bb.Max.Y-bb.Min.Y)*12], reverse=True)
            return "{0:.0f}x{1:.0f}\"".format(dims[0], dims[1])
    except Exception:
        pass
    return "--"

def _kw_matches(family_name, kw1, kw2, kw3, kw4):
    """kw1 AND (kw2 OR kw3 OR kw4). If no OR keywords set, kw1 alone is enough."""
    name = family_name.lower()
    kw1 = (kw1 or "").lower().strip()
    if not kw1 or kw1 not in name:
        return False
    ors = [k.lower().strip() for k in (kw2, kw3, kw4) if k and k.strip()]
    if ors:
        return any(k in name for k in ors)
    return True

# Shape-hint keywords -- if a rule's OR keywords contain any of these,
# it is considered a rect-only or round-only rule for disambiguation.
_RECT_HINTS  = {"rect", "rec", "rectangular", "square"}
_ROUND_HINTS = {"round", "rd", "rnd", "circ", "circular", "pipe"}

def _rule_shape_hint(rule):
    """Return 'rect', 'round', or None based on OR keywords in the rule."""
    ors = {k.lower().strip() for k in
           (rule.get("kw2",""), rule.get("kw3",""), rule.get("kw4",""))
           if k and k.strip()}
    if ors & _RECT_HINTS:
        return "rect"
    if ors & _ROUND_HINTS:
        return "round"
    return None

def get_active_rule(config, source_elem, source_data=None):
    """Return the first matching rule dict, or a shape-aware fallback.

    Three-pass matching:
      1. Category/family match + shape-consistent OR keywords (best match)
      2. Category/family match with no shape hint (generic category rule)
      3. Shape-only fallback -- no category match but OR keywords agree with
         source shape. Handles the common case where Match text is wrong but
         OR keywords correctly describe rect vs round.
    """
    rules = config.get("rules", [])
    shape = (source_data or {}).get("shape", "")

    if source_elem is not None:
        fam = _safe_family_name(source_elem).lower()
        cat = _safe_category(source_elem).lower()

        # Pass 1: category match + shape-consistent keywords
        for rule in rules:
            if not rule.get("enabled", True):
                continue
            match = rule.get("match", "").lower().strip()
            if not match or not (match in fam or match in cat):
                continue
            hint = _rule_shape_hint(rule)
            if hint is None:
                continue
            if shape and hint != shape:
                continue
            return rule

        # Pass 2: category match, no shape hint
        for rule in rules:
            if not rule.get("enabled", True):
                continue
            match = rule.get("match", "").lower().strip()
            if match and (match in fam or match in cat):
                return rule

    # Pass 3: shape-only fallback -- ignore Match text, pick rule whose
    # OR keywords agree with source shape. Handles misconfigured Match text.
    if shape:
        for rule in rules:
            if not rule.get("enabled", True):
                continue
            hint = _rule_shape_hint(rule)
            if hint == shape:
                return rule

    # Hard fallback -- sleeve keyword only
    return {"kw1": "sleeve", "kw2": "", "kw3": "", "kw4": ""}

def rule_label(rule):
    kw1 = (rule.get("kw1", "sleeve") or "sleeve").strip()
    ors = [k.strip() for k in (rule.get("kw2",""), rule.get("kw3",""), rule.get("kw4","")) if k and k.strip()]
    if ors:
        return "{0} AND ({1})".format(kw1, " OR ".join(ors))
    return kw1

def scan_sleeves(log_fn, rule=None, source_shape=None):
    """Scan for sleeve FamilyInstances matching rule keywords.

    source_shape: 'rect' or 'round'. When provided and the rule has no
    shape-specific OR keywords, a family-name shape filter is applied as
    a safety net so rect content never picks up round sleeves and vice versa.
    """
    if rule is None:
        rule = {"kw1": "sleeve", "kw2": "", "kw3": "", "kw4": ""}
    kw1 = rule.get("kw1", "sleeve") or "sleeve"
    kw2 = rule.get("kw2", "")
    kw3 = rule.get("kw3", "")
    kw4 = rule.get("kw4", "")
    label = rule_label(rule)

    # Determine if rule already targets a specific shape
    rule_hint = _rule_shape_hint(rule)
    # Apply safety shape filter when source shape is known and rule has no hint
    # (or rule hint matches -- never filter against a matching hint)
    apply_shape_filter = (source_shape in ("rect", "round") and
                          rule_hint is None)

    results = []
    log_fn("{0} --- Scanning sleeves  rule: {1}  shape: {2} ---".format(
        _ts(), label, source_shape or "any"))
    try:
        collector = (DB.FilteredElementCollector(doc)
                     .OfClass(DB.FamilyInstance)
                     .WhereElementIsNotElementType())
        count = 0
        for elem in collector:
            count += 1
            try:
                fam_name = _safe_family_name(elem)
                if not _kw_matches(fam_name, kw1, kw2, kw3, kw4):
                    continue
                # Safety shape filter: when rule has no shape hint, check
                # family name for shape keywords to avoid cross-contamination
                if apply_shape_filter:
                    fam_lower = fam_name.lower()
                    if source_shape == "rect":
                        # Reject family names that strongly suggest round
                        round_kws = {"round", "rd_", "_rd", "rnd", "circ",
                                     "pipe", "round"}
                        if any(k in fam_lower for k in round_kws):
                            continue
                    elif source_shape == "round":
                        # Reject family names that strongly suggest rect
                        rect_kws = {"rect", "rec_", "_rec", "rectangular",
                                    "square", "duct_wall"}
                        if any(k in fam_lower for k in rect_kws):
                            continue
                results.append(elem)
            except Exception:
                pass
        log_fn("{0}   Checked {1} instances, found {2} sleeves".format(
            _ts(), count, len(results)))
    except Exception as ex:
        log_fn("{0}   Scan ERROR: {1}".format(_ts(), str(ex)))
    return results

# ── Param writer ──────────────────────────────────────────────────
def write_param(elem, param_name, value_ft):
    # Exact match first
    p = elem.LookupParameter(param_name)
    # Case-insensitive fallback
    if p is None:
        name_lower = param_name.lower()
        try:
            for candidate in elem.Parameters:
                try:
                    if candidate.Definition.Name.lower() == name_lower:
                        p = candidate
                        break
                except Exception:
                    pass
        except Exception:
            pass
    if p is None:
        return False, "not found: '{0}'".format(param_name)
    if p.IsReadOnly:
        return False, "read-only"
    st = p.StorageType
    try:
        if st == DB.StorageType.Double:
            p.Set(value_ft)
        elif st == DB.StorageType.Integer:
            p.Set(int(round(value_ft * 12)))
        elif st == DB.StorageType.String:
            p.Set("{0}\"".format(int(round(value_ft * 12))))
        else:
            return False, "unsupported type"
        return True, "ok ({0})".format(p.Definition.Name)
    except Exception as ex:
        return False, str(ex)

# ── Param detector ────────────────────────────────────────────────
_WIDTH_KEYS  = ("width",)
_HEIGHT_KEYS = ("height", "depth")
_DIA_KEYS    = ("diameter", "dia")
NONE_LABEL   = "-- none --"
CONFIG_NONE  = "__none__"   # sentinel: user explicitly chose none in config

def detect_size_params(elem, must_include=None):
    """Scan writable Double params in plausible size range.
    must_include: list of param names to always include regardless of value
    (used so saved config names appear even when their current value is 0).
    """
    candidates = []
    must_include = [n.lower() for n in (must_include or [])]
    auto_w = auto_h = auto_d = ""
    try:
        for p in elem.Parameters:
            try:
                if p.IsReadOnly:
                    continue
                if p.StorageType != DB.StorageType.Double:
                    continue
                if not p.HasValue:
                    continue
                name = p.Definition.Name
                if not name:
                    continue
                v = p.AsDouble()
                # Include if in plausible range OR if it's a must-include name
                if 0.04 <= v <= 10.0 or name.lower() in must_include:
                    candidates.append(name)
            except Exception:
                pass
    except Exception:
        pass
    candidates.sort()
    for name in candidates:
        nl = name.lower()
        if not auto_w and any(k in nl for k in _WIDTH_KEYS):
            auto_w = name
        if not auto_h and any(k in nl for k in _HEIGHT_KEYS):
            auto_h = name
        if not auto_d and any(k in nl for k in _DIA_KEYS):
            auto_d = name
    return candidates, auto_w, auto_h, auto_d

def get_oversize_for_shape(config, shape):
    """Return oversize in feet for the given shape (rect or round)."""
    if shape == "rect":
        return config.get("oversize_rect_in", 2.0) / 12.0
    return config.get("oversize_round_in", 2.0) / 12.0

# ── Pick filter ───────────────────────────────────────────────────
def _bic(name):
    try:
        return getattr(DB.BuiltInCategory, name)
    except AttributeError:
        return None

CONTENT_CAT_IDS = set()
for _n in ["OST_FabricationDuctwork", "OST_FabricationPipework",
           "OST_DuctCurves", "OST_PipeCurves",
           "OST_FlexDuctCurves", "OST_FlexPipeCurves"]:
    _r = _bic(_n)
    if _r is not None:
        CONTENT_CAT_IDS.add(int(_r))

class SourceFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return elem.Category.Id.IntegerValue in CONTENT_CAT_IDS
        except Exception:
            return False
    def AllowReference(self, reference, position):
        return False

# ── Row model ─────────────────────────────────────────────────────
class SleeveRow(object):
    def __init__(self, elem_id, family, type_name, category, size, elev):
        self.ElemId      = elem_id
        self.Family      = family
        self.TypeName    = type_name
        self.Category    = category
        self.Size        = size
        self.Elev        = elev
        self.Selected    = True
        # Per-sleeve saved param choices (None = not yet set by user)
        self.SavedWidth  = None
        self.SavedHeight = None
        self.SavedDia    = None

# ── Rule model ───────────────────────────────────────────────────
class RuleRow(object):
    """One content-to-sleeve mapping rule.
    Kw1 = AND (must be present). Kw2/3/4 = OR (any one must match if set).
    Logic: Kw1 AND (Kw2 OR Kw3 OR Kw4)
    """
    def __init__(self, match_text, kw1, kw2="", kw3="", kw4="", enabled=True):
        self.MatchText = match_text
        self.Kw1       = kw1
        self.Kw2       = kw2
        self.Kw3       = kw3
        self.Kw4       = kw4
        self.Enabled   = enabled
# ── Crash log ─────────────────────────────────────────────────────
LOG_FILE = os.path.join(os.path.dirname(__file__), "debug_log.txt")

def _ts():
    return datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]

class CrashLog(object):
    def __init__(self, path):
        self._path = path
        try:
            with open(path, "a") as fh:
                fh.write("\n" + "="*70 + "\n")
                fh.write("UpdateSleeve  session  {0}\n".format(
                    datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")))
                fh.write("="*70 + "\n")
                fh.flush()
        except Exception:
            pass

    def log(self, text):
        try:
            with open(self._path, "a") as fh:
                fh.write(text + "\n")
                fh.flush()
        except Exception:
            pass

CRASH_LOG = CrashLog(LOG_FILE)

# ── Window ────────────────────────────────────────────────────────
class UpdateSleeveWindow(object):

    def __init__(self, xaml_file, config, source_elem, source_data):
        reader = StreamReader(xaml_file)
        self._window = XamlReader.Load(reader.BaseStream)
        reader.Close()
        self._config      = config
        self._source_elem = source_elem
        self._source_data = source_data
        self._sleeve_rows = ObservableCollection[object]()
        self._rule_rows   = ObservableCollection[object]()
        self._active_rule = get_active_rule(config, source_elem)
        self._last_row         = None
        self._suppress_sel_change  = False
        self._suppress_combo_change = False

        self._get_controls()
        self._wire_events()
        self._populate_source_card()
        self._build_rule_rows()
        self.SleeveGrid.ItemsSource = self._sleeve_rows
        self.RulesGrid.ItemsSource  = self._rule_rows
        self._run_scan()
        self._update_preview()

    def _get_controls(self):
        f = self._window.FindName
        self.DiaCombo             = f("DiaCombo")
        self.HeightCombo          = f("HeightCombo")
        self.KeywordLabel         = f("KeywordLabel")
        self.OversizeBox          = f("OversizeBox")
        self.ParamCount           = f("ParamCount")
        self.PreviewLabel         = f("PreviewLabel")
        self.RulesGrid            = f("RulesGrid")
        self.SleeveGrid           = f("SleeveGrid")
        self.SrcCat               = f("SrcCat")
        self.SrcElev              = f("SrcElev")
        self.SrcFamily            = f("SrcFamily")
        self.SrcShape             = f("SrcShape")
        self.SrcSize              = f("SrcSize")
        self.StatusBar            = f("StatusBar")
        self.WidthCombo           = f("WidthCombo")

    def _wire_events(self):
        self.WidthCombo.SelectionChanged  += lambda s, e: self.WidthCombo_SelectionChanged(s, e)
        self.HeightCombo.SelectionChanged += lambda s, e: self.HeightCombo_SelectionChanged(s, e)
        self.DiaCombo.SelectionChanged    += lambda s, e: self.DiaCombo_SelectionChanged(s, e)
        self.OversizeBox.TextChanged      += lambda s, e: self.OversizeBox_TextChanged(s, e)
        self.SleeveGrid.SelectionChanged  += lambda s, e: self.SleeveGrid_SelectionChanged(s, e)
        f = self._window.FindName
        f("BtnApply").Click      += lambda s, e: self.Apply_Click(s, e)
        f("BtnExport").Click     += lambda s, e: self.Export_Click(s, e)
        f("BtnImport").Click     += lambda s, e: self.Import_Click(s, e)
        f("BtnAddRule").Click    += lambda s, e: self.AddRule_Click(s, e)
        f("BtnRemoveRule").Click += lambda s, e: self.RemoveRule_Click(s, e)
        f("BtnSaveRules").Click  += lambda s, e: self.SaveRules_Click(s, e)

    def ShowDialog(self):
        self._window.ShowDialog()

    def Close(self):
        self._window.Close()

    def _set_status(self, msg):
        self.StatusBar.Text = msg

    def _update_preview(self):
        """Show source dims + oversize so user sees what will be written."""
        try:
            os_in  = safe_float(self.OversizeBox.Text, 0.0)
            os_ft  = os_in / 12.0
            src    = self._source_data
            w  = src.get("Width")
            h  = src.get("Height")
            d  = src.get("Diameter")
            if w is not None and h is not None:
                self.PreviewLabel.Text = (
                    "Will write: {0:.0f}x{1:.0f}in"
                    "  (source {2:.0f}x{3:.0f}in + {4:.1f}in OS)".format(
                        (w + os_ft)*12, (h + os_ft)*12,
                        w*12, h*12, os_in))
            elif d is not None:
                self.PreviewLabel.Text = (
                    "Will write: {0:.0f}in dia"
                    "  (source {1:.0f}in + {2:.1f}in OS)".format(
                        (d + os_ft)*12, d*12, os_in))
            else:
                self.PreviewLabel.Text = "Source has no size data."
        except Exception:
            self.PreviewLabel.Text = ""

    def OversizeBox_TextChanged(self, sender, e):
        self._update_preview()

    def _save_current_row_params(self):
        """Immediately persist combo state for the currently selected row's family,
        and refresh the Size column so it reflects the newly configured params."""
        row = getattr(self, "_last_row", None)
        if row is None:
            row = self.SleeveGrid.SelectedItem
        if row is None:
            return
        w = self._selected_name(self.WidthCombo)
        h = self._selected_name(self.HeightCombo)
        d = self._selected_name(self.DiaCombo)
        row.SavedWidth  = w
        row.SavedHeight = h
        row.SavedDia    = d
        set_family_params(self._config, row.Family,
            self._selected_name_for_save(self.WidthCombo),
            self._selected_name_for_save(self.HeightCombo),
            self._selected_name_for_save(self.DiaCombo))
        save_config(self._config)
        # Refresh size display for this row using the new param names
        try:
            elem = doc.GetElement(DB.ElementId(row.ElemId))
            if elem is not None:
                row.Size = _sleeve_size(elem, w, h, d)
                self._refresh_grid()
        except Exception:
            pass

    def WidthCombo_SelectionChanged(self, sender, e):
        if getattr(self, '_suppress_combo_change', False):
            return
        self._save_current_row_params()
        self._update_preview()

    def HeightCombo_SelectionChanged(self, sender, e):
        if getattr(self, '_suppress_combo_change', False):
            return
        self._save_current_row_params()
        self._update_preview()

    def DiaCombo_SelectionChanged(self, sender, e):
        if getattr(self, '_suppress_combo_change', False):
            return
        self._save_current_row_params()
        self._update_preview()

    def _build_rule_rows(self):
        self._rule_rows.Clear()
        for rule in self._config.get("rules", []):
            self._rule_rows.Add(RuleRow(
                rule.get("match",   ""),
                rule.get("kw1",     "sleeve"),
                rule.get("kw2",     ""),
                rule.get("kw3",     ""),
                rule.get("kw4",     ""),
                rule.get("enabled", True),
            ))

    def AddRule_Click(self, sender, e):
        self._rule_rows.Add(RuleRow("", "sleeve", "", "", "", True))

    def RemoveRule_Click(self, sender, e):
        sel = self.RulesGrid.SelectedItem
        if sel is not None:
            self._rule_rows.Remove(sel)

    def SaveRules_Click(self, sender, e):
        self.RulesGrid.CommitEdit()
        rules = []
        for row in self._rule_rows:
            match = (row.MatchText or "").strip()
            kw1   = (row.Kw1 or "sleeve").strip()
            if match:
                rules.append({
                    "match":   match,
                    "kw1":     kw1,
                    "kw2":     (row.Kw2 or "").strip(),
                    "kw3":     (row.Kw3 or "").strip(),
                    "kw4":     (row.Kw4 or "").strip(),
                    "enabled": bool(row.Enabled),
                })
        self._config["rules"] = rules
        save_config(self._config)
        self._run_scan()
        self._set_status("Rules saved. Active rule: {0}".format(rule_label(self._active_rule)))

    def Export_Click(self, sender, e):
        dlg = System.Windows.Forms.SaveFileDialog()
        dlg.Title      = "Export UpdateSleeve Settings"
        dlg.Filter     = "JSON files (*.json)|*.json"
        dlg.FileName   = "sleeve_config.json"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            try:
                self._config.update({
                    "param_width":  self._selected_name_for_save(self.WidthCombo),
                    "param_height": self._selected_name_for_save(self.HeightCombo),
                    "param_dia":    self._selected_name_for_save(self.DiaCombo),
                    "{0}".format("oversize_rect_in" if self._source_data.get("shape","rect")=="rect" else "oversize_round_in"):
                    safe_float(self.OversizeBox.Text, 2.0),
                })
                with open(dlg.FileName, "w") as fh:
                    json.dump(self._config, fh, indent=2)
                self._set_status("Settings exported to {0}".format(dlg.FileName))
            except Exception as ex:
                self._set_status("Export failed: {0}".format(str(ex)))

    def Import_Click(self, sender, e):
        dlg = System.Windows.Forms.OpenFileDialog()
        dlg.Title  = "Import UpdateSleeve Settings"
        dlg.Filter = "JSON files (*.json)|*.json"
        if dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK:
            try:
                with open(dlg.FileName, "r") as fh:
                    cfg = json.load(fh)
                for k, v in DEFAULT_CONFIG.items():
                    if k not in cfg:
                        cfg[k] = v
                self._config = cfg
                save_config(cfg)
                # Repopulate combos with imported selections
                if self.SleeveGrid.Items.Count > 0:
                    row = self.SleeveGrid.SelectedItem or self._sleeve_rows[0]
                    elem = doc.GetElement(DB.ElementId(row.ElemId))
                    if elem:
                        self._repopulate_combos_for_elem(elem)
                shape = self._source_data.get("shape", "rect")
                os_key = "oversize_rect_in" if shape == "rect" else "oversize_round_in"
                self.OversizeBox.Text = str(cfg.get(os_key, 2.0))
                self._update_preview()
                self._set_status("Settings imported from {0}".format(dlg.FileName))
            except Exception as ex:
                self._set_status("Import failed: {0}".format(str(ex)))

    def _populate_source_card(self):
        d = self._source_data
        # Show product name + family for fabrication parts (e.g. "Straight  Pipe - B88...")
        product = d.get("product", "")
        family  = d.get("family", "--")
        self.SrcFamily.Text = "{0}  {1}".format(product, family).strip() if product else family
        self.SrcShape.Text  = d.get("shape_label", d["shape"].capitalize())
        self.SrcSize.Text   = d["display_size"]
        self.SrcElev.Text   = d["display_elev"]
        self.SrcCat.Text    = d["category"]

    # ── Scan ──────────────────────────────────────────────────────
    def _run_scan(self):
        self._active_rule = get_active_rule(self._config, self._source_elem, self._source_data)
        self.KeywordLabel.Text = "Active rule: {0}".format(
            rule_label(self._active_rule))
        self._sleeve_rows.Clear()
        for elem in scan_sleeves(CRASH_LOG.log, self._active_rule,
                                  self._source_data.get('shape')):
            try:
                fam  = _safe_family_name(elem)
                fw, fh, fd = get_family_params(self._config, fam)
                self._sleeve_rows.Add(SleeveRow(
                    elem.Id.IntegerValue,
                    fam,
                    _safe_type_name(elem),
                    _safe_category(elem),
                    _sleeve_size(elem, fw, fh, fd),
                    _sleeve_elevation(elem),
                ))
            except Exception:
                pass
        n = len(self._sleeve_rows)
        msg = "{0} sleeve(s) found.".format(n) if n else \
              "No sleeves found -- family name must contain 'sleeve'."
        self._set_status(msg)
        # Set oversize box
        shape  = self._source_data.get("shape", "rect")
        os_key = "oversize_rect_in" if shape == "rect" else "oversize_round_in"
        self.OversizeBox.Text = str(self._config.get(os_key, 2.0))
        # Select first row -- fires SelectionChanged which populates combos
        if self.SleeveGrid.Items.Count > 0:
            self.SleeveGrid.SelectedIndex = 0

    def SleeveGrid_SelectionChanged(self, sender, e):
        if getattr(self, "_suppress_sel_change", False):
            return

        # Save combos to the row we are leaving
        prev = getattr(self, "_last_row", None)
        if prev is not None:
            w = self._selected_name(self.WidthCombo)
            h = self._selected_name(self.HeightCombo)
            d = self._selected_name(self.DiaCombo)
            prev.SavedWidth  = w
            prev.SavedHeight = h
            prev.SavedDia    = d
            # Also persist per-family so next session remembers
            set_family_params(self._config, prev.Family, 
                self._selected_name_for_save(self.WidthCombo),
                self._selected_name_for_save(self.HeightCombo),
                self._selected_name_for_save(self.DiaCombo))
            save_config(self._config)

        row = self.SleeveGrid.SelectedItem
        self._last_row = row
        if row is None:
            return
        elem = doc.GetElement(DB.ElementId(row.ElemId))
        if elem is None:
            return
        self._repopulate_combos_for_elem(elem, row)

    def _repopulate_combos_for_elem(self, elem, row=None):
        # Resolve per-family params first, then per-row overrides
        fam_name = row.Family if row else ""
        cfg_w, cfg_h, cfg_d = get_family_params(self._config, fam_name)
        row_w = row.SavedWidth  if row and row.SavedWidth  is not None else None
        row_h = row.SavedHeight if row and row.SavedHeight is not None else None
        row_d = row.SavedDia    if row and row.SavedDia    is not None else None

        # Always include saved names so they appear even when value=0
        must = [n for n in (row_w, row_h, row_d, cfg_w, cfg_h, cfg_d)
                if n and n != CONFIG_NONE]
        candidates, auto_w, auto_h, auto_d = detect_size_params(elem, must)

        self._suppress_combo_change = True
        for combo in (self.WidthCombo, self.HeightCombo, self.DiaCombo):
            combo.Items.Clear()
            combo.Items.Add(NONE_LABEL)
        for name in candidates:
            self.WidthCombo.Items.Add(name)
            self.HeightCombo.Items.Add(name)
            self.DiaCombo.Items.Add(name)

        self.ParamCount.Text = "{0} writable size params on this sleeve.".format(
            len(candidates))

        def _set_combo(combo, target):
            """Reliable combo selection -- iterates Items directly to avoid
            IronPython 2.7 type-coercion issues with list.index() on .NET strings."""
            for i in range(combo.Items.Count):
                if str(combo.Items[i]) == str(target):
                    combo.SelectedIndex = i
                    return
            combo.SelectedIndex = 0

        def _pick(combo, per_row, cfg, auto):
            # per_row == None    never touched this session -> use cfg/auto
            # per_row == ""      explicitly set to none this session
            # cfg == CONFIG_NONE saved explicit none
            # cfg == ""          never configured -> use auto-detect
            if per_row is not None:
                if per_row == "":
                    combo.SelectedIndex = 0
                else:
                    _set_combo(combo, per_row)
                return
            if cfg == CONFIG_NONE:
                combo.SelectedIndex = 0
                return
            if cfg and cfg in candidates:
                _set_combo(combo, cfg)
                return
            if auto and auto in candidates:
                _set_combo(combo, auto)
                return
            combo.SelectedIndex = 0

        _pick(self.WidthCombo,  row_w, cfg_w, auto_w)
        _pick(self.HeightCombo, row_h, cfg_h, auto_h)
        _pick(self.DiaCombo,    row_d, cfg_d, auto_d)
        self._suppress_combo_change = False

    def Rescan_Click(self, sender, e):
        self._run_scan()

    # ── Param detection ───────────────────────────────────────────
    def _selected_name(self, combo):
        item = combo.SelectedItem
        if item is None or str(item) == NONE_LABEL:
            return ""
        return str(item).strip()

    def _selected_name_for_save(self, combo):
        """Returns CONFIG_NONE sentinel when none selected, so config
        can distinguish explicit-none from never-configured."""
        item = combo.SelectedItem
        if item is None or str(item) == NONE_LABEL:
            return CONFIG_NONE
        return str(item).strip()

    # ── Grid controls ─────────────────────────────────────────────
    def _refresh_grid(self):
        self._suppress_sel_change = True
        sel = self.SleeveGrid.SelectedItem
        tmp = self.SleeveGrid.ItemsSource
        self.SleeveGrid.ItemsSource = None
        self.SleeveGrid.ItemsSource = tmp
        if sel is not None:
            self.SleeveGrid.SelectedItem = sel
        self._suppress_sel_change  = False
        self._suppress_combo_change = False

    def _sel(self):
        return self.SleeveGrid.SelectedItem

    # ── Apply ─────────────────────────────────────────────────────
    def Apply_Click(self, sender, e):
        # Save current combo state to the selected row before writing
        cur = self.SleeveGrid.SelectedItem
        if cur is not None:
            cur.SavedWidth  = self._selected_name(self.WidthCombo)
            cur.SavedHeight = self._selected_name(self.HeightCombo)
            cur.SavedDia    = self._selected_name(self.DiaCombo)
            self._last_row  = cur

        targets = [r for r in self._sleeve_rows if r.Selected]
        if not targets:
            self._set_status("Nothing selected.")
            return

        # Combos reflect the currently SELECTED row.
        # For other rows use their per-row saved choices, falling back to
        # config defaults if the user never clicked on them.
        cfg_w    = self._config.get("param_width",  "")
        cfg_h    = self._config.get("param_height", "")
        cfg_d    = self._config.get("param_dia",    "")
        oversize = safe_float(self.OversizeBox.Text, 2.0) / 12.0
        src      = self._source_data

        # Get source centerpoint for move
        src_pt = None
        try:
            loc = self._source_elem.Location
            if hasattr(loc, "Curve"):
                c = loc.Curve
                p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                src_pt = DB.XYZ((p0.X+p1.X)/2, (p0.Y+p1.Y)/2, (p0.Z+p1.Z)/2)
            elif hasattr(loc, "Point"):
                src_pt = loc.Point
        except Exception:
            pass

        n_ok = n_fail = 0
        lines = []

        t = DB.Transaction(doc, "UpdateSleeve - Apply")
        t.Start()
        try:
            for row in targets:
                elem = doc.GetElement(DB.ElementId(row.ElemId))
                if elem is None:
                    lines.append("SKIP  [{0}] not found".format(row.ElemId))
                    n_fail += 1
                    continue

                # Use per-row saved param choices; fall back to config defaults
                p_width  = row.SavedWidth  if row.SavedWidth  is not None else cfg_w
                p_height = row.SavedHeight if row.SavedHeight is not None else cfg_h
                p_dia    = row.SavedDia    if row.SavedDia    is not None else cfg_d

                # Move perpendicular to duct run (Z + cross-axis, not along duct)
                if src_pt is not None:
                    try:
                        cur_loc = elem.Location
                        if hasattr(cur_loc, "Point"):
                            cur_pt = cur_loc.Point
                        else:
                            c = cur_loc.Curve
                            p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                            cur_pt = DB.XYZ((p0.X+p1.X)/2, (p0.Y+p1.Y)/2, (p0.Z+p1.Z)/2)

                        full_delta = DB.XYZ(
                            src_pt.X - cur_pt.X,
                            src_pt.Y - cur_pt.Y,
                            src_pt.Z - cur_pt.Z)

                        # Subtract component along duct axis
                        try:
                            dc = self._source_elem.Location.Curve
                            dp0 = dc.GetEndPoint(0)
                            dp1 = dc.GetEndPoint(1)
                            raw = DB.XYZ(dp1.X-dp0.X, dp1.Y-dp0.Y, dp1.Z-dp0.Z)
                            length = raw.GetLength()
                            if length > 0.001:
                                ux = raw.X / length
                                uy = raw.Y / length
                                uz = raw.Z / length
                                dot = (full_delta.X*ux +
                                       full_delta.Y*uy +
                                       full_delta.Z*uz)
                                move = DB.XYZ(
                                    full_delta.X - dot*ux,
                                    full_delta.Y - dot*uy,
                                    full_delta.Z - dot*uz)
                            else:
                                move = full_delta
                        except Exception:
                            move = full_delta

                        DB.ElementTransformUtils.MoveElement(doc, elem.Id, move)
                        lines.append("MOVED [{0}]  {1}".format(row.ElemId, row.Family))
                    except Exception as ex:
                        lines.append("MOVE FAIL [{0}]: {1}".format(row.ElemId, str(ex)))

                # Size writes
                for key, pname in (("Width", p_width), ("Height", p_height),
                                   ("Diameter", p_dia)):
                    if not pname or src.get(key) is None:
                        continue
                    val = src[key] + oversize
                    ok, msg = write_param(elem, pname, val)
                    lines.append("  {0}  {1} -> {2} = {3:.2f}\"".format(
                        "OK  " if ok else "FAIL",
                        key, pname, val*12 if ok else 0.0))
                    if not ok:
                        lines[-1] += "  ({0})".format(msg)
                        n_fail += 1
                    else:
                        n_ok += 1

            t.Commit()

            # Refresh Size column using configured param names
            for row in targets:
                try:
                    elem = doc.GetElement(DB.ElementId(row.ElemId))
                    if elem is not None:
                        fw, fh, fd = get_family_params(self._config, row.Family)
                        row.Size = _sleeve_size(elem, fw, fh, fd)
                except Exception:
                    pass
            self._refresh_grid()

        except Exception as ex:
            t.RollBack()
            self._set_status("ERROR: {0}".format(str(ex)))
            self._set_status("Error -- rolled back.")
            return

        # Save per-family param choices for every target row
        shape  = self._source_data.get("shape", "rect")
        os_key = "oversize_rect_in" if shape == "rect" else "oversize_round_in"
        cur_row = self.SleeveGrid.SelectedItem
        cur_w = self._selected_name_for_save(self.WidthCombo)
        cur_h = self._selected_name_for_save(self.HeightCombo)
        cur_d = self._selected_name_for_save(self.DiaCombo)
        # Save current row's family params
        if cur_row:
            set_family_params(self._config, cur_row.Family, cur_w, cur_h, cur_d)
        # Save per-row params for all other targets using their SavedWidth/H/D
        for row in targets:
            if cur_row and row.ElemId == cur_row.ElemId:
                continue
            rw = row.SavedWidth  if row.SavedWidth  is not None else cur_w
            rh = row.SavedHeight if row.SavedHeight is not None else cur_h
            rd = row.SavedDia    if row.SavedDia    is not None else cur_d
            set_family_params(self._config, row.Family,
                rw if rw != "" else cur_w,
                rh if rh != "" else cur_h,
                rd if rd != "" else cur_d)
        self._config[os_key] = safe_float(self.OversizeBox.Text, 2.0)
        save_config(self._config)

        summary = "Moved: {0}  Writes OK: {1}  Failed: {2}".format(
            len(targets), n_ok, n_fail)
        self._set_status(summary + "\n\n" + "\n".join(lines))
        self._set_status(summary)
        CRASH_LOG.log("{0} Apply -- {1}".format(_ts(), summary))
        self.Close()

# ── Launch dialog ─────────────────────────────────────────────────
class LaunchWindow(object):
    ACTION_CANCEL      = "cancel"
    ACTION_VIEWPORT    = "viewport"
    ACTION_PART_SLEEVE = "part_sleeve"

    def __init__(self, xaml_file):
        reader = StreamReader(xaml_file)
        self._window = XamlReader.Load(reader.BaseStream)
        reader.Close()
        self.action = self.ACTION_CANCEL
        # Wire all events in Python -- no event attrs in XAML
        f = self._window.FindName
        f("BtnCancel").Click       += lambda s, e: self._on_cancel(s, e)
        f("BtnPickViewport").Click  += lambda s, e: self._on_viewport(s, e)
        f("BtnPickPartSleeve").Click += lambda s, e: self._on_part_sleeve(s, e)

    def ShowDialog(self):
        self._window.ShowDialog()

    def Close(self):
        self._window.Close()

    def _on_cancel(self, s, e):
        self.action = self.ACTION_CANCEL
        self._window.Close()

    def _on_viewport(self, s, e):
        self.action = self.ACTION_VIEWPORT
        self._window.Close()

    def _on_part_sleeve(self, s, e):
        self.action = self.ACTION_PART_SLEEVE
        self._window.Close()

def _do_pick_part_sleeve(config):
    """Pick content, pick sleeve, apply move + size instantly."""
    # Pick content
    try:
        ref = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.Element,
            SourceFilter(),
            "UpdateSleeve -- Pick Part + Sleeve: pick a duct or pipe")
        source_elem = doc.GetElement(ref.ElementId)
        source_data = extract_source_data(source_elem)
        CRASH_LOG.log("{0} Part+Sleeve src id={1}  {2}".format(
            _ts(), source_elem.Id.IntegerValue, source_data["display_size"]))
    except Exception:
        return

    # Pick sleeve (any element)
    try:
        ref2 = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.Element,
            "UpdateSleeve -- Pick Part + Sleeve: pick the sleeve")
        sleeve_elem = doc.GetElement(ref2.ElementId)
        CRASH_LOG.log("{0} Part+Sleeve slv id={1}".format(
            _ts(), sleeve_elem.Id.IntegerValue))
    except Exception:
        return

    # OS from config, shape-aware
    shape    = source_data.get("shape", "rect")
    oversize = get_oversize_for_shape(config, shape)

    # Use per-family params, fall back to auto-detect
    slv_fam_name = _safe_family_name(sleeve_elem)
    p_width, p_height, p_dia = get_family_params(config, slv_fam_name)
    # Strip CONFIG_NONE sentinel
    if p_width  == CONFIG_NONE: p_width  = ""
    if p_height == CONFIG_NONE: p_height = ""
    if p_dia    == CONFIG_NONE: p_dia    = ""
    # Auto-detect any that are blank
    if not p_width or not p_height or not p_dia:
        _, auto_w, auto_h, auto_d = detect_size_params(sleeve_elem)
        if not p_width  and shape == "rect":  p_width  = auto_w
        if not p_height and shape == "rect":  p_height = auto_h
        if not p_dia    and shape != "rect":  p_dia    = auto_d

    # Source centerpoint
    src_pt = None
    try:
        loc = source_elem.Location
        if hasattr(loc, "Curve"):
            c = loc.Curve
            p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
            src_pt = DB.XYZ((p0.X+p1.X)/2, (p0.Y+p1.Y)/2, (p0.Z+p1.Z)/2)
        elif hasattr(loc, "Point"):
            src_pt = loc.Point
    except Exception:
        pass

    lines = []
    t = DB.Transaction(doc, "UpdateSleeve - Part+Sleeve")
    t.Start()
    try:
        # Move
        if src_pt is not None:
            try:
                cur_loc = sleeve_elem.Location
                if hasattr(cur_loc, "Point"):
                    cur_pt = cur_loc.Point
                else:
                    c = cur_loc.Curve
                    p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                    cur_pt = DB.XYZ((p0.X+p1.X)/2, (p0.Y+p1.Y)/2, (p0.Z+p1.Z)/2)
                full_delta = DB.XYZ(
                    src_pt.X - cur_pt.X,
                    src_pt.Y - cur_pt.Y,
                    src_pt.Z - cur_pt.Z)
                try:
                    dc = source_elem.Location.Curve
                    dp0 = dc.GetEndPoint(0)
                    dp1 = dc.GetEndPoint(1)
                    raw = DB.XYZ(dp1.X-dp0.X, dp1.Y-dp0.Y, dp1.Z-dp0.Z)
                    length = raw.GetLength()
                    if length > 0.001:
                        ux = raw.X / length
                        uy = raw.Y / length
                        uz = raw.Z / length
                        dot  = full_delta.X*ux + full_delta.Y*uy + full_delta.Z*uz
                        move = DB.XYZ(full_delta.X - dot*ux,
                                      full_delta.Y - dot*uy,
                                      full_delta.Z - dot*uz)
                    else:
                        move = full_delta
                except Exception:
                    move = full_delta
                DB.ElementTransformUtils.MoveElement(doc, sleeve_elem.Id, move)
                lines.append("MOVED  delta ({0:.3f}, {1:.3f}, {2:.3f}) ft".format(
                    move.X, move.Y, move.Z))
            except Exception as ex:
                lines.append("MOVE FAIL: {0}".format(str(ex)))

        # Write size params
        for key, pname in (("Width", p_width), ("Height", p_height),
                           ("Diameter", p_dia)):
            if not pname or source_data.get(key) is None:
                continue
            val = source_data[key] + oversize
            ok, msg = write_param(sleeve_elem, pname, val)
            lines.append("{0}  {1} -> {2} = {3:.2f}in".format(
                "OK  " if ok else "FAIL", key, pname, val*12))
            if not ok:
                lines[-1] += "  ({0})".format(msg)

        t.Commit()
        CRASH_LOG.log("{0} Part+Sleeve done: {1}".format(_ts(), "; ".join(lines)))
    except Exception as ex:
        t.RollBack()
        CRASH_LOG.log("{0} Part+Sleeve ERROR: {1}".format(_ts(), str(ex)))

# ── Entry point ───────────────────────────────────────────────────
CRASH_LOG.log("{0} === SESSION ===  doc={1}".format(_ts(), doc.Title))

config      = load_config()
launch_xaml = os.path.join(os.path.dirname(__file__), "launch.xaml")
launch      = LaunchWindow(launch_xaml)
launch.ShowDialog()

if launch.action == LaunchWindow.ACTION_PART_SLEEVE:
    _do_pick_part_sleeve(config)

elif launch.action == LaunchWindow.ACTION_VIEWPORT:
    source_elem = None
    source_data = None
    try:
        ref = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.Element,
            SourceFilter(),
            "UpdateSleeve -- Pick a duct or pipe")
        source_elem = doc.GetElement(ref.ElementId)
        source_data = extract_source_data(source_elem)
        CRASH_LOG.log("{0} Picked id={1}  {2}".format(
            _ts(), source_elem.Id.IntegerValue, source_data["display_size"]))
    except Exception as ex:
        CRASH_LOG.log("{0} Pick cancelled: {1}".format(_ts(), str(ex)))

    if source_elem is not None and source_data is not None:
        xaml_path = os.path.join(os.path.dirname(__file__), "window.xaml")
        win = UpdateSleeveWindow(xaml_path, config, source_elem, source_data)
        win.ShowDialog()
