# -*- coding: utf-8 -*-
"""
NW Parameter Formula Tool
Revit 2022-2026 | PyRevit | IronPython 2.7

Writes formula-based string values to Revit element parameters.
Profiles define what to collect, how to filter, and what to write.
Works on any Revit element type — not limited to hangers.

PROFILE SCHEMA
--------------
{
  "name":           "Profile display name",
  "notes":          "Optional description shown in the run bar",
  "filter_logic":   "AND" | "OR",
  "allow_nested":   false,            // exclude nested sub-components
  "filter_rules": [
    {"type": "category_eq",           "value": "Generic Models"},
    {"type": "category_contains",     "value": "Pipe"},
    {"type": "param_has_value",       "param": "CP_Service Abbv"},
    {"type": "param_value_contains",  "param": "CP_Service Abbv", "value": "MECH"},
    {"type": "elem_name_contains",    "value": "Hanger"},
    {"type": "family_name_contains",  "value": "Beam Clamp"}
  ],
  "group_by_param": "CP_Service Abbv",
  "outputs": [
    {
      "target_param": "GTP_PointNumber_0",
      "label":        "Point Number",
      "sep":          "-",
      "seq_width":    3,
      "slots": [
        {"key": "project_number",        "label": "Job #",   "color": "#7b4ea0",
         "unit_type": "string",          "read_format": "str"},
        {"key": "param:CP_Service Abbv", "label": "Service", "color": "#1a7a8a",
         "unit_type": "string",          "read_format": "str"},
        {"key": "SEQUENCE",              "label": "###",     "color": "#1a8a4a",
         "unit_type": "string",          "read_format": "str"}
      ]
    }
  ]
}

EXECUTION PIPELINE
------------------
1. Window: user picks profile + source + mode, clicks Run
2. Window closes, execute() runs:
   a. Collect  — pick_box_select / collect_from_selection / collect_visible_elements
   b. Filter   — filter_elements_with_reasons() (nested check + rule match, all sources)
   c. Group    — by group_by_param
   d. Write    — Transaction: build formula string, Set() each output param
   e. Report   — _show_run_report() with consistent MessageBox popups
"""

# ==========================================================
# IMPORTS
# WPF/CLR, pyRevit, Revit DB — all top-level to avoid nested
# import issues with IronPython closures
# ==========================================================

import os
import re
import json

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import System
import System.Windows
import System.Windows.Controls
import System.Windows.Media
import System.Windows.Input

from System.Windows import (
    Window, MessageBox, MessageBoxButton, MessageBoxResult,
    ResizeMode, SystemParameters,
    DragDropEffects, DataObject,
    GridLength, Thickness, HorizontalAlignment, VerticalAlignment,
    SizeToContent, WindowStartupLocation, TextWrapping,
    CornerRadius, FontWeights, RoutedEventHandler, GridUnitType,
    Visibility
)
from System.Windows.Controls import (
    ListBoxItem,
    Grid, ColumnDefinition,
    DockPanel, StackPanel, WrapPanel, ScrollViewer,
    ListBox, TextBox, TextBlock, Button,
    RadioButton, CheckBox, Separator, Border,
    GridSplitter, Orientation, ScrollBarVisibility, Dock,
    ComboBox, ComboBoxItem,
    TabControl, TabItem
)
from System.Windows.Input import Cursors
from System.Windows.Media import SolidColorBrush, Color, Brushes
from System.Windows.Media import FontFamily as WpfFontFamily
from Microsoft.Win32 import OpenFileDialog, SaveFileDialog

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    BuiltInParameter, Transaction, UnitUtils,
    FilteredElementCollector,
    ElementCategoryFilter, LogicalOrFilter
)
from Autodesk.Revit.UI.Selection import ObjectType as PickObjectType, ISelectionFilter

logger = script.get_logger()

def double_infinity():
    """Return WPF's 'no constraint' value for MaxHeight/MaxWidth."""
    try:
        return System.Double.PositiveInfinity
    except Exception:
        return 9999

_REVIT_CATEGORY_NAMES = []   # populated lazily on first use

def get_revit_category_names(doc):
    """
    Return a sorted list of all user-visible Revit category names.
    Cached after first call.
    """
    global _REVIT_CATEGORY_NAMES
    if _REVIT_CATEGORY_NAMES:
        return _REVIT_CATEGORY_NAMES
    names = []
    try:
        cats = doc.Settings.Categories
        for cat in cats:
            try:
                n = cat.Name
                if n:
                    names.append(n)
            except Exception:
                pass
    except Exception:
        pass
    _REVIT_CATEGORY_NAMES = sorted(set(names))
    return _REVIT_CATEGORY_NAMES

# ==========================================================
# PATHS
# BUNDLE_DIR  = pushbutton folder (division profiles live here)
# SHARED_DIR  = bundle/profiles/  (requires admin write access)
# USER_DIR    = %APPDATA%/pyRevit/NW_Tools/Renumber Tool/profiles/
# ==========================================================

BUNDLE_DIR = os.path.dirname(__file__)
SHARED_DIR = os.path.join(BUNDLE_DIR, "profiles")
APPDATA    = os.environ.get("APPDATA", os.path.expanduser("~"))
USER_DIR   = os.path.join(APPDATA, "pyRevit", "NW_Tools", "Renumber Tool", "profiles")

for _d in (SHARED_DIR, USER_DIR):
    if not os.path.exists(_d):
        try:
            os.makedirs(_d)
        except Exception:
            pass

ROUND_DENOM = 16

# ==========================================================
# PARAMETER TOKEN COLORS
# Used in the palette and slot chips to indicate unit type
# ==========================================================

COLOR_PROJECT  = "#7b4ea0"   # project_number / project_name tokens
COLOR_LENGTH   = "#2a7ae2"   # length unit type parameters
COLOR_TEXT     = "#555555"   # generic text
COLOR_SEQUENCE = "#1a8a4a"   # auto-increment sequence token
COLOR_STATIC   = "#888888"   # static literal text tokens
COLOR_CUSTOM   = "#1a7a8a"   # manually added / other parameters
COLOR_ASK      = "#c07000"   # ASK: user-input-at-run-time tokens

# ==========================================================
# UTILITIES
# ==========================================================

def _safe_filename(name):
    return re.sub(r'[^\w\- ]', '_', name).strip()

# (No built-in profiles — use Import or Save to Division to add shared profiles)

# ==========================================================
# REVIT ELEMENT HELPERS
# Raw element/selection access — no filtering applied here
# ==========================================================

def get_selected_elements(doc):
    sel = revit.get_selection()
    try:
        ids = list(sel.element_ids)
        return [doc.GetElement(eid) for eid in ids]
    except Exception:
        pass
    try:
        items = list(sel)
        if items and hasattr(items[0], "IntegerValue"):
            return [doc.GetElement(x) for x in items]
        return items
    except Exception:
        return []

# ==========================================================
# ELEMENT COLLECTION
# Three sources: pick_box_select / collect_from_selection /
# collect_visible_elements. All return raw elements — filtering
# is applied by filter_elements_with_reasons() in execute().
# Category pre-filtering at collector level keeps Python fast
# on large models (never loads excluded categories).
# ==========================================================

def _extract_category_ids(doc, profile):
    """
    Parse filter_rules for category_eq / category_contains rules.
    Returns a list of Revit CategoryId objects that match.
    Falls back to empty list (meaning: no pre-filter, collect all).
    """
    rules = profile.get("filter_rules", [])
    cat_values = []
    for r in rules:
        rtype = r.get("type", "")
        val   = r.get("value", "").strip().lower()
        if rtype in ("category_eq", "category_contains") and val:
            cat_values.append((rtype, val))

    if not cat_values:
        return []

    from Autodesk.Revit.DB import Category
    matched_ids = []
    try:
        cats = doc.Settings.Categories
        for cat in cats:
            cat_name_lower = cat.Name.lower()
            for rtype, val in cat_values:
                if rtype == "category_eq" and cat_name_lower == val:
                    matched_ids.append(cat.Id)
                    break
                elif rtype == "category_contains" and val in cat_name_lower:
                    matched_ids.append(cat.Id)
                    break
    except Exception as ex:
        logger.warning("_extract_category_ids failed: {}".format(ex))

    return matched_ids

def _collector_with_categories(doc, scope_id, cat_ids):
    """
    Build a FilteredElementCollector scoped to a view or selection,
    pre-filtered to the given category IDs.
    scope_id: view ElementId (or None for whole-doc collector).
    cat_ids: list of CategoryId. If empty, no category filter applied.
    """
    if scope_id is not None:
        col = FilteredElementCollector(doc, scope_id)
    else:
        col = FilteredElementCollector(doc)

    col = col.WhereElementIsNotElementType()

    if not cat_ids:
        return col

    if len(cat_ids) == 1:
        col = col.WherePasses(ElementCategoryFilter(cat_ids[0]))
    else:
        # Union multiple category filters with LogicalOrFilter
        filters = [ElementCategoryFilter(cid) for cid in cat_ids]
        while len(filters) > 1:
            # LogicalOrFilter takes exactly two filters; chain them
            filters = [LogicalOrFilter(filters[i], filters[i+1])
                       for i in range(0, len(filters)-1, 2)] +                       (filters[-1:] if len(filters) % 2 else [])
        col = col.WherePasses(filters[0])

    return col

def collect_visible_elements(doc, uidoc, profile=None):
    """
    Return elements visible in the active view.
    If profile has category rules, pre-filters at collector level
    so large views don't load thousands of irrelevant elements.
    """
    try:
        view   = uidoc.ActiveView
        cat_ids = _extract_category_ids(doc, profile) if profile else []
        col    = _collector_with_categories(doc, view.Id, cat_ids)
        result = list(col)
        if cat_ids:
            logger.debug("Pre-filtered collector: {} cats, {} elements".format(
                len(cat_ids), len(result)))
        return result
    except Exception as ex:
        logger.warning("collect_visible_elements failed: {}".format(ex))
        return []

def collect_from_selection(doc, uidoc, profile=None):
    """
    Return currently selected elements, optionally pre-filtered
    by category at the collector level using a selection-scoped collector.
    """
    sel = revit.get_selection()
    try:
        ids = list(sel.element_ids)
    except Exception:
        try:
            items = list(sel)
            ids = items if items and hasattr(items[0], "IntegerValue") else []
        except Exception:
            ids = []

    if not ids:
        return []

    # If no profile or no category rules, just return selected elements directly
    cat_ids = _extract_category_ids(doc, profile) if profile else []
    if not cat_ids:
        return [doc.GetElement(eid) for eid in ids]

    # Use a selection-scoped collector for efficient category pre-filtering
    try:
        from Autodesk.Revit.DB import ElementId
        col = FilteredElementCollector(doc, ids)
        col = col.WhereElementIsNotElementType()
        filters = [ElementCategoryFilter(cid) for cid in cat_ids]
        while len(filters) > 1:
            filters = [LogicalOrFilter(filters[i], filters[i+1])
                       for i in range(0, len(filters)-1, 2)] +                       (filters[-1:] if len(filters) % 2 else [])
        col = col.WherePasses(filters[0])
        return list(col)
    except Exception as ex:
        logger.warning("collect_from_selection category filter failed: {} — falling back".format(ex))
        return [doc.GetElement(eid) for eid in ids]

def _make_selection_filter(profile):
    """
    Build a proper .NET ISelectionFilter using clr.
    IronPython requires explicit interface implementation via clr
    for APIs like PickElementsByRectangle that check the interface
    at the CLR level rather than duck-typing.
    Returns None if interface wiring fails (graceful fallback).
    """
    if not profile:
        return None
    try:
        import clr as _clr
        from Autodesk.Revit.UI.Selection import ISelectionFilter as _ISF

        class _Filter(_ISF):
            def __init__(self, p):
                self._p = p

            def AllowElement(self, elem):
                try:
                    allow_nested = self._p.get("allow_nested", False)
                    if not is_top_level_instance(elem, allow_nested):
                        return False
                    kept, _ = filter_elements([elem], self._p)
                    return len(kept) > 0
                except Exception:
                    return True

            def AllowReference(self, ref, pos):
                return False

        return _Filter(profile)
    except Exception as ex:
        logger.warning("_make_selection_filter failed: {} — no live filter".format(ex))
        return None

def pick_box_select(uidoc, profile=None):
    """
    Open Revit's standard multi-select mode with the Finish button.
    Uses PickObjects(ObjectType.Element) which gives the normal Revit
    selection UX: click elements one at a time, press Finish (green
    checkmark in top-left ribbon) or Esc to cancel.

    If profile has filter rules, a live ISelectionFilter grays out
    non-matching elements. Falls back to post-filtering if the live
    filter can't be wired.
    """
    has_rules  = bool(profile and profile.get("filter_rules"))
    sel_filter = _make_selection_filter(profile) if has_rules else None

    if has_rules:
        prompt = "Select elements — only matching elements are enabled. Click Finish when done."
    else:
        prompt = "Select elements, then click Finish."

    # ── Attempt 1: PickObjects with live ISelectionFilter ───────────
    # Note: live filter grays out non-matching elements in the UI but
    # is NOT authoritative — filter_elements() always runs post-pick
    # to ensure nested components and non-matching elements are removed.
    if sel_filter is not None:
        try:
            refs = uidoc.Selection.PickObjects(
                PickObjectType.Element,
                sel_filter,
                prompt
            )
            return [uidoc.Document.GetElement(r) for r in refs]
        except Exception as ex:
            logger.warning("PickObjects(filter) failed: {}".format(ex))

    # ── Attempt 2: PickObjects without filter, post-filter result ───
    try:
        refs = uidoc.Selection.PickObjects(PickObjectType.Element, prompt)
        return [uidoc.Document.GetElement(r) for r in refs]
    except Exception as ex:
        logger.warning("PickObjects failed: {}".format(ex))
        return []

def is_top_level_instance(elem, allow_nested=False):
    """
    Returns True if the element should be included based on nesting.
    allow_nested=False (default): only top-level instances pass.
    allow_nested=True: nested sub-components also pass.
    """
    if allow_nested:
        return True
    try:
        from Autodesk.Revit.DB import FamilyInstance
        if not isinstance(elem, FamilyInstance):
            return True   # not a family instance — always include
        sc = elem.SuperComponent
        return sc is None
    except Exception:
        return True   # on error, don't exclude

def lookup_param(elem, name):
    """
    Look up a named parameter on an element.
    Checks instance params, GetParameters (shared params),
    then element type params for RFA families.
    """
    if elem is None or not name:
        return None
    # 1. Instance LookupParameter
    try:
        p = elem.LookupParameter(name)
        if p is not None:
            return p
    except Exception:
        pass
    # 2. GetParameters — catches shared params LookupParameter can miss
    try:
        params = elem.GetParameters(name)
        if params and params.Count > 0:
            return params[0]
    except Exception:
        pass
    # 3. Element type — for type-level params on RFA/Generic Model families
    try:
        etype = elem.Document.GetElement(elem.GetTypeId())
        if etype is not None:
            p = etype.LookupParameter(name)
            if p is not None:
                return p
            params = etype.GetParameters(name)
            if params and params.Count > 0:
                return params[0]
    except Exception:
        pass
    return None

def param_text(param):
    """Return readable text for a parameter. Returns '' for truly empty params."""
    if not param:
        return ""
    try:
        val = param.AsString()
        if val:
            return val.strip()
        # AsString() returned None or "" — check storage type before
        # falling through to AsInteger, because integer storage on an
        # empty string param returns 0 which looks like a real value.
        storage = param.StorageType.ToString()
        if storage == "String":
            return ""   # definitively empty string param
    except Exception:
        pass
    try:
        val = param.AsValueString()
        if val:
            return val.strip()
    except Exception:
        pass
    try:
        # Only use AsInteger for genuinely integer-storage params
        storage = param.StorageType.ToString()
        if storage == "Integer":
            return str(param.AsInteger())
    except Exception:
        pass
    return ""

def param_is_empty(param):
    """
    True if the parameter has no meaningful user-visible value.
    Used by fill_missing to decide whether to write.
    More reliable than checking param_text() == '' because it
    avoids the AsInteger()==0 false-positive on blank string params.
    """
    if not param:
        return True
    try:
        storage = param.StorageType.ToString()
        if storage == "String":
            val = param.AsString()
            return val is None or val.strip() == ""
        if storage == "Integer":
            # 0 may be a valid value for non-text params, but for our
            # output params (GTP_PointNumber_0 etc.) they are strings.
            # Return True only if AsString is also empty.
            val = param.AsString()
            return val is None or val.strip() == ""
        if storage == "Double":
            return False  # doubles always have a value
        # ElementId storage
        val = param.AsString()
        return val is None or val.strip() == ""
    except Exception:
        return True

def get_job_number(doc):
    pi = doc.ProjectInformation
    try:
        if pi.Number and pi.Number.strip():
            return pi.Number.strip()
    except Exception:
        pass
    try:
        param = pi.get_Parameter(BuiltInParameter.PROJECT_NUMBER)
        value = param_text(param)
        if value:
            return value
    except Exception:
        pass
    value = forms.ask_for_string(
        default="",
        prompt="ProjectInformation.Number is blank.\nEnter Job Number:",
        title="Job Number Required"
    )
    if not value:
        script.exit()
    return value.strip()

# ==========================================================
# ELEMENT ATTRIBUTE HELPERS
# ==========================================================

def elem_category_name(elem):
    try:
        return elem.Category.Name or ""
    except Exception:
        return ""

def elem_name(elem):
    try:
        return elem.Name or ""
    except Exception:
        return ""

def elem_family_name(elem):
    try:
        return elem.Symbol.FamilyName or ""
    except Exception:
        try:
            return elem.FamilyName or ""
        except Exception:
            return ""

# ==========================================================
# FILTER SYSTEM
# filter_elements()             — lightweight, used by ISelectionFilter
# filter_elements_with_reasons()— used by execute(), returns skip log
# Both check is_top_level_instance() before profile rules.
# ==========================================================

RULE_TYPES = [
    ("category_eq",          "Category equals"),
    ("category_contains",    "Category contains"),
    ("param_has_value",      "Has parameter"),
    ("param_value_contains", "Param value contains"),
    ("elem_name_contains",   "Element name contains"),
    ("family_name_contains", "Family name contains"),
]

def _rule_matches(elem, rule):
    """Return True if elem satisfies a single rule dict."""
    rtype = rule.get("type", "")
    value = rule.get("value", "").strip().lower()

    if rtype == "category_eq":
        return elem_category_name(elem).lower() == value

    if rtype == "category_contains":
        return value in elem_category_name(elem).lower()

    if rtype == "param_has_value":
        pname = rule.get("param", "").strip()
        if not pname:
            return False
        p = lookup_param(elem, pname)
        return p is not None and not param_is_empty(p)

    if rtype == "param_value_contains":
        pname = rule.get("param", "").strip()
        if not pname:
            return False
        pval = param_text(lookup_param(elem, pname)).lower()
        return value in pval

    if rtype == "elem_name_contains":
        return value in elem_name(elem).lower()

    if rtype == "family_name_contains":
        return value in elem_family_name(elem).lower()

    return True  # unknown rule type — pass through

def filter_elements(elements, profile):
    """
    Apply profile filter_rules with filter_logic AND/OR.
    Always excludes nested sub-components unless allow_nested is set.
    Returns (kept_list, skipped_count).
    """
    rules = profile.get("filter_rules", [])
    logic = profile.get("filter_logic", "AND").upper()
    allow_nested = profile.get("allow_nested", False)

    kept = []
    skipped = 0
    for elem in elements:
        if elem is None:
            continue
        # First gate: nested sub-component check
        if not is_top_level_instance(elem, allow_nested):
            skipped += 1
            continue
        # Second gate: profile filter rules
        if rules:
            if logic == "OR":
                match = any(_rule_matches(elem, r) for r in rules)
            else:
                match = all(_rule_matches(elem, r) for r in rules)
            if not match:
                skipped += 1
                continue
        kept.append(elem)

    return kept, skipped

def filter_elements_with_reasons(elements, profile):
    """
    Like filter_elements but returns detailed skip reasons per element.
    Used by execute() to populate the skip report.
    Returns (kept_list, [(elem_id_str, reason), ...])
    """
    rules = profile.get("filter_rules", [])
    logic = profile.get("filter_logic", "AND").upper()
    allow_nested = profile.get("allow_nested", False)

    kept = []
    skip_log = []

    for elem in elements:
        if elem is None:
            continue
        eid = str(elem.Id.IntegerValue)

        # Nested sub-component
        if not is_top_level_instance(elem, allow_nested):
            cat = elem_category_name(elem) or "?"
            skip_log.append((eid, "Nested sub-component  [{}]".format(cat)))
            continue

        # Profile filter rules
        if rules:
            if logic == "OR":
                match = any(_rule_matches(elem, r) for r in rules)
            else:
                match = all(_rule_matches(elem, r) for r in rules)
            if not match:
                cat = elem_category_name(elem) or "?"
                skip_log.append((eid, "Failed filter rules  [{}]".format(cat)))
                continue

        kept.append(elem)

    return kept, skip_log

def get_all_param_names(elements):
    """
    Scan provided elements (all of them if <= 5, else first 5)
    and return sorted list of all parameter names found.
    For re-scan we pass a single element so this is fast.
    """
    names = set()
    sample = elements if len(elements) <= 5 else elements[:5]
    for elem in sample:
        if elem is None:
            continue
        try:
            for p in elem.Parameters:
                try:
                    n = p.Definition.Name
                    if n:
                        names.add(n)
                except Exception:
                    pass
        except Exception:
            pass
    return sorted(names)

def pick_one_element(uidoc):
    """
    Open a Revit PickObject prompt and return the picked element (or None).
    """
    try:
        ref = uidoc.Selection.PickObject(
            PickObjectType.Element,
            "Pick one element to scan its parameters"
        )
        return uidoc.Document.GetElement(ref)
    except Exception:
        return None

# ==========================================================
# LENGTH / UNIT CONVERSION
# Converts Revit internal feet to human-readable strings.
# Used by _read_param_with_format() for 'length' unit type.
# ==========================================================

def _gcd(a, b):
    while b:
        a, b = b, a % b
    return a

def _round_frac(value, denom):
    return int(round(value * denom)) / float(denom)

def _inches_to_string(inches):
    inches = _round_frac(inches, ROUND_DENOM)
    whole = int(inches)
    frac = inches - whole
    num = int(round(frac * ROUND_DENOM))
    if num == 0:
        return '{}"'.format(whole)
    g = _gcd(num, ROUND_DENOM)
    num //= g
    den = ROUND_DENOM // g
    if whole == 0:
        return '{}/{}"'.format(num, den)
    return '{} {}/{}"'.format(whole, num, den)

def _parse_inches_component(value):
    value = value.strip()
    m = re.match(r"(\d+)\s+(\d+)/(\d+)", value)
    if m:
        return float(m.group(1)) + float(m.group(2)) / float(m.group(3))
    m = re.match(r"(\d+)/(\d+)", value)
    if m:
        return float(m.group(1)) / float(m.group(2))
    m = re.match(r"^\d+$", value)
    if m:
        return float(m.group(0))
    return None

def length_param_to_inches_string(param):
    if not param:
        return ""
    try:
        if param.StorageType.ToString() == "Double":
            feet = param.AsDouble()
            try:
                from Autodesk.Revit.DB import UnitTypeId
                inches = UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Inches)
            except Exception:
                try:
                    from Autodesk.Revit.DB import DisplayUnitType
                    inches = UnitUtils.ConvertFromInternalUnits(
                        feet, DisplayUnitType.DUT_DECIMAL_INCHES)
                except Exception:
                    inches = feet * 12.0
            return _inches_to_string(inches)
    except Exception:
        pass
    raw = param_text(param)
    if not raw:
        return ""
    cleaned = re.sub(r"\s+", " ", raw.replace("-", " ")).strip()
    m = re.search(r"(-?\d+)\s*'\s*(.*)", cleaned)
    if m:
        feet = int(m.group(1))
        remainder = m.group(2).replace('"', '').strip()
        iv = _parse_inches_component(remainder)
        if iv is None:
            return raw
        return _inches_to_string(feet * 12.0 + iv)
    iv = _parse_inches_component(cleaned.replace('"', ''))
    if iv is None:
        return raw
    return _inches_to_string(iv)

# ==========================================================
# UNIT TYPE DETECTION + FORMAT SYSTEM
# _detect_unit_type()     — 4-tier detection: ParameterType,
#                           UnitType, GetDataType, AsValueString
# _read_param_with_format()— reads a param with the given format code
# FORMAT_OPTIONS          — ordered list of (code, label) per type
# DEFAULT_FORMAT          — default format code per unit type
# ==========================================================

# Format codes per unit category
# Each category has an ordered list of (code, label) pairs.
# The chip badge cycles through these on click.
FORMAT_OPTIONS = {
    "length":  [("in_frac",  'in"'),   ("in_dec",  "in."),  ("ft_in",   "ft-in"), ("ft_dec",  "ft."),  ("revit",   "rev")],
    "area":    [("val",      "val"),   ("revit",   "rev")],
    "volume":  [("val",      "val"),   ("revit",   "rev")],
    "flow":    [("val",      "val"),   ("revit",   "rev")],
    "angle":   [("val",      "val"),   ("revit",   "rev")],
    "force":   [("val",      "val"),   ("revit",   "rev")],
    "yesno":   [("tf",       "T/F"),   ("yn",      "Y/N"),   ("int",     "0/1")],
    "integer": [("int",      "int")],
    "string":  [("str",      "str")],
    "other":   [("revit",    "rev"),   ("val",     "val")],
}

DEFAULT_FORMAT = {
    "length":  "in_frac",
    "area":    "val",
    "volume":  "val",
    "flow":    "val",
    "angle":   "val",
    "force":   "val",
    "yesno":   "tf",
    "integer": "int",
    "string":  "str",
    "other":   "revit",
}

def _detect_unit_type(param):
    """
    Detect semantic unit category. Tries 4 sources:
    1. ParameterType enum  2. UnitType enum
    3. GetDataType/ForgeTypeId  4. AsValueString heuristic
    Unidentified Double defaults to "length" (most MEP doubles are dimensional).
    """
    if param is None:
        return "string"
    try:
        storage = param.StorageType.ToString()
        if storage == "String":
            return "string"
        if storage == "Integer":
            try:
                defn = param.Definition
                pt = ""
                try: pt = str(defn.ParameterType)
                except Exception: pass
                if not pt:
                    try: pt = str(defn.GetDataType())
                    except Exception: pass
                if any(x in pt for x in ["YesNo","Boolean","yesno","bool"]):
                    return "yesno"
            except Exception:
                pass
            return "integer"
        if storage == "Double":
            # Source 1: ParameterType
            try:
                defn = param.Definition
                pt = str(defn.ParameterType).lower()
                if any(x in pt for x in ["length","size","distance","elevation","thickness","height","width","depth","radius","diameter"]):
                    return "length"
                if "area" in pt: return "area"
                if "volume" in pt: return "volume"
                if any(x in pt for x in ["flow","airflow"]): return "flow"
                if any(x in pt for x in ["angle","rotation"]): return "angle"
                if any(x in pt for x in ["force","weight","mass","stress","moment","structural"]): return "force"
            except Exception:
                pass
            # Source 2+3: UnitType / GetDataType
            unit_str = ""
            try:
                defn = param.Definition
                unit_str = str(defn.UnitType)
            except Exception: pass
            if not unit_str or unit_str in ("UT_Undefined","undefined",""):
                try:
                    defn = param.Definition
                    unit_str = str(defn.GetDataType())
                except Exception: pass
            if unit_str:
                u = unit_str.lower()
                if any(x in u for x in ["ut_length","length","millimeter","centimeter","meter","feet","inch","slope","offset"]):
                    return "length"
                if any(x in u for x in ["area"]): return "area"
                if any(x in u for x in ["volume"]): return "volume"
                if any(x in u for x in ["flow","airflow","hvacairflow","pipingflow"]): return "flow"
                if any(x in u for x in ["angle"]): return "angle"
                if any(x in u for x in ["force","weight","mass","stress","moment","structural"]): return "force"
            # Source 4: AsValueString heuristic
            try:
                vs = (param.AsValueString() or "").strip()
                if vs.endswith('"') or vs.endswith("'") or "'- " in vs: return "length"
                if any(vs.endswith(x) for x in ["SF","m2","m\u00b2"]): return "area"
                if any(vs.endswith(x) for x in ["CF","m3","m\u00b3"]): return "volume"
                if any(vs.endswith(x) for x in ["CFM","L/s"]): return "flow"
                if vs.endswith("\u00b0"): return "angle"
            except Exception: pass
            # Default for unidentified Double — most MEP doubles are lengths
            return "length"
        return "other"
    except Exception:
        return "string"

def _read_param_with_format(param, read_format, doc=None):
    """
    Read a parameter and format it according to read_format code.
    Falls back to param_text() on any error.
    """
    if param is None:
        return ""
    try:
        storage = param.StorageType.ToString()

        # ── String ──────────────────────────────────────────────────
        if storage == "String" or read_format == "str":
            return param_text(param)

        # ── YesNo / Boolean ─────────────────────────────────────────
        if read_format in ("tf", "yn", "int"):
            try:
                val = param.AsInteger()
                if read_format == "tf":  return "True"  if val else "False"
                if read_format == "yn":  return "Yes"   if val else "No"
                if read_format == "int": return str(val)
            except Exception:
                return param_text(param)

        # ── Integer ─────────────────────────────────────────────────
        if read_format == "int" or storage == "Integer":
            try:
                return str(param.AsInteger())
            except Exception:
                return param_text(param)

        # ── Revit display string (any double) ────────────────────────
        if read_format == "revit":
            try:
                vs = param.AsValueString()
                return vs.strip() if vs else ""
            except Exception:
                return param_text(param)

        # -- Value only (strip unit suffix from AsValueString) -----------
        if read_format == "val":
            try:
                vs = param.AsValueString()
                if vs:
                    import re as _re
                    cleaned = _re.sub(r'\s+[^\d.,]+$', '', vs.strip()).strip()
                    return cleaned if cleaned else vs.strip()
            except Exception:
                pass
            try:
                return "{:.4f}".format(param.AsDouble()).rstrip('0').rstrip('.')
            except Exception:
                return param_text(param)

        # ── Length formats ───────────────────────────────────────────
        if storage == "Double":
            try:
                feet = param.AsDouble()
                try:
                    from Autodesk.Revit.DB import UnitTypeId
                    inches_total = UnitUtils.ConvertFromInternalUnits(feet, UnitTypeId.Inches)
                except Exception:
                    try:
                        from Autodesk.Revit.DB import DisplayUnitType
                        inches_total = UnitUtils.ConvertFromInternalUnits(
                            feet, DisplayUnitType.DUT_DECIMAL_INCHES)
                    except Exception:
                        inches_total = feet * 12.0

                if read_format == "in_frac":
                    return _inches_to_string(inches_total)

                if read_format == "in_dec":
                    return "{:.4f}".format(inches_total).rstrip('0').rstrip('.')

                if read_format == "ft_dec":
                    return "{:.4f}".format(feet).rstrip('0').rstrip('.')

                if read_format == "ft_in":
                    whole_ft = int(feet)
                    rem_in   = (feet - whole_ft) * 12.0
                    return "{}\'- {}".format(whole_ft, _inches_to_string(rem_in))

            except Exception:
                pass

        # Fallback
        return param_text(param)

    except Exception:
        return param_text(param)

def get_all_param_info(elements):
    """
    Scan elements and return list of dicts:
      {name, storage_type, unit_type, default_format}
    Used to build the token palette with auto-detected formats.
    """
    seen = {}   # name -> info dict
    sample = elements if len(elements) <= 5 else elements[:5]
    for elem in sample:
        if elem is None:
            continue
        try:
            for p in elem.Parameters:
                try:
                    n = p.Definition.Name
                    if not n or n in seen:
                        continue
                    ut = _detect_unit_type(p)
                    seen[n] = {
                        "name":           n,
                        "storage_type":   p.StorageType.ToString(),
                        "unit_type":      ut,
                        "default_format": DEFAULT_FORMAT.get(ut, "str"),
                    }
                except Exception:
                    pass
        except Exception:
            pass
    return sorted(seen.values(), key=lambda x: x["name"])

# ==========================================================
# FORMULA RESOLUTION
# resolve_slot_value()     — resolves one slot against a real element
# build_string_from_slots()— builds the full formula string
# preview_from_slots()     — builds a sample preview string for the UI
# ==========================================================

def resolve_slot_value(slot, elem, doc, seq_number=None, seq_width=3, user_inputs=None):
    key = slot.get("key", "")
    if key == "SEQUENCE":
        if seq_number is not None:
            return str(seq_number).zfill(seq_width)
        return "###"
    if key.startswith("ASK:"):
        if user_inputs is not None:
            return user_inputs.get(key, "")
        # Preview / live-element mode — show placeholder
        label = key[len("ASK:"):]
        return "<{}>".format(label)
    if key == "project_number":
        return get_job_number(doc)
    if key == "project_name":
        try:
            return doc.ProjectInformation.Name.strip()
        except Exception:
            return ""
    if key.startswith("static:"):
        return key[len("static:"):]
    if key.startswith("param:"):
        pname = key[len("param:"):]
        param = lookup_param(elem, pname)
        read_format = slot.get("read_format", None)
        # Legacy support: is_length=True maps to in_frac
        if read_format is None:
            if slot.get("is_length", False):
                read_format = "in_frac"
            else:
                # Auto-detect from param if no format set
                ut = _detect_unit_type(param)
                read_format = DEFAULT_FORMAT.get(ut, "str")
        return _read_param_with_format(param, read_format, doc)
    return ""

def build_string_from_slots(slots, sep, elem, doc, seq_number=None, seq_width=3, user_inputs=None):
    parts = []
    for slot in slots:
        val = resolve_slot_value(slot, elem, doc, seq_number, seq_width, user_inputs=user_inputs)
        if val:
            parts.append(val)
    return sep.join(parts)

# ==========================================================
# SEQUENCING
# get_next_sequence() scans existing param values to find the
# highest existing sequence number and returns next+1.
# Used for fill_missing and start_highest run modes.
# ==========================================================

def get_next_sequence(elems, prefix, sep, target_param):
    """Scan existing values in target_param to find highest sequence under prefix."""
    max_seq = 0
    if not prefix:
        pattern = r"^(\d+)$"
    else:
        pattern = r"^" + re.escape(prefix) + re.escape(sep) + r"(\d+)$"
    for e in elems:
        current = param_text(lookup_param(e, target_param))
        m = re.match(pattern, current)
        if m:
            seq = int(m.group(1))
            if seq > max_seq:
                max_seq = seq
    return max_seq + 1

# ==========================================================
# PROFILE I/O
# Profiles are stored as JSON files in SHARED_DIR or USER_DIR.
# list_profiles() merges both dirs; division profiles get readonly=True.
# save_profile() / delete_profile() raise on permission errors.
# ==========================================================

def list_profiles():
    profiles = []

    def _load_dir(directory, readonly):
        if not os.path.isdir(directory):
            return
        for fname in sorted(os.listdir(directory)):
            if not fname.lower().endswith(".json"):
                continue
            fpath = os.path.join(directory, fname)
            try:
                with open(fpath, "r") as f:
                    data = json.load(f)
                data["readonly"]  = readonly
                data["_filepath"] = fpath
                profiles.append(data)
            except Exception as e:
                logger.warning("Could not load profile {}: {}".format(fname, e))

    _load_dir(SHARED_DIR, True)
    _load_dir(USER_DIR,   False)
    return profiles

def save_profile(profile_data, destination="user"):
    """
    destination: "user"     -> USER_DIR  (always works)
                 "division" -> SHARED_DIR (may raise PermissionError)
    """
    fname = _safe_filename(profile_data["name"]) + ".json"
    folder = SHARED_DIR if destination == "division" else USER_DIR
    fpath = os.path.join(folder, fname)
    to_save = {k: v for k, v in profile_data.items()
               if k not in ("readonly", "_filepath")}
    with open(fpath, "w") as f:
        json.dump(to_save, f, indent=2)
    return fpath

def delete_profile(profile_data):
    """
    Delete the profile JSON file from disk.
    Works for both user and division profiles.
    Division profiles require write access to the bundle folder — raises
    PermissionError or OSError if access is denied, caller handles it.
    """
    fp = profile_data.get("_filepath", "")
    if not fp:
        raise ValueError("Profile has no file path — it may not be saved yet.")
    if not os.path.exists(fp):
        raise ValueError("Profile file not found:\n{}".format(fp))
    os.remove(fp)   # raises PermissionError if no write access

# ==========================================================
# PREVIEW HELPERS
# ==========================================================

SAMPLE_VALUES = {
    "project_number": "22-1234",
    "project_name":   "Sample Project",
    "SEQUENCE":       "001",
}

def preview_from_slots(slots, sep, seq_width=3, extra_samples=None):
    """
    Build a preview string from slots using sample values.
    Uses extra_samples (from real element scan) when available,
    falls back to SAMPLE_VALUES. Shows read_format in the slot label
    so the user knows what format will be applied.
    """
    samples = dict(SAMPLE_VALUES)
    if extra_samples:
        samples.update(extra_samples)
    parts = []
    for slot in slots:
        key = slot.get("key", "")
        if key == "SEQUENCE":
            parts.append("001".zfill(seq_width))
        elif key.startswith("ASK:"):
            label = key[len("ASK:"):]
            parts.append("<{}>".format(label))
        elif key.startswith("static:"):
            v = key[len("static:"):]
            if v:
                parts.append(v)
        else:
            v = samples.get(key)
            if v:
                # Show format suffix in preview so user knows which format applies
                fmt = slot.get("read_format")
                unit_type = slot.get("unit_type", "string")
                if fmt and fmt != "str" and unit_type != "string":
                    fmt_opts = FORMAT_OPTIONS.get(unit_type, [])
                    badge = fmt
                    for code, lbl in fmt_opts:
                        if code == fmt:
                            badge = lbl
                            break
                    parts.append("{}[{}]".format(v, badge))
                else:
                    parts.append(v)
            else:
                # No sample — show param name stub so preview isn't just ?
                if key.startswith("param:"):
                    pname = key[len("param:"):]
                    parts.append("<{}>".format(pname[:12]))
                else:
                    parts.append("?")
    return sep.join(parts)

# ==========================================================
# CHIP FACTORY  (module-level so nested fns can use it)
# ==========================================================

def _hex_to_color(hex_str):
    h = hex_str.lstrip("#")
    if len(h) == 6:
        return Color.FromRgb(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    return Color.FromRgb(0x55, 0x55, 0x55)

def make_chip(token, removable=False, refresh_cb=None):
    """
    Build a WPF Border chip for a token dict.
    Removable chips show:  Label  |  Full Name  [fmt]  x
    refresh_cb: called after format cycle to rebuild the slot list.
    """
    color = token.get("color", COLOR_TEXT)
    bdr = Border()
    bdr.Margin = Thickness(2, 2, 4, 2)
    bdr.CornerRadius = CornerRadius(4)
    bdr.Background = SolidColorBrush(_hex_to_color(color))
    bdr.Cursor = Cursors.Hand
    bdr.Tag = token

    # Full name for tooltip always
    full_name = token.get("desc") or token.get("key", "")
    if full_name.startswith("param:"):
        full_name = full_name[len("param:"):]
    bdr.ToolTip = full_name if full_name else token.get("label", "")

    inner = StackPanel()
    inner.Orientation = Orientation.Horizontal

    if removable:
        # Param name label — full_name is the param name, label is short alias
        # Show param name directly (no duplicate label|name pattern)
        display = full_name if full_name else token.get("label", "?")
        # Truncate for display, full name always in tooltip
        chip_lbl = TextBlock()
        chip_lbl.Text = display[:28] + ("..." if len(display) > 28 else "")
        chip_lbl.FontFamily = WpfFontFamily("Consolas")
        chip_lbl.FontSize = 11
        chip_lbl.FontWeight = FontWeights.Bold
        chip_lbl.Foreground = Brushes.White
        chip_lbl.VerticalAlignment = VerticalAlignment.Center
        chip_lbl.Margin = Thickness(6, 0, 2, 0)
        inner.Children.Add(chip_lbl)

        # Format badge — clickable TextBlock cycles read format
        unit_type   = token.get("unit_type", "string")
        read_format = token.get("read_format", DEFAULT_FORMAT.get(unit_type, "str"))
        fmt_options = FORMAT_OPTIONS.get(unit_type, FORMAT_OPTIONS["string"])

        if len(fmt_options) > 1:
            cur_badge = read_format
            for code, flbl in fmt_options:
                if code == read_format:
                    cur_badge = flbl
                    break
            all_labels_str = " | ".join("{}: {}".format(fl, c) for c, fl in fmt_options)
            fmt_tb = TextBlock()
            fmt_tb.Text = "[{}]".format(cur_badge)
            fmt_tb.Tag = token
            fmt_tb.FontSize = 9
            fmt_tb.FontFamily = WpfFontFamily("Consolas")
            fmt_tb.Foreground = SolidColorBrush(Color.FromArgb(255, 255, 230, 80))
            fmt_tb.VerticalAlignment = VerticalAlignment.Center
            fmt_tb.Cursor = Cursors.Hand
            fmt_tb.ToolTip = "Click to cycle format  |  {}".format(all_labels_str)
            fmt_tb.Margin = Thickness(2, 0, 4, 0)

            def _make_cycle(tok, fopts, ftb, cb):
                def _on_click(s, e):
                    codes = [c for c, _ in fopts]
                    cur = tok.get("read_format", codes[0])
                    try: idx = codes.index(cur)
                    except ValueError: idx = 0
                    tok["read_format"] = codes[(idx + 1) % len(codes)]
                    for code, flbl in fopts:
                        if code == tok["read_format"]:
                            ftb.Text = "[{}]".format(flbl)
                            break
                    if cb: cb()
                    e.Handled = True
                return _on_click

            fmt_tb.MouseLeftButtonUp += _make_cycle(token, fmt_options, fmt_tb, refresh_cb)
            inner.Children.Add(fmt_tb)

        # x remove
        xbtn = Button()
        xbtn.Content = "x"
        xbtn.Tag = token
        xbtn.FontSize = 9
        xbtn.FontWeight = FontWeights.Bold
        xbtn.Foreground = Brushes.White
        xbtn.Background = Brushes.Transparent
        xbtn.BorderThickness = Thickness(0)
        xbtn.Cursor = Cursors.Hand
        xbtn.Padding = Thickness(2, 0, 5, 1)
        xbtn.VerticalAlignment = VerticalAlignment.Center
        xbtn.ToolTip = "Remove"
        inner.Children.Add(xbtn)
        bdr.Height = 26
    else:
        # Palette chip — compact, just label
        lbl = TextBlock()
        lbl.Text = token.get("label", "?")
        lbl.FontFamily = WpfFontFamily("Consolas")
        lbl.FontSize = 11
        lbl.Foreground = Brushes.White
        lbl.VerticalAlignment = VerticalAlignment.Center
        lbl.Margin = Thickness(6, 0, 6, 0)
        inner.Children.Add(lbl)
        bdr.Height = 22

    bdr.Child = inner
    return bdr

# ==========================================================
# OUTPUT SLOT EDITOR
# One of these per output parameter in the profile.
# ==========================================================

class OutputSlotEditor(object):
    """
    A self-contained UI block for one output parameter formula.
    Owns: target param TextBox, separator TextBox, seq_width TextBox,
          slot drop zone (WrapPanel), and a remove-self Button.
    """

    def __init__(self, on_change_cb, on_remove_cb, output_data=None,
                 doc=None, preview_elem_cb=None):
        """
        on_change_cb:    callable() — called whenever formula changes
        on_remove_cb:    callable(self) — called when user clicks Remove
        doc:             Revit Document (for live preview resolution)
        preview_elem_cb: callable() -> Element|None — returns current preview element
        output_data:     optional dict to pre-populate from
        """
        self._on_change     = on_change_cb
        self._on_remove     = on_remove_cb
        self._doc           = doc
        self._preview_elem_cb = preview_elem_cb
        self.slots = []

        if output_data:
            self.slots = []
            seen_ask_keys = set()
            for s in output_data.get("slots", []):
                slot = dict(s)
                if "_slot_id" not in slot:
                    slot["_slot_id"] = id(slot)
                # Fix duplicate ASK: keys from old profiles — give each a unique key
                key = slot.get("key", "")
                if key.startswith("ASK:"):
                    base_label = key[len("ASK:"):]
                    label = base_label
                    counter = 2
                    while "ASK:{}".format(label) in seen_ask_keys:
                        label = "{}{}".format(base_label, counter)
                        counter += 1
                    new_key = "ASK:{}".format(label)
                    if new_key != key:
                        slot["key"]   = new_key
                        slot["label"] = "?{}".format(label)
                        slot["desc"]  = "Ask: {}".format(label)
                    seen_ask_keys.add(new_key)
                self.slots.append(slot)

        self._build()

        if output_data:
            self._loading = True
            self.tb_target.Text   = output_data.get("target_param", "")
            self.tb_label.Text    = output_data.get("label", "")
            self.tb_sep.Text      = output_data.get("sep", "-")
            self.tb_seq_w.Text    = str(output_data.get("seq_width", 3))
            self._loading = False
            self._refresh_slots()

    def _build(self):
        # Outer card border
        outer = Border()
        outer.BorderBrush = SolidColorBrush(_hex_to_color("#CCCCCC"))
        outer.BorderThickness = Thickness(1)
        outer.CornerRadius = CornerRadius(4)
        outer.Margin = Thickness(0, 0, 0, 8)
        outer.Padding = Thickness(8, 6, 8, 6)
        outer.Background = SolidColorBrush(_hex_to_color("#FAFAFA"))

        sp = StackPanel()
        outer.Child = sp

        # Header row: label + target param + remove button
        hdr = Grid()
        hdr.Margin = Thickness(0, 0, 0, 4)
        c0 = ColumnDefinition()
        c1 = ColumnDefinition()
        c1.Width = GridLength(1, GridUnitType.Star)
        c2 = ColumnDefinition()
        c2.Width = GridLength.Auto
        hdr.ColumnDefinitions.Add(c0)
        hdr.ColumnDefinitions.Add(c1)
        hdr.ColumnDefinitions.Add(c2)

        lbl_name = TextBlock()
        lbl_name.Text = "Label:"
        lbl_name.VerticalAlignment = VerticalAlignment.Center
        lbl_name.Margin = Thickness(0, 0, 4, 0)
        Grid.SetColumn(lbl_name, 0)

        self.tb_label = TextBox()
        self.tb_label.Height = 24
        self.tb_label.Padding = Thickness(4, 0, 4, 0)
        self.tb_label.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_label.FontFamily = WpfFontFamily("Consolas")
        self.tb_label.ToolTip = "Friendly label for this output (display only)"
        self.tb_label.TextChanged += self._changed
        Grid.SetColumn(self.tb_label, 1)

        btn_remove = Button()
        btn_remove.Content = "Remove"
        btn_remove.Height = 24
        btn_remove.Padding = Thickness(8, 0, 8, 0)
        btn_remove.Margin = Thickness(6, 0, 0, 0)
        btn_remove.VerticalAlignment = VerticalAlignment.Center
        btn_remove.Click += self._on_remove_click
        Grid.SetColumn(btn_remove, 2)

        hdr.Children.Add(lbl_name)
        hdr.Children.Add(self.tb_label)
        hdr.Children.Add(btn_remove)
        sp.Children.Add(hdr)

        # Target param row
        tp_row = StackPanel()
        tp_row.Orientation = Orientation.Horizontal
        tp_row.Margin = Thickness(0, 0, 0, 4)
        lbl_tp = TextBlock()
        lbl_tp.Text = "Write to param:"
        lbl_tp.VerticalAlignment = VerticalAlignment.Center
        lbl_tp.Margin = Thickness(0, 0, 6, 0)
        self.tb_target = ComboBox()
        self.tb_target.IsEditable = True
        self.tb_target.IsTextSearchEnabled = True
        self.tb_target.Width = 260
        self.tb_target.Height = 24
        self.tb_target.FontFamily = WpfFontFamily("Consolas")
        self.tb_target.ToolTip = (
            "Parameter name to write the result into.\n"
            "Type or pick from the list — scan an element to populate."
        )
        self.tb_target.SelectionChanged += self._changed
        # Also fire _changed on keyboard edits inside the editable box
        self.tb_target.AddHandler(
            TextBox.TextChangedEvent,
            RoutedEventHandler(self._changed)
        )
        tp_row.Children.Add(lbl_tp)
        tp_row.Children.Add(self.tb_target)
        sp.Children.Add(tp_row)

        # Sep + seq width row
        opt_row = StackPanel()
        opt_row.Orientation = Orientation.Horizontal
        opt_row.Margin = Thickness(0, 0, 0, 4)
        lbl_sep = TextBlock()
        lbl_sep.Text = "Separator:"
        lbl_sep.VerticalAlignment = VerticalAlignment.Center
        lbl_sep.Margin = Thickness(0, 0, 6, 0)
        self.tb_sep = TextBox()
        self.tb_sep.Width = 50
        self.tb_sep.Height = 24
        self.tb_sep.Padding = Thickness(4, 0, 4, 0)
        self.tb_sep.FontFamily = WpfFontFamily("Consolas")
        self.tb_sep.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_sep.Text = "-"
        self.tb_sep.TextChanged += self._changed
        lbl_sq = TextBlock()
        lbl_sq.Text = "  Seq digits:"
        lbl_sq.VerticalAlignment = VerticalAlignment.Center
        lbl_sq.Margin = Thickness(0, 0, 6, 0)
        self.tb_seq_w = TextBox()
        self.tb_seq_w.Width = 40
        self.tb_seq_w.Height = 24
        self.tb_seq_w.Padding = Thickness(4, 0, 4, 0)
        self.tb_seq_w.FontFamily = WpfFontFamily("Consolas")
        self.tb_seq_w.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_seq_w.Text = "3"
        self.tb_seq_w.TextChanged += self._changed
        opt_row.Children.Add(lbl_sep)
        opt_row.Children.Add(self.tb_sep)
        opt_row.Children.Add(lbl_sq)
        opt_row.Children.Add(self.tb_seq_w)
        sp.Children.Add(opt_row)

        # Drop zone
        drop_bdr = Border()
        drop_bdr.BorderBrush = SolidColorBrush(_hex_to_color("#BBBBBB"))
        drop_bdr.BorderThickness = Thickness(1)
        drop_bdr.CornerRadius = CornerRadius(3)
        drop_bdr.MinHeight = 34
        drop_bdr.Background = SolidColorBrush(_hex_to_color("#F2F2F2"))
        drop_bdr.Padding = Thickness(4, 2, 4, 2)
        drop_bdr.AllowDrop = True
        drop_bdr.ToolTip = "Drag parameters here to build the formula"
        drop_bdr.Drop += self._on_drop
        drop_bdr.DragOver += self._on_dragover
        self._drop_bdr = drop_bdr

        self._ic_slots = WrapPanel()
        self._ic_slots.Margin = Thickness(0)
        drop_bdr.Child = self._ic_slots
        sp.Children.Add(drop_bdr)

        # Static text quick-add row
        static_row = StackPanel()
        static_row.Orientation = Orientation.Horizontal
        static_row.Margin = Thickness(0, 4, 0, 4)
        lbl_st = TextBlock()
        lbl_st.Text = "Add text:"
        lbl_st.VerticalAlignment = VerticalAlignment.Center
        lbl_st.Margin = Thickness(0, 0, 6, 0)
        lbl_st.FontSize = 11
        lbl_st.Foreground = Brushes.DimGray
        self._tb_static = TextBox()
        self._tb_static.Width = 120
        self._tb_static.Height = 24
        self._tb_static.Padding = Thickness(4, 0, 4, 0)
        self._tb_static.FontFamily = WpfFontFamily("Consolas")
        self._tb_static.VerticalContentAlignment = VerticalAlignment.Center
        self._tb_static.ToolTip = "Type literal text to insert into the formula (e.g. 'H' for a prefix)"
        btn_add_st = Button()
        btn_add_st.Content = "Add"
        btn_add_st.Height = 24
        btn_add_st.Padding = Thickness(8, 0, 8, 0)
        btn_add_st.Margin = Thickness(4, 0, 0, 0)
        btn_add_st.ToolTip = "Insert this text as a fixed value in the formula"
        btn_add_st.Click += self._on_add_static
        static_row.Children.Add(lbl_st)
        static_row.Children.Add(self._tb_static)
        static_row.Children.Add(btn_add_st)
        sp.Children.Add(static_row)

        # Preview row
        prev_row = StackPanel()
        prev_row.Orientation = Orientation.Horizontal
        prev_row.Margin = Thickness(0, 2, 0, 0)
        lbl_prev = TextBlock()
        lbl_prev.Text = "Preview: "
        lbl_prev.FontWeight = FontWeights.SemiBold
        lbl_prev.VerticalAlignment = VerticalAlignment.Center
        self.tb_preview = TextBlock()
        self.tb_preview.FontFamily = WpfFontFamily("Consolas")
        self.tb_preview.TextWrapping = TextWrapping.Wrap
        self.tb_preview.VerticalAlignment = VerticalAlignment.Center
        prev_row.Children.Add(lbl_prev)
        prev_row.Children.Add(self.tb_preview)
        sp.Children.Add(prev_row)

        # Wire button clicks — ClickEvent bubbles up through WrapPanel to drop_bdr
        drop_bdr.AddHandler(
            Button.ClickEvent,
            RoutedEventHandler(self._on_slot_remove_click)
        )
        # Wire drag-to-reorder using the standard WPF two-event pattern:
        #   MouseDown  -> record candidate element + start position
        #   MouseMove  -> if threshold crossed while button held, fire DoDragDrop
        #   MouseUp    -> clear drag state
        self._drag_start_pos = None
        self._drag_source_el = None
        self._ic_slots.PreviewMouseLeftButtonDown += self._on_slot_drag_start
        self._ic_slots.PreviewMouseMove           += self._on_slot_drag_move
        self._ic_slots.PreviewMouseLeftButtonUp   += self._on_slot_drag_cancel

        self.root = outer
        self._loading = False

    def _safe_seq_width(self):
        try:
            return max(1, int(self.tb_seq_w.Text))
        except Exception:
            return 3

    def _refresh_slots(self):
        self._ic_slots.Children.Clear()
        for slot in self.slots:
            chip = make_chip(slot, removable=True, refresh_cb=self._after_format_change)
            self._ic_slots.Children.Add(chip)
        self._refresh_preview()


    def _after_format_change(self):
        # Called by format badge cycle -- rebuild chips then notify parent window
        self._refresh_slots()
        if self._on_change:
            self._on_change()
    def _refresh_preview(self):
        sep = self.tb_sep.Text
        sw  = self._safe_seq_width()
        # Try live element preview first
        elem = None
        if self._preview_elem_cb:
            try:
                elem = self._preview_elem_cb()
            except Exception:
                elem = None
        if elem is not None and self._doc is not None:
            try:
                parts = []
                for slot in self.slots:
                    val = resolve_slot_value(slot, elem, self._doc,
                                            seq_number=1, seq_width=sw)
                    parts.append(val if val else "?")
                self.tb_preview.Text = sep.join(parts)
                return
            except Exception:
                pass
        # Fallback to sample-based preview
        self.tb_preview.Text = preview_from_slots(self.slots, sep, sw)

    def _changed(self, sender, e):
        if self._loading:
            return
        self._refresh_preview()
        self._on_change()

    def _on_remove_click(self, sender, e):
        self._on_remove(self)

    def _on_dragover(self, sender, e):
        if e.Data.GetDataPresent("ParamToken"):
            e.Effects = DragDropEffects.Copy
        elif e.Data.GetDataPresent("SlotMove"):
            e.Effects = DragDropEffects.Move
        else:
            e.Effects = DragDropEffects.None_
        e.Handled = True

    def _on_drop(self, sender, e):
        # ── Slot reorder (drag from within this drop zone) ───────────
        if e.Data.GetDataPresent("SlotMove"):
            dragged = e.Data.GetData("SlotMove")
            sid = dragged.get("_slot_id")
            if sid is None:
                return
            # Find source index
            src_idx = None
            for i, s in enumerate(self.slots):
                if s.get("_slot_id") == sid:
                    src_idx = i
                    break
            if src_idx is None:
                return
            # Determine drop position from mouse X relative to slot chips
            drop_idx = len(self.slots)  # default: append at end
            try:
                drop_pos = e.GetPosition(self._ic_slots)
                for i in range(self._ic_slots.Children.Count):
                    chip = self._ic_slots.Children[i]
                    chip_pos = chip.TranslatePoint(
                        System.Windows.Point(0, 0), self._ic_slots)
                    chip_width = chip.ActualWidth
                    if drop_pos.X < chip_pos.X + chip_width / 2.0:
                        drop_idx = i
                        break
            except Exception:
                drop_idx = len(self.slots)
            # Move: remove from src, insert at dst
            slot = self.slots.pop(src_idx)
            if drop_idx > src_idx:
                drop_idx -= 1
            self.slots.insert(min(drop_idx, len(self.slots)), slot)
            self._refresh_slots()
            self._on_change()
            return

        # ── New token drop from palette ───────────────────────────────
        if not e.Data.GetDataPresent("ParamToken"):
            return
        token = e.Data.GetData("ParamToken")
        if token.get("key") == "SEQUENCE":
            if any(s.get("key") == "SEQUENCE" for s in self.slots):
                return
        new_slot = dict(token)
        new_slot["_slot_id"] = id(new_slot)
        self.slots.append(new_slot)
        self._refresh_slots()
        self._on_change()

    def _find_chip_border(self, hit_el):
        """Walk up the visual tree from hit_el to the chip Border.
        Returns the Border whose Tag is a slot dict, or None.
        Stops at Buttons and format-badge TextBlocks (those have a Tag dict).
        """
        el = hit_el
        while el is not None:
            if el is self._ic_slots:
                return None
            if isinstance(el, Button):
                return None
            if isinstance(el, TextBlock):
                tag = getattr(el, 'Tag', None)
                if isinstance(tag, dict) and "key" in tag:
                    return None  # format badge -- don't drag
                el = getattr(el, 'Parent', None)
                continue
            if isinstance(el, Border):
                tag = getattr(el, 'Tag', None)
                if isinstance(tag, dict) and "key" in tag:
                    return el
            el = getattr(el, 'Parent', None)
        return None

    def _on_slot_drag_start(self, sender, e):
        """Record candidate chip and mouse-down position.
        DoDragDrop fires in _on_slot_drag_move once threshold is crossed.
        """
        self._drag_start_pos = None
        self._drag_source_el = None
        pos = e.GetPosition(self._ic_slots)
        hit = self._ic_slots.InputHitTest(pos)
        chip = self._find_chip_border(hit)
        if chip is not None:
            self._drag_start_pos = pos
            self._drag_source_el = chip
            self._ic_slots.CaptureMouse()

    def _on_slot_drag_move(self, sender, e):
        """Fire DoDragDrop once cursor has moved past the WPF drag threshold."""
        if self._drag_source_el is None or self._drag_start_pos is None:
            return
        if e.LeftButton != System.Windows.Input.MouseButtonState.Pressed:
            self._drag_start_pos = None
            self._drag_source_el = None
            return
        cur = e.GetPosition(self._ic_slots)
        dx = abs(cur.X - self._drag_start_pos.X)
        dy = abs(cur.Y - self._drag_start_pos.Y)
        if (dx > SystemParameters.MinimumHorizontalDragDistance or
                dy > SystemParameters.MinimumVerticalDragDistance):
            chip = self._drag_source_el
            self._drag_start_pos = None
            self._drag_source_el = None
            try:
                self._ic_slots.ReleaseMouseCapture()
            except Exception:
                pass
            data_obj = DataObject("SlotMove", chip.Tag)
            from System.Windows import DragDrop
            DragDrop.DoDragDrop(chip, data_obj, DragDropEffects.Move)

    def _on_slot_drag_cancel(self, sender, e):
        """Clear drag state on mouse-up (click without drag)."""
        self._drag_start_pos = None
        self._drag_source_el = None
        try:
            self._ic_slots.ReleaseMouseCapture()
        except Exception:
            pass

    def _on_slot_remove_click(self, sender, e):
        btn = e.OriginalSource
        if not isinstance(btn, Button):
            return
        if btn.Content != "x":
            return
        slot = getattr(btn, 'Tag', None)
        if slot is None:
            return

        # Find and remove by _slot_id (handles duplicates)
        sid = slot.get("_slot_id")
        idx = None
        if sid is not None:
            for i, s in enumerate(self.slots):
                if s.get("_slot_id") == sid:
                    idx = i
                    break
        if idx is None:
            key = slot.get("key", "")
            for i, s in enumerate(self.slots):
                if s.get("key") == key:
                    idx = i
                    break
        if idx is None:
            return
        del self.slots[idx]

        self._refresh_slots()
        self._on_change()
        e.Handled = True

    def _on_add_static(self, sender, e):
        text = self._tb_static.Text
        if not text:
            return
        slot = {
            "key":       "static:{}".format(text),
            "label":     text,
            "color":     COLOR_STATIC,
            "unit_type": "string",
            "read_format": "str",
            "_slot_id":  id(slot),
        }
        slot["_slot_id"] = id(slot)
        self.slots.append(slot)
        self._tb_static.Text = ""
        self._refresh_slots()
        self._on_change()

    def get_data(self):
        # Strip internal _slot_id before saving — it's runtime-only
        clean_slots = [{k: v for k, v in s.items() if k != "_slot_id"}
                       for s in self.slots]
        return {
            "target_param": self.tb_target.Text.strip(),
            "label":        self.tb_label.Text.strip(),
            "slots":        clean_slots,
            "sep":          self.tb_sep.Text,
            "seq_width":    self._safe_seq_width(),
        }

    def update_preview_samples(self, extra_samples):
        self.tb_preview.Text = preview_from_slots(
            self.slots, self.tb_sep.Text, self._safe_seq_width(), extra_samples
        )

    def populate_target_dropdown(self, param_info_list):
        """Populate the Write-to-param dropdown with String-storage param names.
        param_info_list: list of dicts from get_all_param_info().
        Preserves whatever the user has already typed.
        """
        current = self.tb_target.Text
        self.tb_target.Items.Clear()
        for info in param_info_list:
            if info.get("storage_type", "") == "String":
                self.tb_target.Items.Add(info["name"])
        # Restore typed/loaded value
        self.tb_target.Text = current

# ==========================================================
# MAIN WINDOW
# ==========================================================

# ==========================================================
# FILTER RULE ROW
# One UI row per filter rule in the profile editor.
# ==========================================================

RULE_TYPE_OPTIONS = [
    ("category_eq",          "Category equals"),
    ("category_contains",    "Category contains"),
    ("param_has_value",      "Has parameter"),
    ("param_value_contains", "Param value contains"),
    ("elem_name_contains",   "Element name contains"),
    ("family_name_contains", "Family name contains"),
]

class FilterRuleRow(object):
    """
    A single filter-rule row:  [Type dropdown]  [Value textbox]  ([Param textbox])  [x]
    The param textbox only shows for param_value_contains.
    """

    def __init__(self, on_change_cb, on_remove_cb, rule_data=None, doc=None):
        self._on_change = on_change_cb
        self._on_remove = on_remove_cb
        self._doc       = doc   # used to populate category list
        self._loading   = False   # set True during profile load to suppress callbacks
        self._build()
        if rule_data:
            self._loading = True
            rtype = rule_data.get("type", "category_eq")
            self._set_type(rtype)
            self._set_value(rule_data.get("value", ""))
            self.tb_param.Text = rule_data.get("param", "")
            self._loading = False
            self._update_value_control()

    def _build(self):
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.Margin = Thickness(0, 0, 0, 4)

        # Rule type dropdown
        self.cb_type = ComboBox()
        self.cb_type.Width = 190
        self.cb_type.Height = 26
        self.cb_type.Margin = Thickness(0, 0, 6, 0)
        for key, label in RULE_TYPE_OPTIONS:
            item = ComboBoxItem()
            item.Content = label
            item.Tag = key
            self.cb_type.Items.Add(item)
        self.cb_type.SelectedIndex = 0
        # SelectionChanged wired after all controls are built below

        # Param name box (only shown for param_value_contains)
        self.tb_param = TextBox()
        self.tb_param.Width = 150
        self.tb_param.Height = 26
        self.tb_param.Padding = Thickness(4, 0, 4, 0)
        self.tb_param.FontFamily = WpfFontFamily("Consolas")
        self.tb_param.VerticalContentAlignment = VerticalAlignment.Center
        # Tooltip is set dynamically by _update_value_control based on rule type
        self.tb_param.Margin = Thickness(0, 0, 4, 0)
        self.tb_param.Visibility = Visibility.Collapsed
        self.tb_param.TextChanged += self._changed

        # Category picker — editable ComboBox (shown for category rules)
        self.cb_value = ComboBox()
        self.cb_value.Width = 200
        self.cb_value.Height = 26
        self.cb_value.IsEditable = True        # user can type freely
        self.cb_value.IsTextSearchEnabled = True
        self.cb_value.FontFamily = WpfFontFamily("Consolas")
        self.cb_value.Margin = Thickness(0, 0, 6, 0)
        self.cb_value.ToolTip = "Pick a Revit category from the list, or type to search / enter a custom value"
        self.cb_value.Visibility = Visibility.Visible
        self.cb_value.SelectionChanged += self._on_cb_value_text_changed
        self._populate_category_combo()

        # Plain text value box — shown for non-category rules
        self.tb_value = TextBox()
        self.tb_value.Width = 200
        self.tb_value.Height = 26
        self.tb_value.Padding = Thickness(4, 0, 4, 0)
        self.tb_value.FontFamily = WpfFontFamily("Consolas")
        self.tb_value.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_value.ToolTip = "Value to match (case-insensitive)"
        self.tb_value.Margin = Thickness(0, 0, 6, 0)
        self.tb_value.Visibility = Visibility.Collapsed
        self.tb_value.TextChanged += self._changed

        btn_x = Button()
        btn_x.Content = "x"
        btn_x.Width = 24
        btn_x.Height = 26
        btn_x.FontSize = 10
        btn_x.FontWeight = FontWeights.Bold
        btn_x.Padding = Thickness(0)
        btn_x.ToolTip = "Remove this rule"
        btn_x.Click += self._on_remove_click

        row.Children.Add(self.cb_type)
        row.Children.Add(self.tb_param)
        row.Children.Add(self.cb_value)
        row.Children.Add(self.tb_value)
        row.Children.Add(btn_x)
        self.root = row
        # Wire SelectionChanged now all controls exist
        self.cb_type.SelectionChanged += self._on_type_changed
        self._update_value_control()
        self._loading = False

    def _populate_category_combo(self):
        self.cb_value.Items.Clear()
        cats = []
        if self._doc:
            try:
                cats = get_revit_category_names(self._doc)
            except Exception:
                pass
        for name in cats:
            self.cb_value.Items.Add(name)

    def _get_value(self):
        """Return current value regardless of which control is active."""
        rtype = self._get_type()
        if rtype in ("category_eq", "category_contains"):
            return (self.cb_value.Text or "").strip()
        return self.tb_value.Text.strip()

    def _set_value(self, val):
        rtype = self._get_type()
        if rtype in ("category_eq", "category_contains"):
            self.cb_value.Text = val
        else:
            self.tb_value.Text = val

    def _update_value_control(self):
        """Show the right controls depending on rule type:
          category_eq/contains     → category dropdown only
          has_parameter            → param name only (no value)
          param_value_contains     → param name + value text
          elem/family name         → value text only
        """
        rtype = self._get_type()
        is_cat        = rtype in ("category_eq", "category_contains")
        needs_param   = rtype in ("param_has_value", "param_value_contains")
        needs_value   = rtype not in ("category_eq", "category_contains", "param_has_value")

        self.cb_value.Visibility = Visibility.Visible   if is_cat      else Visibility.Collapsed
        self.tb_value.Visibility = Visibility.Visible   if needs_value else Visibility.Collapsed
        self.tb_param.Visibility = Visibility.Visible   if needs_param else Visibility.Collapsed

        # Update tooltip on param box to reflect rule type
        if rtype == "param_has_value":
            self.tb_param.ToolTip = "Parameter name — element passes if this param exists and has a value"
        else:
            self.tb_param.ToolTip = "Exact parameter name to check on the element"

        # param_has_value needs no value field — existence check only
        if rtype == "param_has_value":
            self.tb_value.Visibility = Visibility.Collapsed

    def _on_cb_value_text_changed(self, sender, e):
        if not self._loading:
            self._changed(sender, e)

    def _get_type(self):
        item = self.cb_type.SelectedItem
        if item is None:
            return "category_eq"
        return item.Tag

    def _set_type(self, rtype):
        for i in range(self.cb_type.Items.Count):
            item = self.cb_type.Items.GetItemAt(i)
            if hasattr(item, 'Tag') and item.Tag == rtype:
                self.cb_type.SelectedIndex = i
                return
        self.cb_type.SelectedIndex = 0

    def _update_param_visibility(self):
        # Kept for back-compat; delegates to combined method
        self._update_value_control()

    def _on_type_changed(self, sender, e):
        self._update_value_control()
        if not self._loading:
            self._changed(sender, e)

    def _changed(self, sender, e):
        if not self._loading:
            self._on_change(sender, e)

    def _on_remove_click(self, sender, e):
        self._on_remove(self)

    def get_data(self):
        rtype = self._get_type()
        data = {
            "type":  rtype,
            "value": self._get_value() if rtype not in ("param_has_value",) else "",
        }
        if rtype in ("param_has_value", "param_value_contains"):
            data["param"] = self.tb_param.Text.strip()
        return data

# ==========================================================
# CHECKBOX SELECTION WINDOW
# Used for selective export/import of filter rules and outputs.
# ==========================================================

class _CheckboxSelectWindow(Window):
    """
    A simple modal window showing grouped checkboxes.
    sections: list of (section_name, [item_label, ...])
    After ShowDialog(), check .confirmed and .is_checked(section, index).
    """

    def __init__(self, title, sections):
        self.confirmed = False
        self._checks = {}   # (section_name, idx) -> CheckBox

        self.Title = title
        self.Width = 480
        self.SizeToContent = SizeToContent.Height
        self.ShowInTaskbar = False
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        outer = StackPanel(); outer.Margin = Thickness(14, 12, 14, 12)
        self.Content = outer

        for section_name, items in sections:
            if not items:
                continue
            hdr = TextBlock(); hdr.Text = section_name; hdr.FontWeight = FontWeights.SemiBold
            hdr.Margin = Thickness(0, 6, 0, 4)
            outer.Children.Add(hdr)
            for i, label in enumerate(items):
                cb = CheckBox()
                cb.Content = label if label else "(unnamed)"
                cb.IsChecked = True
                cb.Margin = Thickness(8, 2, 0, 2)
                cb.ToolTip = "Include this item in the export/import"
                outer.Children.Add(cb)
                self._checks[(section_name, i)] = cb

        sep = Separator(); sep.Margin = Thickness(0, 10, 0, 8)
        outer.Children.Add(sep)

        btn_row = StackPanel(); btn_row.Orientation = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right
        btn_ok = Button(); btn_ok.Content = "OK"; btn_ok.Width = 80; btn_ok.Height = 28; btn_ok.Margin = Thickness(0,0,8,0)
        btn_ok.IsDefault = True
        btn_ok.Click += self._ok
        btn_cancel = Button(); btn_cancel.Content = "Cancel"; btn_cancel.Width = 80; btn_cancel.Height = 28
        btn_cancel.IsCancel = True
        btn_cancel.Click += self._cancel
        btn_row.Children.Add(btn_ok); btn_row.Children.Add(btn_cancel)
        outer.Children.Add(btn_row)

    def _ok(self, sender, e):
        self.confirmed = True
        self.Close()

    def _cancel(self, sender, e):
        self.confirmed = False
        self.Close()

    def is_checked(self, section_name, idx):
        cb = self._checks.get((section_name, idx))
        if cb is None:
            return True   # not shown = always include
        return bool(cb.IsChecked)

class _DestinationPickerWindow(Window):
    """
    Clean destination picker replacing the Yes/No/Cancel dialog.
    Returns self.destination = 'user' | 'division' | None (cancelled).
    """
    def __init__(self, profile_count=1):
        self.destination = None
        self.Title = "Choose Save Location"
        self.Width = 380
        self.SizeToContent = SizeToContent.Height
        self.ShowInTaskbar = False
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.NoResize

        outer = StackPanel(); outer.Margin = Thickness(16, 14, 16, 14)
        self.Content = outer

        hdr = TextBlock()
        hdr.Text = "Where should {} profile(s) be saved?".format(profile_count)
        hdr.FontWeight = FontWeights.SemiBold
        hdr.Margin = Thickness(0, 0, 0, 12)
        hdr.TextWrapping = TextWrapping.Wrap
        outer.Children.Add(hdr)

        def _option_btn(title, subtitle, dest):
            bdr = Border()
            bdr.BorderBrush = SolidColorBrush(_hex_to_color("#C0CCDD"))
            bdr.BorderThickness = Thickness(1)
            bdr.CornerRadius = CornerRadius(4)
            bdr.Padding = Thickness(12, 8, 12, 8)
            bdr.Margin = Thickness(0, 0, 0, 8)
            bdr.Cursor = Cursors.Hand
            bdr.Background = SolidColorBrush(_hex_to_color("#F4F8FF"))
            sp = StackPanel()
            t1 = TextBlock(); t1.Text = title; t1.FontWeight = FontWeights.SemiBold
            t2 = TextBlock(); t2.Text = subtitle; t2.FontSize = 10
            t2.Foreground = Brushes.DimGray; t2.TextWrapping = TextWrapping.Wrap
            sp.Children.Add(t1); sp.Children.Add(t2)
            bdr.Child = sp
            def _click(s, e, d=dest):
                self.destination = d
                self.Close()
            bdr.MouseLeftButtonUp += _click
            return bdr

        outer.Children.Add(_option_btn(
            "My Profiles",
            "Saved to your AppData folder — always available, only visible to you.",
            "user"))
        outer.Children.Add(_option_btn(
            "Division Folder",
            "Saved to the shared bundle folder. Visible to all team members. Requires write access to the extension directory.",
            "division"))

        sep = Separator(); sep.Margin = Thickness(0, 4, 0, 8)
        outer.Children.Add(sep)

        btn_cancel = Button()
        btn_cancel.Content = "Cancel"
        btn_cancel.Height = 28
        btn_cancel.Width = 80
        btn_cancel.HorizontalAlignment = HorizontalAlignment.Right
        btn_cancel.IsCancel = True
        btn_cancel.Click += lambda s, e: self.Close()
        outer.Children.Add(btn_cancel)

class ParamFormulaWindow(Window):

    def __init__(self, doc, uidoc):
        self.doc      = doc
        self.uidoc    = uidoc
        self._profiles = []
        self._output_editors = []   # list of OutputSlotEditor
        self._dirty        = False
        self._loading      = False
        self._pending_clean = False   # suppresses spurious dirty after load
        self._active_editor   = None  # last output editor clicked for click-to-add
        self._preview_elem    = None  # element used for live formula preview
        self._scanned_tokens  = []  # from selection scan
        self._extra_samples   = {}
        self._scanned_param_info = []  # for populating target dropdowns
        self._filter_rule_rows = []  # list of FilterRuleRow UI objects

        self._build_ui()
        self._load_profiles()
        self._scan_selection_auto()
        self._refresh_token_palette()
        # Populate target dropdowns on all editors loaded from profile
        if self._scanned_param_info:
            for ed in self._output_editors:
                ed.populate_target_dropdown(self._scanned_param_info)
        # Auto-set preview element from current selection
        try:
            sel_elems = get_selected_elements(self.doc)
            if sel_elems:
                self._set_preview_elem(sel_elems[0])
        except Exception:
            pass
        self.Closing += self._on_closing

    # ----------------------------------------------------------
    # UI BUILD
    # ----------------------------------------------------------

    def _build_ui(self):
        self.Title              = "NW Param & Number"
        self.Width              = 960
        self.MinWidth           = 700
        self.MinHeight          = 100
        self.SizeToContent      = SizeToContent.Height   # shrinks/grows with content
        self.ShowInTaskbar      = False
        self.ResizeMode         = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen

        # ── Root: plain StackPanel — SizeToContent.Height handles sizing ──
        # No outer ScrollViewer: it would suppress SizeToContent behaviour.
        # The editor panel has its own internal ScrollViewer.
        outer = StackPanel()
        outer.Margin = Thickness(10)
        self.Content = outer

        # ══════════════════════════════════════════════════════════════════
        # QUICK-RUN BAR  (always visible)
        # ══════════════════════════════════════════════════════════════════
        qr_bdr = Border()
        qr_bdr.Background = SolidColorBrush(_hex_to_color("#F0F4FA"))
        qr_bdr.BorderBrush = SolidColorBrush(_hex_to_color("#C8D4E8"))
        qr_bdr.BorderThickness = Thickness(1)
        qr_bdr.CornerRadius = CornerRadius(4)
        qr_bdr.Padding = Thickness(10, 8, 10, 8)
        qr_bdr.Margin = Thickness(0, 0, 0, 6)
        outer.Children.Add(qr_bdr)

        qr_sp = StackPanel()
        qr_bdr.Child = qr_sp

        # Row 1: Profile picker + expand toggle
        row1 = Grid()
        qc0 = ColumnDefinition()
        qc1 = ColumnDefinition(); qc1.Width = GridLength.Auto
        row1.ColumnDefinitions.Add(qc0)
        row1.ColumnDefinitions.Add(qc1)
        row1.Margin = Thickness(0, 0, 0, 6)

        profile_sp = StackPanel(); profile_sp.Orientation = Orientation.Horizontal
        lbl_p = TextBlock(); lbl_p.Text = "Profile:"; lbl_p.VerticalAlignment = VerticalAlignment.Center; lbl_p.FontWeight = FontWeights.SemiBold; lbl_p.Margin = Thickness(0,0,8,0)
        self.cb_profile_picker = ComboBox(); self.cb_profile_picker.Height = 28; self.cb_profile_picker.MinWidth = 220
        self.cb_profile_picker.ToolTip = "Select a saved profile to use"
        self.cb_profile_picker.SelectionChanged += self._on_quick_profile_changed
        profile_sp.Children.Add(lbl_p)
        profile_sp.Children.Add(self.cb_profile_picker)
        Grid.SetColumn(profile_sp, 0)

        self.btn_expand = Button()
        self.btn_expand.Content = "Edit Profile  ▼"
        self.btn_expand.Height = 28
        self.btn_expand.Padding = Thickness(10, 0, 10, 0)
        self.btn_expand.ToolTip = "Show / hide the full profile editor"
        self.btn_expand.Click += self._on_toggle_expand
        Grid.SetColumn(self.btn_expand, 1)

        self.btn_help = Button()
        self.btn_help.Content = "?"
        self.btn_help.Width = 28
        self.btn_help.Height = 28
        self.btn_help.FontWeight = FontWeights.Bold
        self.btn_help.Margin = Thickness(6, 0, 0, 0)
        self.btn_help.ToolTip = "How to use this tool"
        self.btn_help.Click += self._on_help
        Grid.SetColumn(self.btn_help, 1)

        # put expand + help in same column via StackPanel
        right_btns = StackPanel(); right_btns.Orientation = Orientation.Horizontal
        right_btns.Children.Add(self.btn_expand)
        right_btns.Children.Add(self.btn_help)
        Grid.SetColumn(right_btns, 1)

        row1.Children.Add(profile_sp)
        row1.Children.Add(right_btns)
        qr_sp.Children.Add(row1)

        # Row 2: Source box | Mode box | Run
        row2 = StackPanel()
        row2.Orientation = Orientation.Horizontal
        row2.Margin = Thickness(0, 2, 0, 0)

        def _group_box(label_text, bg="#F0F4FA", border="#C8D4E8"):
            """Labeled pill-box that groups a header + radio options."""
            bdr = Border()
            bdr.Background       = SolidColorBrush(_hex_to_color(bg))
            bdr.BorderBrush      = SolidColorBrush(_hex_to_color(border))
            bdr.BorderThickness  = Thickness(1)
            bdr.CornerRadius     = CornerRadius(4)
            bdr.Padding          = Thickness(8, 4, 8, 4)
            bdr.Margin           = Thickness(0, 0, 8, 0)
            inner = StackPanel(); inner.Orientation = Orientation.Horizontal
            lbl = TextBlock()
            lbl.Text            = label_text
            lbl.FontWeight      = FontWeights.Bold
            lbl.VerticalAlignment = VerticalAlignment.Center
            lbl.Margin          = Thickness(0, 0, 8, 0)
            inner.Children.Add(lbl)
            bdr.Child = inner
            return bdr, inner

        # ── Source box ──────────────────────────────────────────────────
        src_bdr, src_inner = _group_box("Source:")

        self.rb_src_sel  = RadioButton()
        self.rb_src_sel.Content           = "Selection"
        self.rb_src_sel.GroupName         = "ElemSrc"
        self.rb_src_sel.IsChecked         = False
        self.rb_src_sel.Margin            = Thickness(0, 0, 8, 0)
        self.rb_src_sel.VerticalAlignment = VerticalAlignment.Center
        self.rb_src_sel.ToolTip           = "Use your current Revit selection as the input. Filter rules narrow it down to matching elements only."

        self.rb_src_pick = RadioButton()
        self.rb_src_pick.Content           = "Box Select"
        self.rb_src_pick.GroupName         = "ElemSrc"
        self.rb_src_pick.IsChecked         = True
        self.rb_src_pick.Margin            = Thickness(0, 0, 8, 0)
        self.rb_src_pick.VerticalAlignment = VerticalAlignment.Center
        self.rb_src_pick.ToolTip           = "Click elements in Revit one at a time, then press Finish (green checkmark in ribbon). Profile filter rules gray out non-matching elements."

        self.rb_src_view = RadioButton()
        self.rb_src_view.Content           = "Visible in View"
        self.rb_src_view.GroupName         = "ElemSrc"
        self.rb_src_view.VerticalAlignment = VerticalAlignment.Center
        self.rb_src_view.ToolTip           = "Collect all elements visible in the active view. Category rules pre-filter at collector level for speed on large models."

        src_inner.Children.Add(self.rb_src_pick)
        src_inner.Children.Add(self.rb_src_sel)
        src_inner.Children.Add(self.rb_src_view)
        row2.Children.Add(src_bdr)

        # ── Mode box ────────────────────────────────────────────────────
        mode_bdr, mode_inner = _group_box("Mode:")

        # Mode order: Renumber All (most common/destructive, clear intent)
        #             From Highest (safe continuation)
        #             Fill Missing (gentlest — only blank params)
        self.rb_renumber_all = RadioButton()
        self.rb_renumber_all.Content           = "Renumber All"
        self.rb_renumber_all.GroupName         = "RunMode"
        self.rb_renumber_all.IsChecked         = True
        self.rb_renumber_all.Margin            = Thickness(0, 0, 8, 0)
        self.rb_renumber_all.VerticalAlignment = VerticalAlignment.Center
        self.rb_renumber_all.ToolTip           = "Overwrite all output params on every matched element, starting from the sequence start number"

        self.rb_start_highest = RadioButton()
        self.rb_start_highest.Content           = "From Highest"
        self.rb_start_highest.GroupName         = "RunMode"
        self.rb_start_highest.Margin            = Thickness(0, 0, 8, 0)
        self.rb_start_highest.VerticalAlignment = VerticalAlignment.Center
        self.rb_start_highest.ToolTip           = "Find the highest existing number per group and continue from there — skips already-numbered"

        self.rb_fill_missing = RadioButton()
        self.rb_fill_missing.Content           = "Fill Missing"
        self.rb_fill_missing.GroupName         = "RunMode"
        self.rb_fill_missing.Margin            = Thickness(0, 0, 8, 0)
        self.rb_fill_missing.VerticalAlignment = VerticalAlignment.Center
        self.rb_fill_missing.ToolTip           = "Only write to elements that have blank output params — skips already-numbered items"

        # Start at — visible when Renumber All is selected
        lbl_start = TextBlock()
        lbl_start.Text = "Start at:"
        lbl_start.VerticalAlignment = VerticalAlignment.Center
        lbl_start.Margin = Thickness(4, 0, 4, 0)
        lbl_start.FontSize = 11

        self.tb_seq_start = TextBox()
        self.tb_seq_start.Text = "1"
        self.tb_seq_start.Width = 40
        self.tb_seq_start.Height = 26
        self.tb_seq_start.Padding = Thickness(4, 0, 4, 0)
        self.tb_seq_start.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_seq_start.FontFamily = WpfFontFamily("Consolas")
        self.tb_seq_start.ToolTip = "Starting sequence number for Renumber All mode. Type 10 to start at 010, 100 to start at 100, etc."

        self._lbl_start_at = lbl_start
        self._tb_seq_start_container = StackPanel()
        self._tb_seq_start_container.Orientation = Orientation.Horizontal
        self._tb_seq_start_container.Visibility = Visibility.Visible  # Renumber All is default
        self._tb_seq_start_container.Children.Add(lbl_start)
        self._tb_seq_start_container.Children.Add(self.tb_seq_start)

        def _mode_changed(s, e):
            self._tb_seq_start_container.Visibility = (
                Visibility.Visible if self.rb_renumber_all.IsChecked
                else Visibility.Collapsed)

        self.rb_renumber_all.Checked  += _mode_changed
        self.rb_start_highest.Checked += _mode_changed
        self.rb_fill_missing.Checked  += _mode_changed

        mode_inner.Children.Add(self.rb_renumber_all)
        mode_inner.Children.Add(self.rb_start_highest)
        mode_inner.Children.Add(self.rb_fill_missing)
        mode_inner.Children.Add(self._tb_seq_start_container)
        row2.Children.Add(mode_bdr)

        # ── Test Filter + Run ────────────────────────────────────────────
        self.btn_run = Button()
        self.btn_run.Content    = "▶  Run"
        self.btn_run.Height     = 30
        self.btn_run.Padding    = Thickness(14, 0, 14, 0)
        self.btn_run.FontWeight = FontWeights.Bold
        self.btn_run.ToolTip    = "Run the selected profile — collect, filter, and write formula values to matched elements"
        self.btn_run.Click     += self._on_run

        row2.Children.Add(self.btn_run)
        qr_sp.Children.Add(row2)

        # Notes bar — shows profile notes under the run bar
        self.lbl_notes_bar = TextBlock()
        self.lbl_notes_bar.Foreground = Brushes.DimGray
        self.lbl_notes_bar.FontSize = 11
        self.lbl_notes_bar.TextWrapping = TextWrapping.Wrap
        self.lbl_notes_bar.Margin = Thickness(2, 4, 0, 0)
        self.lbl_notes_bar.Visibility = Visibility.Collapsed
        outer.Children.Add(self.lbl_notes_bar)

        # ══════════════════════════════════════════════════════════════════
        # EDITOR PANEL  (collapsible)
        # ══════════════════════════════════════════════════════════════════
        self._editor_bdr = Border()
        self._editor_bdr.BorderBrush = SolidColorBrush(_hex_to_color("#CCCCCC"))
        self._editor_bdr.BorderThickness = Thickness(1)
        self._editor_bdr.CornerRadius = CornerRadius(4)
        self._editor_bdr.Padding = Thickness(10, 8, 10, 8)
        self._editor_bdr.Visibility = Visibility.Collapsed  # hidden by default
        outer.Children.Add(self._editor_bdr)

        # Cap editor height so it never pushes off screen
        # Use 75% of working area height minus the quick-run bar (~120px)
        try:
            screen_h = SystemParameters.WorkArea.Height
            max_editor_h = max(300, screen_h * 0.75 - 120)
        except Exception:
            max_editor_h = 500

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.MaxHeight = max_editor_h
        self._editor_bdr.Child = sv

        sp = StackPanel()
        self._editor_sp = sp
        sv.Content = sp

        # ── Left/Right split: Profile list | tabs ───────────────────────
        editor_grid = Grid()
        editor_grid.MaxHeight = max_editor_h - 20
        eg_c0 = ColumnDefinition(); eg_c0.Width = GridLength(220); eg_c0.MinWidth = 160
        eg_c1 = ColumnDefinition(); eg_c1.Width = GridLength(4)
        eg_c2 = ColumnDefinition(); eg_c2.Width = GridLength(1, GridUnitType.Star)
        editor_grid.ColumnDefinitions.Add(eg_c0)
        editor_grid.ColumnDefinitions.Add(eg_c1)
        editor_grid.ColumnDefinitions.Add(eg_c2)
        sp.Children.Add(editor_grid)

        # ── LEFT: Profile management ─────────────────────────────────────
        left_dp = DockPanel()
        left_dp.Margin = Thickness(0, 0, 4, 0)
        left_dp.LastChildFill = True
        Grid.SetColumn(left_dp, 0)

        lbl_profiles = TextBlock()
        lbl_profiles.Text = "Profiles"
        lbl_profiles.FontWeight = FontWeights.SemiBold
        lbl_profiles.Margin = Thickness(0, 0, 0, 4)
        DockPanel.SetDock(lbl_profiles, Dock.Top)
        left_dp.Children.Add(lbl_profiles)

        # ── Action stack docked to bottom ─────────────────────────────────
        def _lb(label):
            t = TextBlock()
            t.Text = label
            t.FontSize = 10
            t.Foreground = Brushes.DimGray
            t.Margin = Thickness(0, 6, 0, 2)
            return t

        def _full_btn(text, tip):
            b = Button()
            b.Content = text
            b.Height = 26
            b.Padding = Thickness(6, 0, 6, 0)
            b.Margin = Thickness(0, 0, 0, 2)
            b.HorizontalAlignment = HorizontalAlignment.Stretch
            b.ToolTip = tip
            return b

        # Action buttons — vertical stack in left panel
        action_sp = StackPanel()
        action_sp.Margin = Thickness(0, 4, 0, 0)
        DockPanel.SetDock(action_sp, Dock.Bottom)

        def _lb(label):
            t = TextBlock(); t.Text = label; t.FontSize = 10
            t.Foreground = Brushes.DimGray
            t.Margin = Thickness(0, 6, 0, 2)
            return t

        def _full_btn(text, tip):
            b = Button(); b.Content = text; b.Height = 26
            b.Padding = Thickness(6, 0, 6, 0)
            b.Margin = Thickness(0, 0, 0, 2)
            b.HorizontalAlignment = HorizontalAlignment.Stretch
            b.ToolTip = tip
            return b

        action_sp.Children.Add(_lb("Save to:"))
        self.btn_save_user = _full_btn("Save to My Profiles",  "Save to your personal AppData folder — always works")
        self.btn_save_div  = _full_btn("Save to Division",     "Save to shared bundle folder — requires write access")
        action_sp.Children.Add(self.btn_save_user)
        action_sp.Children.Add(self.btn_save_div)

        s1 = Separator(); s1.Margin = Thickness(0, 4, 0, 0)
        action_sp.Children.Add(s1)
        action_sp.Children.Add(_lb("Profile:"))
        self.btn_new    = _full_btn("New Profile",       "Create a blank new profile")
        self.btn_dupe   = _full_btn("Duplicate Profile", "Copy selected profile as a starting point")
        self.btn_delete = _full_btn("Delete Profile",    "Delete selected profile")
        action_sp.Children.Add(self.btn_new)
        action_sp.Children.Add(self.btn_dupe)
        action_sp.Children.Add(self.btn_delete)

        s2 = Separator(); s2.Margin = Thickness(0, 4, 0, 0)
        action_sp.Children.Add(s2)
        action_sp.Children.Add(_lb("Transfer:"))
        self.btn_import = _full_btn("Import Profile...", "Import profiles from a JSON bundle file")
        self.btn_export = _full_btn("Export Profile...", "Export profiles to a JSON bundle file")
        action_sp.Children.Add(self.btn_import)
        action_sp.Children.Add(self.btn_export)

        left_dp.Children.Add(action_sp)

        # Search box — docked Top so it appears ABOVE the list
        self.tb_profile_search = TextBox()
        self.tb_profile_search.Height = 26
        self.tb_profile_search.Padding = Thickness(4, 0, 4, 0)
        self.tb_profile_search.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_profile_search.ToolTip = "Type to filter profiles by name"
        self.tb_profile_search.Margin = Thickness(0, 4, 0, 2)
        self.tb_profile_search.TextChanged += self._on_profile_search
        DockPanel.SetDock(self.tb_profile_search, Dock.Top)
        left_dp.Children.Add(self.tb_profile_search)

        # Profile list — fills remaining space (LastChildFill=True default)
        self.lb_profiles = ListBox()
        self.lb_profiles.FontFamily = WpfFontFamily("Consolas")
        self.lb_profiles.MinHeight = 60
        left_dp.Children.Add(self.lb_profiles)
        editor_grid.Children.Add(left_dp)

        # ── RIGHT: Profile editor ─────────────────────────────────────────
        # ── RIGHT: Two-tab layout ─────────────────────────────────────────
        spl = GridSplitter()
        spl.HorizontalAlignment = HorizontalAlignment.Stretch
        spl.Width = 4
        Grid.SetColumn(spl, 1)
        editor_grid.Children.Add(spl)

        tab_ctrl = TabControl()
        tab_ctrl.Margin = Thickness(4, 0, 0, 0)
        tab_ctrl.MaxHeight = max_editor_h - 20
        Grid.SetColumn(tab_ctrl, 2)
        editor_grid.Children.Add(tab_ctrl)

        def _make_tab(title):
            ti = TabItem()
            ti.Header = title
            ti.FontWeight = FontWeights.SemiBold
            sv = ScrollViewer()
            sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            sp = StackPanel()
            sp.Margin = Thickness(4, 6, 4, 6)
            sv.Content = sp
            ti.Content = sv
            tab_ctrl.Items.Add(ti)
            return sp

        # Tab 1: Profile settings + Element Filter
        rsp1 = _make_tab("Profile & Filter")
        # Tab 2: Output Formulas + Token Palette
        rsp2 = _make_tab("Outputs & Parameters")

        self._right_sp = rsp1   # kept for compatibility

        rsp = rsp1

        # ── SECTION: Profile ────────────────────────────────────────────────
        sec_profile, csp_profile = self._collapsible_section(
            "Profile", start_expanded=True)
        rsp.Children.Add(sec_profile)

        csp_profile.Children.Add(self._section("Name", top=0))
        self.tb_profile_name = self._textbox()
        self.tb_profile_name.ToolTip = "Display name shown in the profile picker dropdown"
        csp_profile.Children.Add(self.tb_profile_name)

        csp_profile.Children.Add(self._section("Notes", top=6))
        self.tb_notes = TextBox()
        self.tb_notes.Height = 48
        self.tb_notes.Padding = Thickness(4, 4, 4, 4)
        self.tb_notes.AcceptsReturn = True
        self.tb_notes.TextWrapping = TextWrapping.Wrap
        self.tb_notes.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.tb_notes.ToolTip = "Optional notes — shown under the run bar when this profile is selected"
        self.tb_notes.TextChanged += self._on_notes_changed
        csp_profile.Children.Add(self.tb_notes)

        # ── SECTION: Element Filter ──────────────────────────────────────────
        sec_filter, csp_filter = self._collapsible_section(
            "Element Filter",
            hint_text="Rules that narrow which elements get processed. Elements must pass these rules to be written.",
            start_expanded=True)
        rsp.Children.Add(sec_filter)

        logic_row = StackPanel(); logic_row.Orientation = Orientation.Horizontal; logic_row.Margin = Thickness(0,0,0,6)
        lbl_lg = TextBlock(); lbl_lg.Text = "Match:"; lbl_lg.VerticalAlignment = VerticalAlignment.Center; lbl_lg.Margin = Thickness(0,0,8,0)
        self.rb_and = RadioButton(); self.rb_and.Content = "AND  (all rules)"; self.rb_and.GroupName = "FilterLogic"; self.rb_and.IsChecked = True; self.rb_and.Margin = Thickness(0,0,16,0)
        self.rb_and.ToolTip = "Element must match every rule to pass"
        self.rb_or = RadioButton(); self.rb_or.Content = "OR  (any rule)"; self.rb_or.GroupName = "FilterLogic"
        self.rb_or.ToolTip = "Element passes if it matches at least one rule"
        logic_row.Children.Add(lbl_lg); logic_row.Children.Add(self.rb_and); logic_row.Children.Add(self.rb_or)
        csp_filter.Children.Add(logic_row)

        self._rules_sp = StackPanel(); self._rules_sp.Margin = Thickness(0,0,0,4)
        csp_filter.Children.Add(self._rules_sp)

        btn_add_rule = Button(); btn_add_rule.Content = "+ Add Rule"; btn_add_rule.Height = 24
        btn_add_rule.HorizontalAlignment = HorizontalAlignment.Left
        btn_add_rule.Padding = Thickness(8,0,8,0); btn_add_rule.Margin = Thickness(0,0,0,4)
        btn_add_rule.ToolTip = "Add a filter rule — category, param value, element name, or family name"
        btn_add_rule.Click += self._on_add_filter_rule
        csp_filter.Children.Add(btn_add_rule)

        grp_row = StackPanel(); grp_row.Orientation = Orientation.Horizontal; grp_row.Margin = Thickness(0,4,0,0)
        lbl_gp = TextBlock(); lbl_gp.Text = "Group by param:"; lbl_gp.VerticalAlignment = VerticalAlignment.Center; lbl_gp.Margin = Thickness(0,0,6,0)
        self.tb_group_by = self._textbox(width=220, mono=True)
        self.tb_group_by.ToolTip = "Group elements by this param before sequencing. Each group gets its own counter. Leave blank = one group."
        grp_row.Children.Add(lbl_gp); grp_row.Children.Add(self.tb_group_by)
        csp_filter.Children.Add(grp_row)

        # Nested components setting
        nested_row = StackPanel()
        nested_row.Orientation = Orientation.Horizontal
        nested_row.Margin = Thickness(0, 6, 0, 0)
        self.cb_allow_nested = CheckBox()
        self.cb_allow_nested.Content = "Include nested sub-components"
        self.cb_allow_nested.IsChecked = False
        self.cb_allow_nested.VerticalAlignment = VerticalAlignment.Center
        self.cb_allow_nested.ToolTip = (
            "Unchecked (default): only top-level family instances are collected. "
            "Nested geometry inside families is ignored. "
            "Checked: sub-components inside families are also included. "
            "Use when your target elements are nested inside host families.")
        self.cb_allow_nested.Click += self._on_any_change
        nested_row.Children.Add(self.cb_allow_nested)
        csp_filter.Children.Add(nested_row)

        # ── SECTION: Output Formulas ─────────────────────────────────────────
        rsp = rsp2   # Switch to Tab 2 for Outputs + Tokens

        sec_outputs, csp_outputs = self._collapsible_section(
            "Output Formulas",
            hint_text="Each card writes a formula to one parameter. Drag parameters from the palette below into the drop zone.",
            start_expanded=True)
        rsp.Children.Add(sec_outputs)

        self._outputs_sp = StackPanel()
        csp_outputs.Children.Add(self._outputs_sp)

        btn_add_out = Button(); btn_add_out.Content = "+ Add Output Parameter"; btn_add_out.Height = 26
        btn_add_out.HorizontalAlignment = HorizontalAlignment.Left
        btn_add_out.Padding = Thickness(10,0,10,0); btn_add_out.Margin = Thickness(0,0,0,4)
        btn_add_out.ToolTip = "Add a new output formula card"
        btn_add_out.Click += self._on_add_output
        csp_outputs.Children.Add(btn_add_out)

        # Preview element bar — shows which element previews are based on
        prev_bar = Border()
        prev_bar.Background = SolidColorBrush(_hex_to_color("#EEF2FA"))
        prev_bar.BorderBrush = SolidColorBrush(_hex_to_color("#C8D4E8"))
        prev_bar.BorderThickness = Thickness(1)
        prev_bar.CornerRadius = CornerRadius(3)
        prev_bar.Padding = Thickness(8, 4, 8, 4)
        prev_bar.Margin = Thickness(0, 0, 0, 4)
        prev_bar_row = StackPanel(); prev_bar_row.Orientation = Orientation.Horizontal
        lbl_prev_icon = TextBlock(); lbl_prev_icon.Text = "Preview: "
        lbl_prev_icon.FontWeight = FontWeights.SemiBold
        lbl_prev_icon.VerticalAlignment = VerticalAlignment.Center
        lbl_prev_icon.Margin = Thickness(0,0,4,0)
        self.lbl_preview_elem = TextBlock()
        self.lbl_preview_elem.Text = "No element selected — using sample values"
        self.lbl_preview_elem.Foreground = Brushes.DimGray
        self.lbl_preview_elem.FontSize = 11
        self.lbl_preview_elem.VerticalAlignment = VerticalAlignment.Center
        self.lbl_preview_elem.TextWrapping = TextWrapping.NoWrap
        btn_change_prev = Button()
        btn_change_prev.Content = "Change..."
        btn_change_prev.Height = 22
        btn_change_prev.Padding = Thickness(6,0,6,0)
        btn_change_prev.Margin = Thickness(8,0,0,0)
        btn_change_prev.FontSize = 10
        btn_change_prev.ToolTip = "Pick a specific element to use as the preview reference"
        btn_change_prev.Click += self._on_change_preview_elem
        prev_bar_row.Children.Add(lbl_prev_icon)
        prev_bar_row.Children.Add(self.lbl_preview_elem)
        prev_bar_row.Children.Add(btn_change_prev)
        prev_bar.Child = prev_bar_row
        csp_outputs.Children.Add(prev_bar)

        # ── SECTION: Token Palette ────────────────────────────────────────────
        sec_tokens, csp_tokens = self._collapsible_section(
            "Parameter Palette",
            hint_text="Parameters are the building blocks of your formulas. Scan an element to populate. Double-click or drag into a drop zone.",
            start_expanded=True)
        rsp.Children.Add(sec_tokens)

        # Color legend
        legend_sp = StackPanel(); legend_sp.Orientation = Orientation.Horizontal; legend_sp.Margin = Thickness(0,0,0,6)
        def _dot(color, label):
            row = StackPanel(); row.Orientation = Orientation.Horizontal; row.Margin = Thickness(0,0,10,0)
            d = Border(); d.Width = 10; d.Height = 10; d.CornerRadius = CornerRadius(5)
            d.Background = SolidColorBrush(_hex_to_color(color)); d.Margin = Thickness(0,0,4,0); d.VerticalAlignment = VerticalAlignment.Center
            t = TextBlock(); t.Text = label; t.FontSize = 10; t.Foreground = Brushes.DimGray; t.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(d); row.Children.Add(t)
            return row
        legend_sp.Children.Add(_dot(COLOR_PROJECT,  "Project"))
        legend_sp.Children.Add(_dot(COLOR_SEQUENCE, "Sequence"))
        legend_sp.Children.Add(_dot(COLOR_ASK,      "Ask User"))
        legend_sp.Children.Add(_dot(COLOR_LENGTH,   "Length"))
        legend_sp.Children.Add(_dot("#2a8a6a",      "Area"))
        legend_sp.Children.Add(_dot("#2a6a8a",      "Volume"))
        legend_sp.Children.Add(_dot("#4a7a2a",      "Yes/No"))
        legend_sp.Children.Add(_dot("#8a4a2a",      "Force"))
        legend_sp.Children.Add(_dot(COLOR_CUSTOM,   "Other"))
        csp_tokens.Children.Add(legend_sp)

        # Scan row
        scan_row = StackPanel(); scan_row.Orientation = Orientation.Horizontal; scan_row.Margin = Thickness(0,0,0,4)
        self.tb_scan_status = TextBlock(); self.tb_scan_status.Foreground = Brushes.DimGray; self.tb_scan_status.FontSize = 11; self.tb_scan_status.VerticalAlignment = VerticalAlignment.Center; self.tb_scan_status.Margin = Thickness(0,0,8,0)
        btn_rescan = Button(); btn_rescan.Content = "Pick Element to Scan"; btn_rescan.Height = 24
        btn_rescan.Padding = Thickness(8,0,8,0)
        btn_rescan.ToolTip = "Click one element in Revit to scan its parameters into the palette"
        btn_rescan.Click += self._on_rescan
        scan_row.Children.Add(self.tb_scan_status); scan_row.Children.Add(btn_rescan)
        csp_tokens.Children.Add(scan_row)

        # ── Static text token — type + click to add ──────────────────────
        # Type any literal text, click Add to insert into the active formula.
        static_row = StackPanel()
        static_row.Orientation = Orientation.Horizontal
        static_row.Margin = Thickness(0, 0, 0, 6)

        static_lbl = TextBlock()
        static_lbl.Text = "Static text:"
        static_lbl.FontSize = 10
        static_lbl.Foreground = Brushes.DimGray
        static_lbl.VerticalAlignment = VerticalAlignment.Center
        static_lbl.Margin = Thickness(0, 0, 6, 0)

        self._tb_static_chip = TextBox()
        self._tb_static_chip.Width = 100
        self._tb_static_chip.Height = 26
        self._tb_static_chip.Padding = Thickness(4, 0, 4, 0)
        self._tb_static_chip.FontFamily = WpfFontFamily("Consolas")
        self._tb_static_chip.FontWeight = FontWeights.Bold
        self._tb_static_chip.VerticalContentAlignment = VerticalAlignment.Center
        self._tb_static_chip.ToolTip = "Type a prefix, suffix, or any fixed text (e.g. H, MECH-, .)"
        self._tb_static_chip.Text = ""

        btn_add_static = Button()
        btn_add_static.Content = "Add to Formula"
        btn_add_static.Height = 26
        btn_add_static.Padding = Thickness(8, 0, 8, 0)
        btn_add_static.Margin = Thickness(4, 0, 0, 0)
        btn_add_static.ToolTip = "Add this text as a fixed value in the last active formula"
        btn_add_static.Click += self._on_static_chip_add

        static_row.Children.Add(static_lbl)
        static_row.Children.Add(self._tb_static_chip)
        static_row.Children.Add(btn_add_static)
        csp_tokens.Children.Add(static_row)

        # Ask User row — prompts for a value at run time
        ask_row = StackPanel(); ask_row.Orientation = Orientation.Horizontal; ask_row.Margin = Thickness(0, 0, 0, 4)
        ask_lbl = TextBlock(); ask_lbl.Text = "Ask label:"; ask_lbl.VerticalAlignment = VerticalAlignment.Center; ask_lbl.Margin = Thickness(0, 0, 6, 0)
        self._tb_ask_label = TextBox()
        self._tb_ask_label.Width = 130
        self._tb_ask_label.Height = 26
        self._tb_ask_label.Padding = Thickness(4, 0, 4, 0)
        self._tb_ask_label.FontFamily = WpfFontFamily("Consolas")
        self._tb_ask_label.FontWeight = FontWeights.Bold
        self._tb_ask_label.VerticalContentAlignment = VerticalAlignment.Center
        self._tb_ask_label.ToolTip = (
            "Label shown in the prompt dialog at run time.\n"
            "SAME label = one prompt, value reused everywhere that label appears.\n"
            "DIFFERENT label = separate prompt for each unique label.\n"
            "e.g. 'Phase' and 'Zone' = two prompts; 'Phase' twice = one prompt."
        )
        self._tb_ask_label.Text = "Value"
        btn_add_ask = Button()
        btn_add_ask.Content = "Add Ask"
        btn_add_ask.Height = 26
        btn_add_ask.Padding = Thickness(8, 0, 8, 0)
        btn_add_ask.Margin = Thickness(4, 0, 0, 0)
        btn_add_ask.Background = SolidColorBrush(_hex_to_color(COLOR_ASK))
        btn_add_ask.Foreground = Brushes.White
        btn_add_ask.ToolTip = "Add a run-time user input prompt to the last active formula"
        btn_add_ask.Click += self._on_ask_chip_add
        ask_row.Children.Add(ask_lbl)
        ask_row.Children.Add(self._tb_ask_label)
        ask_row.Children.Add(btn_add_ask)
        csp_tokens.Children.Add(ask_row)

        # Search box
        self.tb_param_search = TextBox()
        self.tb_param_search.Height = 26; self.tb_param_search.Padding = Thickness(4,0,4,0)
        self.tb_param_search.VerticalContentAlignment = VerticalAlignment.Center
        self.tb_param_search.FontFamily = WpfFontFamily("Consolas")
        self.tb_param_search.ToolTip = "Type to filter the parameter list"
        self.tb_param_search.Margin = Thickness(0,0,0,4)
        self.tb_param_search.TextChanged += self._on_token_search
        csp_tokens.Children.Add(self.tb_param_search)

        # Token ListBox
        self._ic_tokens = ListBox()
        self._ic_tokens.Height = 160; self._ic_tokens.FontFamily = WpfFontFamily("Consolas")
        self._ic_tokens.FontSize = 11; self._ic_tokens.Margin = Thickness(0,0,0,4)
        self._ic_tokens.SelectionMode = System.Windows.Controls.SelectionMode.Single
        self._ic_tokens.ToolTip = "Double-click to add to the last active formula. Drag into a drop zone."
        self._ic_tokens.MouseDoubleClick += self._on_token_dblclick
        self._ic_tokens.PreviewMouseLeftButtonDown += self._on_token_dragstart
        csp_tokens.Children.Add(self._ic_tokens)

        self.lbl_param_target = TextBlock()
        self.lbl_param_target.Foreground = Brushes.DimGray
        self.lbl_param_target.FontSize = 10; self.lbl_param_target.Margin = Thickness(0,0,0,6)
        self.lbl_param_target.Text = "Double-click a parameter to add it to the first output formula"
        csp_tokens.Children.Add(self.lbl_param_target)

        # Manual param add — name + unit type dropdown
        manual_outer = StackPanel(); manual_outer.Margin = Thickness(0,4,0,0)

        manual_row1 = StackPanel(); manual_row1.Orientation = Orientation.Horizontal; manual_row1.Margin = Thickness(0,0,0,3)
        lbl_m = TextBlock(); lbl_m.Text = "Add param:"; lbl_m.VerticalAlignment = VerticalAlignment.Center; lbl_m.Margin = Thickness(0,0,6,0)
        self.tb_manual_param = self._textbox(width=180, mono=True)
        self.tb_manual_param.ToolTip = "Exact Revit parameter name to add"
        manual_row1.Children.Add(lbl_m); manual_row1.Children.Add(self.tb_manual_param)
        manual_outer.Children.Add(manual_row1)

        manual_row2 = StackPanel(); manual_row2.Orientation = Orientation.Horizontal
        lbl_ut = TextBlock(); lbl_ut.Text = "Unit type:"; lbl_ut.VerticalAlignment = VerticalAlignment.Center; lbl_ut.Margin = Thickness(0,0,6,0)
        self.cb_manual_unit = ComboBox()
        self.cb_manual_unit.Height = 26; self.cb_manual_unit.Width = 130
        self.cb_manual_unit.Margin = Thickness(0,0,6,0)
        self.cb_manual_unit.ToolTip = "How this parameter's value should be read and formatted"
        for ut_code, ut_label in [
            ("auto",    "Auto-detect"),
            ("length",  "Length"),
            ("area",    "Area"),
            ("volume",  "Volume"),
            ("flow",    "Flow rate"),
            ("angle",   "Angle"),
            ("force",   "Force / Weight"),
            ("yesno",   "Yes / No"),
            ("integer", "Integer"),
            ("string",  "Text / String"),
        ]:
            ci = ComboBoxItem(); ci.Content = ut_label; ci.Tag = ut_code
            self.cb_manual_unit.Items.Add(ci)
        self.cb_manual_unit.SelectedIndex = 0
        btn_add_manual = Button(); btn_add_manual.Content = "Add to Palette"
        btn_add_manual.Height = 26; btn_add_manual.Padding = Thickness(8,0,8,0)
        btn_add_manual.ToolTip = "Add this parameter to the palette with the selected unit type"
        btn_add_manual.Click += self._on_add_manual_token
        manual_row2.Children.Add(lbl_ut); manual_row2.Children.Add(self.cb_manual_unit)
        manual_row2.Children.Add(btn_add_manual)
        manual_outer.Children.Add(manual_row2)
        csp_tokens.Children.Add(manual_outer)

        # Wire all buttons
        self.btn_new.Click         += self._on_new_profile
        self.btn_dupe.Click        += self._on_dupe_profile
        self.btn_delete.Click      += self._on_delete_profile
        self.btn_import.Click      += self._on_import
        self.btn_export.Click      += self._on_export
        self.btn_save_user.Click   += self._on_save_user
        self.btn_save_div.Click    += self._on_save_division
        self.lb_profiles.SelectionChanged += self._on_profile_selected
        self.lb_profiles.MouseDoubleClick  += self._on_profile_rename
        self.tb_profile_name.TextChanged  += self._on_any_change

    def _on_help(self, sender, e):
        help_win = _HelpWindow()
        help_win.Owner = self
        help_win.ShowDialog()

    def _on_dupe_profile(self, sender, e):
        item = self.lb_profiles.SelectedItem
        if item is None:
            MessageBox.Show("Select a profile to duplicate.", "Nothing Selected", MessageBoxButton.OK)
            return
        p = item.Tag
        import copy as _copy
        duped = _copy.deepcopy({k: v for k, v in p.items()
                                if k not in ("readonly", "_filepath")})
        duped["name"] = "Copy of " + duped.get("name", "Profile")
        try:
            save_profile(duped, "user")
            self._load_profiles()
            # Select the new duplicate
            for i in range(self.lb_profiles.Items.Count):
                li = self.lb_profiles.Items.GetItemAt(i)
                if hasattr(li, "Tag") and li.Tag.get("name") == duped["name"]:
                    self.lb_profiles.SelectedIndex = i
                    break
        except Exception as ex:
            MessageBox.Show("Duplicate failed:\n{}".format(ex), "Error", MessageBoxButton.OK)

    def _on_profile_search(self, sender, e):
        """Filter the profile listbox as the user types."""
        query = self.tb_profile_search.Text.strip().lower()
        self.lb_profiles.Items.Clear()
        for p in self._profiles:
            if query and query not in p["name"].lower():
                continue
            prefix = "[D] " if p.get("readonly") else ""
            item = ListBoxItem()
            item.Content = prefix + p["name"]
            item.Tag = p
            item.FontFamily = WpfFontFamily("Consolas")
            self.lb_profiles.Items.Add(item)

    def _on_notes_changed(self, sender, e):
        if self._loading:
            return
        self._update_notes_bar()
        self._dirty = True

    def _update_notes_bar(self):
        """Show/hide the notes label under the quick-run bar."""
        try:
            notes = self.tb_notes.Text.strip()
            if notes:
                self.lbl_notes_bar.Text = notes
                self.lbl_notes_bar.Visibility = Visibility.Visible
            else:
                self.lbl_notes_bar.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _section(self, text, top=8):
        t = TextBlock()
        t.Text = text
        t.FontWeight = FontWeights.SemiBold
        t.Margin = Thickness(0, top, 0, 3)
        return t

    def _hint(self, text):
        t = TextBlock()
        t.Text = text
        t.Foreground = Brushes.DimGray
        t.FontSize = 11
        t.TextWrapping = TextWrapping.Wrap
        t.Margin = Thickness(0, 0, 0, 4)
        return t

    def _textbox(self, width=None, mono=False, default=""):
        tb = TextBox()
        tb.Height = 26
        tb.Padding = Thickness(4, 0, 4, 0)
        tb.VerticalContentAlignment = VerticalAlignment.Center
        if mono:
            tb.FontFamily = WpfFontFamily("Consolas")
        if width:
            tb.Width = width
        tb.Text = default
        return tb

    def _collapsible_section(self, title, hint_text=None, start_expanded=True):
        """
        Returns (outer_sp, content_sp, toggle_btn).
        outer_sp: add to parent
        content_sp: add children to this
        toggle_btn: the clickable header button
        """
        outer = StackPanel()
        outer.Margin = Thickness(0, 8, 0, 0)

        # Header row: toggle arrow + title
        hdr = Border()
        hdr.Background = SolidColorBrush(_hex_to_color("#E8EEF8"))
        hdr.BorderBrush = SolidColorBrush(_hex_to_color("#C0CCDD"))
        hdr.BorderThickness = Thickness(0, 0, 0, 1)
        hdr.Padding = Thickness(6, 4, 6, 4)
        hdr.Cursor = Cursors.Hand
        hdr.Margin = Thickness(0, 0, 0, 0)

        hdr_sp = StackPanel(); hdr_sp.Orientation = Orientation.Horizontal
        arrow = TextBlock()
        arrow.Text = "▼" if start_expanded else "▶"
        arrow.FontSize = 9
        arrow.VerticalAlignment = VerticalAlignment.Center
        arrow.Margin = Thickness(0, 0, 6, 0)
        arrow.Foreground = SolidColorBrush(_hex_to_color("#445566"))
        lbl = TextBlock()
        lbl.Text = title
        lbl.FontWeight = FontWeights.SemiBold
        lbl.FontSize = 12
        lbl.VerticalAlignment = VerticalAlignment.Center
        hdr_sp.Children.Add(arrow)
        hdr_sp.Children.Add(lbl)
        hdr.Child = hdr_sp

        # Content panel
        content_bdr = Border()
        content_bdr.Padding = Thickness(4, 6, 4, 6)
        content_bdr.Visibility = Visibility.Visible if start_expanded else Visibility.Collapsed
        content_sp = StackPanel()
        if hint_text:
            hint = self._hint(hint_text)
            content_sp.Children.Add(hint)
        content_bdr.Child = content_sp

        # Wire toggle
        def _make_toggle(arr, bdr_content):
            def _toggle(s, e):
                if bdr_content.Visibility == Visibility.Visible:
                    bdr_content.Visibility = Visibility.Collapsed
                    arr.Text = "▶"
                else:
                    bdr_content.Visibility = Visibility.Visible
                    arr.Text = "▼"
            return _toggle
        hdr.MouseLeftButtonUp += _make_toggle(arrow, content_bdr)

        outer.Children.Add(hdr)
        outer.Children.Add(content_bdr)
        return outer, content_sp

    def _on_any_change(self, sender, e):
        if self._loading:
            return
        # If a profile was just loaded, ignore the first wave of
        # change events that fire as WPF settles its bindings.
        if getattr(self, '_pending_clean', False):
            self._pending_clean = False
            self._dirty = False
            return
        self._dirty = True

    def _on_toggle_expand(self, sender, e):
        if self._editor_bdr.Visibility == Visibility.Collapsed:
            self._editor_bdr.Visibility = Visibility.Visible
            self.btn_expand.Content = "Edit Profile  ▲"
            # Let window grow to content, capped at 90% of screen height
            try:
                cap = SystemParameters.WorkArea.Height * 0.90
                self.MaxHeight = cap
            except Exception:
                self.MaxHeight = 900
            self.SizeToContent = SizeToContent.Height
        else:
            self._editor_bdr.Visibility = Visibility.Collapsed
            self.btn_expand.Content = "Edit Profile  ▼"
            # Snap back to minimal height
            self.MaxHeight = double_infinity()
            self.SizeToContent = SizeToContent.Height

    def _on_quick_profile_changed(self, sender, e):
        item = self.cb_profile_picker.SelectedItem
        if item is None:
            return
        p = item.Tag
        if p is None:
            return
        # Sync the full editor list selection too
        for i in range(self.lb_profiles.Items.Count):
            lb_item = self.lb_profiles.Items.GetItemAt(i)
            if hasattr(lb_item, 'Tag') and lb_item.Tag is p:
                self._loading = True
                self.lb_profiles.SelectedIndex = i
                self._loading = False
                break
        self._apply_profile(p)

    # ----------------------------------------------------------
    # FILTER RULES
    # ----------------------------------------------------------

    def _on_add_filter_rule(self, sender, e):
        row = FilterRuleRow(
            on_change_cb=self._on_any_change,
            on_remove_cb=self._on_rule_removed,
            doc=self.doc
        )
        self._filter_rule_rows.append(row)
        self._rules_sp.Children.Add(row.root)
        self._dirty = True

    def _on_rule_removed(self, row):
        if row in self._filter_rule_rows:
            self._filter_rule_rows.remove(row)
        if row.root in self._rules_sp.Children:
            self._rules_sp.Children.Remove(row.root)
        self._dirty = True

    def _clear_filter_rules(self):
        for row in list(self._filter_rule_rows):
            if row.root in self._rules_sp.Children:
                self._rules_sp.Children.Remove(row.root)
        self._filter_rule_rows = []

    def _get_filter_logic(self):
        return "AND" if self.rb_and.IsChecked else "OR"

    def _set_filter_logic(self, logic):
        if logic == "OR":
            self.rb_or.IsChecked = True
        else:
            self.rb_and.IsChecked = True

    def _get_filter_rules(self):
        return [row.get_data() for row in self._filter_rule_rows]

    # ----------------------------------------------------------
    # SELECTION SCAN
    # ----------------------------------------------------------

    def _build_tokens_from_param_names(self, param_names, sample_elem):
        tokens = [
            {"key": "project_number", "label": "Job #",      "color": COLOR_PROJECT,
             "desc": "ProjectInformation.Number", "unit_type": "string", "read_format": "str"},
            {"key": "project_name",   "label": "Proj Name",  "color": COLOR_PROJECT,
             "desc": "ProjectInformation.Name",   "unit_type": "string", "read_format": "str"},
            {"key": "SEQUENCE",       "label": "###",         "color": COLOR_SEQUENCE,
             "desc": "Auto-increment sequence",   "unit_type": "string", "read_format": "str"},
            {"key": "ASK:Value",      "label": "Ask User...", "color": COLOR_ASK,
             "desc": "Ask: Value",
             "unit_type": "string", "read_format": "str"},
        ]
        extras = {}

        # Use get_all_param_info for auto-detection if we have a sample element
        param_info = {}
        if sample_elem:
            for info in get_all_param_info([sample_elem]):
                param_info[info["name"]] = info

        for pname in param_names:
            info = param_info.get(pname, {})
            unit_type    = info.get("unit_type", "string")
            read_format  = info.get("default_format", "str")
            storage_type = info.get("storage_type", "String")

            # Color by unit type for quick visual identification
            color_map = {
                "length":  COLOR_LENGTH,
                "area":    "#2a8a6a",
                "volume":  "#2a6a8a",
                "flow":    "#2a5a9a",
                "angle":   "#6a2a8a",
                "force":   "#8a4a2a",
                "yesno":   "#4a7a2a",
                "integer": "#7a7a2a",
                "string":  COLOR_CUSTOM,
                "other":   COLOR_CUSTOM,
            }

            tokens.append({
                "key":          "param:{}".format(pname),
                "label":        pname[:22] + ("..." if len(pname) > 22 else ""),
                "color":        color_map.get(unit_type, COLOR_CUSTOM),
                "desc":         pname,
                "unit_type":    unit_type,
                "read_format":  read_format,
                "storage_type": storage_type,
            })
            if sample_elem:
                p = lookup_param(sample_elem, pname)
                val = _read_param_with_format(p, read_format)
                if val:
                    extras["param:{}".format(pname)] = val[:20]
        return tokens, extras

    def _scan_selection_auto(self):
        """Scan the current Revit selection on window open."""
        elements = get_selected_elements(self.doc)
        if not elements:
            self.tb_scan_status.Text = "No active selection — use 'Pick Element to Scan'."
            self._scanned_tokens     = []
            self._extra_samples      = {}
            self._scanned_param_info = []
            return
        param_names = get_all_param_names(elements)
        self._scanned_tokens, self._extra_samples = self._build_tokens_from_param_names(
            param_names, elements[0])
        self._scanned_param_info = get_all_param_info(elements)
        self.tb_scan_status.Text = "Auto-scanned {} params from {} elements.".format(
            len(param_names), len(elements))

    def _on_rescan(self, sender, e):
        """Hide window, let user pick one element, scan it, restore window."""
        self.Hide()
        try:
            elem = pick_one_element(self.uidoc)
        except Exception:
            elem = None
        self.Show()

        if elem is None:
            self.tb_scan_status.Text = "Pick cancelled — palette unchanged."
            return

        try:
            cat  = elem_category_name(elem)
            name = elem_name(elem)
        except Exception:
            cat  = ""
            name = "?"
        param_names = get_all_param_names([elem])
        self._scanned_tokens, self._extra_samples = self._build_tokens_from_param_names(
            param_names, elem)
        self._scanned_param_info = get_all_param_info([elem])
        self.tb_scan_status.Text = "Scanned {} params from: {} / {}".format(
            len(param_names), cat, name)
        self._refresh_token_palette()
        for ed in self._output_editors:
            ed.update_preview_samples(self._extra_samples)
            ed.populate_target_dropdown(self._scanned_param_info)

    def _refresh_token_palette(self, query=""):
        """Rebuild the token ListBox, optionally filtered by search query."""
        self._ic_tokens.Items.Clear()
        tokens = list(self._scanned_tokens)
        if not tokens:
            tokens = [
                {"key": "project_number", "label": "Job #",      "color": COLOR_PROJECT,  "desc": "ProjectInformation.Number"},
                {"key": "project_name",   "label": "Proj Name",  "color": COLOR_PROJECT,  "desc": "ProjectInformation.Name"},
                {"key": "SEQUENCE",       "label": "###",         "color": COLOR_SEQUENCE, "desc": "Auto-increment sequence"},
                {"key": "ASK:Value",      "label": "Ask User...", "color": COLOR_ASK,      "desc": "Ask: Value"},
            ]

        q = query.lower().strip()
        for token in tokens:
            full = token.get("desc") or token.get("key", "")
            if full.startswith("param:"):
                full = full[len("param:"):]
            label = token.get("label", "")
            # Filter by search query
            if q and q not in full.lower() and q not in label.lower():
                continue

            # Build a ListBoxItem with colored dot + full name
            item = ListBoxItem()
            item.Tag = token
            item.Padding = Thickness(4, 2, 4, 2)

            row = StackPanel(); row.Orientation = Orientation.Horizontal
            # Colored dot
            dot = Border()
            dot.Width = 10; dot.Height = 10
            dot.CornerRadius = CornerRadius(5)
            dot.Background = SolidColorBrush(_hex_to_color(token.get("color", COLOR_TEXT)))
            dot.Margin = Thickness(0, 0, 6, 0)
            dot.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(dot)

            # Full name text
            lbl = TextBlock()
            display = full if full else label
            lbl.Text = display
            lbl.FontFamily = WpfFontFamily("Consolas")
            lbl.FontSize = 11
            lbl.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(lbl)

            item.Content = row
            self._ic_tokens.Items.Add(item)

    def _on_token_search(self, sender, e):
        query = self.tb_param_search.Text
        self._refresh_token_palette(query)

    def _on_token_dblclick(self, sender, e):
        """Double-click a token in the list to add it to the active/first output editor."""
        item = self._ic_tokens.SelectedItem
        if item is None:
            return
        token = getattr(item, 'Tag', None)
        if token is None:
            return
        # Add to the last focused output editor, or first one available
        target_ed = getattr(self, '_active_editor', None)
        if target_ed is None and self._output_editors:
            target_ed = self._output_editors[0]
        if target_ed is None:
            return
        if token.get("key") == "SEQUENCE":
            if any(s.get("key") == "SEQUENCE" for s in target_ed.slots):
                return
        new_slot = dict(token)
        new_slot["_slot_id"] = id(new_slot)
        target_ed.slots.append(new_slot)
        target_ed._refresh_slots()
        target_ed._on_change()
        # Update hint label
        try:
            name = target_ed.tb_label.Text or target_ed.tb_target.Text or "formula"
            self.lbl_param_target.Text = "Added to: {}".format(name)
        except Exception:
            pass

    def _on_add_manual_token(self, sender, e):
        pname = self.tb_manual_param.Text.strip()
        if not pname:
            return
        # Get selected unit type from dropdown
        ut_code = "auto"
        try:
            sel = self.cb_manual_unit.SelectedItem
            if sel is not None:
                ut_code = sel.Tag
        except Exception:
            ut_code = "auto"

        # If auto, try to detect from the live element
        if ut_code == "auto":
            ut_code = "string"
            if self._preview_elem:
                try:
                    p = self._preview_elem.LookupParameter(pname)
                    if p:
                        ut_code = _detect_unit_type(p)
                except Exception:
                    pass

        read_format = DEFAULT_FORMAT.get(ut_code, "str")
        color_map = {
            "length": COLOR_LENGTH, "area": "#2a8a6a", "volume": "#2a6a8a",
            "flow": "#2a5a9a", "angle": "#6a2a8a", "force": "#8a4a2a",
            "yesno": "#4a7a2a", "integer": "#7a7a2a",
            "string": COLOR_CUSTOM, "other": COLOR_CUSTOM,
        }
        token = {
            "key":         "param:{}".format(pname),
            "label":       pname[:22] + ("..." if len(pname) > 22 else ""),
            "color":       color_map.get(ut_code, COLOR_CUSTOM),
            "desc":        pname,
            "unit_type":   ut_code,
            "read_format": read_format,
        }
        existing_keys = [t["key"] for t in self._scanned_tokens]
        if token["key"] not in existing_keys:
            self._scanned_tokens.append(token)
        # Sample value
        sample = "<{}>".format(pname[:10])
        if self._preview_elem:
            try:
                p = self._preview_elem.LookupParameter(pname)
                if p:
                    v = _read_param_with_format(p, read_format)
                    if v:
                        sample = v[:20]
            except Exception:
                pass
        self._extra_samples["param:{}".format(pname)] = sample
        self._refresh_token_palette()
        self.tb_manual_param.Text = ""
        # Reset to auto
        self.cb_manual_unit.SelectedIndex = 0

    # ----------------------------------------------------------
    # DRAG FROM PALETTE
    # ----------------------------------------------------------

    def _on_static_chip_add(self, sender, e):
        """Add the static text token to the last active output formula."""
        text = self._tb_static_chip.Text.strip()
        if not text:
            return
        token = {
            "key":         "static:{}".format(text),
            "label":       text,
            "color":       COLOR_STATIC,
            "desc":        "Static text: {}".format(text),
            "unit_type":   "string",
            "read_format": "str",
        }
        target_ed = getattr(self, '_active_editor', None)
        if target_ed is None and self._output_editors:
            target_ed = self._output_editors[0]
        if target_ed is None:
            MessageBox.Show(
                "No output formula exists yet. Add an output formula card first.",
                "No Formula", MessageBoxButton.OK)
            return
        new_slot = dict(token)
        new_slot["_slot_id"] = id(new_slot)
        target_ed.slots.append(new_slot)
        target_ed._refresh_slots()
        target_ed._on_change()
        try:
            name = target_ed.tb_label.Text or target_ed.tb_target.Text or "formula"
            self.lbl_param_target.Text = "Added '{}' to: {}".format(text, name)
        except Exception:
            pass

    def _on_ask_chip_add(self, sender, e):
        """Add an ASK: user-input token to the last active output formula."""
        label = self._tb_ask_label.Text.strip()
        if not label:
            label = "Value"

        target_ed = getattr(self, '_active_editor', None)
        if target_ed is None and self._output_editors:
            target_ed = self._output_editors[0]
        if target_ed is None:
            MessageBox.Show(
                "No output formula exists yet. Add an output formula card first.",
                "No Formula", MessageBoxButton.OK)
            return

        # If this label already exists in the formula, auto-suffix to keep it unique.
        # ASK:Phase already present -> try ASK:Phase2, Phase3, etc.
        # This ensures each chip fires its own prompt at run time.
        existing_keys = {s.get("key", "") for s in target_ed.slots}
        base_label = label
        counter = 2
        while "ASK:{}".format(label) in existing_keys:
            label = "{}{}".format(base_label, counter)
            counter += 1

        key = "ASK:{}".format(label)
        token = {
            "key":         key,
            "label":       "?{}".format(label),
            "color":       COLOR_ASK,
            "desc":        "Ask: {}".format(label),
            "unit_type":   "string",
            "read_format": "str",
        }
        new_slot = dict(token)
        new_slot["_slot_id"] = id(new_slot)
        target_ed.slots.append(new_slot)
        target_ed._refresh_slots()
        target_ed._on_change()
        # Update the label box to reflect the actual label used, so the
        # user can see it was renamed and can edit it if needed.
        self._tb_ask_label.Text = label
        try:
            name = target_ed.tb_label.Text or target_ed.tb_target.Text or "formula"
            self.lbl_param_target.Text = "Added '?{}' to: {}".format(label, name)
        except Exception:
            pass

    def _on_token_dragstart(self, sender, e):
        pos = e.GetPosition(self._ic_tokens)
        hit = self._ic_tokens.InputHitTest(pos)
        el = hit
        while el is not None:
            tag = getattr(el, 'Tag', None)
            if tag is not None and isinstance(tag, dict):
                data_obj = DataObject("ParamToken", tag)
                from System.Windows import DragDrop
                DragDrop.DoDragDrop(el, data_obj, DragDropEffects.Copy)
                return
            # ListBoxItem tag
            if isinstance(el, ListBoxItem):
                tag = getattr(el, 'Tag', None)
                if tag is not None and isinstance(tag, dict):
                    data_obj = DataObject("ParamToken", tag)
                    from System.Windows import DragDrop
                    DragDrop.DoDragDrop(el, data_obj, DragDropEffects.Copy)
                    return
            el = getattr(el, 'Parent', None)

    # ----------------------------------------------------------
    # OUTPUT EDITORS
    # ----------------------------------------------------------

    def _on_add_output(self, sender, e):
        self._add_output_editor(None)
        self._dirty = True

    def _add_output_editor(self, output_data):
        ed = OutputSlotEditor(
            on_change_cb=self._on_output_changed,
            on_remove_cb=self._on_output_removed,
            output_data=output_data,
            doc=self.doc,
            preview_elem_cb=lambda: self._preview_elem,
        )
        self._output_editors.append(ed)
        self._outputs_sp.Children.Add(ed.root)
        # Track which editor was last interacted with for click-to-add
        def _make_focus_handler(editor):
            def _on_focus(s, ev):
                self._active_editor = editor
                try:
                    name = editor.tb_label.Text or editor.tb_target.Text or "formula"
                    self.lbl_param_target.Text = "Double-click a parameter to add to: {}".format(name)
                except Exception:
                    pass
            return _on_focus
        ed.root.GotFocus += _make_focus_handler(ed)
        ed._drop_bdr.GotFocus += _make_focus_handler(ed)
        ed._drop_bdr.MouseLeftButtonDown += _make_focus_handler(ed)
        # If a scan has already happened, pre-populate the target dropdown now
        if getattr(self, '_scanned_param_info', None):
            ed.populate_target_dropdown(self._scanned_param_info)
        return ed

    def _on_output_changed(self):
        if not self._loading:
            self._dirty = True

    def _on_output_removed(self, editor):
        if editor in self._output_editors:
            self._output_editors.remove(editor)
        if editor.root in self._outputs_sp.Children:
            self._outputs_sp.Children.Remove(editor.root)
        self._dirty = True

    # ----------------------------------------------------------
    # PREVIEW ELEMENT
    # ----------------------------------------------------------

    def _set_preview_elem(self, elem):
        """Set the element used for live formula preview and refresh all cards."""
        self._preview_elem = elem
        if elem is None:
            self.lbl_preview_elem.Text = "No element — using sample values"
        else:
            try:
                cat  = elem_category_name(elem) or ""
                name = elem_name(elem) or "element"
                self.lbl_preview_elem.Text = "{} / {} (id {})".format(
                    cat, name, elem.Id.IntegerValue)
            except Exception:
                self.lbl_preview_elem.Text = "Element id {}".format(
                    elem.Id.IntegerValue)
        # Refresh all output editors — _refresh_slots rebuilds chips
        # AND calls _refresh_preview, keeping format badges live
        for ed in self._output_editors:
            ed._refresh_slots()

    def _on_change_preview_elem(self, sender, e):
        """Pick a specific element to use as the live preview reference."""
        self.Hide()
        try:
            elem = pick_one_element(self.uidoc)
        except Exception:
            elem = None
        self.Show()
        if elem is not None:
            self._set_preview_elem(elem)

    def _clear_output_editors(self):
        for ed in list(self._output_editors):
            if ed.root in self._outputs_sp.Children:
                self._outputs_sp.Children.Remove(ed.root)
        self._output_editors = []

    # ----------------------------------------------------------
    # PROFILE LIST
    # ----------------------------------------------------------

    def _load_profiles(self):
        self._profiles = list_profiles()

        # Rebuild listbox
        self.lb_profiles.Items.Clear()
        for p in self._profiles:
            prefix = "[D] " if p.get("readonly") else ""
            item = ListBoxItem()
            item.Content = prefix + p["name"]
            item.Tag = p
            item.FontFamily = WpfFontFamily("Consolas")
            self.lb_profiles.Items.Add(item)

        # Rebuild quick-run combo — remember current selection
        prev_name = None
        if self.cb_profile_picker.SelectedItem is not None:
            pi = self.cb_profile_picker.SelectedItem
            prev_name = getattr(pi, 'Tag', {}).get('name')

        self.cb_profile_picker.Items.Clear()
        select_idx = 0
        for i, p in enumerate(self._profiles):
            item = ComboBoxItem()
            item.Content = ("[D] " if p.get("readonly") else "") + p["name"]
            item.Tag = p
            item.FontFamily = WpfFontFamily("Consolas")
            self.cb_profile_picker.Items.Add(item)
            if p["name"] == prev_name:
                select_idx = i

        if self._profiles:
            self._loading = True
            self.cb_profile_picker.SelectedIndex = select_idx
            self.lb_profiles.SelectedIndex = select_idx
            self._loading = False
            self._apply_profile(self._profiles[select_idx])

    def _on_profile_rename(self, sender, e):
        """Double-click a profile in the list to rename it in place."""
        item = self.lb_profiles.SelectedItem
        if item is None:
            return
        p = item.Tag
        old_name = p.get("name", "")

        # Simple input dialog using a small WPF window
        dlg = Window()
        dlg.Title = "Rename Profile"
        dlg.Width = 360
        dlg.SizeToContent = SizeToContent.Height
        dlg.WindowStartupLocation = WindowStartupLocation.CenterOwner
        dlg.Owner = self
        dlg.ResizeMode = ResizeMode.NoResize
        dlg.ShowInTaskbar = False

        outer = StackPanel(); outer.Margin = Thickness(14, 12, 14, 12)
        dlg.Content = outer

        lbl = TextBlock(); lbl.Text = "New name for '{}'".format(old_name)
        lbl.TextWrapping = TextWrapping.Wrap
        lbl.Margin = Thickness(0, 0, 0, 8)
        outer.Children.Add(lbl)

        tb = TextBox(); tb.Text = old_name; tb.Height = 28
        tb.Padding = Thickness(4, 0, 4, 0)
        tb.VerticalContentAlignment = VerticalAlignment.Center
        tb.FontFamily = WpfFontFamily("Consolas")
        tb.Margin = Thickness(0, 0, 0, 10)
        tb.SelectAll()
        outer.Children.Add(tb)

        btn_row = StackPanel(); btn_row.Orientation = Orientation.Horizontal
        btn_row.HorizontalAlignment = HorizontalAlignment.Right
        btn_ok = Button(); btn_ok.Content = "Rename"; btn_ok.Width = 80; btn_ok.Height = 28
        btn_ok.Margin = Thickness(0, 0, 8, 0); btn_ok.IsDefault = True
        btn_cancel = Button(); btn_cancel.Content = "Cancel"; btn_cancel.Width = 80; btn_cancel.Height = 28
        btn_cancel.IsCancel = True
        btn_row.Children.Add(btn_ok); btn_row.Children.Add(btn_cancel)
        outer.Children.Add(btn_row)

        result = {"ok": False}

        def _ok(s, ev):
            result["ok"] = True
            dlg.Close()

        def _cancel(s, ev):
            dlg.Close()

        btn_ok.Click     += _ok
        btn_cancel.Click += _cancel

        # Focus the textbox when dialog opens
        dlg.Loaded += lambda s, ev: tb.Focus()
        dlg.ShowDialog()

        if not result["ok"]:
            return

        new_name = tb.Text.strip()
        if not new_name or new_name == old_name:
            return

        # Update the profile data and re-save to same file
        fp = p.get("_filepath", "")
        if not fp:
            MessageBox.Show(
                "Cannot rename — profile has no saved file path.\nSave it first.",
                "Not Saved", MessageBoxButton.OK)
            return

        is_division = p.get("readonly", False)
        try:
            p_data = {k: v for k, v in p.items()
                      if k not in ("readonly", "_filepath")}
            p_data["name"] = new_name
            # Write back to same file
            with open(fp, "w") as f:
                json.dump(p_data, f, indent=2)
            self._load_profiles()
            # Re-select the renamed profile
            for i in range(self.lb_profiles.Items.Count):
                li = self.lb_profiles.Items.GetItemAt(i)
                if hasattr(li, "Tag") and li.Tag.get("name") == new_name:
                    self.lb_profiles.SelectedIndex = i
                    break
        except Exception as ex:
            MessageBox.Show(
                "Rename failed:\n{}".format(ex),
                "Error", MessageBoxButton.OK)

    def _on_profile_selected(self, sender, e):
        if self._loading:
            return
        item = self.lb_profiles.SelectedItem
        if item is None:
            return
        p = item.Tag
        # Sync combo picker
        for i in range(self.cb_profile_picker.Items.Count):
            ci = self.cb_profile_picker.Items.GetItemAt(i)
            if hasattr(ci, 'Tag') and ci.Tag is p:
                self._loading = True
                self.cb_profile_picker.SelectedIndex = i
                self._loading = False
                break
        self._apply_profile(p)

    def _apply_profile(self, p):
        self._loading = True
        self.tb_profile_name.Text = p.get("name", "")
        self.tb_notes.Text        = p.get("notes", "")
        self._set_filter_logic(p.get("filter_logic", "AND"))
        self.tb_group_by.Text = p.get("group_by_param", "")

        self._clear_filter_rules()
        for rule in p.get("filter_rules", []):
            row = FilterRuleRow(
                on_change_cb=self._on_any_change,
                on_remove_cb=self._on_rule_removed,
                rule_data=rule,
                doc=self.doc
            )
            self._filter_rule_rows.append(row)
            self._rules_sp.Children.Add(row.root)

        self._clear_output_editors()
        for od in p.get("outputs", []):
            self._add_output_editor(od)

        self.cb_allow_nested.IsChecked = bool(p.get("allow_nested", False))
        self._loading = False
        # Reset AFTER editors finish loading — editor callbacks may have
        # fired _on_any_change even with _loading=True if they use their
        # own internal _loading flag.
        self._dirty = False
        # Force a second reset on next event loop tick via a flag
        self._pending_clean = True
        self._update_notes_bar()

    def _current_profile_data(self):
        return {
            "name":           self.tb_profile_name.Text.strip(),
            "notes":          self.tb_notes.Text.strip(),
            "filter_logic":   self._get_filter_logic(),
            "filter_rules":   self._get_filter_rules(),
            "group_by_param": self.tb_group_by.Text.strip(),
            "allow_nested":   bool(self.cb_allow_nested.IsChecked),
            "outputs":        [ed.get_data() for ed in self._output_editors],
        }

    # ----------------------------------------------------------
    # PROFILE CRUD
    # ----------------------------------------------------------

    def _on_new_profile(self, sender, e):
        self._loading = True
        self.tb_profile_name.Text = "New Profile"
        self.tb_notes.Text        = ""
        self.cb_allow_nested.IsChecked = False
        self._set_filter_logic("AND")
        self.tb_group_by.Text = ""
        self._clear_filter_rules()
        self._clear_output_editors()
        self._loading = False
        self._dirty = True
        self.lb_profiles.SelectedIndex = -1
        self._update_notes_bar()

    def _on_delete_profile(self, sender, e):
        item = self.lb_profiles.SelectedItem
        if item is None:
            return
        p = item.Tag
        is_division = p.get("readonly", False)
        loc = "division folder" if is_division else "My Profiles"
        extra = ""
        if is_division:
            extra = "\n\nThis is a division profile. You need write access to the bundle folder."
        r = MessageBox.Show(
            "Delete '{}' from {}?{}".format(p["name"], loc, extra),
            "Confirm Delete", MessageBoxButton.YesNo)
        if r == MessageBoxResult.Yes:
            try:
                delete_profile(p)
                self._load_profiles()
            except Exception as ex:
                MessageBox.Show(
                    "Could not delete profile:\n{}\n\nYou may not have write access to the division folder.".format(ex),
                    "Delete Failed", MessageBoxButton.OK)

    # ----------------------------------------------------------
    # SAVE HELPERS
    # ----------------------------------------------------------

    def _do_save(self, destination):
        name = self.tb_profile_name.Text.strip()

        if not name:
            MessageBox.Show(
                "Profile name cannot be blank. Enter a name before saving.",
                "Name Required", MessageBoxButton.OK)
            return

        safe = _safe_filename(name)
        if not safe:
            MessageBox.Show(
                "Profile name contains only special characters and cannot"
                " be saved as a file. Use letters or numbers.",
                "Invalid Name", MessageBoxButton.OK)
            return

        data = self._current_profile_data()
        folder = SHARED_DIR if destination == "division" else USER_DIR
        target_fpath = os.path.join(folder, safe + ".json")

        current_fp = ""
        sel = self.lb_profiles.SelectedItem
        if sel and hasattr(sel, "Tag"):
            current_fp = sel.Tag.get("_filepath", "")

        if os.path.exists(target_fpath) and target_fpath != current_fp:
            try:
                with open(target_fpath) as _f:
                    conflict = json.load(_f)
                conflict_name = conflict.get("name", target_fpath)
            except Exception:
                conflict_name = target_fpath
            msg = "Cannot save - a profile already exists at that location.\n"
            msg += "Existing: {}\n".format(conflict_name)
            msg += "File: {}\n\n".format(target_fpath)
            msg += "Rename your profile to something unique before saving."
            MessageBox.Show(msg, "Name Conflict", MessageBoxButton.OK)
            return

        try:
            fpath = save_profile(data, destination)
            self._dirty = False
            self._pending_clean = False
            self._load_profiles()
            label = "My Profiles" if destination == "user" else "Division"
            MessageBox.Show(
                "Profile {} saved to {}.".format(name, label),
                "Saved", MessageBoxButton.OK)
        except (IOError, OSError) as ex:
            if destination == "division":
                MessageBox.Show(
                    "Access denied to division folder.\n"
                    "Path: {}\nError: {}".format(SHARED_DIR, ex),
                    "Access Denied", MessageBoxButton.OK)
            else:
                MessageBox.Show("Save failed:\n{}".format(ex), "Error", MessageBoxButton.OK)
        except Exception as ex:
            MessageBox.Show("Save failed:\n{}".format(ex), "Error", MessageBoxButton.OK)

    def _on_save_user(self, sender, e):
        self._do_save("user")

    def _on_save_division(self, sender, e):
        r = MessageBox.Show(
            "Save to the shared Division folder?\n"
            "This requires write access to:\n{}\n\n"
            "If you don't have access, the save will fail with an access error.".format(SHARED_DIR),
            "Save to Division", MessageBoxButton.YesNo
        )
        if r == MessageBoxResult.Yes:
            self._do_save("division")

    def _on_import(self, sender, e):
        """Import one or more profiles from a JSON bundle or single-profile file."""
        dlg = OpenFileDialog()
        dlg.Filter = "JSON Files (*.json)|*.json"
        dlg.Title  = "Import Profiles"
        if dlg.ShowDialog() is not True:
            return

        # Load file — supports bundle {"profiles": [...]} or single profile {name:...}
        try:
            with open(dlg.FileName, "r") as f:
                data = json.load(f)
        except Exception as ex:
            MessageBox.Show("Could not read file:\n{}".format(ex), "Error", MessageBoxButton.OK)
            return

        if "profiles" in data:
            candidates = data["profiles"]   # bundle format
        elif "name" in data:
            candidates = [data]             # single profile format
        else:
            MessageBox.Show(
                "Unrecognised file format.\nExpected a profile bundle or single profile JSON.",
                "Error", MessageBoxButton.OK)
            return

        # Strip readonly/filepath flags
        for p in candidates:
            p.pop("readonly", None)
            p.pop("_filepath", None)

        if not candidates:
            MessageBox.Show("No profiles found in file.", "Empty File", MessageBoxButton.OK)
            return

        # Let user pick which profiles to import
        sel_win = _CheckboxSelectWindow(
            title="Import: Select Profiles to Import",
            sections=[("Profiles", [p.get("name", "Unnamed") for p in candidates])]
        )
        sel_win.ShowDialog()
        if not sel_win.confirmed:
            return

        chosen = [p for i, p in enumerate(candidates)
                  if sel_win.is_checked("Profiles", i)]
        if not chosen:
            MessageBox.Show("No profiles selected.", "Nothing Selected", MessageBoxButton.OK)
            return

        # Destination — use proper picker UI
        dest_win = _DestinationPickerWindow(len(chosen))
        dest_win.ShowDialog()
        if dest_win.destination is None:
            return
        destination = dest_win.destination

        # Check for name conflicts with existing profiles
        existing_names = {p["name"] for p in list_profiles()}
        saved = []
        skipped = []
        failed = []

        for p in chosen:
            pname = p.get("name", "Unnamed")
            # Handle duplicate name
            if pname in existing_names:
                r = MessageBox.Show(
                    "A profile named '{}' already exists.\n\n"
                    "Overwrite  →  Yes\n"
                    "Keep both (rename)  →  No\n"
                    "Skip this profile  →  Cancel".format(pname),
                    "Duplicate Profile Name",
                    MessageBoxButton.YesNoCancel)
                if r == MessageBoxResult.Cancel:
                    skipped.append(pname)
                    continue
                elif r == MessageBoxResult.No:
                    # Rename: append (2), (3) etc.
                    base = pname
                    n = 2
                    while pname in existing_names:
                        pname = "{} ({})".format(base, n)
                        n += 1
                    p = dict(p)
                    p["name"] = pname
                # Yes = overwrite (save_profile will overwrite by filename)

            try:
                save_profile(p, destination)
                saved.append(pname)
                existing_names.add(pname)
            except Exception as ex:
                failed.append("{}: {}".format(pname, ex))

        self._load_profiles()
        dest_label = "My Profiles" if destination == "user" else "Division"
        lines = ["Imported {} profile(s) to {}.".format(len(saved), dest_label)]
        if skipped:
            lines.append("\nSkipped: {}".format(", ".join(skipped)))
        if failed:
            lines.append("\nFailed:\n" + "\n".join(failed))
        MessageBox.Show("\n".join(lines), "Import Complete", MessageBoxButton.OK)

    def _on_export(self, sender, e):
        """Export one or more profiles to a single JSON bundle file."""
        all_profiles = list_profiles()
        if not all_profiles:
            MessageBox.Show("No profiles to export.", "Nothing to Export", MessageBoxButton.OK)
            return

        # Pick which profiles to export
        sel_win = _CheckboxSelectWindow(
            title="Export: Select Profiles to Export",
            sections=[
                ("Profiles", [p["name"] for p in all_profiles])
            ]
        )
        sel_win.ShowDialog()
        if not sel_win.confirmed:
            return

        chosen = [p for i, p in enumerate(all_profiles)
                  if sel_win.is_checked("Profiles", i)]
        if not chosen:
            MessageBox.Show("No profiles selected.", "Nothing Selected", MessageBoxButton.OK)
            return

        # Strip internal keys before saving
        bundle = [
            {k: v for k, v in p.items() if k not in ("readonly", "_filepath")}
            for p in chosen
        ]

        dlg = SaveFileDialog()
        dlg.Filter   = "NW Profile Bundle (*.json)|*.json"
        dlg.FileName = "NW_Profiles.json"
        dlg.Title    = "Export Profiles"
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        dlg.InitialDirectory = downloads if os.path.exists(downloads) else os.path.expanduser("~")
        if dlg.ShowDialog() is not True:
            return

        try:
            with open(dlg.FileName, "w") as f:
                json.dump({"profiles": bundle}, f, indent=2)
            MessageBox.Show(
                "Exported {} profile(s) to:\n{}".format(len(bundle), dlg.FileName),
                "Export Complete", MessageBoxButton.OK)
        except Exception as ex:
            MessageBox.Show("Export failed:\n{}".format(ex), "Error", MessageBoxButton.OK)

    # ----------------------------------------------------------
    # CLOSE GUARD
    # ----------------------------------------------------------

    def _on_closing(self, sender, e):
        # DialogResult=True means Run was clicked intentionally — skip guard.
        # Also skip if _dirty was never actually set by user action.
        if self.DialogResult == True:
            return
        if self._dirty:
            r = MessageBox.Show("Unsaved changes. Discard and close?", "Unsaved", MessageBoxButton.YesNo)
            if r != MessageBoxResult.Yes:
                e.Cancel = True

    # ----------------------------------------------------------
    # RUN
    # ----------------------------------------------------------

    def _on_run(self, sender, e):
        profile = self._current_profile_data()
        if not profile["name"]:
            MessageBox.Show("Enter a profile name.", "Missing Name", MessageBoxButton.OK)
            return
        if not profile["outputs"]:
            MessageBox.Show("Add at least one output formula.", "No Outputs", MessageBoxButton.OK)
            return
        for od in profile["outputs"]:
            if not od["target_param"]:
                MessageBox.Show("One or more outputs is missing a 'Write to param' name.", "Missing Target", MessageBoxButton.OK)
                return

        if self.rb_renumber_all.IsChecked:
            run_mode = "renumber_all"
        elif self.rb_start_highest.IsChecked:
            run_mode = "start_highest"
        else:
            run_mode = "fill_missing"

        if self.rb_src_sel.IsChecked:
            elem_source = "selection"
        elif self.rb_src_pick.IsChecked:
            elem_source = "pick"
        else:
            elem_source = "view"

        # Sequence start number (only used for renumber_all)
        try:
            seq_start = max(1, int(self.tb_seq_start.Text.strip()))
        except Exception:
            seq_start = 1

        self._profile_to_run = profile
        self.run_mode        = run_mode
        self.elem_source     = elem_source
        self.seq_start       = seq_start
        self.DialogResult    = True
        self.Close()

# ==========================================================
# EXECUTION PIPELINE
# Called after the window closes with DialogResult=True.
# Stages:
#   1  Collect   — source-specific collection (pick/selection/view)
#   2  Filter    — filter_elements_with_reasons() — same for all sources
#   3  Job #     — from ProjectInformation or user prompt
#   4  Group     — by group_by_param
#   4b Check     — warn if output params are non-String storage
#   5  Write     — single Transaction, Set() each output param
#   6  Report    — _show_run_report() consistent MessageBox popup
# ==========================================================

def execute(doc, uidoc, profile, run_mode, elem_source, seq_start=1):
    """
    Main execution pipeline — called after the tool window closes.
    doc:         Revit Document
    uidoc:       Revit UIDocument
    profile:     dict — current profile from _current_profile_data()
    run_mode:    fill_missing | start_highest | renumber_all
    elem_source: pick | selection | view
    """
    # ── Stage 1: Collect elements ───────────────────────────────────
    # Each collector internally extracts category IDs for pre-filtering.

    if elem_source == "pick":
        # Window is already closed — Revit can accept pick input.
        # ISelectionFilter grays out non-matching elements in the UI (best-effort).
        # filter_elements() at line ~4000 is the authoritative gate for all sources.
        elements = pick_box_select(uidoc, profile)
        if not elements:
            return   # Esc pressed or nothing picked

    elif elem_source == "view":
        # Category pre-filter at collector level — never loads excluded cats
        elements = collect_visible_elements(doc, uidoc, profile)
        if not elements:
            MessageBox.Show("No visible elements found in the active view.", "Nothing Found", MessageBoxButton.OK)
            return

    else:  # "selection"
        # Category pre-filter scoped to the selection IDs
        elements = collect_from_selection(doc, uidoc, profile)
        if not elements:
            MessageBox.Show(
                "Nothing in your selection matched the profile's category rules.\n"
                "Select elements first, or check your filter rules.",
                "Nothing Matched", MessageBoxButton.OK)
            return

    # ── Stage 2: Filter — uniform for all collection sources ───────────
    kept, pre_skip_log = filter_elements_with_reasons(elements, profile)

    # Warn if no filter rules and using view — could be thousands of elements
    if elem_source == "view" and not profile.get("filter_rules"):
        r = MessageBox.Show(
            "No filter rules are defined.\n"
            "This will process ALL {} visible elements.\n\n"
            "Continue?".format(len(kept)),
            "No Filter Rules", MessageBoxButton.YesNo
        )
        if r != MessageBoxResult.Yes:
            return

    if not kept:
        detail = ""
        if pre_skip_log:
            reasons = {}
            for _, reason in pre_skip_log:
                reasons[reason] = reasons.get(reason, 0) + 1
            detail = "\n\n" + "\n".join("  {}  x{}".format(r, n) for r, n in sorted(reasons.items(), key=lambda x: -x[1]))
        MessageBox.Show(
            "No elements to process.\n\n"
            "Collected:  {}\n"
            "Excluded:   {}{}".format(
                len(elements), len(pre_skip_log), detail),
            "Nothing to Process", MessageBoxButton.OK)
        return

    # ── Stage 3: Job Number ──────────────────────────────────────────
    job_number = get_job_number(doc)
    SAMPLE_VALUES["project_number"] = job_number

    # ── Stage 3b: Collect and prompt for ASK: tokens ─────────────────
    # Gather unique ASK: keys across all output slots (order-preserving).
    # Matching keys share a single prompt — only asked once per label.
    ask_keys_seen = []
    for od in profile["outputs"]:
        for slot in od.get("slots", []):
            k = slot.get("key", "")
            if k.startswith("ASK:") and k not in ask_keys_seen:
                ask_keys_seen.append(k)

    user_inputs = {}
    for k in ask_keys_seen:
        label = k[len("ASK:"):]
        val = forms.ask_for_string(
            default="",
            prompt="Enter value for:  {}".format(label),
            title=label
        )
        if val is None:   # user cancelled
            return
        user_inputs[k] = val.strip()

    # ── Stage 4: Group by param ──────────────────────────────────────
    group_param = profile.get("group_by_param", "").strip()
    groups = {}
    for elem in kept:
        if group_param:
            key = param_text(lookup_param(elem, group_param)) or "UNK"
        else:
            key = "ALL"
        groups.setdefault(key, []).append(elem)

    outputs = profile["outputs"]

    # ── Stage 4b: Output param storage type check ────────────────────
    # Warn before writing if any output param is non-String storage.
    # This tool writes string values — Set(string) fails on Integer/Double.
    sample_elem = kept[0]
    non_string_warnings = []
    for od in outputs:
        target = od.get("target_param", "").strip()
        if not target:
            continue
        p = lookup_param(sample_elem, target)
        if p is None:
            continue
        try:
            storage = p.StorageType.ToString()
            if storage != "String":
                non_string_warnings.append("  '{}' is {} storage".format(target, storage))
        except Exception:
            pass
    if non_string_warnings:
        r = MessageBox.Show(
            "Warning: {} output param(s) are not String storage:\n\n{}\n\n"
            "Writing string formula results to non-string params will fail.\n"
            "Continue anyway?".format(len(non_string_warnings), "\n".join(non_string_warnings)),
            "Non-String Output Params", MessageBoxButton.YesNo)
        if r != MessageBoxResult.Yes:
            return

    # ── Stage 5: Transaction ─────────────────────────────────────────
    t = Transaction(doc, "NW Param & Number")
    t.Start()
    written = 0
    # skip_log starts with elements excluded by filter (nested/rule failures)
    skip_log = list(pre_skip_log)

    try:
        for group_key, elems in groups.items():
            elems_sorted = sorted(elems, key=lambda x: x.Id.IntegerValue)

            for od in outputs:
                target  = od["target_param"]
                slots   = od["slots"]
                sep     = od["sep"]
                sw      = od.get("seq_width", 3)

                pre_slots = [s for s in slots if s.get("key") != "SEQUENCE"]
                prefix = build_string_from_slots(pre_slots, sep, elems_sorted[0], doc,
                                                 seq_number=None, seq_width=sw,
                                                 user_inputs=user_inputs)
                has_seq = any(s.get("key") == "SEQUENCE" for s in slots)

                if has_seq:
                    if run_mode == "renumber_all":
                        next_seq = seq_start   # user-defined start (default 1)
                    else:
                        next_seq = get_next_sequence(elems_sorted, prefix, sep, target)
                else:
                    next_seq = None

                for elem in elems_sorted:
                    eid = str(elem.Id.IntegerValue)
                    p_out = lookup_param(elem, target)
                    if p_out is None:
                        skip_log.append((eid, "Param not found: {}".format(target)))
                        continue

                    if run_mode == "fill_missing":
                        if not param_is_empty(p_out):
                            existing = param_text(p_out)
                            skip_log.append((eid, "Already filled: '{}' = '{}'".format(target, existing[:30])))
                            continue

                    value = build_string_from_slots(
                        slots, sep, elem, doc,
                        seq_number=next_seq, seq_width=sw,
                        user_inputs=user_inputs
                    )
                    try:
                        p_out.Set(value)
                        written += 1
                    except Exception as write_ex:
                        skip_log.append((eid, "Write error on '{}': {}".format(target, write_ex)))

                    if has_seq and next_seq is not None:
                        next_seq += 1

        t.Commit()

    except Exception as ex:
        logger.exception(ex)
        t.RollBack()
        MessageBox.Show("Transaction failed:\n{}".format(ex), "Error", MessageBoxButton.OK)
        return

    # ── Stage 6: Report ─────────────────────────────────────────────
    # elements_processed = elements that passed all filters (kept)
    _show_run_report(written, skip_log, elements_processed=len(kept))

# ==========================================================
# HELP WINDOW
# ==========================================================

HELP_TEXT = """NW Param & Number — Quick Reference

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
QUICK-RUN BAR
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Profile        Select a saved profile from the dropdown.
               Notes (if any) appear below the run bar.
Edit Profile   Expand / collapse the full profile editor.

Source
  Pick (Finish)    Click elements in Revit one at a time, then
                   press Finish (green checkmark in ribbon).
                   Filter rules gray out non-matching elements.
  Selection        Uses your current Revit selection as input.
                   Filter rules narrow it to matching elements.
  Visible in View  Collects visible elements in the active view.
                   Category rules pre-filter at collector level
                   so large models stay fast.

Mode
  Renumber All    Overwrites all matched elements starting from
                  the "Start at" number (default 1). Use for
                  first-time numbering or a full reset.
  From Highest    Finds the highest existing number per group
                  and continues from there. Skips already-filled.
  Fill Missing    Only writes to elements where the output
                  parameter is blank. Leaves existing values.

Start at        (Renumber All only) Sets the first sequence
                number. Enter 10 to start at 010, 100 for 100.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PROFILES  (left panel)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
[D] prefix  Division profile — stored in the bundle folder.
            Requires write access to delete or overwrite.
no prefix   Your personal profile — stored in AppData.

Double-click a profile name to rename it.

Save to My Profiles   AppData — always writable.
Save to Division      Bundle folder — requires admin access.
New Profile           Create a blank profile.
Duplicate Profile     Copy the selected profile as a base.
Delete Profile        Delete. Division profiles need write access.
Import Profile...     Load profiles from a JSON bundle file.
Export Profile...     Save selected profiles to a JSON bundle file.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ELEMENT FILTER  (Profile & Filter tab)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Controls which collected elements get processed.
  AND = element must match ALL rules
  OR  = element passes if it matches ANY rule

Rule types:
  Category equals       Exact Revit category name match.
  Category contains     Partial category name match.
                        Both show a searchable category dropdown.
  Has parameter         Passes if the named param exists and
                        has a non-blank value.
  Param value contains  Passes if the named param's value
                        contains the specified text.
  Element name contains The element Name field contains text.
  Family name contains  The Family Name contains text.

Group by param    Splits matched elements into sub-groups
                  before sequencing. Each group gets its own
                  counter independently.
                  E.g. CP_Service Abbv: MECH→001-003, PLMB→001-002
                  Leave blank = all elements share one counter.

Include nested    OFF (default): only top-level family instances
sub-components    are processed. Nested geometry inside host
                  families is excluded. Correct for hanger RFAs
                  and equipment families where shared params live
                  on the parent instance.
                  ON: nested sub-components are also included.
                  Use when your target elements are nested inside
                  a host family.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMULAS  (Outputs & Parameters tab)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Each card writes a formula result to one Revit parameter.
Add as many cards as the profile needs.

Write to param   Exact Revit parameter name to write.
                 Must be String storage type.
Separator        Character placed between each slot (default: -)
Seq digits       Digit width of the sequence number (3 = 001)
Add text         Type a literal prefix, suffix, or fixed word
                 and click Add to insert it as a slot.
Preview          Shows real resolved values from the preview
                 element. Auto-sets from selection on open.
                 Click Change... to pick a different element.

Parameter chips in the drop zone:
  Shows the parameter name and [fmt] format badge.
  Click [fmt] to cycle through available read formats.
  Drag chips to reorder within the formula.
  Click x to remove a chip.

Output params must be String storage. A warning appears
before the transaction if any are not.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PARAMETER PALETTE  (Outputs & Parameters tab)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Built-in parameters (always present):
  Job #       ProjectInformation.Number
  Proj Name   ProjectInformation.Name
  ###         Auto-incrementing sequence number
  Ask User... Prompts for a typed value at run time

Ask User token:
  Type a label in the "Ask label" box (e.g. Phase, Zone)
  and click "Add Ask" to insert a run-time input prompt.
  The user is prompted once per unique label when Run is
  clicked. Use the same label in multiple formulas to reuse
  the same answer. Use different labels to get different
  values (e.g. ASK:Phase and ASK:Zone = two prompts).
  Preview shows <Label> as a placeholder.

Scan parameters: click "Pick Element to Scan" to load all
parameters from one element into the list. Color indicates
unit type (see legend dots at top of palette).

Double-click a parameter to add it to the last active formula.
Drag a parameter from the list into a formula drop zone.

Static text: type text in the "Text token" box and click
"Add to Formula" to insert a fixed value (H, MECH-, etc.).

Add param manually:
  1. Type the exact parameter name
  2. Select the unit type (or leave Auto-detect)
  3. Click "Add to Palette"
  Auto-detect uses the preview element if one is selected.

Format options per unit type:
  Length   in" (3/4")   in. (0.75)   ft-in (0'-9")   ft. (0.75)   rev
  Area     val (125.3)  rev (125.3 SF)
  Volume   val  rev
  Flow     val  rev
  Angle    val  rev
  Yes/No   T/F  Y/N  0/1
  Integer  int (raw number)
  String   str (as-is)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMPORT / EXPORT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Export  Select one or more profiles to include, then save
        as a JSON bundle file (NW_Profiles.json).
Import  Open a JSON bundle file, select which profiles to
        bring in, then choose My Profiles or Division.
        Handles duplicate names: overwrite, rename, or skip.
"""

class _HelpWindow(Window):
    def __init__(self):
        self.Title = "NW Param & Number — Help"
        self.Width = 560
        self.Height = 620
        self.MinWidth = 400
        self.MinHeight = 300
        self.ShowInTaskbar = False
        self.WindowStartupLocation = WindowStartupLocation.CenterOwner
        self.ResizeMode = ResizeMode.CanResizeWithGrip

        outer = StackPanel(); outer.Margin = Thickness(12, 10, 12, 10)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Height = 520

        tb = TextBlock()
        tb.Text = HELP_TEXT
        tb.FontFamily = WpfFontFamily("Consolas")
        tb.FontSize = 11
        tb.TextWrapping = TextWrapping.Wrap
        tb.Margin = Thickness(0, 0, 8, 0)
        sv.Content = tb

        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Width = 80
        btn_close.Height = 28
        btn_close.Margin = Thickness(0, 8, 0, 0)
        btn_close.HorizontalAlignment = HorizontalAlignment.Right
        btn_close.IsDefault = True
        btn_close.IsCancel  = True
        btn_close.Click += lambda s, e: self.Close()

        outer.Children.Add(sv)
        outer.Children.Add(btn_close)
        self.Content = outer

# ==========================================================
# SKIP REPORT
# ==========================================================

def _show_run_report(written, skip_log, elements_processed=None):
    """
    Show a consistent run summary popup.

    written           = number of parameter Set() calls that succeeded
    skip_log          = list of (elem_id_str, reason) for everything skipped
    elements_processed = number of elements that passed all filters and were
                         attempted (optional — shown when > 0)

    Verbiage is element-centric, not write-count-centric, to avoid
    confusion when one element has multiple output parameters.
    """
    total_skipped = len(skip_log)

    # Deduplicate reason counts for summary
    reason_counts = {}
    for eid, reason in skip_log:
        reason_counts[reason] = reason_counts.get(reason, 0) + 1

    # Separate nested exclusions from genuine rule failures
    nested_count = sum(v for k, v in reason_counts.items() if "Nested" in k)
    rule_fail_count = sum(v for k, v in reason_counts.items() if "Failed filter" in k)
    already_filled_count = sum(v for k, v in reason_counts.items() if "Already filled" in k)
    write_error_count = sum(v for k, v in reason_counts.items() if "Write error" in k or "not found" in k)

    if total_skipped == 0:
        msg = "Run complete.\n\n"
        if elements_processed:
            msg += "Elements processed:  {}\n".format(elements_processed)
        msg += "Parameters written:  {}".format(written)
        MessageBox.Show(msg, "Run Complete", MessageBoxButton.OK)
        return

    lines = ["Run complete.\n"]
    if elements_processed:
        lines.append("Elements processed:  {}".format(elements_processed))
    lines.append("Parameters written:  {}".format(written))
    lines.append("")

    if already_filled_count:
        lines.append("Already had values (skipped):  {}".format(already_filled_count))
    if nested_count:
        lines.append("Nested sub-components (excluded):  {}".format(nested_count))
    if rule_fail_count:
        lines.append("Did not match filter rules:  {}".format(rule_fail_count))
    if write_error_count:
        lines.append("Write errors:  {}".format(write_error_count))

    # Any other reasons
    other_count = total_skipped - nested_count - rule_fail_count - already_filled_count - write_error_count
    if other_count:
        lines.append("Other skips:  {}".format(other_count))

    lines.append("")
    lines.append("Open detail report to see element IDs?")

    r = MessageBox.Show(
        "\n".join(lines),
        "Run Complete",
        MessageBoxButton.YesNo
    )
    if r == MessageBoxResult.Yes:
        rpt_win = _SkipReportWindow(skip_log)
        rpt_win.ShowDialog()

class _SkipReportWindow(Window):
    def __init__(self, skip_log):
        self.Title = "Skip Report"
        self.Width = 620
        self.Height = 480
        self.MinWidth = 400
        self.MinHeight = 260
        self.ShowInTaskbar = False
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.ResizeMode = ResizeMode.CanResizeWithGrip

        outer = StackPanel(); outer.Margin = Thickness(12, 10, 12, 10)

        hdr = TextBlock()
        hdr.Text = "{} element(s) not written:".format(len(skip_log))
        hdr.FontWeight = FontWeights.SemiBold
        hdr.Margin = Thickness(0, 0, 0, 6)
        outer.Children.Add(hdr)

        sv = ScrollViewer()
        sv.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.Height = 360

        lines_tb = TextBlock()
        lines_tb.FontFamily = WpfFontFamily("Consolas")
        lines_tb.FontSize = 11
        lines_tb.TextWrapping = TextWrapping.Wrap

        lines = []
        for eid, reason in skip_log:
            lines.append("ID {:>10}   {}".format(eid, reason))
        lines_tb.Text = "\n".join(lines)
        sv.Content = lines_tb

        btn_close = Button()
        btn_close.Content = "Close"
        btn_close.Width = 80
        btn_close.Height = 28
        btn_close.Margin = Thickness(0, 8, 0, 0)
        btn_close.HorizontalAlignment = HorizontalAlignment.Right
        btn_close.IsDefault = True
        btn_close.IsCancel  = True
        btn_close.Click += lambda s, e: self.Close()

        outer.Children.Add(sv)
        outer.Children.Add(btn_close)
        self.Content = outer

# ==========================================================
# ENTRY POINT
# ==========================================================

doc   = revit.doc
uidoc = revit.uidoc
win   = ParamFormulaWindow(doc, uidoc)
win.ShowDialog()

if getattr(win, '_profile_to_run', None) and getattr(win, 'run_mode', None):
    execute(
        doc, uidoc,
        win._profile_to_run,
        win.run_mode,
        getattr(win, 'elem_source', 'selection'),
        seq_start=getattr(win, 'seq_start', 1)
    )