# -*- coding: utf-8 -*-
# pyRevit IronPython 2.7
# ITM Hanger BOM Exporter - SI Tools V3 output contract (headers/order/formatting) preserved
# Native-only: NO Naviate / NO BIMrx
# Revit 2022-2026
#
# Bug fixes applied from diagnostic dump analysis (21 Feb 2026):
#   1. Rod diameter: accept Loose+SupportRod ancillary (clevis uses Loose, not Hanger)
#   2. CP_Rod Middle Length: return Length A for single-rod hangers (NorCal clevis)
#   3. Rod Length_2: try Length B FIRST, Drop Rod Length #2 only as fallback
#   4. Rod offsets: return abs() value (fabrication convention is negative)
#   5. Elevation F.F. - B.O.P.: use FABRICATION_OFFSET_PARAM (relative to level)
#   6. CP_Hanger Size: use Diameter dimension (hanger rated size, not pipe supported width)
#   7. CP_Construction Area: guard against empty/null returning "0"
#   8. Service: NorCal/NorCalTrap/SoCal expect full Fabrication Service name for ITM
#   9. E-BOH: Height + Rod Extn Above for clevis; strut length for trapeze (matches SoCal DLL logic)

from pyrevit import revit, forms

from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (
    BuiltInCategory, BuiltInParameter, FamilyInstance, FamilySymbol, Family, Element,
    LocationPoint, SpecTypeId, UnitTypeId, Units, FormatOptions, UnitSystem,
    UnitFormatUtils, StorageType
)

try:
    from Autodesk.Revit.DB import FabricationPart
except Exception:
    FabricationPart = None

from datetime import datetime

from System import Drawing
from System.Windows.Forms import (
    Form, ComboBox, Button, Label, DialogResult,
    FormBorderStyle, FormStartPosition, AnchorStyles, ComboBoxStyle
)

# =============================================================================
# WinForms Division Picker (dropdown like SI Tools)
# =============================================================================

def pick_division_winforms(options, default_value=None, title="Selection", prompt="Select your Division:"):
    form = Form()
    form.Text = title
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.StartPosition = FormStartPosition.CenterScreen
    form.MinimizeBox = False
    form.MaximizeBox = False
    form.ShowInTaskbar = False
    form.ClientSize = Drawing.Size(420, 120)
    form.Font = Drawing.SystemFonts.MessageBoxFont

    lbl = Label()
    lbl.Text = prompt
    lbl.AutoSize = True
    lbl.Location = Drawing.Point(12, 12)
    lbl.Font = Drawing.SystemFonts.MessageBoxFont

    cb = ComboBox()
    cb.DropDownStyle = ComboBoxStyle.DropDownList
    cb.Location = Drawing.Point(12, 38)
    cb.Size = Drawing.Size(396, 24)
    cb.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
    cb.Font = Drawing.SystemFonts.MessageBoxFont

    for o in options:
        cb.Items.Add(o)

    if default_value and default_value in options:
        cb.SelectedItem = default_value
    else:
        cb.SelectedIndex = 0

    ok = Button()
    ok.Text = "OK"
    ok.Location = Drawing.Point(252, 74)
    ok.Size = Drawing.Size(75, 28)
    ok.DialogResult = DialogResult.OK
    ok.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
    ok.Font = Drawing.SystemFonts.MessageBoxFont

    cancel = Button()
    cancel.Text = "Cancel"
    cancel.Location = Drawing.Point(333, 74)
    cancel.Size = Drawing.Size(75, 28)
    cancel.DialogResult = DialogResult.Cancel
    cancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right
    cancel.Font = Drawing.SystemFonts.MessageBoxFont

    form.Controls.Add(lbl)
    form.Controls.Add(cb)
    form.Controls.Add(ok)
    form.Controls.Add(cancel)
    form.AcceptButton = ok
    form.CancelButton = cancel

    result = form.ShowDialog()
    if result == DialogResult.OK:
        return cb.SelectedItem
    return None


# =============================================================================
# ITM selection filter (STRICT: Fabrication Hangers only)
# =============================================================================

class ItmHangerSelectionFilter(ISelectionFilter):
    """Allow ONLY FabricationPart hangers (prevents selecting fabrication pipework)."""

    def AllowElement(self, element):
        try:
            if FabricationPart is None or not isinstance(element, FabricationPart):
                return False

            cat = element.Category
            if not cat:
                return False

            try:
                if cat.Id.IntegerValue != int(BuiltInCategory.OST_FabricationHangers):
                    return False
            except Exception:
                cname = (cat.Name or "").lower()
                if "fabrication hangers" not in cname and "mep fabrication hangers" not in cname:
                    return False

            # Secondary guard: must have hanger-like dims or rod info
            try:
                defs = element.GetDimensions()
                for d in defs:
                    nm = (d.Name or "").strip().lower()
                    if nm in ("length a", "bearer extn", "width"):
                        return True
            except Exception:
                pass

            try:
                ri = element.GetRodInfo()
                _ = ri.RodCount
                return True
            except Exception:
                return False

        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


# =============================================================================
# Parameter helpers
# =============================================================================

def param_string_or_value_string(p):
    if not p:
        return ""
    try:
        s = p.AsString()
        if s is not None:
            return s
    except Exception:
        pass
    try:
        s = p.AsValueString()
        if s is not None:
            return s
    except Exception:
        pass
    return ""


# =============================================================================
# SI Tools helper equivalents (HangerBOMV3Service)
# =============================================================================

def hanger_symbol_has_si_version_parameter(hanger):
    # ITM: always False; all pname_map lambdas take the legacy path
    try:
        sym = hanger.Symbol
        if not sym:
            return False
        return sym.LookupParameter("SI_Version") is not None
    except Exception:
        return False


def get_si_version_string(hanger):
    p = get_param(hanger, "SI_Version")
    v = ""
    if p:
        try:
            v = (p.AsString() or "").strip()
        except Exception:
            try:
                v = (p.AsValueString() or "").strip()
            except Exception:
                v = ""
    return v if v else "Legacy"


def of_si_version(hanger, version_text):
    return get_si_version_string(hanger) == version_text


def is_socal_specialty_hanger(hanger):
    """SoCal DLL checks Symbol param 'Model' for Trapeze/C Channel/Tube Steel/Angle Iron.
    For ITM FabricationParts: detect trapeze via native dimensions (Bearer Extn / Width).
    This covers the same intent — specialty = multi-rod strut-based hangers."""
    # ITM path: detect by fabrication dimensions
    try:
        if FabricationPart is not None and isinstance(hanger, FabricationPart):
            return _is_trapeze(hanger)
    except Exception:
        pass
    # RFA path: original DLL logic
    try:
        sym = hanger.Symbol
        if not sym:
            return False
        p = sym.LookupParameter("Model")
        if not p:
            return False
        model = (p.AsString() or p.AsValueString() or "").strip()
        return model in ("Trapeze", "C Channel", "Tube Steel", "Angle Iron")
    except Exception:
        return False


def get_family_name(h):
    # ITM: ProductLongDescription is the stable label
    try:
        if FabricationPart is not None and isinstance(h, FabricationPart):
            return getattr(h, "ProductLongDescription", "") or ""
    except Exception:
        pass
    try:
        return h.Symbol.Family.Name
    except Exception:
        return ""


def get_type_name(h):
    try:
        if FabricationPart is not None and isinstance(h, FabricationPart):
            return ""
    except Exception:
        pass
    try:
        return h.Symbol.Name
    except Exception:
        try:
            return h.Name
        except Exception:
            return ""


def get_family_and_type(h):
    fam = get_family_name(h)
    typ = get_type_name(h)
    if fam and typ:
        return "{}:{}".format(fam, typ)
    return fam or typ or ""


# =============================================================================
# Length formatting (mirrors LengthFormattingUtility)
# =============================================================================

VALUE_CATEGORY = SpecTypeId.Length

def make_units(display_unit_type, accuracy):
    units = Units(UnitSystem.Imperial)
    fmt = FormatOptions(display_unit_type)
    fmt.Accuracy = accuracy
    units.SetFormatOptions(VALUE_CATEGORY, fmt)
    return units

UNITS_DECIMAL_INCHES_0_001 = make_units(UnitTypeId.Inches, 0.001)
UNITS_DECIMAL_FEET_0_001 = make_units(UnitTypeId.Feet, 0.001)
UNITS_FEET_FRACTIONAL_INCHES_1_96 = make_units(UnitTypeId.FeetFractionalInches, 1.0 / 96.0)
UNITS_FRACTIONAL_INCHES_0_125 = make_units(UnitTypeId.FractionalInches, 0.125)
UNITS_DECIMAL_INCHES_1 = make_units(UnitTypeId.Inches, 1.0)
UNITS_FRACTIONAL_INCHES_1 = make_units(UnitTypeId.FractionalInches, 1.0)

def format_length(units, value_internal_feet):
    try:
        return UnitFormatUtils.Format(units, VALUE_CATEGORY, value_internal_feet, False)
    except Exception:
        return ""


# =============================================================================
# CSV sanitization (matches EvaluateFormattedCellValue: Replace(",", ":"))
# =============================================================================

def safe_cell(val):
    try:
        if val is None:
            s = ""
        elif isinstance(val, basestring):
            s = val
        else:
            s = str(val)
        return s.replace(",", ":")
    except Exception:
        return ""


# =============================================================================
# MappingItem (replicates HangerBOMMappingItem)
# =============================================================================

class MappingItem(object):
    def __init__(self, header, parameter_name=None, pname_map=None, vfmt=None):
        self.Header = header
        self.ParameterName = parameter_name
        self.ParameterNameMappingFunc = pname_map
        self.ValueFormattingFunc = vfmt or self.vfmt_string_or_valuestring

    def _eval_param_name(self, hanger):
        pname = None
        if self.ParameterNameMappingFunc:
            try:
                pname = self.ParameterNameMappingFunc(hanger)
            except Exception:
                pname = None
        return pname or self.ParameterName

    def _get_parameter(self, hanger):
        pname = self._eval_param_name(hanger)
        if pname in ("Family", "Type", "Family and Type"):
            return None
        return get_param(hanger, pname)

    def eval(self, hanger):
        try:
            v = self.ValueFormattingFunc(hanger, self)
            return safe_cell(v)
        except Exception:
            return ""

    # ---- ValueFormattingFunctions equivalents ----

    @staticmethod
    def vfmt_string_or_valuestring(hanger, item):
        pname = item._eval_param_name(hanger)
        if pname == "Family":
            return get_family_name(hanger)
        if pname == "Type":
            return get_type_name(hanger)
        if pname == "Family and Type":
            return get_family_and_type(hanger)
        return param_string_or_value_string(item._get_parameter(hanger))

    @staticmethod
    def vfmt_raw_value_to_string(hanger, item):
        p = item._get_parameter(hanger)
        if not p:
            return ""
        try:
            if p.StorageType.ToString() == "Double":
                return str(p.AsDouble())
            if p.StorageType.ToString() == "Integer":
                return str(p.AsInteger())
            if p.StorageType.ToString() == "String":
                return p.AsString() or ""
            if p.StorageType.ToString() == "ElementId":
                eid = p.AsElementId()
                return str(eid.IntegerValue) if eid else ""
        except Exception:
            pass
        return param_string_or_value_string(p)

    @staticmethod
    def vfmt_three_decimal_inches(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_DECIMAL_INCHES_0_001, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_three_decimal_feet(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_DECIMAL_FEET_0_001, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_feet_fractional_inches_round_to_1_8(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_FEET_FRACTIONAL_INCHES_1_96, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_fractional_inches_round_to_1_8(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_FRACTIONAL_INCHES_0_125, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_decimal_whole_inches(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_DECIMAL_INCHES_1, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_fractional_whole_inches(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return format_length(UNITS_FRACTIONAL_INCHES_1, p.AsDouble())
        except Exception:
            return ""

    @staticmethod
    def vfmt_si_version(hanger, item):
        return get_si_version_string(hanger)

    @staticmethod
    def vfmt_element_id(hanger, item):
        try:
            return str(hanger.Id.IntegerValue)
        except Exception:
            return ""

    @staticmethod
    def vfmt_location_string(hanger, item):
        location_string = "No Location Point"
        try:
            loc = hanger.Location
            lp = loc if isinstance(loc, LocationPoint) else None
            if lp and lp.Point:
                p = lp.Point
                location_string = "({} {} {})".format(p.X, p.Y, p.Z)
        except Exception:
            pass
        return location_string

    @staticmethod
    def vfmt_one_count(hanger, item):
        return "1"

    @staticmethod
    def vfmt_integer_or_zero(hanger, item):
        p = item._get_parameter(hanger)
        try:
            return str(p.AsInteger()) if p else "0"
        except Exception:
            return "0"

    @staticmethod
    def vfmt_mad_revit_rod_length_decimal_inches(hanger, item):
        def dbl(name):
            p = get_param(hanger, name)
            try:
                return p.AsDouble() if p else 0.0
            except Exception:
                return 0.0
        longest = max(dbl("CP_Rod Left Length"), dbl("CP_Rod Middle Length"), dbl("CP_Rod Right Length"))
        return format_length(UNITS_DECIMAL_INCHES_0_001, longest)

    @staticmethod
    def vfmt_mad_revit_rod_length_ft_fractional_inches(hanger, item):
        def dbl(name):
            p = get_param(hanger, name)
            try:
                return p.AsDouble() if p else 0.0
            except Exception:
                return 0.0
        longest = max(dbl("CP_Rod Left Length"), dbl("CP_Rod Middle Length"), dbl("CP_Rod Right Length"))
        return format_length(UNITS_FEET_FRACTIONAL_INCHES_1_96, longest)

    @staticmethod
    def vfmt_mtw_longest_rod_length_ft(hanger, item):
        def dbl(name):
            p = get_param(hanger, name)
            try:
                return p.AsDouble() if p else 0.0
            except Exception:
                return 0.0
        longest = max(dbl("CP_Rod Left Length"), dbl("CP_Rod Middle Length"), dbl("CP_Rod Right Length"))
        return format_length(UNITS_DECIMAL_FEET_0_001, longest)


# Common mapping items (from HangerBOMMappingItem.Common)
COMMON_HANGER_VERSION = MappingItem("SI_Version", vfmt=MappingItem.vfmt_si_version)
COMMON_ELEMENT_ID = MappingItem("Element Id", vfmt=MappingItem.vfmt_element_id)
COMMON_ELEMENT_LOCATION = MappingItem("Element Location", vfmt=MappingItem.vfmt_location_string)


# =============================================================================
# ITM native resolver (FabricationPart) — NO Naviate/BIMrx dependencies
# Returns FakeParam objects that mimic Revit Parameter API used by MappingItem vfmt funcs.
# =============================================================================

class _FakeStorageType(object):
    def __init__(self, name):
        self._name = name
    def ToString(self):
        return self._name

class FakeParam(object):
    """Minimal Parameter-like adapter for MappingItem formatting funcs."""
    def __init__(self, storage_type_name, value, doc_for_valuestring=None):
        self.StorageType = _FakeStorageType(storage_type_name)
        self._v = value
        self._doc = doc_for_valuestring

    def AsDouble(self):
        try:
            return float(self._v)
        except Exception:
            return 0.0

    def AsInteger(self):
        try:
            return int(self._v)
        except Exception:
            return 0

    def AsString(self):
        try:
            if self.StorageType.ToString() == "String":
                return self._v if self._v is not None else ""
        except Exception:
            pass
        return None

    def AsValueString(self):
        try:
            if self.StorageType.ToString() == "Double" and self._doc is not None:
                return UnitFormatUtils.Format(self._doc.GetUnits(), SpecTypeId.Length, float(self._v), False)
        except Exception:
            pass
        try:
            s = self.AsString()
            return s if s is not None else ""
        except Exception:
            return ""


# =============================================================================
# ITM native helpers — fabrication dimension / rod / ancillary extraction
# =============================================================================

def _param_text(elem, pname):
    """Read a string/valuestring from a real Revit parameter on the element."""
    try:
        p = elem.LookupParameter(pname)
        if not p:
            return ""
        try:
            vs = p.AsValueString()
            if vs is not None:
                return vs
        except Exception:
            pass
        try:
            s = p.AsString()
            if s is not None:
                return s
        except Exception:
            pass
        try:
            return str(p.AsInteger())
        except Exception:
            pass
    except Exception:
        pass
    return ""


def _bip_text(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if not p:
            return ""
        try:
            vs = p.AsValueString()
            if vs is not None:
                return vs
        except Exception:
            pass
        try:
            s = p.AsString()
            if s is not None:
                return s
        except Exception:
            pass
    except Exception:
        pass
    return ""


def _dim_ft(part, dim_name):
    """Get a fabrication dimension value in internal feet by exact name match."""
    target = (dim_name or "").strip().lower()
    if not target:
        return None
    try:
        defs = part.GetDimensions()
        for d in defs:
            try:
                if (d.Name or "").strip().lower() == target:
                    return part.GetDimensionValue(d)
            except Exception:
                continue
    except Exception:
        pass
    return None


def _has_dim(part, dim_name):
    return _dim_ft(part, dim_name) is not None


def _is_trapeze(part):
    return _has_dim(part, "Bearer Extn") or _has_dim(part, "Width")


def _rod_count(part):
    try:
        ri = part.GetRodInfo()
        return int(ri.RodCount)
    except Exception:
        return 0


# ---- Rod diameter (FIX #1: accept both Hanger and Loose SupportRod) ----

def _support_rod_dia_ft(part):
    """Get support rod diameter from ancillary usage.
    Trapeze uses UsageType=Hanger, Clevis uses UsageType=Loose.
    Both have Type=SupportRod."""
    try:
        usage = part.GetPartAncillaryUsage()
        if usage:
            for anc in usage:
                try:
                    tp = (str(anc.Type) or "").lower()
                    if tp == "supportrod":
                        dia_ft = anc.AncillaryWidthOrDiameter
                        if dia_ft and abs(dia_ft) > 1e-6:
                            return dia_ft
                except Exception:
                    continue
    except Exception:
        pass
    return None


# ---- Rod lengths ----

def _rod_len_left_ft(part):
    """Rod 1 / Left rod length = Length A."""
    return _dim_ft(part, "Length A")


def _rod_len_right_ft(part):
    """Rod 2 / Right rod length.
    FIX #3: Try Length B FIRST (= correct rod 2 length on trapeze).
    Drop Rod Length #2 is an internal partial value — only use as last resort."""
    v = _dim_ft(part, "Length B")
    if v is not None and v > 1e-9:
        return v
    v = _dim_ft(part, "Drop Rod Length #2")
    if v is not None and v > 1e-9:
        return v
    return None


def _rod_len_middle_ft(part):
    """FIX #2: For single-rod hangers (clevis), the one rod IS the 'middle' rod.
    NorCal non-trapeze maps its single rod length column to CP_Rod Middle Length."""
    rc = _rod_count(part)
    if rc == 1:
        # Single rod = Length A = "middle" rod in NorCal schema
        return _dim_ft(part, "Length A")
    # Multi-rod: no distinct middle rod on ITM trapeze (left + right only)
    return None


# ---- Rod offsets (FIX #4: abs() for negative convention) ----

def _rod_offset_left_ft(part):
    if not _is_trapeze(part):
        return None
    v = _dim_ft(part, "Left Rod Offset")
    return abs(v) if v is not None else None


def _rod_offset_right_ft(part):
    if not _is_trapeze(part):
        return None
    v = _dim_ft(part, "Right Rod Offset")
    return abs(v) if v is not None else None


# ---- Strut length (trapeze) ----

def _strut_len_ft(part):
    if not _is_trapeze(part):
        return None
    w = _dim_ft(part, "Width")
    ext = _dim_ft(part, "Bearer Extn")
    if w is None or ext is None:
        return None
    return w + 2.0 * ext


# ---- Hanger size (FIX #6: use Diameter dimension) ----

def _hanger_size_ft(part):
    """Hanger rated size. On clevis: Diameter dim (e.g. 12").
    On trapeze: Diameter if present, else Supported Width."""
    v = _dim_ft(part, "Diameter")
    if v is not None and abs(v) > 1e-9:
        return v
    v = _dim_ft(part, "Supported Width")
    if v is not None and abs(v) > 1e-9:
        return v
    return None


# ---- E-BOH / Fab Height (FIX #9: trapeze = strut length, clevis = Height + Rod Extn Above) ----

def _eboh_ft(part):
    """E-BOH / fab height.
    Trapeze: strut length (Width + 2*Bearer Extn) — matches SoCal DLL logic.
    Clevis/other: Height + Rod Extn Above (total vertical extent of hanger hardware)."""
    if _is_trapeze(part):
        return _strut_len_ft(part)
    h = _dim_ft(part, "Height") or 0.0
    a = _dim_ft(part, "Rod Extn Above") or 0.0
    total = h + a
    return total if total > 1e-9 else None


# ---- Host nominal diameter (pipe size w/o insulation) ----

def _host_nominal_dia_ft(part):
    v = _dim_ft(part, "Supported Width")
    if v is not None and abs(v) > 1e-9:
        return v
    v = _dim_ft(part, "Supported Depth")
    if v is not None and abs(v) > 1e-9:
        return v
    v = _dim_ft(part, "Diameter")
    if v is not None and abs(v) > 1e-9:
        return v
    return None


# ---- Elevation F.F. - B.O.P. (FIX #5: FABRICATION_OFFSET_PARAM) ----

def _elev_ffbop_ft(part):
    """Bottom of pipe elevation relative to reference level.
    Uses FABRICATION_OFFSET_PARAM which is the part's offset from its reference level.
    Diagnostic confirmed: clevis=-0.510417 ft=-6.125" (expected -6.13), trapeze=-0.505208 ft=-6.0625" (expected -6.06)."""
    try:
        p = part.get_Parameter(BuiltInParameter.FABRICATION_OFFSET_PARAM)
        if p:
            return p.AsDouble()
    except Exception:
        pass
    return None


# ---- Full fabrication service name (for ITM: "PIPE: Hose Bib Water") ----

def _full_service_name(part):
    """Return the full Fabrication Service string (e.g. 'PIPE: Hose Bib Water').
    LookupParameter('Fabrication Service') has StorageType=Integer (it's an ElementId internally),
    but AsValueString() returns the display name."""
    try:
        p = part.LookupParameter("Fabrication Service")
        if p:
            vs = p.AsValueString()
            if vs:
                return vs
    except Exception:
        pass
    return _bip_text(part, BuiltInParameter.FABRICATION_SERVICE_NAME)


# =============================================================================
# itm_get_param: Map SI Tools CP_* parameter names → ITM native sources
# Returns FakeParam objects compatible with MappingItem vfmt functions.
# =============================================================================

def itm_get_param(part, pname):
    p = (pname or "").strip()

    # ---- string fields ----
    if p in ("Mark", "Comments", "Assembly Name"):
        return FakeParam("String", _param_text(part, p), doc)

    if p in ("Item Number", "CP_Item Number"):
        v = _param_text(part, "Item Number") or getattr(part, "ProductCode", "") or ""
        return FakeParam("String", v, doc)

    if p in ("Reference Level", "Base Level"):
        return FakeParam("String", _param_text(part, "Reference Level"), doc)

    if p == "SI_Version":
        return FakeParam("String", "Legacy", doc)

    # ---- CP_Construction Area (FIX #7: guard empty/null) ----
    if p == "CP_Construction Area":
        v = _param_text(part, "CP_Construction Area")
        # Guard: if param is missing or returns empty-ish, return blank
        if not v or v.strip() in ("", "0"):
            # Double-check: is it actually the string "0" from an unset param, or a real value?
            # Only suppress if the raw parameter doesn't exist or is truly empty
            try:
                rp = part.LookupParameter("CP_Construction Area")
                if rp and rp.StorageType == StorageType.String:
                    raw = rp.AsString()
                    return FakeParam("String", raw or "", doc)
            except Exception:
                pass
            return FakeParam("String", "", doc)
        return FakeParam("String", v, doc)

    # ---- Service fields (FIX #8: ITM uses full Fabrication Service for all divisions) ----
    if p in ("CP_Service", "CP_Service Name"):
        return FakeParam("String", _full_service_name(part), doc)

    if p == "CP_Service Abbv":
        # For ITM hangers, use full service name to match expected BOM output
        # (SI Tools RFA hangers use abbreviation, but ITM field expects full name)
        return FakeParam("String", _full_service_name(part), doc)

    if p == "Fabrication Service Abbreviation":
        return FakeParam("String", _param_text(part, "Fabrication Service Abbreviation"), doc)

    if p in ("Fabrication Service", "Fabrication Service Name"):
        return FakeParam("String", _full_service_name(part), doc)

    # ---- numeric/length fields (double, internal feet) ----
    if p in ("CP_Rod Diameter", "Rod Diameter", "Rod Size"):
        return FakeParam("Double", _support_rod_dia_ft(part) or 0.0, doc)

    if p == "CP_Rod Left Length":
        return FakeParam("Double", _rod_len_left_ft(part) or 0.0, doc)

    if p == "CP_Rod Middle Length":
        # FIX #2: single-rod hangers return Length A as "middle"
        return FakeParam("Double", _rod_len_middle_ft(part) or 0.0, doc)

    if p == "CP_Rod Right Length":
        return FakeParam("Double", _rod_len_right_ft(part) or 0.0, doc)

    if p in ("CP_Rod Left Offset From End Desired", "Rod Left Offset From End Desired"):
        return FakeParam("Double", _rod_offset_left_ft(part) or 0.0, doc)

    if p in ("CP_Rod Right Offset From End Desired", "Rod Right Offset From End Desired"):
        return FakeParam("Double", _rod_offset_right_ft(part) or 0.0, doc)

    # Attachment placeholders (future: hand-typed project parameter)
    if p in (
        "CP_Left Attachment Annotation", "CP_Middle Attachment Annotation", "CP_Right Attachment Annotation",
        "CP_Rod Left Attachment", "CP_Rod Middle Attachment", "CP_Rod Right Attachment",
        "CP_Left Attachment", "CP_Middle Attachment", "CP_Right Attachment"
    ):
        return FakeParam("String", "", doc)

    if p in ("CP_Rod QTY", "CP_Rod Count"):
        return FakeParam("Integer", _rod_count(part), doc)

    if p in ("CP_Strut QTY",):
        # Not directly available on ITM; return the real param if it exists, else 0
        try:
            rp = part.LookupParameter("CP_Strut QTY")
            if rp:
                return FakeParam("Integer", rp.AsInteger(), doc)
        except Exception:
            pass
        return FakeParam("Integer", 0, doc)

    if p in ("CP_Trapeze Width Actual", "CP_Length", "Unistrut Width"):
        return FakeParam("Double", _strut_len_ft(part) or 0.0, doc)

    if p in ("CP_Hung Object Bottom Elev", "CP_Bottom Elevation"):
        # FIX #5: use FABRICATION_OFFSET_PARAM for relative-to-level elevation
        return FakeParam("Double", _elev_ffbop_ft(part) or 0.0, doc)

    if p in ("CP_Fab Height", "CP_Hanger Fab Height"):
        # FIX #9: trapeze=strut, clevis=Height+Rod Extn Above
        return FakeParam("Double", _eboh_ft(part) or 0.0, doc)

    if p in ("CP_Host Nominal Diameter", "Host Nominal Diameter"):
        return FakeParam("Double", _host_nominal_dia_ft(part) or 0.0, doc)

    if p == "CP_Hanger Size":
        # FIX #6: use Diameter dimension (hanger rated size), not pipe supported width
        return FakeParam("Double", _hanger_size_ft(part) or 0.0, doc)

    if p == "CP_Coating Type":
        return FakeParam("String", _param_text(part, "CP_Coating Type"), doc)

    if p in ("CP_Hanger Type Abbreviation", "CP_Prefix"):
        return FakeParam("String", _param_text(part, p), doc)

    if p in ("CP_Cable Qty",):
        return FakeParam("String", _param_text(part, p), doc)

    if p in ("CP_Total Cable Order Length",):
        try:
            rp = part.LookupParameter(p)
            if rp:
                return FakeParam("Double", rp.AsDouble(), doc)
        except Exception:
            pass
        return FakeParam("Double", 0.0, doc)

    # ---- fallback: try direct param lookup ----
    try:
        val = _param_text(part, p)
        if val:
            return FakeParam("String", val, doc)
    except Exception:
        pass

    return None


def get_param(elem, pname):
    """Return a real Revit Parameter (RFA) or a FakeParam (ITM) for the requested name."""
    if not pname:
        return None

    # ITM mode
    try:
        if FabricationPart is not None and isinstance(elem, FabricationPart):
            return itm_get_param(elem, pname)
    except Exception:
        pass

    # RFA fallback
    try:
        p = elem.LookupParameter(pname)
        if p:
            return p
    except Exception:
        pass
    try:
        if isinstance(elem, FamilyInstance) and elem.Symbol:
            return elem.Symbol.LookupParameter(pname)
    except Exception:
        pass
    return None


# =============================================================================
# Division Enums
# =============================================================================

DIV_MAD = "MAD"
DIV_MTW = "MTW"
DIV_NORCAL = "NorCal"
DIV_NORCALTRAP = "NorCalTrap"
DIV_PNW = "PNW"
DIV_SOCAL = "SoCal"


# =============================================================================
# Mappings (unchanged from SI Tools DLL — ITM adaptation is in itm_get_param)
# =============================================================================

def mappings_MAD():
    return [
        COMMON_HANGER_VERSION,
        COMMON_ELEMENT_ID,
        COMMON_ELEMENT_LOCATION,

        MappingItem("MARK", "Mark"),
        MappingItem("ITEM NUMBER", "CP_Item Number"),
        MappingItem("SPOOL", "Assembly Name"),

        MappingItem(
            "HANGER PREFIX",
            "CP_Hanger Type Abbreviation",
            pname_map=lambda h: "CP_Hanger Type Abbreviation" if hanger_symbol_has_si_version_parameter(h) else "CP_Prefix"
        ),

        MappingItem("HANGER NAME", "Family"),
        MappingItem("HANGER TYPE", "Type"),
        MappingItem("SERVICE", "CP_Service Abbv"),
        MappingItem("LEVEL", "Base Level"),
        MappingItem("AREA", "CP_Construction Area"),

        MappingItem("HANGER SIZE (DECIMAL IN.)", "CP_Hanger Size", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("HANGER SIZE (IN.)", "CP_Hanger Size", vfmt=MappingItem.vfmt_fractional_inches_round_to_1_8),

        MappingItem(
            "TRAPEZE LENGTH (DECIMAL IN.)",
            "CP_Trapeze Width Actual",
            pname_map=lambda h: "CP_Trapeze Width Actual" if hanger_symbol_has_si_version_parameter(h) else "CP_Length",
            vfmt=MappingItem.vfmt_decimal_whole_inches
        ),
        MappingItem(
            "TRAPEZE LENGTH (IN.)",
            "CP_Trapeze Width Actual",
            pname_map=lambda h: "CP_Trapeze Width Actual" if hanger_symbol_has_si_version_parameter(h) else "CP_Length",
            vfmt=MappingItem.vfmt_fractional_whole_inches
        ),

        MappingItem("STRUT QTY", "CP_Strut QTY", vfmt=MappingItem.vfmt_integer_or_zero),

        MappingItem("ROD SIZE (DECIMAL IN.)", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("ROD SIZE (FT. IN.)", "CP_Rod Diameter", vfmt=MappingItem.vfmt_fractional_inches_round_to_1_8),

        MappingItem("REVIT ROD LENGTH (DECIMAL IN.)", vfmt=MappingItem.vfmt_mad_revit_rod_length_decimal_inches),
        MappingItem("REVIT ROD LENGTH (FT. IN.)", vfmt=MappingItem.vfmt_mad_revit_rod_length_ft_fractional_inches),

        MappingItem(
            "HANGER ELEVATION (FT.)",
            "CP_Hung Object Bottom Elev",
            pname_map=lambda h: "CP_Hung Object Bottom Elev" if hanger_symbol_has_si_version_parameter(h) else "CP_Bottom Elevation",
            vfmt=MappingItem.vfmt_three_decimal_feet
        ),
        MappingItem(
            "HANGER ELEVATION (FT. IN.)",
            "CP_Hung Object Bottom Elev",
            pname_map=lambda h: "CP_Hung Object Bottom Elev" if hanger_symbol_has_si_version_parameter(h) else "CP_Bottom Elevation",
            vfmt=MappingItem.vfmt_feet_fractional_inches_round_to_1_8
        ),

        MappingItem(
            "REVIT HANGER FAB (DECIMAL IN.)",
            "CP_Fab Height",
            pname_map=lambda h: "CP_Fab Height" if hanger_symbol_has_si_version_parameter(h) else "CP_Hanger Fab Height",
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),
        MappingItem(
            "REVIT HANGER FAB (IN.)",
            "CP_Fab Height",
            pname_map=lambda h: "CP_Fab Height" if hanger_symbol_has_si_version_parameter(h) else "CP_Hanger Fab Height",
            vfmt=MappingItem.vfmt_fractional_inches_round_to_1_8
        ),

        MappingItem(
            "ROD QTY",
            "CP_Rod QTY",
            pname_map=lambda h: "CP_Rod QTY" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Count",
            vfmt=MappingItem.vfmt_integer_or_zero
        ),

        MappingItem("COMMENTS", "Comments"),
        MappingItem("LEFT ATTACHMENT TYPE", "CP_Rod Left Attachment"),
        MappingItem("MIDDLE ATTACHMENT TYPE", "CP_Rod Middle Attachment"),
        MappingItem("RIGHT ATTACHMENT TYPE", "CP_Rod Right Attachment"),
    ]


def mappings_MTW():
    return [
        COMMON_HANGER_VERSION,
        COMMON_ELEMENT_ID,
        COMMON_ELEMENT_LOCATION,

        MappingItem("Spool", "Assembly Name"),
        MappingItem("Item Number", "CP_Item Number"),
        MappingItem("Hanger Size (in)", "CP_Hanger Size", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Hanger Size (ft)", "CP_Hanger Size"),

        MappingItem("Rod Size", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Family", "Family"),

        MappingItem(
            "Fab Length",
            "CP_Fab Height",
            pname_map=lambda h: "CP_Fab Height" if hanger_symbol_has_si_version_parameter(h) else "CP_Hanger Fab Height",
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Rod Length Left", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length Middle", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length Right", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Rod Length Left (ft)", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_feet),
        MappingItem("Rod Length Middle (ft)", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_feet),
        MappingItem("Rod Length Right (ft)", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_feet),

        MappingItem(
            "Bottom Elevation",
            "CP_Hung Object Bottom Elev",
            pname_map=lambda h: "CP_Hung Object Bottom Elev" if hanger_symbol_has_si_version_parameter(h) else "CP_Bottom Elevation"
        ),

        MappingItem("Service", "CP_Service"),

        MappingItem(
            "Upper Attachment Left",
            "CP_Left Attachment Annotation",
            pname_map=lambda h: "CP_Left Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Left Attachment"
        ),
        MappingItem(
            "Upper Attachment Mid",
            "CP_Middle Attachment Annotation",
            pname_map=lambda h: "CP_Middle Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Middle Attachment"
        ),
        MappingItem(
            "Upper Attachment Right",
            "CP_Right Attachment Annotation",
            pname_map=lambda h: "CP_Right Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Right Attachment"
        ),

        MappingItem(
            "Rod Count",
            "CP_Rod QTY",
            pname_map=lambda h: "CP_Rod QTY" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Count",
            vfmt=MappingItem.vfmt_integer_or_zero
        ),

        MappingItem("Count", vfmt=MappingItem.vfmt_one_count),
        MappingItem("Level", "Base Level"),
        MappingItem("Comments", "Comments"),

        MappingItem(
            "Unistrut Width",
            "CP_Trapeze Width Actual",
            pname_map=lambda h: "CP_Trapeze Width Actual" if hanger_symbol_has_si_version_parameter(h) else "CP_Length"
        ),
        MappingItem(
            "Unistrut Width (ft)",
            "CP_Trapeze Width Actual",
            pname_map=lambda h: "CP_Trapeze Width Actual" if hanger_symbol_has_si_version_parameter(h) else "CP_Length",
            vfmt=MappingItem.vfmt_three_decimal_feet
        ),

        MappingItem("Longest Rod Length (ft)", vfmt=MappingItem.vfmt_mtw_longest_rod_length_ft),
        MappingItem("Cable Count", "CP_Cable Qty"),
        MappingItem("Cable Length", "CP_Total Cable Order Length", vfmt=MappingItem.vfmt_three_decimal_feet),
    ]


def mappings_NorCal():
    return [
        COMMON_HANGER_VERSION,

        MappingItem("Item No", "CP_Item Number"),
        MappingItem("Rod Size (decimal inches)", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Middle Length (decimal inches)", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Description", "CP_Hanger Size", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem(
            "Attachment Type Middle",
            "CP_Middle Attachment Annotation",
            pname_map=lambda h: "CP_Middle Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Middle Attachment"
        ),

        MappingItem("Source", "Family"),

        MappingItem(
            "E-BOH (decimal inches)",
            "CP_Fab Height",
            pname_map=lambda h: "CP_Fab Height" if hanger_symbol_has_si_version_parameter(h) else "CP_Hanger Fab Height",
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Level", "Base Level"),
        MappingItem("Area", "CP_Construction Area"),
        MappingItem("Service", "CP_Service Abbv"),

        MappingItem(
            "Elevation F.F. - B.O.P. (decimal inches)",
            "CP_Hung Object Bottom Elev",
            pname_map=lambda h: "CP_Hung Object Bottom Elev" if hanger_symbol_has_si_version_parameter(h) else "CP_Bottom Elevation",
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Coating Type", "CP_Coating Type"),
    ]


def mappings_NorCalTrap():
    def left_attach_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Left Attachment Annotation"
        elif of_si_version(h, "1.0"):
            return "CP_Left Attachment Annotation"
        elif of_si_version(h, "Legacy"):
            return "CP_Rod Left Attachment"
        return None

    def left_offset_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Rod Left Offset From End Desired"
        elif of_si_version(h, "1.0"):
            return "Rod Left Offset From End Desired"
        elif of_si_version(h, "Legacy"):
            return "Rod Left Offset From End Desired"
        return None

    def right_attach_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Right Attachment Annotation"
        elif of_si_version(h, "1.0"):
            return "CP_Right Attachment Annotation"
        elif of_si_version(h, "Legacy"):
            return "CP_Rod Right Attachment"
        return None

    def right_offset_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Rod Right Offset From End Desired"
        elif of_si_version(h, "1.0"):
            return "Rod Right Offset From End Desired"
        elif of_si_version(h, "Legacy"):
            return "Rod Right Offset From End Desired"
        return None

    def unistrut_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Trapeze Width Actual"
        elif of_si_version(h, "1.0"):
            return "CP_Trapeze Width Actual"
        elif of_si_version(h, "Legacy"):
            return "CP_Length"
        return None

    def elev_param(h):
        if of_si_version(h, "2.0"):
            return "CP_Hung Object Bottom Elev"
        elif of_si_version(h, "1.0"):
            return "CP_Hung Object Bottom Elev"
        elif of_si_version(h, "Legacy"):
            return "CP_Bottom Elevation"
        return None

    return [
        COMMON_HANGER_VERSION,

        MappingItem("Item No", "CP_Item Number"),

        MappingItem("Rod Size_1", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length_1", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Attachment Type_1", "CP_Left Attachment Annotation", pname_map=left_attach_param),

        MappingItem(
            "Rod Offset from End_1",
            "CP_Rod Left Offset From End Desired",
            pname_map=left_offset_param,
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Rod Size_2", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length_2", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Attachment Type_2", "CP_Right Attachment Annotation", pname_map=right_attach_param),

        MappingItem(
            "Rod Offset from End_2",
            "CP_Rod Right Offset From End Desired",
            pname_map=right_offset_param,
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Source", "Family and Type"),

        MappingItem("Unistrut Width", "CP_Trapeze Width Actual", pname_map=unistrut_param, vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Level", "Base Level"),
        MappingItem("Area", "CP_Construction Area"),
        MappingItem("Service", "CP_Service Abbv"),

        MappingItem("Elevation F.F. - B.O.P.", "CP_Hung Object Bottom Elev", pname_map=elev_param, vfmt=MappingItem.vfmt_three_decimal_inches),
    ]


def mappings_PNW():
    return [
        COMMON_HANGER_VERSION,
        COMMON_ELEMENT_ID,
        COMMON_ELEMENT_LOCATION,

        MappingItem("Spool", "Assembly Name"),
        MappingItem("Item Number", "CP_Item Number"),
        MappingItem("Hanger Size (in)", "CP_Hanger Size", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Hanger Size (ft)", "CP_Hanger Size"),

        MappingItem("Rod Size", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Family", "Family"),

        MappingItem(
            "Fab Length",
            "CP_Fab Height",
            pname_map=lambda h: "CP_Fab Height" if hanger_symbol_has_si_version_parameter(h) else "CP_Hanger Fab Height",
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Rod Length (in) Left", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length (in) Middle", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length (in) Right", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Rod Length (ft) Left", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_feet),
        MappingItem("Rod Length (ft) Middle", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_feet),
        MappingItem("Rod Length (ft) Right", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_feet),

        MappingItem(
            "Bottom Elevation",
            "CP_Hung Object Bottom Elev",
            pname_map=lambda h: "CP_Hung Object Bottom Elev" if hanger_symbol_has_si_version_parameter(h) else "CP_Bottom Elevation"
        ),

        MappingItem("Service", "CP_Service"),

        MappingItem(
            "Upper Attachment Left",
            "CP_Left Attachment Annotation",
            pname_map=lambda h: "CP_Left Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Left Attachment"
        ),
        MappingItem(
            "Upper Attachment Middle",
            "CP_Middle Attachment Annotation",
            pname_map=lambda h: "CP_Middle Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Middle Attachment"
        ),
        MappingItem(
            "Upper Attachment Right",
            "CP_Right Attachment Annotation",
            pname_map=lambda h: "CP_Right Attachment Annotation" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Right Attachment"
        ),

        MappingItem(
            "Rod Count",
            "CP_Rod QTY",
            pname_map=lambda h: "CP_Rod QTY" if hanger_symbol_has_si_version_parameter(h) else "CP_Rod Count",
            vfmt=MappingItem.vfmt_integer_or_zero
        ),

        MappingItem("Count", vfmt=MappingItem.vfmt_one_count),
        MappingItem("Level", "Base Level"),
        MappingItem("Comments", "Comments"),
    ]


def mappings_SoCal():
    def eboh_or_unistrut_param(h):
        if hanger_symbol_has_si_version_parameter(h):
            return "CP_Trapeze Width Actual" if is_socal_specialty_hanger(h) else "CP_Fab Height"
        else:
            return "CP_Length" if is_socal_specialty_hanger(h) else "CP_Hanger Fab Height"

    return [
        COMMON_HANGER_VERSION,
        COMMON_ELEMENT_ID,
        COMMON_ELEMENT_LOCATION,

        MappingItem("Item No", "CP_Item Number"),
        MappingItem("Rod Size", "CP_Rod Diameter", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length Left", "CP_Rod Left Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length Middle", "CP_Rod Middle Length", vfmt=MappingItem.vfmt_three_decimal_inches),
        MappingItem("Rod Length Right", "CP_Rod Right Length", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem("Description", "CP_Hanger Size", vfmt=MappingItem.vfmt_three_decimal_inches),

        MappingItem(
            "Pipe Size (w/o Insulation)",
            "CP_Host Nominal Diameter",
            pname_map=lambda h: "CP_Host Nominal Diameter" if hanger_symbol_has_si_version_parameter(h) else "Host Nominal Diameter",
            vfmt=MappingItem.vfmt_three_decimal_feet
        ),

        MappingItem("Source", "Family"),

        MappingItem(
            "E-BOH or Unistrut Width",
            "CP_Trapeze Width Actual",
            pname_map=eboh_or_unistrut_param,
            vfmt=MappingItem.vfmt_three_decimal_inches
        ),

        MappingItem("Level", "Base Level"),
        MappingItem("Area", "CP_Construction Area"),
        MappingItem("Service", "CP_Service Name"),
        MappingItem("Elevation F.F. - B.O.P.", "CP_Hung Object Bottom Elev",
                    vfmt=MappingItem.vfmt_three_decimal_inches),
    ]


# =============================================================================
# Division dropdown values
# =============================================================================

DIVISION_UI = [
    "NorCal Division Trapeze",
    "Mid Atlantic Division",
    "Mountain West Division",
    "NorCal Division Non Trapeze",
    "Pacific Northwest Division",
    "SoCal Division",
]

DIVISION_UI_TO_ENUM_AND_MAPPING = {
    "Mid Atlantic Division": (DIV_MAD, mappings_MAD),
    "Mountain West Division": (DIV_MTW, mappings_MTW),
    "NorCal Division Non Trapeze": (DIV_NORCAL, mappings_NorCal),
    "NorCal Division Trapeze": (DIV_NORCALTRAP, mappings_NorCalTrap),
    "Pacific Northwest Division": (DIV_PNW, mappings_PNW),
    "SoCal Division": (DIV_SOCAL, mappings_SoCal),
}


# =============================================================================
# CSV writing helpers
# =============================================================================

def write_line_binary(fh, s):
    if s is None:
        s = ""
    if not isinstance(s, basestring):
        s = str(s)
    fh.write(s.encode("utf-8"))
    fh.write("\r\n")


def default_filename_like_si_tools(doc_title):
    now = datetime.now()
    day = str(now.day)
    mon = now.strftime("%b")
    yr = now.strftime("%Y")
    timepart = now.strftime("%H%M%S")
    return "{} {}{}{} {}".format(doc_title, day, mon, yr, timepart)


# =============================================================================
# Main execution
# =============================================================================

uidoc = revit.uidoc
doc = revit.doc

if FabricationPart is None:
    forms.alert("FabricationPart API not available in this Revit build.", title="ITM Export")
    raise SystemExit

# 1) Pick ITM hangers
try:
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        ItmHangerSelectionFilter(),
        "Select ITM hangers to export"
    )
except Exception:
    forms.alert("Cancelled.")
    raise SystemExit

selected = []
for r in refs:
    e = doc.GetElement(r)
    try:
        if isinstance(e, FabricationPart):
            selected.append(e)
    except Exception:
        pass

forms.alert("There were {} hangers selected.".format(len(selected)), title="Count")
if not selected:
    raise SystemExit

# 2) Division dropdown
division_choice = pick_division_winforms(
    DIVISION_UI,
    default_value="NorCal Division Trapeze",
    title="Selection",
    prompt="Select your Division:"
)
if not division_choice:
    forms.alert("Cancelled.")
    raise SystemExit

if division_choice not in DIVISION_UI_TO_ENUM_AND_MAPPING:
    forms.alert(
        "Division mapping not wired yet for:\n{}".format(division_choice),
        title="ITM Export"
    )
    raise SystemExit

division_enum, mapping_fn = DIVISION_UI_TO_ENUM_AND_MAPPING[division_choice]
mapping_items = mapping_fn()

# 3) Build output
utc_stamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%SZ")
title_row = "{} Hanger BOM:,{},{}".format(division_enum, doc.Title, utc_stamp)

header_row = ",".join([safe_cell(m.Header) for m in mapping_items])

# 4) Save file dialog
default_name = default_filename_like_si_tools(doc.Title)
save_path = forms.save_file(file_ext="csv", default_name=default_name)
if not save_path:
    forms.alert("Cancelled.")
    raise SystemExit

# 5) Write CSV
with open(save_path, "wb") as f:
    write_line_binary(f, title_row)
    write_line_binary(f, header_row)
    for h in selected:
        row = ",".join([m.eval(h) for m in mapping_items])
        write_line_binary(f, row)
    write_line_binary(f, "TOTAL HANGERS: {}".format(len(selected)))

forms.alert("Exported:\n{}".format(save_path), title="Done")
