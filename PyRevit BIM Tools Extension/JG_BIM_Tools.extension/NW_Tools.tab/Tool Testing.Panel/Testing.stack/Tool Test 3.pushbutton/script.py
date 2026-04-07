# -*- coding: utf-8 -*-
# pyRevit IronPython 2.7
# Hanger BOM Exporter (SI Tools V3 behavior replicated from ILSpy)

from pyrevit import revit, forms

from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (
    BuiltInCategory, FamilyInstance, FamilySymbol, Family, Element,
    LocationPoint, SpecTypeId, UnitTypeId, Units, FormatOptions, UnitSystem,
    UnitFormatUtils
)

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

    # Make it look like normal Windows dialogs (fixes "font too big" complaints)
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
# Hanger placement logic (mirrors HangerPlacementService.IsGenericModelHanger)
# =============================================================================

def _resolve_family(elem):
    fam = None
    if isinstance(elem, FamilyInstance):
        try:
            fam = elem.Symbol.Family
        except Exception:
            fam = None
    elif isinstance(elem, FamilySymbol):
        try:
            fam = elem.Family
        except Exception:
            fam = None
    elif isinstance(elem, Family):
        fam = elem
    return fam


def is_generic_model_hanger(elem):
    """
    DLL logic:
      familyName.Contains("Hanger") && category.IsBuiltInCategory(OST_GenericModel)
    """
    fam = _resolve_family(elem)
    if not fam:
        return False
    try:
        fam_name = fam.Name or ""
        cat = fam.FamilyCategory
        if not cat:
            return False
        is_generic_model = (cat.Id.IntegerValue == int(BuiltInCategory.OST_GenericModel))  # -2000151
        return ("Hanger" in fam_name) and is_generic_model
    except Exception:
        return False


class HangerRefreshSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        if not isinstance(element, FamilyInstance):
            return False
        return is_generic_model_hanger(element)

    def AllowReference(self, reference, point):
        return False


# =============================================================================
# Parameter helpers (mimic ParameterUtility.GetParameterFromInstanceOrType)
# =============================================================================

def get_param(elem, pname):
    if not pname:
        return None
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
    try:
        return h.Symbol.Family.Name
    except Exception:
        return ""


def get_type_name(h):
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
            # Best-effort raw conversion
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
# Division Enums (string names to match SI_CorporateDivision.ToString())
# =============================================================================

DIV_MAD = "MAD"
DIV_MTW = "MTW"
DIV_NORCAL = "NorCal"
DIV_NORCALTRAP = "NorCalTrap"
DIV_PNW = "PNW"
DIV_SOCAL = "SoCal"


# =============================================================================
# Mappings (ported directly from your ILSpy dump)
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
    # PNW mapping list matches your ILSpy dump (very similar to MTW)
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
        # From ILSpy:
        # Has SI_Version ?
        #   IsSoCalSpecialtyHanger ? "CP_Trapeze Width Actual" : "CP_Fab Height"
        # else
        #   IsSoCalSpecialtyHanger ? "CP_Length" : "CP_Hanger Fab Height"
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
        MappingItem("Elevation F.F. - B.O.P.", "CP_Hung Object Bottom Elev"),
    ]


# =============================================================================
# Division dropdown values (UI labels) + mapping to division enum + mapping list
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
# CSV writing helpers (IronPython-safe; enforces Windows CRLF)
# =============================================================================

def write_line_binary(fh, s):
    if s is None:
        s = ""
    if not isinstance(s, basestring):
        s = str(s)
    fh.write(s.encode("utf-8"))
    fh.write("\r\n")


def default_filename_like_si_tools(doc_title):
    # SI Tools: base.Document.Title + " " + DateTime.Now.ToString("dMMMyyyy HHmmss")
    # .NET "d" => day without leading zero. We'll reproduce that exactly.
    now = datetime.now()
    day = str(now.day)  # no leading zero
    mon = now.strftime("%b")  # Feb, Mar, etc (depends on OS culture; matches typical SI environment)
    yr = now.strftime("%Y")
    timepart = now.strftime("%H%M%S")
    return "{} {}{}{} {}".format(doc_title, day, mon, yr, timepart)


# =============================================================================
# Main execution (mirrors HangerBOMExecute pipeline)
# =============================================================================

uidoc = revit.uidoc
doc = revit.doc

# 1) Pick hangers
try:
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        HangerRefreshSelectionFilter(),
        "Select hangers to export"
    )
except Exception:
    forms.alert("Cancelled.")
    raise SystemExit

selected = []
for r in refs:
    e = doc.GetElement(r)
    if isinstance(e, FamilyInstance) and is_generic_model_hanger(e):
        selected.append(e)

forms.alert("There were {} hangers selected.".format(len(selected)), title="Count")
if not selected:
    raise SystemExit

# 2) Division dropdown (WinForms)
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
        "Division mapping not wired yet for:\n{}\n\nTell me which SI_CorporateDivision this should map to and I'll plug it in."
        .format(division_choice),
        title="Tool Test 3"
    )
    raise SystemExit

division_enum, mapping_fn = DIVISION_UI_TO_ENUM_AND_MAPPING[division_choice]
mapping_items = mapping_fn()

# 3) Build output text EXACTLY like SI Tools
# SI Tools: selectedDivision.ToString() + " Hanger BOM:," + Document.Title + "," + DateTime.Now.ToString("u")
utc_stamp = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%SZ")  # "u" style
title_row = "{} Hanger BOM:,{},{}".format(division_enum, doc.Title, utc_stamp)

header_row = ",".join([safe_cell(m.Header) for m in mapping_items])

# 4) Save file dialog (pyRevit) - matches typical behavior; file content is what matters for pivot
default_name = default_filename_like_si_tools(doc.Title)
save_path = forms.save_file(file_ext="csv", default_name=default_name)
if not save_path:
    forms.alert("Cancelled.")
    raise SystemExit

# 5) Write CSV (binary; CRLF)
with open(save_path, "wb") as f:
    write_line_binary(f, title_row)
    write_line_binary(f, header_row)
    for h in selected:
        row = ",".join([m.eval(h) for m in mapping_items])
        write_line_binary(f, row)
    write_line_binary(f, "TOTAL HANGERS: {}".format(len(selected)))

forms.alert("Exported:\n{}".format(save_path), title="Done")