# -*- coding: utf-8 -*-
"""
NW BOM - Column Settings
========================

Profile-based manager for user-defined extra export columns.

Profiles are stored as individual JSON files under:
    AppData/Roaming/pyRevit/NW_BOM/profiles/

The active profile name is tracked in:
    AppData/Roaming/pyRevit/NW_BOM/active_profile.json

Only one profile is active at a time. The BOM run button reads whichever
profile is currently active. If no profile is active, only the standard
built-in columns are exported.

No Revit transaction is opened -- this is pure UI / file I/O.
"""

from pyrevit import revit, forms
import clr
import os
import json
import shutil

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

import System.Windows.Forms as SWF
import System.Drawing as SD

Form              = SWF.Form
Label             = SWF.Label
TextBox           = SWF.TextBox
ComboBox          = SWF.ComboBox
Button            = SWF.Button
ListBox           = SWF.ListBox
DialogResult      = SWF.DialogResult
MessageBox        = SWF.MessageBox
MessageBoxButtons = SWF.MessageBoxButtons
MessageBoxIcon    = SWF.MessageBoxIcon
OpenFileDialog    = SWF.OpenFileDialog
SaveFileDialog    = SWF.SaveFileDialog
SelectionMode     = SWF.SelectionMode
FormBorderStyle   = SWF.FormBorderStyle
StartPosition     = SWF.FormStartPosition
AnchorStyles      = SWF.AnchorStyles
ToolTip           = SWF.ToolTip

Size      = SD.Size
Point     = SD.Point
Color     = SD.Color
Font      = SD.Font
FontStyle = SD.FontStyle


# =============================================================================
# DATA TYPES AND UNIT TABLES
# =============================================================================

EXTRA_COL_TYPES = [
    "text",
    "number",
    "integer",
    "boolean",
    "length",
    "weight / mass",
    "area",
    "volume",
    "flow rate",
    "pressure",
    "temperature",
    "angle",
    "velocity",
    "force",
]

# For each type that needs a unit picker: list of (display_label, internal_key)
# Revit internal storage units:
#   length   -> decimal feet
#   mass     -> lbs (pounds-mass)
#   area     -> sq ft
#   volume   -> cubic feet
#   flow     -> cubic feet per second
#   pressure -> lb/ft2 (PSF)
#   temp     -> Fahrenheit
#   angle    -> radians
#   velocity -> feet per second
#   force    -> lbf (pound-force)
EXTRA_COL_UNITS = {
    "length": [
        ("Feet-Inches fractional  (e.g. 12'-6 1/2\")",  "ft_in_frac"),
        ("Decimal feet  (e.g. 12.542)",                  "ft_decimal"),
        ("Decimal inches  (e.g. 150.500)",               "in_decimal"),
        ("Millimeters  (e.g. 3823)",                     "mm"),
        ("Centimeters  (e.g. 382.3)",                    "cm"),
        ("Meters  (e.g. 3.823)",                         "m"),
    ],
    "weight / mass": [
        ("Pounds  (lbs)",                                "lbs"),
        ("Kilograms  (kg)",                              "kg"),
        ("Tonnes / metric tons  (t)",                    "tonne"),
        ("Ounces  (oz)",                                 "oz"),
        ("Grams  (g)",                                   "g"),
    ],
    "area": [
        ("Square feet  (ft2)",                           "sqft"),
        ("Square inches  (in2)",                         "sqin"),
        ("Square meters  (m2)",                          "sqm"),
        ("Square millimeters  (mm2)",                    "sqmm"),
        ("Square centimeters  (cm2)",                    "sqcm"),
    ],
    "volume": [
        ("Cubic feet  (ft3)",                            "cuft"),
        ("Cubic inches  (in3)",                          "cuin"),
        ("Cubic meters  (m3)",                           "cum"),
        ("Liters  (L)",                                  "liters"),
        ("US Gallons  (gal)",                            "usgal"),
        ("Milliliters  (mL)",                            "ml"),
    ],
    "flow rate": [
        ("Cubic feet per second  (ft3/s)",               "cfs"),
        ("Cubic feet per minute  (CFM)",                 "cfm"),
        ("Gallons per minute  (GPM)",                    "gpm"),
        ("Liters per second  (L/s)",                     "lps"),
        ("Liters per minute  (L/min)",                   "lpm"),
        ("Cubic meters per hour  (m3/h)",                "cmh"),
    ],
    "pressure": [
        ("Pounds per sq inch  (PSI)",                    "psi"),
        ("Pounds per sq foot  (PSF)",                    "psf"),
        ("Pascals  (Pa)",                                "pa"),
        ("Kilopascals  (kPa)",                           "kpa"),
        ("Bar",                                          "bar"),
        ("Inches of water column  (inWC)",               "inwc"),
        ("Inches of mercury  (inHg)",                    "inhg"),
    ],
    "temperature": [
        ("Fahrenheit  (F)",                              "degF"),
        ("Celsius  (C)",                                 "degC"),
        ("Kelvin  (K)",                                  "kelvin"),
        ("Rankine  (R)",                                 "rankine"),
    ],
    "angle": [
        ("Degrees  (deg)",                               "degrees"),
        ("Radians  (rad)",                               "radians"),
        ("Gradians  (grad)",                             "gradians"),
    ],
    "velocity": [
        ("Feet per second  (ft/s)",                      "fps"),
        ("Feet per minute  (FPM)",                       "fpm"),
        ("Miles per hour  (MPH)",                        "mph"),
        ("Meters per second  (m/s)",                     "mps"),
        ("Kilometers per hour  (km/h)",                  "kmh"),
    ],
    "force": [
        ("Pound-force  (lbf)",                           "lbf"),
        ("Kilonewtons  (kN)",                            "kn"),
        ("Newtons  (N)",                                 "newtons"),
        ("Kilogram-force  (kgf)",                        "kgf"),
    ],
}

# Revit internal -> display unit conversion factors.
# None = handled by special-case code in script.py
UNIT_FACTORS = {
    "ft_in_frac":  None,
    "ft_decimal":  1.0,
    "in_decimal":  12.0,
    "mm":          304.8,
    "cm":          30.48,
    "m":           0.3048,
    "lbs":         1.0,
    "kg":          0.453592,
    "tonne":       0.000453592,
    "oz":          16.0,
    "g":           453.592,
    "sqft":        1.0,
    "sqin":        144.0,
    "sqm":         0.092903,
    "sqmm":        92903.04,
    "sqcm":        929.0304,
    "cuft":        1.0,
    "cuin":        1728.0,
    "cum":         0.0283168,
    "liters":      28.3168,
    "usgal":       7.48052,
    "ml":          28316.8,
    "cfs":         1.0,
    "cfm":         60.0,
    "gpm":         448.831,
    "lps":         28.3168,
    "lpm":         1699.01,
    "cmh":         101.941,
    "psi":         0.00694444,
    "psf":         1.0,
    "pa":          47.8803,
    "kpa":         0.0478803,
    "bar":         0.000478803,
    "inwc":        0.192221,
    "inhg":        0.014139,
    "degF":        None,   # direct value
    "degC":        None,   # (F-32)*5/9
    "kelvin":      None,   # (F-32)*5/9 + 273.15
    "rankine":     None,   # F + 459.67
    "degrees":     57.2958,
    "radians":     1.0,
    "gradians":    63.6620,
    "fps":         1.0,
    "fpm":         60.0,
    "mph":         0.681818,
    "mps":         0.3048,
    "kmh":         1.09728,
    "lbf":         1.0,
    "kn":          0.00444822,
    "newtons":     4.44822,
    "kgf":         0.453592,
}

TYPE_NOTES = {
    "text":          "Read as string. Works for most text and shared parameters.",
    "number":        "Read as decimal number. No unit conversion applied.",
    "integer":       "Read as whole number (good for count parameters).",
    "boolean":       "Integer parameter: 1 = Yes, 0 = No.",
    "length":        "Revit stores length as decimal feet internally.",
    "weight / mass": "Revit stores mass as pounds (lbs) internally.",
    "area":          "Revit stores area as square feet internally.",
    "volume":        "Revit stores volume as cubic feet internally.",
    "flow rate":     "Revit stores flow as cubic feet per second (CFS) internally.",
    "pressure":      "Revit stores pressure as lb/ft2 (PSF) internally.",
    "temperature":   "Revit stores temperature as Fahrenheit internally.",
    "angle":         "Revit stores angles as radians internally.",
    "velocity":      "Revit stores velocity as feet per second internally.",
    "force":         "Revit stores force as pound-force (lbf) internally.",
}


# =============================================================================
# PROFILE STORAGE PATHS
# =============================================================================

def _profiles_dir():
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    folder = os.path.join(appdata, "pyRevit", "NW_BOM", "profiles")
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder)
        except OSError:
            pass
    return folder


def _active_profile_path():
    appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
    return os.path.join(appdata, "pyRevit", "NW_BOM", "active_profile.json")


def _profile_path(profile_name):
    safe = profile_name.replace("/", "-").replace("\\", "-").replace(":", "-")
    return os.path.join(_profiles_dir(), safe + ".json")


# =============================================================================
# JSON HELPERS  (IronPython 2.7 safe -- binary open, explicit UTF-8)
# =============================================================================

def _read_json(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        raw = raw[3:]
    return json.loads(raw.decode("utf-8", errors="replace"))


def _write_json(path, data):
    text = json.dumps(data, indent=2, ensure_ascii=True)
    with open(path, "wb") as fh:
        fh.write(text.encode("utf-8"))


# =============================================================================
# PROFILE I/O
# =============================================================================

def list_profiles():
    folder = _profiles_dir()
    names = []
    try:
        for fn in os.listdir(folder):
            if fn.lower().endswith(".json"):
                names.append(fn[:-5])
    except OSError:
        pass
    return sorted(names, key=lambda s: s.lower())


def get_active_profile():
    path = _active_profile_path()
    if not os.path.isfile(path):
        return None
    try:
        data = _read_json(path)
        return data.get("active") or None
    except Exception:
        return None


def set_active_profile(name):
    path = _active_profile_path()
    parent = os.path.dirname(path)
    if not os.path.isdir(parent):
        try:
            os.makedirs(parent)
        except OSError:
            pass
    _write_json(path, {"active": name})


def load_profile(profile_name):
    path = _profile_path(profile_name)
    if not os.path.isfile(path):
        return []
    try:
        data = _read_json(path)
        return [ExtraColumnConfig.from_dict(d) for d in data.get("extra_columns", [])]
    except Exception as ex:
        MessageBox.Show(
            "Could not load profile '{}':\n\n{}".format(profile_name, ex),
            "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Warning
        )
        return []


def save_profile(profile_name, configs):
    path = _profile_path(profile_name)
    _write_json(path, {"extra_columns": [c.to_dict() for c in configs]})


def delete_profile(profile_name):
    path = _profile_path(profile_name)
    if os.path.isfile(path):
        os.remove(path)
    if get_active_profile() == profile_name:
        set_active_profile(None)


def rename_profile(old_name, new_name):
    old_path = _profile_path(old_name)
    new_path = _profile_path(new_name)
    if os.path.isfile(old_path):
        os.rename(old_path, new_path)
    if get_active_profile() == old_name:
        set_active_profile(new_name)


def duplicate_profile(source_name, new_name):
    src = _profile_path(source_name)
    dst = _profile_path(new_name)
    if os.path.isfile(src):
        shutil.copy2(src, dst)


def load_active_columns():
    """Public entry point used by script.py on each BOM run."""
    active = get_active_profile()
    if not active:
        return []
    return load_profile(active)


# =============================================================================
# DATA MODEL
# =============================================================================

class ExtraColumnConfig(object):
    def __init__(self, param_name="", display_name="", data_type="text", unit_key=None):
        self.param_name   = param_name
        self.display_name = display_name
        self.data_type    = data_type
        self.unit_key     = unit_key

    def to_dict(self):
        return {
            "param_name":   self.param_name,
            "display_name": self.display_name,
            "data_type":    self.data_type,
            "unit_key":     self.unit_key,
        }

    @staticmethod
    def from_dict(d):
        return ExtraColumnConfig(
            param_name   = d.get("param_name", ""),
            display_name = d.get("display_name", ""),
            data_type    = d.get("data_type", "text"),
            unit_key     = d.get("unit_key", None),
        )

    def label(self):
        unit_suffix = " [{}]".format(self.unit_key) if self.unit_key else ""
        return "{display}   <-   {param}   ({dtype}{unit})".format(
            display = self.display_name or "(no display name)",
            param   = self.param_name   or "(no param name)",
            dtype   = self.data_type,
            unit    = unit_suffix,
        )


# =============================================================================
# PARAMETER BROWSER
# =============================================================================

def collect_param_names(doc):
    names = set()
    try:
        it = doc.ParameterBindings.ForwardIterator()
        it.Reset()
        while it.MoveNext():
            try:
                names.add(it.Key.Name)
            except Exception:
                pass
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import FilteredElementCollector
        collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
        count = 0
        for el in collector:
            if count > 500:
                break
            try:
                for p in el.Parameters:
                    try:
                        n = p.Definition.Name
                        if n:
                            names.add(n)
                    except Exception:
                        pass
                count += 1
            except Exception:
                pass
    except Exception:
        pass
    return sorted(names, key=lambda s: s.lower())


# =============================================================================
# HELP TEXT
# =============================================================================

HELP_PROFILE_MANAGER = (
    "PROFILE MANAGER  --  How to use\n"
    "================================\n"
    "\n"
    "WHAT IS A PROFILE?\n"
    "A profile is a named set of extra columns appended to every BOM\n"
    "export row. You can have as many profiles as you want -- one per\n"
    "job, one per scope, one for racks, one for hangers, etc.\n"
    "\n"
    "Only ONE profile is active at a time. The BOM run button always\n"
    "uses whichever profile is currently active. If no profile is\n"
    "active, only the standard built-in columns are exported.\n"
    "\n"
    "BUTTONS:\n"
    "  New profile            Create a blank profile and give it a name.\n"
    "  Duplicate              Copy the selected profile under a new name.\n"
    "  Rename                 Rename the selected profile.\n"
    "  Delete                 Permanently delete the selected profile.\n"
    "  Set active             Make this profile the one the BOM uses.\n"
    "  Deactivate all         Run the BOM with no extra columns.\n"
    "  Edit columns           Open the column editor for the selected profile.\n"
    "  Import profile         Load a profile from a file someone shared.\n"
    "  Export profile         Save the selected profile to share with others.\n"
    "  Set active and close   Activate the selected profile and close in\n"
    "                         one click -- ready to run the BOM immediately.\n"
    "\n"
    "THE ACTIVE BADGE:\n"
    "  The currently active profile shows [ACTIVE] next to its name.\n"
    "  The column count is also shown so you can see what is in each profile.\n"
    "  Only one profile can be active at a time.\n"
    "\n"
    "VALIDATION:\n"
    "  When you run the BOM, it checks whether your configured parameter\n"
    "  names exist on the selected elements. If a name is not found it\n"
    "  warns you before export so you can fix a typo before it wastes time.\n"
    "\n"
    "TYPICAL WORKFLOW:\n"
    "  1. Click 'New profile' and name it (e.g. 'Racks' or 'Job 1234').\n"
    "  2. Click 'Edit columns' to add the parameters you want.\n"
    "  3. Click 'Set active and close' -- then run the BOM button.\n"
    "  4. Extra columns appear at the right of the output.\n"
    "  5. Switch profiles anytime without losing any of them.\n"
)

HELP_COLUMN_EDITOR = (
    "COLUMN EDITOR  --  How to use\n"
    "==============================\n"
    "\n"
    "WHAT IS A COLUMN?\n"
    "Each column you add here appears as an extra column at the right\n"
    "of the BOM export, after all the standard built-in columns.\n"
    "\n"
    "The column reads a Revit parameter by name from each element and\n"
    "writes the value (with your chosen unit conversion) into the TSV.\n"
    "\n"
    "BUTTONS:\n"
    "  Add column       Add a new extra column to this profile.\n"
    "  Edit             Edit the selected column settings.\n"
    "  Remove           Remove the selected column.\n"
    "  Move up/down     Change the order columns appear in the export.\n"
    "  Save profile     Save all changes to this profile.\n"
    "\n"
    "ADDING A COLUMN:\n"
    "  Revit parameter name\n"
    "      The exact Revit parameter name (case-sensitive).\n"
    "      Click Browse to pick from parameters in the open model.\n"
    "      The tool checks the instance first, then the type element.\n"
    "\n"
    "  Display name\n"
    "      The column header in the TSV/Excel output.\n"
    "      Defaults to the parameter name if left blank.\n"
    "\n"
    "  Data type\n"
    "      What kind of value the parameter holds.\n"
    "        text        plain string, no conversion\n"
    "        number      raw decimal\n"
    "        integer     whole number (good for count params)\n"
    "        boolean     1=Yes, 0=No\n"
    "        length      Revit stores as decimal feet\n"
    "        weight      Revit stores as pounds (lbs)\n"
    "        area        Revit stores as square feet\n"
    "        volume      Revit stores as cubic feet\n"
    "        flow rate   Revit stores as cubic feet/second\n"
    "        pressure    Revit stores as lb/ft2 (PSF)\n"
    "        temperature Revit stores as Fahrenheit\n"
    "        angle       Revit stores as radians\n"
    "\n"
    "  Unit / format\n"
    "      Pick the output unit. Conversion from Revit internal\n"
    "      units is applied automatically.\n"
    "\n"
    "NOTES:\n"
    "  - Missing parameters get a blank cell (not N/A).\n"
    "  - Derived rows (hanger rods) always get blank cells.\n"
    "  - Column order in this list = column order in the export.\n"
    "  - Click 'Save profile' after editing.\n"
)


class HelpDialog(Form):
    """Help dialog styled like the Export Mode dialog -- plain labels, OK button."""

    def __init__(self, title, text):
        super(HelpDialog, self).__init__()
        self.Text            = title
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = StartPosition.CenterParent
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        pad  = 20
        lw   = 500    # label width
        lh   = 20     # line height -- must be >= font height + descent
        y    = pad

        lines = text.split("\n")
        for line in lines:
            lbl = Label()
            lbl.Text      = line if line else " "
            lbl.Location  = Point(pad, y)
            lbl.Size      = Size(lw, lh)
            lbl.ForeColor = Color.Black
            self.Controls.Add(lbl)
            y += lh

        y += 12  # gap before button

        btn = Button()
        btn.Text     = "OK"
        btn.Size     = Size(88, 28)
        btn.Location = Point((lw + pad * 2) / 2 - 44, y)
        btn.Click   += lambda s, e: self.Close()
        self.Controls.Add(btn)
        self.AcceptButton = btn
        self.CancelButton = btn

        self.ClientSize = Size(lw + pad * 2, y + btn.Height + pad)


# =============================================================================
# PARAMETER PICKER DIALOG
# =============================================================================

class ParamPickerDialog(Form):
    def __init__(self, names):
        super(ParamPickerDialog, self).__init__()
        self._all_names    = names
        self.selected_name = None
        self._build_ui()

    def _build_ui(self):
        self.Text            = "Select Parameter"
        self.Size            = Size(420, 520)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = StartPosition.CenterParent
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        pad = 12

        lbl = Label()
        lbl.Text     = "Search:"
        lbl.Location = Point(pad, pad)
        lbl.Size     = Size(52, 18)
        self.Controls.Add(lbl)

        self._txt_search = TextBox()
        self._txt_search.Location    = Point(pad + 58, pad - 2)
        self._txt_search.Size        = Size(316, 24)
        self._txt_search.TextChanged += self._on_search
        self.Controls.Add(self._txt_search)

        self._lst = ListBox()
        self._lst.Location      = Point(pad, pad + 34)
        self._lst.Size          = Size(380, 378)
        self._lst.SelectionMode = SelectionMode.One
        self._lst.DoubleClick  += self._on_ok
        self.Controls.Add(self._lst)
        self._refresh_list("")

        btn_ok = Button()
        btn_ok.Text     = "Select"
        btn_ok.Location = Point(pad, pad + 422)
        btn_ok.Size     = Size(88, 28)
        btn_ok.Click   += self._on_ok
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text     = "Cancel"
        btn_cancel.Location = Point(pad + 100, pad + 422)
        btn_cancel.Size     = Size(88, 28)
        btn_cancel.Click   += self._on_cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

    def _refresh_list(self, query):
        self._lst.Items.Clear()
        q = query.lower()
        for n in self._all_names:
            if not q or q in n.lower():
                self._lst.Items.Add(n)

    def _on_search(self, sender, args):
        self._refresh_list(self._txt_search.Text)

    def _on_ok(self, sender, args):
        if self._lst.SelectedItem:
            self.selected_name = self._lst.SelectedItem
            self.DialogResult  = DialogResult.OK
            self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# =============================================================================
# EDIT / ADD COLUMN DIALOG
# =============================================================================

class EditColumnDialog(Form):
    def __init__(self, config=None, doc=None):
        super(EditColumnDialog, self).__init__()
        self._doc    = doc
        self._result = None
        self._build_ui()
        if config:
            self._populate(config)
        else:
            self._on_type_changed(None, None)

    def _build_ui(self):
        self.Text            = "Edit Column"
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = StartPosition.CenterParent
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        pad      = 16
        content_w = 460   # width of content area
        total_w   = content_w + pad * 2
        row       = pad + 10

        # Help button top-right (position computed from total_w)
        btn_help = Button()
        btn_help.Text     = "?"
        btn_help.Size     = Size(28, 28)
        btn_help.Location = Point(total_w - 36, 8)
        btn_help.Click   += lambda s, e: HelpDialog("Column Editor - Help", HELP_COLUMN_EDITOR).ShowDialog(self)
        self.Controls.Add(btn_help)

        # Param name
        lbl = Label()
        lbl.Text     = "Revit parameter name:"
        lbl.Location = Point(pad, row)
        lbl.Size     = Size(220, 18)
        self.Controls.Add(lbl)
        row += 22

        self._txt_param = TextBox()
        self._txt_param.Location = Point(pad, row)
        self._txt_param.Size     = Size(content_w - 100, 24)
        self.Controls.Add(self._txt_param)

        btn_browse = Button()
        btn_browse.Text     = "Browse..."
        btn_browse.Location = Point(pad + content_w - 92, row - 2)
        btn_browse.Size     = Size(90, 28)
        btn_browse.Click   += self._on_browse
        self.Controls.Add(btn_browse)
        row += 40

        # Display name
        lbl2 = Label()
        lbl2.Text     = "Display name (column header in export):"
        lbl2.Location = Point(pad, row)
        lbl2.Size     = Size(380, 18)
        self.Controls.Add(lbl2)
        row += 22

        self._txt_display = TextBox()
        self._txt_display.Location = Point(pad, row)
        self._txt_display.Size     = Size(content_w, 24)
        self.Controls.Add(self._txt_display)
        row += 40

        # Data type
        lbl3 = Label()
        lbl3.Text     = "Data type:"
        lbl3.Location = Point(pad, row)
        lbl3.Size     = Size(90, 18)
        self.Controls.Add(lbl3)

        self._cmb_type = ComboBox()
        self._cmb_type.Location             = Point(pad + 100, row - 2)
        self._cmb_type.Size                 = Size(220, 24)
        self._cmb_type.DropDownStyle        = SWF.ComboBoxStyle.DropDownList
        for t in EXTRA_COL_TYPES:
            self._cmb_type.Items.Add(t)
        self._cmb_type.SelectedIndex         = 0
        self._cmb_type.SelectedIndexChanged += self._on_type_changed
        self.Controls.Add(self._cmb_type)
        row += 40

        # Unit
        self._lbl_unit = Label()
        self._lbl_unit.Text     = "Unit / format:"
        self._lbl_unit.Location = Point(pad, row)
        self._lbl_unit.Size     = Size(90, 18)
        self.Controls.Add(self._lbl_unit)

        self._cmb_unit = ComboBox()
        self._cmb_unit.Location      = Point(pad + 100, row - 2)
        self._cmb_unit.Size          = Size(content_w - 100, 24)
        self._cmb_unit.DropDownStyle = SWF.ComboBoxStyle.DropDownList
        self.Controls.Add(self._cmb_unit)
        row += 40

        # Note
        self._lbl_note = Label()
        self._lbl_note.Text      = ""
        self._lbl_note.Location  = Point(pad, row)
        self._lbl_note.Size      = Size(content_w, 36)
        self._lbl_note.ForeColor = Color.Black
        self.Controls.Add(self._lbl_note)
        row += 48

        # OK / Cancel
        btn_ok = Button()
        btn_ok.Text     = "OK"
        btn_ok.Size     = Size(88, 30)
        btn_ok.Location = Point(pad, row)
        btn_ok.Click   += self._on_ok
        self.Controls.Add(btn_ok)
        self.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text     = "Cancel"
        btn_cancel.Size     = Size(88, 30)
        btn_cancel.Location = Point(pad + 100, row)
        btn_cancel.Click   += self._on_cancel
        self.Controls.Add(btn_cancel)
        self.CancelButton = btn_cancel

        self.ClientSize = Size(total_w, row + 30 + pad)

    def _populate(self, config):
        self._txt_param.Text   = config.param_name
        self._txt_display.Text = config.display_name
        for i, t in enumerate(EXTRA_COL_TYPES):
            if t == config.data_type:
                self._cmb_type.SelectedIndex = i
                break
        self._on_type_changed(None, None)
        units = EXTRA_COL_UNITS.get(config.data_type, [])
        for i, (_, key) in enumerate(units):
            if key == config.unit_key:
                self._cmb_unit.SelectedIndex = i
                break

    def _on_type_changed(self, sender, args):
        dtype = self._cmb_type.SelectedItem or "text"
        units = EXTRA_COL_UNITS.get(dtype, [])
        self._cmb_unit.Items.Clear()
        for lbl, _ in units:
            self._cmb_unit.Items.Add(lbl)
        has_units = bool(units)
        self._lbl_unit.Enabled = has_units
        self._cmb_unit.Enabled = has_units
        if has_units and self._cmb_unit.Items.Count > 0:
            self._cmb_unit.SelectedIndex = 0
        self._lbl_note.Text = TYPE_NOTES.get(dtype, "")

    def _on_browse(self, sender, args):
        doc = self._doc
        if doc is None:
            try:
                doc = revit.doc
            except Exception:
                doc = None
        if doc is None:
            MessageBox.Show(
                "No active Revit document found.\nType the parameter name manually.",
                "Parameter Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        names = collect_param_names(doc)
        if not names:
            MessageBox.Show("No parameters found in the current document.",
                            "Parameter Browser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        dlg = ParamPickerDialog(names)
        if dlg.ShowDialog(self) == DialogResult.OK and dlg.selected_name:
            self._txt_param.Text = dlg.selected_name
            if not self._txt_display.Text.strip():
                self._txt_display.Text = dlg.selected_name

    def _on_ok(self, sender, args):
        param_name   = self._txt_param.Text.strip()
        display_name = self._txt_display.Text.strip()
        dtype        = self._cmb_type.SelectedItem or "text"
        if not param_name:
            MessageBox.Show("Parameter name is required.", "Validation",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if not display_name:
            display_name = param_name
        unit_key = None
        units = EXTRA_COL_UNITS.get(dtype, [])
        if units and self._cmb_unit.SelectedIndex >= 0:
            unit_key = units[self._cmb_unit.SelectedIndex][1]
        self._result = ExtraColumnConfig(
            param_name   = param_name,
            display_name = display_name,
            data_type    = dtype,
            unit_key     = unit_key,
        )
        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

    @property
    def result(self):
        return self._result


# =============================================================================
# COLUMN EDITOR DIALOG  (per-profile column list)
# =============================================================================

BUILTIN_HEADERS = {
    "Sorting", "Category", "Family", "Type", "ItemNumber",
    "Description", "Size", "Length_ft", "Length_ft_in", "Count",
    "Item count", "Model", "Area", "Material", "Comments", "Spool",
    "Service", "Hanger Number", "Install Type", "ElementId",
}


class ColumnEditorDialog(Form):
    def __init__(self, profile_name, doc=None):
        super(ColumnEditorDialog, self).__init__()
        self._profile_name = profile_name
        self._doc          = doc
        self._configs      = load_profile(profile_name)
        self._dirty        = False
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        self.Text            = "Edit Columns -- {}".format(self._profile_name)
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = StartPosition.CenterParent
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        pad    = 16
        lst_w  = 560   # list box width
        btn_x  = pad + lst_w + pad   # button column x start
        btn_w  = 140
        total_w = btn_x + btn_w + pad  # total client width

        # Help button
        btn_help = Button()
        btn_help.Text     = "?"
        btn_help.Size     = Size(28, 28)
        btn_help.Location = Point(total_w - 36, 8)
        btn_help.Click   += lambda s, e: HelpDialog("Column Editor - Help", HELP_COLUMN_EDITOR).ShowDialog(self)
        self.Controls.Add(btn_help)

        lbl_info = Label()
        lbl_info.Text      = "Profile: {}    |    Extra columns are appended after the standard BOM fields.".format(self._profile_name)
        lbl_info.Location  = Point(pad, pad)
        lbl_info.Size      = Size(total_w - pad * 2, 18)
        lbl_info.ForeColor = Color.Black
        self.Controls.Add(lbl_info)

        lst_top = pad + 30
        lst_h   = 400

        self._lst = ListBox()
        self._lst.Location      = Point(pad, lst_top)
        self._lst.Size          = Size(lst_w, lst_h)
        self._lst.Font          = Font("Courier New", 8)
        self._lst.SelectionMode = SelectionMode.One
        self._lst.DoubleClick  += self._on_edit
        self.Controls.Add(self._lst)

        bh  = 32
        gap = 10
        by  = [lst_top]

        def _btn(label, handler):
            b = Button()
            b.Text     = label
            b.Location = Point(btn_x, by[0])
            b.Size     = Size(btn_w, bh)
            b.Click   += handler
            self.Controls.Add(b)
            by[0] += bh + gap
            return b

        _btn("Add column",  self._on_add)
        _btn("Edit",        self._on_edit)
        _btn("Remove",      self._on_remove)
        by[0] += gap * 2
        _btn("Move up",     self._on_up)
        _btn("Move down",   self._on_down)

        # Bottom button row -- sits below the list
        bottom_y = lst_top + lst_h + pad

        btn_save = Button()
        btn_save.Text     = "Save profile"
        btn_save.Location = Point(pad, bottom_y)
        btn_save.Size     = Size(130, 32)
        btn_save.Click   += self._on_save
        self.Controls.Add(btn_save)
        self.AcceptButton = btn_save

        btn_close = Button()
        btn_close.Text     = "Close"
        btn_close.Location = Point(pad + 146, bottom_y)
        btn_close.Size     = Size(110, 32)
        btn_close.Click   += self._on_close
        self.Controls.Add(btn_close)
        self.CancelButton = btn_close

        self._lbl_status = Label()
        self._lbl_status.Text      = ""
        self._lbl_status.Location  = Point(pad + 270, bottom_y + 8)
        self._lbl_status.Size      = Size(300, 18)
        self._lbl_status.ForeColor = Color.DarkGreen
        self.Controls.Add(self._lbl_status)

        self.ClientSize = Size(total_w, bottom_y + 32 + pad)

    def _refresh_list(self, keep_index=None):
        self._lst.Items.Clear()
        for cfg in self._configs:
            self._lst.Items.Add(cfg.label())
        if keep_index is not None and 0 <= keep_index < self._lst.Items.Count:
            self._lst.SelectedIndex = keep_index
        elif self._lst.Items.Count > 0:
            self._lst.SelectedIndex = 0

    def _check_duplicate(self, param_name, exclude_index=None):
        for i, cfg in enumerate(self._configs):
            if i == exclude_index:
                continue
            if cfg.param_name == param_name:
                return True
        return False

    def _on_add(self, sender, args):
        dlg = EditColumnDialog(doc=self._doc)
        if dlg.ShowDialog(self) == DialogResult.OK and dlg.result:
            cfg = dlg.result
            if self._check_duplicate(cfg.param_name):
                MessageBox.Show("A column for '{}' already exists.".format(cfg.param_name),
                                "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            if cfg.display_name in BUILTIN_HEADERS:
                MessageBox.Show("'{}' conflicts with a built-in column name.".format(cfg.display_name),
                                "Name Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            self._configs.append(cfg)
            self._dirty = True
            self._lbl_status.Text = "Unsaved changes."
            self._refresh_list(keep_index=len(self._configs) - 1)

    def _on_edit(self, sender, args):
        idx = self._lst.SelectedIndex
        if idx < 0:
            return
        dlg = EditColumnDialog(config=self._configs[idx], doc=self._doc)
        if dlg.ShowDialog(self) == DialogResult.OK and dlg.result:
            cfg = dlg.result
            if self._check_duplicate(cfg.param_name, exclude_index=idx):
                MessageBox.Show("A column for '{}' already exists.".format(cfg.param_name),
                                "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            if cfg.display_name in BUILTIN_HEADERS:
                MessageBox.Show("'{}' conflicts with a built-in column name.".format(cfg.display_name),
                                "Name Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                return
            self._configs[idx] = cfg
            self._dirty = True
            self._lbl_status.Text = "Unsaved changes."
            self._refresh_list(keep_index=idx)

    def _on_remove(self, sender, args):
        idx = self._lst.SelectedIndex
        if idx < 0:
            return
        cfg = self._configs[idx]
        res = MessageBox.Show("Remove column '{}'?".format(cfg.display_name or cfg.param_name),
                              "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        if res == DialogResult.Yes:
            del self._configs[idx]
            self._dirty = True
            self._lbl_status.Text = "Unsaved changes."
            new_idx = min(idx, len(self._configs) - 1)
            self._refresh_list(keep_index=new_idx if new_idx >= 0 else None)

    def _on_up(self, sender, args):
        idx = self._lst.SelectedIndex
        if idx <= 0:
            return
        self._configs[idx - 1], self._configs[idx] = self._configs[idx], self._configs[idx - 1]
        self._dirty = True
        self._lbl_status.Text = "Unsaved changes."
        self._refresh_list(keep_index=idx - 1)

    def _on_down(self, sender, args):
        idx = self._lst.SelectedIndex
        if idx < 0 or idx >= len(self._configs) - 1:
            return
        self._configs[idx], self._configs[idx + 1] = self._configs[idx + 1], self._configs[idx]
        self._dirty = True
        self._lbl_status.Text = "Unsaved changes."
        self._refresh_list(keep_index=idx + 1)

    def _on_save(self, sender, args):
        try:
            save_profile(self._profile_name, self._configs)
            self._dirty = False
            self._lbl_status.Text = "Saved."
        except Exception as ex:
            MessageBox.Show("Could not save profile:\n\n{}".format(ex),
                            "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def _on_close(self, sender, args):
        if self._dirty:
            res = MessageBox.Show("You have unsaved changes. Close anyway?",
                                  "Unsaved Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            if res != DialogResult.Yes:
                return
        self.DialogResult = DialogResult.Cancel
        self.Close()


# =============================================================================
# PROFILE MANAGER DIALOG  (top-level UI)
# =============================================================================

class ProfileManagerDialog(Form):
    def __init__(self, doc=None):
        super(ProfileManagerDialog, self).__init__()
        self._doc = doc
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        self.Text            = "NW BOM -- Column Profile Manager"
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.StartPosition   = StartPosition.CenterParent
        self.MaximizeBox     = False
        self.MinimizeBox     = False

        pad    = 16
        lst_w  = 540
        btn_w  = 152
        btn_x  = pad + lst_w + pad
        total_w = btn_x + btn_w + pad
        lst_top = pad + 54
        lst_h   = 400
        bh      = 32
        gap     = 10

        # Active profile status label
        self._lbl_active = Label()
        self._lbl_active.Location = Point(pad, pad)
        self._lbl_active.Size     = Size(total_w - pad * 2 - 40, 20)
        self._lbl_active.Font     = Font("Segoe UI", 9, FontStyle.Bold)
        self.Controls.Add(self._lbl_active)

        lbl_sub = Label()
        lbl_sub.Text      = "Double-click a profile to edit its columns. Only one profile is active at a time."
        lbl_sub.Location  = Point(pad, pad + 26)
        lbl_sub.Size      = Size(total_w - pad * 2, 18)
        lbl_sub.ForeColor = Color.Black
        self.Controls.Add(lbl_sub)

        # Help button top-right -- positioned from right edge with margin
        btn_help = Button()
        btn_help.Text     = "?"
        btn_help.Size     = Size(30, 30)
        btn_help.Location = Point(total_w - 30 - pad, pad - 2)
        btn_help.Click   += lambda s, e: HelpDialog("Profile Manager - Help", HELP_PROFILE_MANAGER).ShowDialog(self)
        self.Controls.Add(btn_help)

        # Profile list
        self._lst = ListBox()
        self._lst.Location      = Point(pad, lst_top)
        self._lst.Size          = Size(lst_w, lst_h)
        self._lst.SelectionMode = SelectionMode.One
        self._lst.Font          = Font("Courier New", 9)
        self._lst.DoubleClick  += self._on_edit_columns
        self.Controls.Add(self._lst)

        # Right-side buttons
        by = [lst_top]

        def _btn(label, handler):
            b = Button()
            b.Text     = label
            b.Location = Point(btn_x, by[0])
            b.Size     = Size(btn_w, bh)
            b.Click   += handler
            self.Controls.Add(b)
            by[0] += bh + gap
            return b

        _btn("New profile",    self._on_new)
        _btn("Duplicate",      self._on_duplicate)
        _btn("Rename",         self._on_rename)
        _btn("Delete",         self._on_delete)
        by[0] += gap * 2
        _btn("Set active",     self._on_set_active)
        _btn("Deactivate all", self._on_deactivate)
        by[0] += gap * 2
        _btn("Edit columns",   self._on_edit_columns)
        by[0] += gap * 2
        _btn("Import profile", self._on_import)
        _btn("Export profile", self._on_export)

        # Bottom button row
        bottom_y = lst_top + lst_h + pad

        btn_close = Button()
        btn_close.Text     = "Close"
        btn_close.Location = Point(pad, bottom_y)
        btn_close.Size     = Size(110, bh)
        btn_close.Click   += lambda s, e: self.Close()
        self.Controls.Add(btn_close)
        self.CancelButton = btn_close

        btn_activate_close = Button()
        btn_activate_close.Text     = "Set active and close"
        btn_activate_close.Location = Point(pad + 122, bottom_y)
        btn_activate_close.Size     = Size(170, bh)
        btn_activate_close.Click   += self._on_set_active_and_close
        self.Controls.Add(btn_activate_close)

        self.ClientSize = Size(total_w, bottom_y + bh + pad)

    def _refresh_list(self, keep_name=None):
        active   = get_active_profile()
        profiles = list_profiles()

        if active and active in profiles:
            self._lbl_active.Text      = "Active profile:  {}".format(active)
            self._lbl_active.ForeColor = Color.DarkGreen
        else:
            self._lbl_active.Text      = "No active profile -- standard BOM columns only"
            self._lbl_active.ForeColor = Color.Black

        self._lst.Items.Clear()
        keep_index = None
        for i, name in enumerate(profiles):
            # Count columns without triggering error dialogs
            try:
                col_count = len(load_profile(name))
                count_str = "  ({} col{})".format(col_count, "s" if col_count != 1 else "")
            except Exception:
                count_str = ""
            badge = "  [ACTIVE]" if name == active else ""
            self._lst.Items.Add("{}{}{}".format(name, badge, count_str))
            if name == keep_name:
                keep_index = i

        if keep_index is not None:
            self._lst.SelectedIndex = keep_index
        elif self._lst.Items.Count > 0:
            self._lst.SelectedIndex = 0

    def _selected_name(self):
        item = self._lst.SelectedItem
        if item is None:
            return None
        # Strip badge and column count suffix -- name ends before first double-space
        name = item.replace("  [ACTIVE]", "")
        # Strip column count "(N col/cols)" suffix
        paren = name.find("  (")
        if paren >= 0:
            name = name[:paren]
        return name.strip()

    def _on_new(self, sender, args):
        name = self._prompt_name("New profile name:", "")
        if not name:
            return
        if name in list_profiles():
            MessageBox.Show("A profile named '{}' already exists.".format(name),
                            "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        save_profile(name, [])
        self._refresh_list(keep_name=name)
        # Offer to jump straight into the column editor
        res = MessageBox.Show(
            "Profile '{}' created.\n\nOpen the column editor now to add columns?".format(name),
            "Edit Columns?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        if res == DialogResult.Yes:
            dlg = ColumnEditorDialog(name, doc=self._doc)
            dlg.ShowDialog(self)
            self._refresh_list(keep_name=name)

    def _on_duplicate(self, sender, args):
        src = self._selected_name()
        if not src:
            return
        name = self._prompt_name("Name for the copy:", src + " copy")
        if not name:
            return
        if name in list_profiles():
            MessageBox.Show("A profile named '{}' already exists.".format(name),
                            "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        duplicate_profile(src, name)
        self._refresh_list(keep_name=name)

    def _on_rename(self, sender, args):
        old = self._selected_name()
        if not old:
            return
        name = self._prompt_name("New name for '{}'".format(old), old)
        if not name or name == old:
            return
        if name in list_profiles():
            MessageBox.Show("A profile named '{}' already exists.".format(name),
                            "Duplicate", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        rename_profile(old, name)
        self._refresh_list(keep_name=name)

    def _on_delete(self, sender, args):
        name = self._selected_name()
        if not name:
            return
        res = MessageBox.Show(
            "Delete profile '{}'?\nThis cannot be undone.".format(name),
            "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        if res == DialogResult.Yes:
            delete_profile(name)
            self._refresh_list()

    def _on_set_active(self, sender, args):
        name = self._selected_name()
        if not name:
            return
        set_active_profile(name)
        self._refresh_list(keep_name=name)

    def _on_deactivate(self, sender, args):
        set_active_profile(None)
        self._refresh_list()

    def _on_set_active_and_close(self, sender, args):
        """Set the selected profile active and close the dialog.

        One-click shortcut: activates the profile then closes so the user
        can immediately run the BOM button without any extra steps.
        """
        name = self._selected_name()
        if not name:
            MessageBox.Show("Select a profile first.",
                            "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        set_active_profile(name)
        self.Close()

    def _on_edit_columns(self, sender, args):
        name = self._selected_name()
        if not name:
            return
        dlg = ColumnEditorDialog(name, doc=self._doc)
        dlg.ShowDialog(self)
        self._refresh_list(keep_name=name)

    def _on_import(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title  = "Import profile"
        dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        try:
            data     = _read_json(dlg.FileName)
            imported = [ExtraColumnConfig.from_dict(d) for d in data.get("extra_columns", [])]
        except Exception as ex:
            MessageBox.Show("Could not read file:\n\n{}".format(ex),
                            "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return
        suggested = os.path.splitext(os.path.basename(dlg.FileName))[0]
        name = self._prompt_name("Profile name for import:", suggested)
        if not name:
            return
        if name in list_profiles():
            res = MessageBox.Show(
                "A profile named '{}' already exists. Overwrite?".format(name),
                "Overwrite?", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            if res != DialogResult.Yes:
                return
        save_profile(name, imported)
        self._refresh_list(keep_name=name)
        MessageBox.Show(
            "Imported {} column(s) into profile '{}'.".format(len(imported), name),
            "Import", MessageBoxButtons.OK, MessageBoxIcon.Information)

    def _on_export(self, sender, args):
        name = self._selected_name()
        if not name:
            return
        configs = load_profile(name)
        if not configs:
            MessageBox.Show("Profile '{}' has no columns to export.".format(name),
                            "Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        dlg = SaveFileDialog()
        dlg.Title    = "Export profile"
        dlg.Filter   = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        dlg.FileName = name + ".json"
        if dlg.ShowDialog() != DialogResult.OK:
            return
        try:
            _write_json(dlg.FileName, {"extra_columns": [c.to_dict() for c in configs]})
            MessageBox.Show("Exported to:\n{}".format(dlg.FileName),
                            "Export", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as ex:
            MessageBox.Show("Could not save file:\n\n{}".format(ex),
                            "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def _prompt_name(self, prompt, default):
        """Inline text input dialog. Returns string or None."""
        result_holder = [None]

        pad = 16
        w   = 360
        y   = pad

        dlg = Form()
        dlg.Text            = "Enter Name"
        dlg.FormBorderStyle = FormBorderStyle.FixedDialog
        dlg.StartPosition   = StartPosition.CenterParent
        dlg.MaximizeBox     = False
        dlg.MinimizeBox     = False

        lbl = Label()
        lbl.Text      = prompt
        lbl.Location  = Point(pad, y)
        lbl.Size      = Size(w, 20)
        lbl.ForeColor = Color.Black
        dlg.Controls.Add(lbl)
        y += 26

        txt = TextBox()
        txt.Text     = default
        txt.Location = Point(pad, y)
        txt.Size     = Size(w, 24)
        dlg.Controls.Add(txt)
        y += 38

        def _ok(s, e):
            result_holder[0] = txt.Text.strip()
            dlg.DialogResult = DialogResult.OK
            dlg.Close()

        btn_ok = Button()
        btn_ok.Text     = "OK"
        btn_ok.Location = Point(pad, y)
        btn_ok.Size     = Size(90, 32)
        btn_ok.Click   += _ok
        dlg.Controls.Add(btn_ok)
        dlg.AcceptButton = btn_ok

        btn_cancel = Button()
        btn_cancel.Text     = "Cancel"
        btn_cancel.Location = Point(pad + 102, y)
        btn_cancel.Size     = Size(90, 32)
        btn_cancel.Click   += lambda s, e: dlg.Close()
        dlg.Controls.Add(btn_cancel)
        dlg.CancelButton = btn_cancel

        # Extra bottom padding so buttons are never clipped on any DPI
        dlg.ClientSize = Size(w + pad * 2, y + 32 + pad + 8)

        # Select all text so user can type immediately without clearing
        txt.SelectAll()
        txt.Focus()

        dlg.ShowDialog(self)
        val = result_holder[0]
        return val if val else None


# =============================================================================
# ENTRY POINT
# =============================================================================

def main():
    try:
        doc = revit.doc
    except Exception:
        doc = None
    dlg = ProfileManagerDialog(doc=doc)
    dlg.ShowDialog()


if __name__ == "__main__":
    main()