# -*- coding: utf-8 -*-
"""
unit_prefs_dialog.py - Unit Preferences Dialog
===============================================
WinForms dialog that lets the user set display units for every
parameter category in one place.

Called from script.py via show_prefs_dialog().
"""

import clr
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System.Drawing import Size, Point, Font as DFont, FontStyle, Color
from System.Windows.Forms import (
    Form, Label, ComboBox, ComboBoxStyle, Button, Panel,
    ScrollableControl, VScrollBar, FlowLayoutPanel,
    AnchorStyles, DockStyle, DialogResult, FormStartPosition,
    FormBorderStyle, TableLayoutPanel, BorderStyle,
    GroupBox, CheckBox,
)

import unit_prefs as up


class UnitPrefsDialog(Form):
    """
    Displays all unit categories with a dropdown for each.
    Organized into logical groups matching Revit's discipline structure.
    """

    # Visual grouping of categories
    GROUPS = [
        ("Length & Geometry", [
            "Length", "Area", "Volume", "Slope", "Angle",
        ]),
        ("Structural", [
            "Mass", "Force", "LinearForce", "AreaForce",
            "Moment", "Stress", "MassDensity",
        ]),
        ("HVAC", [
            "HvacAirflow", "HvacPressure", "Temperature",
            "Velocity", "DuctSize",
        ]),
        ("Piping", [
            "PipingFlow", "PipingPressure",
        ]),
        ("Electrical", [
            "ElectricalPower", "ApparentPower", "Current", "Voltage",
            "Illuminance", "LuminousFlux", "ColorTemperature",
        ]),
        ("Other", [
            "Currency",
        ]),
    ]

    def __init__(self):
        Form.__init__(self)
        self.Text            = "Unit Display Preferences"
        self.Size            = Size(560, 680)
        self.MinimumSize     = Size(480, 480)
        self.StartPosition   = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.Sizable

        self._combos = {}   # category -> ComboBox

        # Header
        hdr = Label()
        hdr.Text     = "Set the display unit for each parameter type.\nSettings are saved per-user and apply to the grid, Excel export, and import."
        hdr.Dock     = DockStyle.Top
        hdr.Height   = 38
        hdr.Font     = DFont("Arial", 9)
        hdr.Padding  = System_Padding(8, 6, 8, 0)

        # Scrollable content panel
        scroll = Panel()
        scroll.Dock       = DockStyle.Fill
        scroll.AutoScroll = True
        scroll.Padding    = System_Padding(8, 4, 8, 4)

        y = 4
        for group_name, cats in self.GROUPS:
            # Group label
            grp_lbl = Label()
            grp_lbl.Text      = group_name
            grp_lbl.Location  = Point(4, y)
            grp_lbl.Size      = Size(520, 20)
            grp_lbl.Font      = DFont("Arial", 9, FontStyle.Bold)
            grp_lbl.ForeColor = Color.FromArgb(47, 61, 78)
            scroll.Controls.Add(grp_lbl)
            y += 22

            for cat in cats:
                cat_def = up.CATEGORIES.get(cat)
                if not cat_def:
                    continue
                opts = cat_def.get("options", [])
                if not opts:
                    continue  # skip dimensionless categories

                # Row: label + combo
                row_lbl = Label()
                row_lbl.Text     = cat_def["display"] + ":"
                row_lbl.Location = Point(16, y + 3)
                row_lbl.Size     = Size(220, 18)
                row_lbl.Font     = DFont("Arial", 9)
                scroll.Controls.Add(row_lbl)

                cbo = ComboBox()
                cbo.Location      = Point(244, y)
                cbo.Size          = Size(270, 22)
                cbo.DropDownStyle = ComboBoxStyle.DropDownList
                cbo.Font          = DFont("Arial", 9)
                cbo.Tag           = cat

                for attr, display_name, label in opts:
                    cbo.Items.Add("{} ({})".format(display_name, label))

                current_idx = up.get_pref_index(cat)
                cbo.SelectedIndex = max(0, min(current_idx, cbo.Items.Count - 1))
                scroll.Controls.Add(cbo)
                self._combos[cat] = cbo
                y += 28

            y += 6  # gap between groups

        # Bottom buttons
        btn_panel = Panel()
        btn_panel.Dock   = DockStyle.Bottom
        btn_panel.Height = 44
        btn_panel.Padding = System_Padding(8, 6, 8, 6)

        btnDefaults = Button()
        btnDefaults.Text     = "Reset to Defaults"
        btnDefaults.Size     = Size(130, 28)
        btnDefaults.Location = Point(8, 8)
        btnDefaults.Click   += self._on_reset

        btnOK = Button()
        btnOK.Text     = "Apply & Close"
        btnOK.Size     = Size(110, 28)
        btnOK.Anchor   = AnchorStyles.Right | AnchorStyles.Bottom
        btnOK.Click   += self._on_ok

        btnCancel = Button()
        btnCancel.Text   = "Cancel"
        btnCancel.Size   = Size(80, 28)
        btnCancel.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
        btnCancel.Click += lambda s, a: self.Close()

        def _layout_btns(s=None, a=None):
            w = btn_panel.ClientSize.Width
            btnCancel.Location = Point(w - 88, 8)
            btnOK.Location     = Point(w - 88 - 118, 8)
        btn_panel.Resize += _layout_btns
        _layout_btns()

        for w in [btnDefaults, btnOK, btnCancel]:
            btn_panel.Controls.Add(w)

        self.Controls.Add(scroll)
        self.Controls.Add(hdr)
        self.Controls.Add(btn_panel)

    def _on_ok(self, s, a):
        """Apply all combo selections to prefs and save."""
        for cat, cbo in self._combos.items():
            idx = cbo.SelectedIndex
            if idx >= 0:
                up.set_pref_index(cat, idx)
        up.save_prefs()
        self.DialogResult = DialogResult.OK
        self.Close()

    def _on_reset(self, s, a):
        """Reset all combos to defaults."""
        for cat, cbo in self._combos.items():
            default = up.CATEGORIES[cat]["default"]
            cbo.SelectedIndex = default


# Handle Padding import difference between IronPython contexts
try:
    from System.Windows.Forms import Padding as System_Padding
except ImportError:
    from System.Windows.Forms import Padding
    System_Padding = Padding


def show_prefs_dialog(parent=None):
    """
    Opens the unit preferences dialog.
    Returns True if the user applied changes.
    """
    from System.Windows.Forms import Application
    Application.EnableVisualStyles()
    dlg = UnitPrefsDialog()
    if parent:
        result = dlg.ShowDialog(parent)
    else:
        result = dlg.ShowDialog()
    return result == DialogResult.OK