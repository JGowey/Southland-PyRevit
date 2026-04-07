# -*- coding: utf-8 -*-
import clr
import os
import sys
import subprocess
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Windows.Forms import Form, Button, DialogResult
from System.Drawing import Point, Size

# Define the form
class UnitSelectionForm(Form):
    def __init__(self):
        self.Text = "Select Export Unit"
        self.Width = 320
        self.Height = 160

        self.selected = None

        self.btn_feet = Button()
        self.btn_feet.Text = "Decimal Feet"
        self.btn_feet.Size = Size(120, 40)
        self.btn_feet.Location = Point(30, 40)
        self.btn_feet.Click += self.select_feet

        self.btn_inches = Button()
        self.btn_inches.Text = "Decimal Inches"
        self.btn_inches.Size = Size(120, 40)
        self.btn_inches.Location = Point(160, 40)
        self.btn_inches.Click += self.select_inches

        self.Controls.Add(self.btn_feet)
        self.Controls.Add(self.btn_inches)

    def select_feet(self, sender, args):
        self.selected = "feet"
        self.DialogResult = DialogResult.OK
        self.Close()

    def select_inches(self, sender, args):
        self.selected = "inches"
        self.DialogResult = DialogResult.OK
        self.Close()

# Show the form
form = UnitSelectionForm()
result = form.ShowDialog()

# Exit if canceled
if result != DialogResult.OK or not form.selected:
    sys.exit()

# Determine script path
this_script_dir = os.path.dirname(__file__)
if form.selected == "feet":
    target_script = os.path.join(this_script_dir, "export_feet.py")
elif form.selected == "inches":
    target_script = os.path.join(this_script_dir, "export_inches.py")
else:
    sys.exit()

# Load and execute target script
with open(target_script, "r") as f:
    script_code = f.read()
exec(script_code)
