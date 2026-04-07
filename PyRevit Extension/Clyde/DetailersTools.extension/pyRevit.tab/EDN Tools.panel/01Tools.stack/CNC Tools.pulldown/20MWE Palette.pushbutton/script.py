# -*- coding: utf-8 -*-
"""
MWE Palette Button
Opens the MWE dockable panel that was registered at startup.
"""
# Allow button to work without a document open
__context__ = 'zero-doc'

import os
from pyrevit import forms

# Log file
LOG_FILE = os.path.join(os.path.dirname(__file__), "button_click.log")

def log(msg):
    """Write to log file."""
    try:
        with open(LOG_FILE, "a") as f:
            import datetime
            f.write("[{}] {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except:
        pass

log("=" * 60)
log("MWE Palette button clicked")

try:
    # The panel ID must match the one in startup.py
    panel_id = "7b9a87b0-5b78-4b2a-8b7c-0f4f9a5e6e11"
    
    log("Opening panel with ID: {}".format(panel_id))
    
    # Open the dockable panel
    forms.open_dockable_panel(panel_id)
    
    log("Panel opened successfully")
    log("=" * 60)
    
    # Show success message
    forms.toast("MWE Palette opened!", title="Success")
    
except Exception as e:
    log("ERROR: {}".format(str(e)))
    import traceback
    log("Traceback:")
    log(traceback.format_exc())
    log("=" * 60)
    
    # Show error to user
    error_msg = str(e)
    
    if "not registered" in error_msg.lower() or "not been created" in error_msg.lower():
        forms.alert(
            "MWE Palette is not registered!\n\n"
            "This usually means the startup script didn't run.\n\n"
            "Solution:\n"
            "1. Make sure 'startup.py' exists in extension root\n"
            "2. Clear pyRevit cache (Settings > Caches > Clear Cache)\n"
            "3. Restart Revit\n\n"
            "Error: {}".format(error_msg),
            title="MWE Palette Error"
        )
    else:
        forms.alert(
            "Error opening MWE Palette:\n\n{}\n\n"
            "Check log: {}".format(error_msg, LOG_FILE),
            title="MWE Palette Error"
        )
