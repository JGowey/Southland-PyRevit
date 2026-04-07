# -*- coding: utf-8 -*-
"""
Extension initialization - registers MWE dockable pane
This file is loaded when pyRevit initializes the extension
"""

import os
import sys

# Log file for diagnostics
LOG_FILE = os.path.join(os.path.dirname(__file__), "init_registration_log.txt")

def log(msg):
    try:
        with open(LOG_FILE, "a") as f:
            import datetime
            f.write("{} - {}\n".format(datetime.datetime.now(), msg))
    except:
        pass

log("="*60)
log("__init__.py is being executed")
log("__file__: {}".format(__file__))

# Check if we have access to __revit__
log("Checking for __revit__...")
try:
    if '__revit__' in dir():
        log("__revit__ IS available!")
        
        # Import and register the pane
        from System import Guid
        from Autodesk.Revit.UI import DockablePaneId
        
        # Import our pane module
        sys.path.insert(0, os.path.dirname(__file__))
        import mwe_pane
        
        uiapp = __revit__
        pane_id = DockablePaneId(mwe_pane.PANE_GUID)
        
        # Try to register
        try:
            uiapp.GetDockablePane(pane_id)
            log("Pane already registered")
        except:
            provider = mwe_pane.MweProvider()
            uiapp.RegisterDockablePane(pane_id, "MWE Palette", provider)
            log("SUCCESS - Pane registered from __init__.py!")
    else:
        log("__revit__ NOT available")
        log("Available: {}".format(sorted([x for x in dir() if not x.startswith('_')])))
except Exception as e:
    log("ERROR: {}".format(str(e)))
    import traceback
    log(traceback.format_exc())

log("__init__.py completed")
log("="*60)