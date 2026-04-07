# -*- coding: utf-8 -*-
"""
MWE dockable pane - Fixed version with better output window handling
This module defines the pane provider but does NOT auto-register.
Registration happens in the hooks/application-startup.py file.
"""

import os, sys, traceback
from System import Guid

# Create a log file to track what's happening
LOG_FILE = os.path.join(os.path.dirname(__file__), "mwe_registration_log.txt")

def log(msg):
    """Write diagnostic message to log file."""
    try:
        with open(LOG_FILE, "a") as f:
            import datetime
            f.write("{} - {}\n".format(datetime.datetime.now(), msg))
    except:
        pass

log("="*60)
log("mwe_pane.py module is being imported")

# Load .NET/WPF assemblies
import clr
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('System.Xml')

from System.Windows import Thickness
from System.Windows.Controls import Button
from System.Windows.Markup import XamlReader
from System.IO import StringReader
from System.Xml import XmlReader
from Autodesk.Revit.UI import IDockablePaneProvider, DockablePaneProviderData, DockablePaneId

# Fixed GUID for the MWE pane
PANE_GUID = Guid("7b9a87b0-5b78-4b2a-8b7c-0f4f9a5e6e11")
log("PANE_GUID: {}".format(PANE_GUID))

def _ext_root():
    """Get the extension root directory."""
    return os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))

EXT_ROOT = _ext_root()
LIB_ROOT = os.path.join(EXT_ROOT, "lib")

# Button definitions - (label, relative path to script)
# All paths point to the actual pulldown scripts (no more MWE_copied duplicates)
# Paths must include 01Tools.stack and match actual folder names in pyRevit.tab
BUTTONS = [
    # Inspection Tools
    ("Element Visibility",     "pyRevit.tab/EDN Tools.panel/01Tools.stack/CNC Tools.pulldown/01Element Visibility.pushbutton/script.py"),
    ("Snoop Selection",        "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/08Snoop Selection.pushbutton/script.py"),
    ("Sync Views",             "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/05SyncViews.pushbutton/script.py"),
    ("Compare",                "pyRevit.tab/EDN Tools.panel/01Tools.stack/CNC Tools.pulldown/02Compare.pushbutton/script.py"),
    ("Check Duct Orientation", "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/07Check Duct Orientation.pushbutton/script.py"),
    # Modification Tools
    ("ReNumber",               "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/01Renumber.pushbutton/script.py"),
    ("Rotate Multiple",        "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/02Rotate Multiple.pushbutton/script.py"),
    ("Place Tags",             "pyRevit.tab/EDN Tools.panel/01Tools.stack/GMYSTYB Tools.pulldown/03Place Tags.pushbutton/script.py"),
    ("Convert",                "pyRevit.tab/EDN Tools.panel/01Tools.stack/CNC Tools.pulldown/06Convert.pushbutton/script.py"),
]

def _abs_from_root(relpath):
    """Convert relative path from extension root to absolute path."""
    return os.path.join(EXT_ROOT, relpath)

def _run_script_with_context(script_path):
    """
    Run a copied pyRevit tool with proper context.
    This function sets up the environment and executes the script.
    """
    from pyrevit import script as _script, forms as _forms, revit as _revit
    from Autodesk.Revit import DB, UI

    bundle_dir  = os.path.dirname(script_path)
    bundle_name = os.path.basename(bundle_dir)

    cwd0 = os.getcwd()
    old_syspath = list(sys.path)
    
    try:
        # Add necessary paths to sys.path
        for p in (bundle_dir, EXT_ROOT, LIB_ROOT):
            if p and p not in sys.path:
                sys.path.insert(0, p)
        os.chdir(bundle_dir)

        # Set up global context for the script
        g = {
            "__name__": "__main__",
            "__file__": script_path,
            "__commandpath__": bundle_dir,
            "__commandname__": bundle_name,
            "__revit__": __revit__,
            "__window__": None,
            "__context__": None,
            "__shiftclick__": False,
            "script": _script,
            "forms": _forms,
            "revit": _revit,
            "DB": DB,
            "UI": UI,
        }

        # Read and execute the script
        import codecs
        with codecs.open(script_path, "r", encoding="utf-8") as f:
            code = f.read()
        exec(compile(code, script_path, "exec"), g, g)

    except Exception:
        err = traceback.format_exc()
        # Always show error in alert since output window may not work
        try:
            _forms.alert(u"MWE error in {}:\n\n{}".format(os.path.basename(script_path), err[:500]))
        except:
            pass
    finally:
        try:
            os.chdir(cwd0)
        except Exception:
            pass
        sys.path[:] = old_syspath

# XAML definition for the palette UI
XAML = """
<ScrollViewer xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
              xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
              VerticalScrollBarVisibility="Auto">
  <StackPanel Margin="10">
    <TextBlock Text="MWE (Martha Wash Edition)" 
               FontSize="16" 
               FontWeight="Bold"
               Margin="0,0,0,12" />
    <TextBlock Text="Click a button below to run a tool:" 
               Margin="0,0,0,8"
               TextWrapping="Wrap" />
  </StackPanel>
</ScrollViewer>
"""

class MweProvider(IDockablePaneProvider):
    """
    IDockablePaneProvider that builds the MWE palette UI.
    This class creates the dockable pane with buttons for each tool.
    """
    def SetupDockablePane(self, data):
        """
        Called by Revit to set up the pane UI.
        Creates a WPF interface with buttons for each tool.
        """
        log("SetupDockablePane called")
        
        # Parse the XAML to create the UI
        xml_reader = XmlReader.Create(StringReader(XAML))
        root = XamlReader.Load(xml_reader)
        stack = root.Content if hasattr(root, "Content") else root

        # Create a button for each tool
        for (label, rel) in BUTTONS:
            fullp = _abs_from_root(rel)
            btn = Button(Content=label)
            btn.Margin = Thickness(0, 2, 0, 2)
            btn.Height = 28
            
            if os.path.isfile(fullp):
                # Create click handler for this button
                def make_handler(p):
                    def _handler(sender, args):
                        _run_script_with_context(p)
                    return _handler
                btn.Click += make_handler(fullp)
                log("Added button: {} -> {}".format(label, fullp))
            else:
                # Disable button if script file is missing
                btn.IsEnabled = False
                btn.Content = u"{}  (missing)".format(label)
                log("Script missing for: {} -> {}".format(label, fullp))
            
            stack.Children.Add(btn)

        data.FrameworkElement = root
        log("SetupDockablePane completed - {} buttons added".format(len(BUTTONS)))

log("mwe_pane.py module loaded successfully (no auto-registration)")
log("="*60)