# -*- coding: utf-8 -*-
"""
MWE Palette - Extension Startup Script (Beautiful Styled Version)
This registers the MWE dockable panel with improved visual design.
"""

import os
import os.path as op
import sys
import traceback
import System
from pyrevit import forms, script, revit
from Autodesk.Revit import DB, UI

# Logging setup
LOG_FILE = op.join(op.dirname(__file__), "mwe_startup.log")

def log(msg):
    """Write to log file."""
    try:
        with open(LOG_FILE, "a") as f:
            import datetime
            f.write("[{}] {}\n".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), msg))
    except:
        pass

def load_tooltip_from_yaml(script_path):
    """Load tooltip from bundle.yaml file. Handles both | and > YAML multi-line syntaxes."""
    try:
        bundle_dir = op.dirname(script_path)
        yaml_path = op.join(bundle_dir, "bundle.yaml")
        
        if not op.isfile(yaml_path):
            return None
        
        # Simple YAML parsing for tooltip field
        with open(yaml_path, "r") as f:
            lines = f.readlines()
        
        # Find tooltip section
        in_tooltip = False
        tooltip_lines = []
        
        for line in lines:
            if line.startswith("tooltip:"):
                in_tooltip = True
                # Check if tooltip is multi-line ("|" or ">" syntax)
                if "|" in line or ">" in line:
                    # Multi-line tooltip starts next line
                    continue
                else:
                    # Single line tooltip
                    return line.split("tooltip:", 1)[1].strip()
            elif in_tooltip:
                # Check if we've reached the next field (non-indented line with colon)
                if line and not line.startswith(" ") and not line.startswith("\t") and ":" in line:
                    break
                # Add line to tooltip (preserve leading whitespace for structure)
                stripped = line.rstrip()  # Remove trailing whitespace only
                if stripped.startswith("  "):
                    # Remove the base indentation (2 spaces from YAML)
                    tooltip_lines.append(stripped[2:])
                elif stripped:
                    tooltip_lines.append(stripped.strip())
                else:
                    # Empty line
                    tooltip_lines.append("")
        
        if tooltip_lines:
            # Remove leading/trailing empty lines
            while tooltip_lines and not tooltip_lines[0]:
                tooltip_lines.pop(0)
            while tooltip_lines and not tooltip_lines[-1]:
                tooltip_lines.pop()
            
            return "\n".join(tooltip_lines)
        
    except Exception as e:
        log("Error loading tooltip from yaml: {}".format(e))
    
    return None

log("=" * 70)
log("MWE PALETTE STARTUP SCRIPT (Styled Version)")
log("__file__ = {}".format(__file__))
log("=" * 70)

# Get paths - use normpath to handle network paths properly
EXT_ROOT = op.normpath(op.dirname(op.abspath(__file__)))
LIB_ROOT = op.join(EXT_ROOT, "lib")

log("EXT_ROOT (normalized): {}".format(EXT_ROOT))
log("LIB_ROOT: {}".format(LIB_ROOT))
log("EXT_ROOT exists: {}".format(op.isdir(EXT_ROOT)))

# Check if pyRevit.tab exists
pyrevit_tab = op.join(EXT_ROOT, "pyRevit.tab")
log("pyRevit.tab path: {}".format(pyrevit_tab))
log("pyRevit.tab exists: {}".format(op.isdir(pyrevit_tab)))

# List contents of EXT_ROOT for debugging
try:
    log("EXT_ROOT contents: {}".format(os.listdir(EXT_ROOT)))
except Exception as e:
    log("ERROR listing EXT_ROOT: {}".format(e))

# Add lib to path
if LIB_ROOT not in sys.path:
    sys.path.insert(0, LIB_ROOT)
    log("Added lib to sys.path")


def build_tool_path(*parts):
    """Build and verify a tool path, trying multiple base locations."""
    # Try multiple possible paths in order of preference
    possible_bases = [
        # Primary: lib/MWE_copied (exists on local AppData installs)
        op.join(EXT_ROOT, "lib", "MWE_copied"),
        # Secondary: pyRevit.tab paths WITH 01Tools.stack
        op.join(EXT_ROOT, "pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown"),
        op.join(EXT_ROOT, "pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown"),
        # Tertiary: NW_Tools.tab paths
        op.join(EXT_ROOT, "NW_Tools.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown"),
        op.join(EXT_ROOT, "NW_Tools.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown"),
    ]
    
    button_name = parts[-2] if len(parts) >= 2 else parts[0]  # e.g., "Element Visibility.pushbutton"
    
    for base in possible_bases:
        # Try direct path first
        direct_path = op.normpath(op.join(EXT_ROOT, *parts))
        if op.isfile(direct_path):
            log("Tool found at: {}".format(direct_path))
            return direct_path
        
        # Try finding the button in this base
        test_path = op.normpath(op.join(base, button_name, "script.py"))
        if op.isfile(test_path):
            log("Tool found at: {}".format(test_path))
            return test_path
    
    # Fallback: return the first attempt path (will show as missing)
    fallback = op.normpath(op.join(EXT_ROOT, *parts))
    log("Tool NOT found, tried: {}".format(button_name))
    return fallback


# Tool definitions - ALL 12 TOOLS referencing EDN Tools directly
# These paths work because the full extension is installed locally
TOOL_CATEGORIES = [
    {
        "name": "Inspection Tools",
        "color": "#4A90E2",  # Blue
        "tools": [
            ("Element Visibility",     build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "01Element Visibility.pushbutton", "script.py")),
            ("Snoop Selection",        build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "02Snoop Selection.pushbutton", "script.py")),
            ("Find Sheet",             build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "03Find Sheet.pushbutton", "script.py")),
            ("Nested Link Scan",       build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "04Nested Link Scan.pushbutton", "script.py")),
            ("Check Duct Orientation", build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "05Check Duct Orientation.pushbutton", "script.py")),
            ("GridManager",            build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "TTMYGH Tools.pulldown", "07GridManager.pushbutton", "script.py")),
        ]
    },
    {
        "name": "Modification Tools",
        "color": "#F5A623",  # Orange
        "tools": [
            ("Renumber",        build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "01Renumber.pushbutton", "script.py")),
            ("Rotate Multiple", build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "02Rotate Multiple.pushbutton", "script.py")),
            ("Place Tags",      build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "03Place Tags.pushbutton", "script.py")),
            ("SyncViews",       build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "05SyncViews.pushbutton", "script.py")),
            ("ToggleCase",      build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "09ToggleCase.pushbutton", "script.py")),
            ("ViewConverter",   build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "10ViewConverter.pushbutton", "script.py")),
            ("BOM Builder",     build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "11BOM_Builder.pushbutton", "script.py")),
            ("UpperAttachment", build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "08UpperAttachment.pushbutton", "script.py")),
            ("HangerSync",      build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "12HangerSync.pushbutton", "script.py")),
            ("HangerSyncPlus",  build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "14HangerSyncPlus.pushbutton", "script.py")),
            ("Assembly Studio", build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "15Assembly Studio.pushbutton", "script.py")),
            ("UpdateSleeve",    build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "GMYSTYB Tools.pulldown", "11UpdateSleeve.pushbutton", "script.py")),
        ]
    },
    {
        "name": "QA / CNC Tools",
        "color": "#27AE60",  # Green
        "tools": [
            ("Compare",            build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "02Compare.pushbutton", "script.py")),
            ("Convert",            build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "04Convert.pushbutton", "script.py")),
            ("Data",               build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "09Data.pushbutton", "script.py")),
            ("NW Shop QA",         build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "06NW Shop QA.pushbutton", "script.py")),
            ("CoilDuct",           build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "07CoilDuct.pushbutton", "script.py")),
            ("Add to CID Database",build_tool_path("pyRevit.tab", "EDN Tools.panel", "01Tools.stack", "CNC Tools.pulldown", "99Add to CID Database.pushbutton", "script.py")),
        ]
    },
]

log("Tool categories built. Total tools: {}".format(sum(len(cat["tools"]) for cat in TOOL_CATEGORIES)))

# Persistent External Event Handler for running tools from dockable panel
SCRIPT_EXECUTOR_HANDLER = None
SCRIPT_EXECUTOR_EVENT = None


class SimpleExecParams:
    """Fallback EXEC_PARAMS when pyRevit's ExecutorParams isn't available."""
    def __init__(self, bundle_dir, bundle_name, ext_root):
        self.command_path = bundle_dir
        self.command_name = bundle_name.replace('.pushbutton', '')
        self.command_mode = None
        self.command_type = None
        self.command_bundle = bundle_dir
        self.command_extension = ext_root
        self.command_uniqueid = None


class PersistentScriptExecutor(UI.IExternalEventHandler):
    """Persistent handler for executing scripts from dockable panel buttons."""
    
    def __init__(self):
        self.script_path = None
        self.bundle_dir = None
        self.bundle_name = None
    
    def set_script(self, script_path):
        """Set the script to be executed."""
        self.script_path = script_path
        self.bundle_dir = op.dirname(script_path)
        self.bundle_name = op.basename(self.bundle_dir)
        log("Handler configured for: {}".format(script_path))
    
    def Execute(self, uiapp):
        """Execute the script in valid Revit API context"""
        if not self.script_path:
            log("ERROR: Execute called but no script_path set")
            return
        
        log("Execute method called - starting script execution in API context")
        
        # Save current state
        cwd0 = os.getcwd()
        old_syspath = list(sys.path)
        
        try:
            # Set up environment
            for p in (self.bundle_dir, EXT_ROOT, LIB_ROOT):
                if p and p not in sys.path:
                    sys.path.insert(0, p)
            os.chdir(self.bundle_dir)
            
            # Set up pyRevit EXEC_PARAMS
            ExecutorParams = None
            for import_attempt in [
                "from pyrevit.loader.sessioninfo import ExecutorParams",
                "from pyrevit.runtime import ExecutorParams",
                "from pyrevit.loader import ExecutorParams"
            ]:
                try:
                    exec(import_attempt)
                    if ExecutorParams is not None:
                        log("Successfully imported ExecutorParams")
                        break
                except:
                    continue
            
            if ExecutorParams is not None:
                try:
                    script.EXEC_PARAMS = ExecutorParams(
                        command_path=self.bundle_dir,
                        command_name=self.bundle_name.replace('.pushbutton', ''),
                        command_mode=None,
                        command_type=None,
                        command_bundle=self.bundle_dir,
                        command_extension=EXT_ROOT,
                        command_uniqueid=None
                    )
                    log("Set up full EXEC_PARAMS for: {}".format(self.bundle_name))
                except Exception as e:
                    log("ExecutorParams initialization failed: {}".format(e))
                    ExecutorParams = None
            
            if ExecutorParams is None:
                log("Using SimpleExecParams fallback")
                script.EXEC_PARAMS = SimpleExecParams(
                    self.bundle_dir,
                    self.bundle_name,
                    EXT_ROOT
                )
            
            # Create execution context
            g = {
                "__name__": "__main__",
                "__file__": self.script_path,
                "__commandpath__": self.bundle_dir,
                "__commandname__": self.bundle_name,
                "__revit__": __revit__,
                "__window__": None,
                "__context__": None,
                "__shiftclick__": False,
                "script": script,
                "forms": forms,
                "revit": revit,
                "DB": DB,
                "UI": UI,
            }
            
            # Set palette context marker for scripts that need to detect execution context
            # This allows scripts like Sync Views to know they're running from the safe palette context
            import pyrevit.coreutils.envvars as envvars
            envvars.set_pyrevit_env_var('MWE_PALETTE_CONTEXT', 'TRUE')
            log("Set MWE_PALETTE_CONTEXT=TRUE for safe modeless operation")
            
            # Execute the script
            log("Reading script file: {}".format(self.script_path))
            with open(self.script_path, "r") as f:
                code = f.read()
            
            log("About to execute script in API context...")
            # Execute in valid Revit API context
            exec(compile(code, self.script_path, "exec"), g, g)
            log("Tool executed successfully in External Event")
            
        except Exception as e:
            err = traceback.format_exc()
            log("EXCEPTION in Execute: {}".format(err))
            try:
                script.get_output().print_md("**Error:**\n```\n{}\n```".format(err))
            except:
                pass
            try:
                forms.alert("Error running tool:\n\n{}".format(str(e)), title="MWE Tool Error")
            except:
                pass
            log("ERROR executing tool: {}".format(err))
        finally:
            # Clear palette context marker
            try:
                import pyrevit.coreutils.envvars as envvars
                envvars.set_pyrevit_env_var('MWE_PALETTE_CONTEXT', '')
                log("Cleared MWE_PALETTE_CONTEXT")
            except:
                pass
            # Restore environment
            try:
                os.chdir(cwd0)
            except:
                pass
            sys.path[:] = old_syspath
            log("Execute method completed (finally block)")
    
    def GetName(self):
        """Return name for External Event"""
        return "MWE_PersistentScriptExecutor"


def run_tool_script(script_path):
    """Execute a tool script using the persistent External Event."""
    global SCRIPT_EXECUTOR_EVENT, SCRIPT_EXECUTOR_HANDLER
    
    log("Running tool: {}".format(script_path))
    
    if not op.isfile(script_path):
        forms.alert("Script not found:\n{}".format(script_path), title="MWE Error")
        log("ERROR: Script not found: {}".format(script_path))
        return
    
    # Check if handler and event exist
    if SCRIPT_EXECUTOR_HANDLER is None or SCRIPT_EXECUTOR_EVENT is None:
        log("ERROR: External Event not initialized!")
        forms.alert(
            "External Event Handler not initialized.\n"
            "This shouldn't happen - please reload pyRevit.",
            title="MWE Error"
        )
        return
    
    try:
        # Set the script to execute
        SCRIPT_EXECUTOR_HANDLER.set_script(script_path)
        
        # Raise the event to execute in API context
        result = SCRIPT_EXECUTOR_EVENT.Raise()
        log("External event raised with result: {}".format(result))
        
        if result != UI.ExternalEventRequest.Accepted:
            log("WARNING: External event not accepted, status: {}".format(result))
            
    except Exception as e:
        log("ERROR raising external event: {}".format(traceback.format_exc()))
        forms.alert("Failed to launch tool:\n\n{}".format(str(e)), title="MWE Error")


# Define the beautiful dockable panel
class MWEDockablePanel(forms.WPFPanel):
    """MWE Dockable Panel - Martha Wash Edition Tool Palette (Styled)"""
    
    panel_title = "MWE Palette"
    panel_id = "7b9a87b0-5b78-4b2a-8b7c-0f4f9a5e6e11"
    
    def __init__(self):
        """Initialize the panel with beautiful styling."""
        log("Initializing MWEDockablePanel (Styled)")
        
        # Import WPF components
        from System.Windows import (
            Thickness, Window, FontWeights, TextWrapping,
            HorizontalAlignment, VerticalAlignment, Point
        )
        from System.Windows.Controls import (
            ScrollViewer, StackPanel, Button, TextBlock,
            ScrollBarVisibility, Border, Separator, Image, Orientation
        )
        from System.Windows.Media import (
            SolidColorBrush, Color, Colors,
            LinearGradientBrush, GradientStop, GradientStopCollection,
            EllipseGeometry
        )
        from System.Windows.Media.Effects import DropShadowEffect
        from System.Windows.Media.Imaging import BitmapImage
        from System import Uri
        
        # Create main scroll container
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
        scroll.Background = SolidColorBrush(Color.FromRgb(245, 245, 245))
        
        # Main stack panel
        main_stack = StackPanel()
        main_stack.Margin = Thickness(0)
        
        # ============================================================
        # HEADER SECTION - Dark Elegant Design for B&W Portrait
        # ============================================================
        header_border = Border()
        
        # Create dark elegant gradient (black to deep purple - complements B&W photo)
        gradient = LinearGradientBrush()
        gradient.StartPoint = System.Windows.Point(0, 0)
        gradient.EndPoint = System.Windows.Point(1, 0)
        
        stops = GradientStopCollection()
        stops.Add(GradientStop(Color.FromRgb(20, 20, 25), 0.0))      # Deep black
        stops.Add(GradientStop(Color.FromRgb(44, 28, 54), 0.4))      # Dark purple
        stops.Add(GradientStop(Color.FromRgb(75, 35, 95), 1.0))      # Rich purple
        gradient.GradientStops = stops
        
        header_border.Background = gradient
        header_border.Padding = Thickness(15, 12, 15, 12)
        
        # Horizontal stack for image + text
        header_h_stack = StackPanel()
        header_h_stack.Orientation = Orientation.Horizontal
        header_h_stack.VerticalAlignment = VerticalAlignment.Center
        
        # Martha Wash Image (dramatic B&W portrait)
        martha_image_path = op.join(EXT_ROOT, "lib", "martha_wash.png")
        if op.isfile(martha_image_path):
            try:
                # Container for image with glow effect
                image_border = Border()
                image_border.Width = 56
                image_border.Height = 56
                image_border.Margin = Thickness(0, 0, 15, 0)
                image_border.VerticalAlignment = VerticalAlignment.Center
                
                # Create circular clip
                clip = EllipseGeometry()
                clip.Center = Point(28, 28)
                clip.RadiusX = 28
                clip.RadiusY = 28
                image_border.Clip = clip
                
                # Add dramatic glow effect
                glow = DropShadowEffect()
                glow.Color = Color.FromRgb(189, 16, 224)  # Magenta glow
                glow.BlurRadius = 20
                glow.ShadowDepth = 0
                glow.Opacity = 0.8
                image_border.Effect = glow
                
                # Load image
                img = Image()
                img.Width = 56
                img.Height = 56
                img.Stretch = System.Windows.Media.Stretch.UniformToFill
                
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(martha_image_path)
                bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
                bitmap.EndInit()
                
                img.Source = bitmap
                image_border.Child = img
                
                header_h_stack.Children.Add(image_border)
                log("Added Martha Wash header image with glow effect")
            except Exception as e:
                log("Could not load Martha Wash image: {}".format(e))
        else:
            log("Martha Wash image not found at: {}".format(martha_image_path))
        
        # Text stack for title and subtitle
        text_stack = StackPanel()
        text_stack.VerticalAlignment = VerticalAlignment.Center
        
        # Title with elegant font
        title = TextBlock()
        title.Text = "MWE Palette"
        title.FontSize = 22
        title.FontWeight = FontWeights.Bold
        title.FontFamily = System.Windows.Media.FontFamily("Segoe UI Light")
        title.Foreground = SolidColorBrush(Colors.White)
        
        # Add subtle text glow
        text_glow = DropShadowEffect()
        text_glow.Color = Color.FromRgb(189, 16, 224)
        text_glow.BlurRadius = 8
        text_glow.ShadowDepth = 0
        text_glow.Opacity = 0.3
        title.Effect = text_glow
        
        text_stack.Children.Add(title)
        
        # Subtitle with disco-inspired accent
        subtitle = TextBlock()
        subtitle.Text = "Martha Wash Edition"
        subtitle.FontSize = 11
        subtitle.FontFamily = System.Windows.Media.FontFamily("Segoe UI")
        subtitle.Foreground = SolidColorBrush(Color.FromRgb(220, 180, 255))  # Light purple
        subtitle.Margin = Thickness(0, 3, 0, 0)
        subtitle.FontStyle = System.Windows.FontStyles.Italic
        subtitle.Opacity = 0.9
        text_stack.Children.Add(subtitle)
        
        header_h_stack.Children.Add(text_stack)
        header_border.Child = header_h_stack
        main_stack.Children.Add(header_border)
        
        # ============================================================
        # TOOLS SECTION -- collapsible headers, 2-column grid
        # ============================================================
        from System.Windows.Controls import Grid as WpfGrid
        from System.Windows import GridLength, GridUnitType
        from System.Windows import Visibility as WpfVisibility
        from System.Windows.Input import Cursors

        tools_stack = StackPanel()
        tools_stack.Margin = Thickness(8, 8, 8, 8)

        for i, category in enumerate(TOOL_CATEGORIES):
            # ── collapsible header ──────────────────────────────────
            hdr_border = Border()
            # Light tint: blend color toward white at 15% opacity
            base_col = self._hex_to_color(category["color"])
            tint_r = int(base_col.R * 0.18 + 255 * 0.82)
            tint_g = int(base_col.G * 0.18 + 255 * 0.82)
            tint_b = int(base_col.B * 0.18 + 255 * 0.82)
            hdr_border.Background = SolidColorBrush(Color.FromRgb(tint_r, tint_g, tint_b))
            hdr_border.BorderBrush = SolidColorBrush(base_col)
            hdr_border.BorderThickness = Thickness(0, 0, 0, 2)
            hdr_border.Padding = Thickness(8, 6, 8, 6)
            hdr_border.Margin = Thickness(0, 8 if i > 0 else 0, 0, 4)
            hdr_border.Cursor = Cursors.Hand

            hdr_row = StackPanel()
            hdr_row.Orientation = Orientation.Horizontal

            arrow = TextBlock()
            arrow.Text = "▼"
            arrow.FontSize = 9
            arrow.Foreground = SolidColorBrush(self._hex_to_color(category["color"]))
            arrow.FontWeight = FontWeights.Bold
            arrow.Margin = Thickness(0, 0, 6, 0)
            arrow.VerticalAlignment = VerticalAlignment.Center
            hdr_row.Children.Add(arrow)

            hdr_label = TextBlock()
            hdr_label.Text = "{0}  ({1})".format(category["name"], len(category["tools"]))
            hdr_label.FontSize = 11
            hdr_label.FontWeight = FontWeights.SemiBold
            hdr_label.Foreground = SolidColorBrush(self._hex_to_color(category["color"]))
            hdr_label.FontSize = 12
            hdr_label.VerticalAlignment = VerticalAlignment.Center
            hdr_row.Children.Add(hdr_label)

            hdr_border.Child = hdr_row

            # ── strict 2-column Grid (equal star columns) ──────────
            tool_grid = WpfGrid()
            tool_grid.Margin = Thickness(0, 0, 0, 2)
            col0 = System.Windows.Controls.ColumnDefinition()
            col0.Width = GridLength(1, GridUnitType.Star)
            col1 = System.Windows.Controls.ColumnDefinition()
            col1.Width = GridLength(1, GridUnitType.Star)
            tool_grid.ColumnDefinitions.Add(col0)
            tool_grid.ColumnDefinitions.Add(col1)

            tools_list = category["tools"]
            num_rows = (len(tools_list) + 1) // 2
            for r in range(num_rows):
                row_def = System.Windows.Controls.RowDefinition()
                row_def.Height = GridLength.Auto
                tool_grid.RowDefinitions.Add(row_def)

            for idx, (label, script_path) in enumerate(tools_list):
                row_idx = idx // 2
                col_idx = idx % 2
                is_right_col = (col_idx == 1)
                btn = self._create_styled_button(label, script_path, category["color"], mirror=is_right_col)
                WpfGrid.SetRow(btn, row_idx)
                WpfGrid.SetColumn(btn, col_idx)
                tool_grid.Children.Add(btn)

            # ── wire collapse toggle ───────────────────────────────
            def make_toggle(grid, arr):
                def toggle(sender, args):
                    if grid.Visibility == WpfVisibility.Visible:
                        grid.Visibility = WpfVisibility.Collapsed
                        arr.Text = "▶"
                    else:
                        grid.Visibility = WpfVisibility.Visible
                        arr.Text = "▼"
                return toggle

            hdr_border.MouseLeftButtonUp += make_toggle(tool_grid, arrow)

            tools_stack.Children.Add(hdr_border)
            tools_stack.Children.Add(tool_grid)

        main_stack.Children.Add(tools_stack)
        
        # ============================================================
        # FOOTER
        # ============================================================
        footer_border = Border()
        footer_border.Background = SolidColorBrush(Color.FromRgb(236, 240, 241))
        footer_border.Padding = Thickness(15, 8, 15, 8)
        footer_border.BorderBrush = SolidColorBrush(Color.FromRgb(189, 195, 199))
        footer_border.BorderThickness = Thickness(0, 1, 0, 0)
        
        footer_text = TextBlock()
        footer_text.Text = "{} tools available".format(
            sum(len(cat["tools"]) for cat in TOOL_CATEGORIES)
        )
        footer_text.FontSize = 10
        footer_text.Foreground = SolidColorBrush(Color.FromRgb(127, 140, 141))
        footer_text.HorizontalAlignment = HorizontalAlignment.Center
        
        footer_border.Child = footer_text
        main_stack.Children.Add(footer_border)
        
        scroll.Content = main_stack
        self.Content = scroll
        
        log("Panel UI created with {} categories".format(len(TOOL_CATEGORIES)))
    
    def _hex_to_color(self, hex_color):
        """Convert hex color to WPF Color."""
        from System.Windows.Media import Color
        hex_color = hex_color.lstrip('#')
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return Color.FromRgb(r, g, b)
    
    def _create_styled_button(self, label, script_path, accent_color, mirror=False):
        """Create a styled button. mirror=True flips accent+icon to right side."""
        from System.Windows import Thickness, HorizontalAlignment, VerticalAlignment, Duration
        from System.Windows.Controls import Button, Border, TextBlock, StackPanel, Orientation, ToolTip, Image
        from System.Windows.Media import SolidColorBrush, Color, Colors
        from System.Windows.Media.Effects import DropShadowEffect
        from System.Windows.Media.Imaging import BitmapImage
        from System.Windows.Media.Animation import ColorAnimation, DoubleAnimation
        from System import Uri
        
        # Button container with hover effect
        btn = Button()
        btn.Height = 34  # Compact for 2-column layout
        btn.Margin = Thickness(0, 0, 2, 2)
        btn.HorizontalContentAlignment = HorizontalAlignment.Left if not mirror else HorizontalAlignment.Right
        btn.Padding = Thickness(6, 0, 6, 0)
        
        # Add subtle shadow effect
        shadow = DropShadowEffect()
        shadow.BlurRadius = 3
        shadow.ShadowDepth = 1
        shadow.Opacity = 0.2
        shadow.Color = Colors.Gray
        btn.Effect = shadow
        
        # Style the button
        if op.isfile(script_path):
            btn.Background = SolidColorBrush(Colors.White)
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(189, 195, 199))
            btn.BorderThickness = Thickness(1)
            
            # Button content -- mirrored layout for right column
            content_stack = StackPanel()
            content_stack.Orientation = Orientation.Horizontal

            # Build accent bar
            accent_bar = Border()
            accent_bar.Width = 3
            accent_bar.Background = SolidColorBrush(self._hex_to_color(accent_color))

            # Build icon
            icon_elem = None
            icon_path = op.join(op.dirname(script_path), "icon.png")
            if op.isfile(icon_path):
                try:
                    img = Image()
                    img.Width = 18
                    img.Height = 18
                    bitmap = BitmapImage()
                    bitmap.BeginInit()
                    bitmap.UriSource = Uri(icon_path)
                    bitmap.CacheOption = System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
                    bitmap.EndInit()
                    img.Source = bitmap
                    icon_elem = img
                    log("Added icon for: {}".format(label))
                except Exception as e:
                    log("Could not load icon for {}: {}".format(label, str(e)))

            # Build label
            label_text = TextBlock()
            label_text.Text = label
            label_text.FontSize = 11
            label_text.TextWrapping = System.Windows.TextWrapping.Wrap
            label_text.VerticalAlignment = VerticalAlignment.Center
            label_text.Foreground = SolidColorBrush(Color.FromRgb(52, 73, 94))

            if not mirror:
                # Left column: [accent] [icon] [label]
                accent_bar.Margin = Thickness(0, 0, 6, 0)
                content_stack.Children.Add(accent_bar)
                if icon_elem:
                    icon_elem.Margin = Thickness(0, 0, 6, 0)
                    content_stack.Children.Add(icon_elem)
                content_stack.Children.Add(label_text)
            else:
                # Right column: [label] [icon] [accent]
                content_stack.Children.Add(label_text)
                if icon_elem:
                    icon_elem.Margin = Thickness(6, 0, 6, 0)
                    content_stack.Children.Add(icon_elem)
                accent_bar.Margin = Thickness(0, 0, 0, 0)
                content_stack.Children.Add(accent_bar)

            btn.Content = content_stack
            
            # Store original colors for hover animation
            btn.Tag = {
                'accent_color': accent_color,
                'accent_bar': accent_bar,
                'label_text': label_text
            }
            
            # Add hover effects
            def on_mouse_enter(sender, args):
                """Animate button on hover"""
                try:
                    # Animate background to light gray
                    bg_animation = ColorAnimation()
                    bg_animation.To = Color.FromRgb(248, 249, 250)
                    bg_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Background.BeginAnimation(SolidColorBrush.ColorProperty, bg_animation)
                    
                    # Animate border to accent color
                    border_animation = ColorAnimation()
                    border_animation.To = self._hex_to_color(sender.Tag['accent_color'])
                    border_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.BorderBrush.BeginAnimation(SolidColorBrush.ColorProperty, border_animation)
                    
                    # Animate shadow depth
                    shadow_animation = DoubleAnimation()
                    shadow_animation.To = 3
                    shadow_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Effect.BeginAnimation(DropShadowEffect.ShadowDepthProperty, shadow_animation)
                    
                    # Brighten accent bar
                    accent_color_obj = self._hex_to_color(sender.Tag['accent_color'])
                    brighter = Color.FromRgb(
                        min(255, accent_color_obj.R + 20),
                        min(255, accent_color_obj.G + 20),
                        min(255, accent_color_obj.B + 20)
                    )
                    accent_animation = ColorAnimation()
                    accent_animation.To = brighter
                    accent_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Tag['accent_bar'].Background.BeginAnimation(SolidColorBrush.ColorProperty, accent_animation)
                except:
                    pass
            
            def on_mouse_leave(sender, args):
                """Animate button when hover ends"""
                try:
                    # Animate background back to white
                    bg_animation = ColorAnimation()
                    bg_animation.To = Colors.White
                    bg_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Background.BeginAnimation(SolidColorBrush.ColorProperty, bg_animation)
                    
                    # Animate border back to gray
                    border_animation = ColorAnimation()
                    border_animation.To = Color.FromRgb(189, 195, 199)
                    border_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.BorderBrush.BeginAnimation(SolidColorBrush.ColorProperty, border_animation)
                    
                    # Animate shadow depth back
                    shadow_animation = DoubleAnimation()
                    shadow_animation.To = 1
                    shadow_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Effect.BeginAnimation(DropShadowEffect.ShadowDepthProperty, shadow_animation)
                    
                    # Restore accent bar color
                    accent_animation = ColorAnimation()
                    accent_animation.To = self._hex_to_color(sender.Tag['accent_color'])
                    accent_animation.Duration = Duration.op_Implicit(System.TimeSpan.FromMilliseconds(150))
                    sender.Tag['accent_bar'].Background.BeginAnimation(SolidColorBrush.ColorProperty, accent_animation)
                except:
                    pass
            
            btn.MouseEnter += on_mouse_enter
            btn.MouseLeave += on_mouse_leave
            
            # Load and add tooltip from bundle.yaml
            tooltip_text = load_tooltip_from_yaml(script_path)
            if tooltip_text:
                tooltip = ToolTip()
                tooltip.Content = tooltip_text
                tooltip.MaxWidth = 400
                
                # Style the tooltip
                tooltip_text_block = TextBlock()
                tooltip_text_block.Text = tooltip_text
                tooltip_text_block.TextWrapping = System.Windows.TextWrapping.Wrap
                tooltip_text_block.Padding = Thickness(8, 6, 8, 6)
                tooltip_text_block.FontSize = 11
                tooltip_text_block.Background = SolidColorBrush(Color.FromRgb(255, 255, 225))  # Light yellow
                tooltip_text_block.Foreground = SolidColorBrush(Color.FromRgb(51, 51, 51))
                
                tooltip.Content = tooltip_text_block
                btn.ToolTip = tooltip
                log("Added tooltip for: {}".format(label))
            
            # Add click handler
            def make_handler(path):
                def handler(sender, args):
                    run_tool_script(path)
                return handler
            
            btn.Click += make_handler(script_path)
            log("Added styled button: {}".format(label))
        else:
            # Disabled style for missing scripts
            btn.Background = SolidColorBrush(Color.FromRgb(236, 240, 241))
            btn.BorderBrush = SolidColorBrush(Color.FromRgb(189, 195, 199))
            btn.BorderThickness = Thickness(1)
            btn.IsEnabled = False
            
            label_text = TextBlock()
            label_text.Text = "{} (missing)".format(label)
            label_text.FontSize = 11
            label_text.Foreground = SolidColorBrush(Color.FromRgb(149, 165, 166))
            label_text.TextAlignment = System.Windows.TextAlignment.Right if mirror else System.Windows.TextAlignment.Left
            btn.Content = label_text
            
            log("Script missing for: {}".format(label))
        
        return btn


# Initialize the persistent External Event Handler (must be done in API context during startup)
try:
    log("Initializing persistent External Event Handler...")
    SCRIPT_EXECUTOR_HANDLER = PersistentScriptExecutor()
    SCRIPT_EXECUTOR_EVENT = UI.ExternalEvent.Create(SCRIPT_EXECUTOR_HANDLER)
    log("External Event Handler initialized successfully")
except Exception as e:
    log("ERROR initializing External Event Handler: {}".format(traceback.format_exc()))
    SCRIPT_EXECUTOR_HANDLER = None
    SCRIPT_EXECUTOR_EVENT = None


# Register the dockable panel
# ============================================================
# FIXED: Handle "already registered" error gracefully
# ============================================================
try:
    log("Registering MWE dockable panel...")
    
    if forms.is_registered_dockable_panel(MWEDockablePanel):
        log("Panel already registered (pyRevit check)")
    else:
        try:
            forms.register_dockable_panel(MWEDockablePanel)
            log("SUCCESS! MWE Panel registered with pyRevit")
        except Exception as reg_error:
            # Check if it's the "already registered" error - this is OK!
            error_str = str(reg_error).lower()
            if "more than once" in error_str or "already" in error_str:
                log("Panel already registered from previous Revit session (this is OK)")
            else:
                # Some other error - re-raise it
                raise
    
    log("=" * 70)
    log("STARTUP SCRIPT COMPLETED SUCCESSFULLY")
    log("=" * 70)
    
except Exception as e:
    log("!" * 70)
    log("ERROR DURING REGISTRATION")
    log("!" * 70)
    log("Error: {}".format(str(e)))
    log("Traceback:")
    log(traceback.format_exc())
    log("=" * 70)
    
    # Only show alert for REAL errors, not "already registered"
    error_str = str(e).lower()
    if "more than once" not in error_str and "already" not in error_str:
        try:
            forms.alert(
                "Failed to register MWE Palette!\n\n"
                "Error: {}\n\n"
                "Check log: {}".format(str(e), LOG_FILE),
                title="MWE Startup Error"
            )
        except:
            pass
