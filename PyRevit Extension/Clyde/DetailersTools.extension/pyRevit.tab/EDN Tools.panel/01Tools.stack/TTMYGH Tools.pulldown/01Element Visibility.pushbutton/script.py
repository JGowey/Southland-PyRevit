# -*- coding: utf-8 -*-
# Element Visibility with WPF UI

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow
from System.Collections.Generic import List
from System.Windows import Visibility, WindowState
from System.Windows.Media import Brushes, SolidColorBrush
from System.Windows.Controls import StackPanel, TextBlock, Border, ScrollViewer, ScrollBarVisibility, Orientation, TextBox
from System.Windows import Thickness, TextWrapping, HorizontalAlignment, VerticalAlignment, Window, FontWeights, CornerRadius, WindowStartupLocation
import System.Windows.Media
import System.Windows
import System.Windows.Input
import os
import json

# --- Core Globals ---
doc   = revit.doc
uidoc = revit.uidoc
cfg   = script.get_config()

# --- Recent Elements History ---
RECENT_ELEMENTS_FILE = os.path.join(os.path.dirname(__file__), 'recent_elements.json')
MAX_RECENT_ELEMENTS = 20

def _load_recent_elements():
    """Load recent elements from JSON file."""
    try:
        if os.path.exists(RECENT_ELEMENTS_FILE):
            with open(RECENT_ELEMENTS_FILE, 'r') as f:
                return json.load(f)
    except:
        pass
    return []

def _save_recent_elements(elements):
    """Save recent elements to JSON file."""
    try:
        with open(RECENT_ELEMENTS_FILE, 'w') as f:
            json.dump(elements[:MAX_RECENT_ELEMENTS], f, indent=2)
    except:
        pass

def _add_to_recent_elements(el):
    """Add an element to recent history."""
    try:
        el_id = el.Id.IntegerValue
        
        # Create entry (just ID)
        entry = {'id': el_id}
        
        # Load existing, remove duplicates, add to front
        recent = _load_recent_elements()
        recent = [e for e in recent if e.get('id') != el_id]
        recent.insert(0, entry)
        
        # Save
        _save_recent_elements(recent)
    except:
        pass


def _e(out, msg):
    out.print_md("> " + msg)


def _name(x):
    try:
        n = x.Name
        return n if isinstance(n, basestring) else str(n)
    except:
        return "<no name>"


def _view_label(v):
    vt = getattr(v, "ViewType", None)
    tag = "TEMPLATE - " if getattr(v, "IsTemplate", False) else ""
    nm  = getattr(v, "Name", "<Unnamed>")
    try:
        nm = unicode(nm)
    except:
        pass
    return u"{0}{1} - {2}".format(tag, vt if vt is not None else "View", nm)


def _get_views(pred):
    views = []
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if pred(v):
                views.append(v)
        except:
            pass
    views.sort(key=lambda x: _view_label(x).lower())
    return views


# -------------------------
# Results Window (works from both dropdown and palette)
# -------------------------
class ResultsWindow(Window):
    """A styled WPF window to display analysis results."""
    
    def __init__(self, title, results_data):
        """
        results_data is a list of tuples: (section_type, content)
        section_type: 'header', 'subheader', 'info', 'ok', 'warning', 'error', 'likely'
        """
        Window.__init__(self)
        self.Title = title
        self.Width = 1060
        self.Height = 910
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 245))
        
        # Main scroll viewer
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Margin = Thickness(0)
        
        # Main stack panel
        main_panel = StackPanel()
        main_panel.Margin = Thickness(20)
        
        # Add sections based on results_data
        for section_type, content in results_data:
            if section_type == 'header':
                self._add_header(main_panel, content)
            elif section_type == 'subheader':
                self._add_subheader(main_panel, content)
            elif section_type == 'info':
                self._add_info(main_panel, content)
            elif section_type == 'info_alert':
                self._add_info_alert(main_panel, content)
            elif section_type == 'ok':
                self._add_check_result(main_panel, content, "OK", (46, 125, 50))  # Green
            elif section_type == 'warning':
                self._add_check_result(main_panel, content, "!", (230, 126, 34))  # Orange
            elif section_type == 'error':
                self._add_check_result(main_panel, content, "X", (192, 57, 43))  # Red
            elif section_type == 'likely':
                self._add_likely_cause(main_panel, content)
        
        # Close button
        from System.Windows.Controls import Button
        btn = Button()
        btn.Content = "Close"
        btn.Width = 120
        btn.Height = 35
        btn.Margin = Thickness(0, 20, 0, 0)
        btn.HorizontalAlignment = HorizontalAlignment.Center
        btn.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 212))
        btn.Foreground = Brushes.White
        btn.Click += self._on_close
        main_panel.Children.Add(btn)
        
        scroll.Content = main_panel
        self.Content = scroll
    
    def _add_header(self, panel, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 20
        tb.FontWeight = FontWeights.Bold
        tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 51, 51))
        tb.Margin = Thickness(0, 0, 0, 15)
        panel.Children.Add(tb)
    
    def _add_subheader(self, panel, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 14
        tb.FontWeight = FontWeights.SemiBold
        tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 51, 51))
        tb.Margin = Thickness(0, 15, 0, 10)
        panel.Children.Add(tb)
    
    def _add_info(self, panel, text):
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 12
        tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(80, 80, 80))
        tb.TextWrapping = TextWrapping.Wrap
        tb.Margin = Thickness(0, 2, 0, 2)
        panel.Children.Add(tb)
    
    def _add_info_alert(self, panel, text):
        """Add an info line that stands out with a warning-style background"""
        border = Border()
        border.Margin = Thickness(0, 4, 0, 4)
        border.Padding = Thickness(10, 6, 10, 6)
        border.CornerRadius = CornerRadius(4)
        border.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 243, 224))  # Orange/yellow background
        
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 12
        tb.FontWeight = FontWeights.SemiBold
        tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(192, 57, 43))  # Red text
        tb.TextWrapping = TextWrapping.Wrap
        border.Child = tb
        panel.Children.Add(border)
    
    def _add_check_result(self, panel, content, status, color_rgb):
        border = Border()
        border.Margin = Thickness(0, 3, 0, 3)
        border.Padding = Thickness(10, 6, 10, 6)
        border.CornerRadius = CornerRadius(4)
        
        # Light background based on status
        if status == "OK":
            border.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(232, 245, 233))
        elif status == "!":
            border.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 243, 224))
        else:
            border.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 238))
        
        sp = StackPanel()
        sp.Orientation = Orientation.Horizontal
        
        # Status indicator
        status_tb = TextBlock()
        status_tb.Text = status
        status_tb.FontSize = 11
        status_tb.FontWeight = FontWeights.Bold
        status_tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(*color_rgb))
        status_tb.Width = 25
        status_tb.VerticalAlignment = VerticalAlignment.Center
        sp.Children.Add(status_tb)
        
        # Check name and message
        name, msg = content if isinstance(content, tuple) else (content, "")
        
        name_tb = TextBlock()
        name_tb.Text = name
        name_tb.FontSize = 11
        name_tb.FontWeight = FontWeights.SemiBold
        name_tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(51, 51, 51))
        name_tb.VerticalAlignment = VerticalAlignment.Center
        sp.Children.Add(name_tb)
        
        if msg:
            msg_tb = TextBlock()
            msg_tb.Text = "  ->  " + msg
            msg_tb.FontSize = 11
            msg_tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100))
            msg_tb.TextWrapping = TextWrapping.Wrap
            msg_tb.VerticalAlignment = VerticalAlignment.Center
            sp.Children.Add(msg_tb)
        
        border.Child = sp
        panel.Children.Add(border)
    
    def _add_likely_cause(self, panel, text):
        border = Border()
        border.Margin = Thickness(0, 5, 0, 5)
        border.Padding = Thickness(12, 10, 12, 10)
        border.CornerRadius = CornerRadius(4)
        border.Background = SolidColorBrush(System.Windows.Media.Color.FromRgb(227, 242, 253))
        border.BorderBrush = SolidColorBrush(System.Windows.Media.Color.FromRgb(33, 150, 243))
        border.BorderThickness = Thickness(0, 0, 0, 2)
        
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 12
        tb.Foreground = SolidColorBrush(System.Windows.Media.Color.FromRgb(13, 71, 161))
        tb.TextWrapping = TextWrapping.Wrap
        
        border.Child = tb
        panel.Children.Add(border)
    
    def _on_close(self, sender, args):
        self.Close()


# -------------------------
# WPF Window for UI
# -------------------------
class ElementVisibilityWindow(WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'window.xaml')
        WPFWindow.__init__(self, xaml_path)
        
        self.selected_element = None
        self.selected_source = None
        self.selected_target = None
        self.recent_elements_data = []  # Initialize early for safety
        
        # Get all views
        self.views = _get_views(lambda v: not v.IsTemplate and not isinstance(v, DB.ViewSchedule))
        self.view_labels = [_view_label(v) for v in self.views]
        
        # Store filtered indices for mapping selection back to original views
        self.source_filtered_indices = list(range(len(self.views)))
        self.target_filtered_indices = list(range(len(self.views)))
        
        # Track last filter text to avoid unnecessary repopulation
        self._last_source_filter = ""
        self._last_target_filter = ""
        
        # Flag to prevent re-entry during filtering
        self._is_filtering = False
        
        # Populate ComboBoxes
        self._populate_source_combo()
        self._populate_target_combo()
        
        # Hook into Loaded event to set up TextBox event handlers
        self.Loaded += self._on_window_loaded
        
        # Check for last element
        last_id_str = cfg.get_option('last_element_id', '')
        if last_id_str:
            try:
                el = doc.GetElement(DB.ElementId(int(last_id_str)))
                if el and not isinstance(el, DB.ElementType):
                    self.btnUseLastElement.Content = "Use Last Element (Id: {})".format(last_id_str)
                    self.btnUseLastElement.Visibility = Visibility.Visible
            except:
                pass
        
        # Load and populate recent elements dropdown
        self._populate_recent_elements()
    
    def _populate_recent_elements(self):
        """Populate the recent elements dropdown."""
        # Always initialize the data list
        self.recent_elements_data = []
        
        try:
            recent = _load_recent_elements()
            
            if recent:
                self.cmbRecentElements.Items.Clear()
                self.cmbRecentElements.Items.Add("-- Select ID --")
                
                for entry in recent:
                    el_id = entry.get('id')
                    
                    # Verify element still exists
                    try:
                        el = doc.GetElement(DB.ElementId(el_id))
                        if el and not isinstance(el, DB.ElementType):
                            self.cmbRecentElements.Items.Add(str(el_id))
                            self.recent_elements_data.append(el_id)
                    except:
                        pass
                
                if len(self.recent_elements_data) > 0:
                    self.recentElementsPanel.Visibility = Visibility.Visible
                    self.cmbRecentElements.SelectedIndex = 0
        except:
            pass
    
    def _on_window_loaded(self, sender, args):
        """After window loads, hook into TextBox SelectionChanged to prevent auto-select."""
        # Get TextBoxes and prevent auto-selection behavior
        source_tb = self._get_combobox_textbox(self.cmbSourceView)
        target_tb = self._get_combobox_textbox(self.cmbTargetView)
        
        if source_tb:
            source_tb.SelectionChanged += self._on_source_textbox_selection
        if target_tb:
            target_tb.SelectionChanged += self._on_target_textbox_selection
    
    def _on_source_textbox_selection(self, sender, args):
        """Prevent auto-selection during filtering."""
        if self._is_filtering and sender.SelectionLength > 0:
            # Reset selection to cursor position only
            cursor = sender.SelectionStart
            sender.SelectionLength = 0
            sender.SelectionStart = cursor
    
    def _on_target_textbox_selection(self, sender, args):
        """Prevent auto-selection during filtering."""
        if self._is_filtering and sender.SelectionLength > 0:
            # Reset selection to cursor position only
            cursor = sender.SelectionStart
            sender.SelectionLength = 0
            sender.SelectionStart = cursor
    
    def _populate_source_combo(self, filter_text=""):
        """Populate source ComboBox, optionally filtering by text."""
        self.cmbSourceView.Items.Clear()
        self.source_filtered_indices = []
        
        filter_lower = filter_text.lower().strip()
        
        for i, label in enumerate(self.view_labels):
            if not filter_lower or filter_lower in label.lower():
                self.cmbSourceView.Items.Add(label)
                self.source_filtered_indices.append(i)
    
    def _populate_target_combo(self, filter_text=""):
        """Populate target ComboBox, optionally filtering by text."""
        self.cmbTargetView.Items.Clear()
        self.target_filtered_indices = []
        
        filter_lower = filter_text.lower().strip()
        
        for i, label in enumerate(self.view_labels):
            if not filter_lower or filter_lower in label.lower():
                self.cmbTargetView.Items.Add(label)
                self.target_filtered_indices.append(i)
    
    def _get_combobox_textbox(self, combo):
        """Get the TextBox part of an editable ComboBox."""
        try:
            # Try Template.FindName for PART_EditableTextBox (standard WPF name)
            tb = combo.Template.FindName("PART_EditableTextBox", combo)
            if tb:
                return tb
        except:
            pass
        
        # Fallback: search visual tree recursively
        try:
            from System.Windows.Media import VisualTreeHelper
            
            def find_textbox(parent):
                count = VisualTreeHelper.GetChildrenCount(parent)
                for i in range(count):
                    child = VisualTreeHelper.GetChild(parent, i)
                    if isinstance(child, TextBox):
                        return child
                    result = find_textbox(child)
                    if result:
                        return result
                return None
            
            return find_textbox(combo)
        except:
            return None
    
    def OnSourceViewKeyUp(self, sender, args):
        """Handle typing in source ComboBox to filter views."""
        # Don't filter on navigation keys
        if args.Key in [System.Windows.Input.Key.Up, System.Windows.Input.Key.Down, 
                        System.Windows.Input.Key.Enter, System.Windows.Input.Key.Escape,
                        System.Windows.Input.Key.Tab]:
            return
        
        # Get the TextBox inside the ComboBox
        tb = self._get_combobox_textbox(self.cmbSourceView)
        if not tb:
            return
            
        current_text = tb.Text or ""
        cursor_pos = tb.SelectionStart
        
        # Only repopulate if filter text actually changed
        if current_text == self._last_source_filter:
            return
        
        self._last_source_filter = current_text
        
        # Set flag to prevent auto-selection during filtering
        self._is_filtering = True
        try:
            # Filter the items (this clears and repopulates)
            self._populate_source_combo(current_text)
            
            # Restore the TextBox text and cursor directly
            tb.Text = current_text
            tb.SelectionStart = cursor_pos
            tb.SelectionLength = 0
        finally:
            self._is_filtering = False
        
        # Open dropdown to show filtered results
        if self.source_filtered_indices:
            self.cmbSourceView.IsDropDownOpen = True
    
    def OnTargetViewKeyUp(self, sender, args):
        """Handle typing in target ComboBox to filter views."""
        # Don't filter on navigation keys
        if args.Key in [System.Windows.Input.Key.Up, System.Windows.Input.Key.Down, 
                        System.Windows.Input.Key.Enter, System.Windows.Input.Key.Escape,
                        System.Windows.Input.Key.Tab]:
            return
        
        # Get the TextBox inside the ComboBox
        tb = self._get_combobox_textbox(self.cmbTargetView)
        if not tb:
            return
            
        current_text = tb.Text or ""
        cursor_pos = tb.SelectionStart
        
        # Only repopulate if filter text actually changed
        if current_text == self._last_target_filter:
            return
        
        self._last_target_filter = current_text
        
        # Set flag to prevent auto-selection during filtering
        self._is_filtering = True
        try:
            # Filter the items (this clears and repopulates)
            self._populate_target_combo(current_text)
            
            # Restore the TextBox text and cursor directly
            tb.Text = current_text
            tb.SelectionStart = cursor_pos
            tb.SelectionLength = 0
        finally:
            self._is_filtering = False
        
        # Open dropdown to show filtered results
        if self.target_filtered_indices:
            self.cmbTargetView.IsDropDownOpen = True
    
    def OnSourceViewGotFocus(self, sender, args):
        """Prevent auto-selection when ComboBox gets focus."""
        tb = self._get_combobox_textbox(self.cmbSourceView)
        if tb:
            try:
                # Move cursor to end without selection
                tb.SelectionStart = len(tb.Text or "")
                tb.SelectionLength = 0
            except:
                pass
    
    def OnTargetViewGotFocus(self, sender, args):
        """Prevent auto-selection when ComboBox gets focus."""
        tb = self._get_combobox_textbox(self.cmbTargetView)
        if tb:
            try:
                # Move cursor to end without selection
                tb.SelectionStart = len(tb.Text or "")
                tb.SelectionLength = 0
            except:
                pass
    
    def OnUseLastElement(self, sender, args):
        last_id_str = cfg.get_option('last_element_id', '')
        if last_id_str:
            try:
                el = doc.GetElement(DB.ElementId(int(last_id_str)))
                if el and not isinstance(el, DB.ElementType):
                    self._set_element(el)
            except:
                pass
    
    def OnRecentElementSelected(self, sender, args):
        """Handle selection from recent elements dropdown."""
        try:
            # Guard against missing data
            if not hasattr(self, 'recent_elements_data') or not self.recent_elements_data:
                return
                
            idx = self.cmbRecentElements.SelectedIndex
            # Index 0 is the placeholder "-- Select Recent Element --"
            if idx > 0 and idx <= len(self.recent_elements_data):
                el_id = self.recent_elements_data[idx - 1]
                el = doc.GetElement(DB.ElementId(el_id))
                if el and not isinstance(el, DB.ElementType):
                    self._set_element(el)
        except:
            pass
    
    def OnPickElement(self, sender, args):
        self.DialogResult = False
        self.Close()
    
    def OnEnterElementId(self, sender, args):
        last_id_str = cfg.get_option('last_element_id', '')
        self.WindowState = WindowState.Minimized
        
        s = forms.ask_for_string(
            default=last_id_str,
            prompt="Enter the INSTANCE ElementId that is VISIBLE in the SOURCE view:",
            title="Element Visibility"
        )
        
        self.WindowState = WindowState.Normal
        self.Activate()
        
        if s and s.strip():
            try:
                eid = DB.ElementId(int(s.strip()))
                el = doc.GetElement(eid)
                if el and not isinstance(el, DB.ElementType):
                    self._set_element(el)
                elif isinstance(el, DB.ElementType):
                    forms.alert("That Id refers to a Type. Please enter an Instance Id.")
                else:
                    forms.alert("No element with Id {}.".format(s))
            except:
                forms.alert("Invalid ElementId.")
    
    def _set_element(self, el):
        self.selected_element = el
        cfg.set_option('last_element_id', str(el.Id.IntegerValue))
        script.save_config()
        
        # Add to recent elements history
        _add_to_recent_elements(el)
        
        try:
            self.txtSelectedElement.Text = "{} (Id: {})".format(_name(el), el.Id.IntegerValue)
            self.selectedElementPanel.Visibility = Visibility.Visible
        except:
            pass
        
        self.cmbSourceView.IsEnabled = True
        self.cmbTargetView.IsEnabled = True
        self._check_analyze_button()
    
    def OnSourceViewChanged(self, sender, args):
        if self.cmbSourceView.SelectedIndex >= 0 and self.cmbSourceView.SelectedIndex < len(self.source_filtered_indices):
            # Map filtered index back to original view index
            original_index = self.source_filtered_indices[self.cmbSourceView.SelectedIndex]
            self.selected_source = self.views[original_index]
            
            if self.selected_target and self.selected_source.Id == self.selected_target.Id:
                forms.alert("Source and Target views can't be the same.")
                self.cmbSourceView.SelectedIndex = -1
                self.selected_source = None
                return
            
            # Show the selected source panel
            try:
                self.txtSelectedSource.Text = _view_label(self.selected_source)
                self.selectedSourcePanel.Visibility = Visibility.Visible
            except:
                pass
            
            self._check_analyze_button()
    
    def OnTargetViewChanged(self, sender, args):
        if self.cmbTargetView.SelectedIndex >= 0 and self.cmbTargetView.SelectedIndex < len(self.target_filtered_indices):
            # Map filtered index back to original view index
            original_index = self.target_filtered_indices[self.cmbTargetView.SelectedIndex]
            self.selected_target = self.views[original_index]
            
            if self.selected_source and self.selected_target.Id == self.selected_source.Id:
                forms.alert("Source and Target views can't be the same.")
                self.cmbTargetView.SelectedIndex = -1
                self.selected_target = None
                return
            
            # Show the selected target panel
            try:
                self.txtSelectedTarget.Text = _view_label(self.selected_target)
                self.selectedTargetPanel.Visibility = Visibility.Visible
            except:
                pass
            
            self._check_analyze_button()
    
    def _check_analyze_button(self):
        if self.selected_element and self.selected_source and self.selected_target:
            self.btnAnalyze.IsEnabled = True
        else:
            self.btnAnalyze.IsEnabled = False
    
    def OnAnalyze(self, sender, args):
        self.DialogResult = True
        self.Close()
    
    def OnCancel(self, sender, args):
        self.DialogResult = False
        self.Close()


# -------------------------
# ALL ORIGINAL CHECK FUNCTIONS - UNCHANGED
# -------------------------
def check_workset_visibility(vw, el):
    try:
        if not doc.IsWorkshared:
            return False, "OK"
        wsid = getattr(el, "WorksetId", None)
        if not wsid or wsid.IntegerValue <= 0:
            return False, "OK"
        wst = doc.GetWorksetTable()
        ws = wst.GetWorkset(wsid)
        ws_name = ws.Name if ws else "<Unnamed Workset>"
        vis = vw.GetWorksetVisibility(wsid)
        if vis == DB.WorksetVisibility.Hidden:
            return True, "Element is on workset **'{}'**, which is **explicitly hidden in this view**.".format(ws_name)
        if vis == DB.WorksetVisibility.UseGlobalSetting:
            try:
                wdef = DB.WorksetDefaultVisibilitySettings.GetSettings(doc)
                default_on = wdef.IsVisible(wsid)
            except:
                default_on = True
            if not default_on:
                return True, ("Element is on workset **'{}'**, which is **not visible by default**, "
                              "and this view uses **Use Global Settings**. Turn it on here or enable its default visibility."
                              ).format(ws_name)
    except Exception as ex:
        return None, "Workset visibility check failed: {}".format(ex)
    return False, "OK"


def check_revit_link_visibility(vw, el):
    try:
        if not isinstance(el, DB.RevitLinkInstance):
            return False, "OK"
        try:
            if vw.IsElementHidden(el.Id):
                return True, "Hidden by **Hide in View -> Elements**."
        except:
            pass

        def _visible_ids(view_id):
            return (DB.FilteredElementCollector(doc, view_id)
                    .WhereElementIsNotElementType()
                    .ToElementIds())

        tg = DB.TransactionGroup(doc, "LinkVisProbe"); tg.Start()
        culprit = False
        try:
            t1 = DB.Transaction(doc, "DupForProbe"); t1.Start()
            dup_id = vw.Duplicate(DB.ViewDuplicateOption.Duplicate); t1.Commit()
            vprobe = doc.GetElement(dup_id)
            before_has = el.Id in _visible_ids(dup_id)
            t2 = DB.Transaction(doc, "RevealOn"); t2.Start()
            vprobe.EnableRevealHiddenMode(); doc.Regenerate(); t2.Commit()
            after_has = el.Id in _visible_ids(dup_id)
            if (not before_has) and after_has:
                culprit = True
        except:
            pass
        finally:
            tg.RollBack()

        if culprit:
            nm = _name(el)
            return True, ("Revit Links tab -> **{}** is **not visible** in this view. "
                          "Turn **Visible** back **ON** in the Revit Links tab.").format(nm)
    except Exception as ex:
        return None, "Revit Link visibility check failed: {}".format(ex)
    return False, "OK"


def check_imported_categories(vw, el, culprits_found=False, source_view=None):
    """
    Check if element visibility is affected by Imported Categories in V/G.
    
    This check is run LAST and only shows warnings if:
    - No other definitive cause was found (culprits_found=False)
    - Element has imported CAD (direct ImportInstance or FamilyInstance with nested imports)
    - There are hidden import categories in the view
    """
    try:
        # --- Case 1: Direct ImportInstance (DWG/DXF in project) ---
        if isinstance(el, DB.ImportInstance):
            cat = getattr(el, "Category", None)
            cat_name = cat.Name if cat else "<Unknown Import>"
            
            # Check 1: Master "Show imported categories in this view" checkbox
            target_imports_hidden = False
            source_imports_hidden = False
            
            try:
                if hasattr(vw, 'AreImportCategoriesHidden'):
                    target_imports_hidden = vw.AreImportCategoriesHidden
            except:
                pass
            
            # Also check source view to compare
            if source_view:
                try:
                    if hasattr(source_view, 'AreImportCategoriesHidden'):
                        source_imports_hidden = source_view.AreImportCategoriesHidden
                except:
                    pass
            
            # If target has imports hidden but source doesn't -> definite cause
            if target_imports_hidden and not source_imports_hidden:
                return (True, "**Imported Categories** are turned OFF in target view (master checkbox unchecked). Import **'{}'** will be hidden.".format(cat_name))
            
            # If target has imports hidden (regardless of source)
            if target_imports_hidden:
                return (True, "**Imported Categories** are turned OFF in V/G (master checkbox unchecked). Import **'{}'** will be hidden.".format(cat_name))
            
            # Check 2: Specific import category checkbox
            if cat:
                target_cat_hidden = False
                source_cat_hidden = False
                
                # Don't gate on CanCategoryBeHidden - templates lock it to False but GetCategoryHidden still works
                try:
                    target_cat_hidden = vw.GetCategoryHidden(cat.Id)
                except:
                    pass
                
                if source_view:
                    try:
                        source_cat_hidden = source_view.GetCategoryHidden(cat.Id)
                    except:
                        pass
                
                # Compare: hidden in target but visible in source
                if target_cat_hidden and not source_cat_hidden:
                    return (True, "Imported category **'{}'** is turned **OFF** in target view but ON in source view. Check Imported Categories in V/G.".format(cat_name))
                
                if target_cat_hidden:
                    return (True, "Imported category **'{}'** is turned **OFF** in Visibility/Graphics -> Imported Categories tab.".format(cat_name))
                
                # Check subcategories (DWG layers)
                hidden_layers = []
                try:
                    subcats = cat.SubCategories
                    if subcats:
                        for subcat in subcats:
                            try:
                                if vw.GetCategoryHidden(subcat.Id):
                                    hidden_layers.append(subcat.Name)
                            except:
                                continue
                except:
                    pass
                
                if hidden_layers:
                    if len(hidden_layers) <= 3:
                        layers_str = ", ".join(["**{}**".format(l) for l in hidden_layers])
                    else:
                        layers_str = ", ".join(["**{}**".format(l) for l in hidden_layers[:3]]) + " and {} more".format(len(hidden_layers) - 3)
                    return (True, "Import **'{}'** has hidden layers: {}. Check layer visibility in V/G.".format(cat_name, layers_str))
            
            # No issues detected with import visibility
            return False, "OK"
        
        # --- Case 2: FamilyInstance - conditional check based on other results ---
        if isinstance(el, DB.FamilyInstance):
            # If another definitive cause was already found, don't add import warnings
            if culprits_found:
                return False, "OK"
            
            fam = None
            fam_name = ""
            try:
                if el.Symbol:
                    fam = el.Symbol.Family
                    fam_name = fam.Name if fam else ""
            except:
                pass
            
            if not fam_name:
                fam_name = "<Unknown Family>"
            
            # Step 1: Check if family contains imported CAD geometry
            family_has_imports = False
            try:
                if fam and fam.IsEditable:
                    fam_doc = doc.EditFamily(fam)
                    if fam_doc:
                        try:
                            import_collector = DB.FilteredElementCollector(fam_doc).OfClass(DB.ImportInstance)
                            family_has_imports = import_collector.GetElementCount() > 0
                        finally:
                            fam_doc.Close(False)
            except:
                family_has_imports = None  # Unknown
            
            # If family doesn't have imports, no need to check further
            if family_has_imports == False:
                return False, "OK"
            
            # Step 2: Check if imported categories are hidden in the view
            # Use AreImportCategoriesHidden property to check the master checkbox
            imports_hidden = False
            
            # Method 1: Check AreImportCategoriesHidden property (corresponds to "Show imported categories in this view" checkbox)
            try:
                if hasattr(vw, 'AreImportCategoriesHidden'):
                    imports_hidden = vw.AreImportCategoriesHidden
            except:
                pass
            
            # Fallback Method 2: Check GetCategoryHidden on OST_ImportObjectStyles
            if not imports_hidden:
                try:
                    cats = doc.Settings.Categories
                    import_styles_cat = cats.get_Item(DB.BuiltInCategory.OST_ImportObjectStyles)
                    if import_styles_cat:
                        imports_hidden = vw.GetCategoryHidden(import_styles_cat.Id)
                except:
                    pass
            
            # Step 3: Only show warning if family has imports AND imports are hidden
            if family_has_imports and imports_hidden:
                return (None, "Family **'{}'** contains imported CAD. **Imported Categories** are turned OFF in V/G. Check **Imports in Families** under Imported Categories tab.".format(fam_name))
            
            # Family has imports but imports are visible - OK
            return False, "OK"
        
        return False, "OK"
        
    except Exception as ex:
        return (None, "Imported Categories check failed: {}".format(ex))


def check_category_hidden(vw, el):
    if isinstance(el, DB.ImportInstance):
        return False, "OK"
    cat = getattr(el, "Category", None)
    if not cat:
        return False, "OK"
    try:
        # Don't gate on CanCategoryBeHidden - templates lock it to False but GetCategoryHidden still works
        if vw.GetCategoryHidden(cat.Id):
            return True, "Model category **'{}'** is turned **OFF** in Visibility/Graphics.".format(cat.Name)
    except:
        pass
    try:
        parent = getattr(cat, "Parent", None)
        if parent and vw.GetCategoryHidden(parent.Id):
            return True, "Parent model category **'{}'** is turned **OFF** in Visibility/Graphics.".format(parent.Name)
    except:
        pass
    return False, "OK"


def check_subcategory_hidden(vw, el, source_view=None):
    """
    Check if any subcategories of the element's category are hidden.
    This catches cases where the main category is visible but subcategories 
    (like "Hidden Lines", "Center Line", etc.) are turned off.
    
    Compares source view (visible) vs target view (missing) to detect differences.
    """
    if isinstance(el, DB.ImportInstance):
        return False, "OK"
    
    cat = getattr(el, "Category", None)
    if not cat:
        return False, "OK"
    
    cat_name = cat.Name
    hidden_subcats = []
    different_subcats = []  # Hidden in target but visible in source
    total_subcats = 0
    
    try:
        # Check if the element's category is actually a subcategory
        # and if so, check if it's hidden while parent is visible
        parent = getattr(cat, "Parent", None)
        if parent:
            try:
                # Try to get hidden state directly - don't gate on CanCategoryBeHidden
                # (templates lock CanCategoryBeHidden=False but GetCategoryHidden still works)
                parent_hidden = False
                subcat_hidden = False
                
                try:
                    parent_hidden = vw.GetCategoryHidden(parent.Id)
                except:
                    pass
                
                try:
                    subcat_hidden = vw.GetCategoryHidden(cat.Id)
                except:
                    pass
                
                parent_visible = not parent_hidden
                
                if parent_visible and subcat_hidden:
                    # Check if it was visible in source
                    source_subcat_visible = True
                    if source_view:
                        try:
                            source_subcat_visible = not source_view.GetCategoryHidden(cat.Id)
                        except:
                            pass
                    
                    if source_subcat_visible:
                        return (True, "Subcategory **'{}'** is turned **OFF** in target view but ON in source view. Parent category **'{}'** is ON.".format(cat.Name, parent.Name))
                    else:
                        return (True, "Subcategory **'{}'** is turned **OFF** in V/G while parent category **'{}'** is ON.".format(cat.Name, parent.Name))
            except:
                pass
        
        # Collect all subcategories to check from multiple sources
        subcats_to_check = []
        seen_ids = set()
        
        # Method 1: Get subcategories from element's category directly
        try:
            if cat.SubCategories:
                for sc in cat.SubCategories:
                    if sc and sc.Name and sc.Id.IntegerValue not in seen_ids:
                        subcats_to_check.append(sc)
                        seen_ids.add(sc.Id.IntegerValue)
        except:
            pass
        
        # Method 2: Get category from document settings
        try:
            doc_cat = doc.Settings.Categories.get_Item(cat.Id)
            if doc_cat and doc_cat.SubCategories:
                for sc in doc_cat.SubCategories:
                    if sc and sc.Name and sc.Id.IntegerValue not in seen_ids:
                        subcats_to_check.append(sc)
                        seen_ids.add(sc.Id.IntegerValue)
        except:
            pass
        
        # Method 3: Try by BuiltInCategory
        try:
            bic = cat.BuiltInCategory if hasattr(cat, 'BuiltInCategory') else None
            if bic and int(bic) > 0:
                doc_cat = doc.Settings.Categories.get_Item(bic)
                if doc_cat and doc_cat.SubCategories:
                    for sc in doc_cat.SubCategories:
                        if sc and sc.Name and sc.Id.IntegerValue not in seen_ids:
                            subcats_to_check.append(sc)
                            seen_ids.add(sc.Id.IntegerValue)
        except:
            pass
        
        # Check each subcategory
        for subcat in subcats_to_check:
            try:
                subcat_name = subcat.Name
                if not subcat_name:
                    continue
                
                total_subcats += 1
                target_hidden = False
                source_hidden = False
                
                # Check target view - try GetCategoryHidden regardless of CanCategoryBeHidden
                # (templates lock CanCategoryBeHidden=False but GetCategoryHidden still works)
                try:
                    target_hidden = vw.GetCategoryHidden(subcat.Id)
                except:
                    pass
                
                # Check source view
                if source_view:
                    try:
                        source_hidden = source_view.GetCategoryHidden(subcat.Id)
                    except:
                        pass
                
                # Track subcategories hidden in target but visible in source
                if target_hidden and not source_hidden:
                    different_subcats.append(subcat_name)
                elif target_hidden:
                    hidden_subcats.append(subcat_name)
                    
            except:
                continue
                
    except:
        pass
    
    # Report subcategories that differ between views (definite cause)
    if different_subcats:
        if len(different_subcats) <= 3:
            subcats_str = ", ".join(["**{}**".format(s) for s in different_subcats])
        else:
            subcats_str = ", ".join(["**{}**".format(s) for s in different_subcats[:3]]) + " and {} more".format(len(different_subcats) - 3)
        return (True, "Subcategories of **'{}'** OFF in target but ON in source: {}. Check V/G -> Model Categories -> expand category.".format(cat_name, subcats_str))
    
    # Check if ALL subcategories are hidden (definite cause for invisibility)
    all_hidden = hidden_subcats and total_subcats > 0 and len(hidden_subcats) == total_subcats
    if all_hidden:
        return (True, "**ALL** subcategories of **'{}'** are hidden in V/G ({} subcategories). Element will be invisible. Check V/G -> Model Categories -> expand category.".format(cat_name, total_subcats))
    
    # Report subcategories hidden in both views (possible cause)
    if hidden_subcats:
        if len(hidden_subcats) <= 3:
            subcats_str = ", ".join(["**{}**".format(s) for s in hidden_subcats])
        else:
            subcats_str = ", ".join(["**{}**".format(s) for s in hidden_subcats[:3]]) + " and {} more".format(len(hidden_subcats) - 3)
        return (None, "Subcategories of **'{}'** hidden in V/G: {}. These may affect element visibility.".format(cat_name, subcats_str))
    
    return False, "OK"


def check_filters(vw, el):
    try:
        hiding = []
        for fid in (vw.GetFilters() or List[DB.ElementId]()):
            vf = doc.GetElement(fid)
            if not vf:
                continue
            try:
                if vf.GetElementFilter().PassesFilter(doc, el.Id) and not vw.GetFilterVisibility(vf.Id):
                    hiding.append(vf.Name)
            except:
                pass
        if hiding:
            return True, "Hidden by view filter(s): " + ", ".join(hiding)
    except:
        pass
    return False, "OK"


def check_discipline_hint(vw, el):
    try:
        if vw.Discipline == DB.ViewDiscipline.Coordination:
            return False, "OK"
        cat = getattr(el, "Category", None)
        if not cat:
            return False, "OK"
        cat_id = cat.Id.IntegerValue
        MAP = {
            DB.ViewDiscipline.Mechanical: ([-2008032, -2008034, -2008160, -2001140, -2000032, -2008193], "Mechanical"),
            DB.ViewDiscipline.Electrical: ([-2008060, -2008050, -2001040, -2001060], "Electrical"),
            DB.ViewDiscipline.Plumbing:   ([-2008040, -2008042, -2001160, -2003300], "Plumbing"),
            DB.ViewDiscipline.Structural: ([-2001320, -2001330, -2001300, -2001352], "Structural"),
        }
        for disc, (cats, name) in MAP.items():
            if cat_id in cats and disc != vw.Discipline:
                return (None, "Element is **{}**, but the view discipline is **{}**. "
                              "Switch to **{}** or **Coordination**.".format(name, str(vw.Discipline), name))
    except:
        pass
    return False, "OK"


def check_phase_hint(vw, el):
    try:
        vp = doc.GetElement(vw.get_Parameter(DB.BuiltInParameter.VIEW_PHASE).AsElementId())
        cp = doc.GetElement(el.get_Parameter(DB.BuiltInParameter.PHASE_CREATED).AsElementId())
        dp = doc.GetElement(el.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED).AsElementId())
        
        if not vp or not cp:
            return False, "OK"
        
        # Build phase order map (chronological: Past→Future)
        phase_order = {}
        try:
            phases = doc.Phases
            for i, phase in enumerate(phases):
                phase_order[phase.Id.IntegerValue] = i
        except:
            # Fallback to ID comparison if doc.Phases not available
            phase_order = None
        
        if phase_order:
            vp_order = phase_order.get(vp.Id.IntegerValue, -1)
            cp_order = phase_order.get(cp.Id.IntegerValue, -1)
            dp_order = phase_order.get(dp.Id.IntegerValue, -1) if dp else -1
            
            if cp_order > vp_order:
                return (None, "Created in a later phase (**{}**) than the view phase (**{}**).".format(cp.Name, vp.Name))
            if dp and dp_order <= vp_order:
                return (None, "Demolished in or before the view phase (**{}**). Phase Filter may be hiding it.".format(vp.Name))
        else:
            # Fallback to old behavior
            if cp.Id.IntegerValue > vp.Id.IntegerValue:
                return (None, "Created in a later phase (**{}**) than the view phase (**{}**).".format(cp.Name, vp.Name))
            if dp and dp.Id.IntegerValue <= vp.Id.IntegerValue:
                return (None, "Demolished in or before the view phase (**{}**). Phase Filter may be hiding it.".format(vp.Name))
    except:
        pass
    return False, "OK"


def check_phase_filter_hint(vw, el):
    """Check if the view's Phase Filter is hiding this element based on its phase status."""
    try:
        # Get the view's phase filter
        pf_param = vw.get_Parameter(DB.BuiltInParameter.VIEW_PHASE_FILTER)
        if not pf_param or not pf_param.HasValue:
            return False, "No phase filter applied"
        
        pf_id = pf_param.AsElementId()
        if pf_id == DB.ElementId.InvalidElementId:
            return False, "No phase filter applied"
        
        phase_filter = doc.GetElement(pf_id)
        if not phase_filter:
            return False, "Could not retrieve phase filter"
        
        # Get view phase
        vp_param = vw.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
        if not vp_param or not vp_param.HasValue:
            return False, "View has no phase set"
        
        view_phase = doc.GetElement(vp_param.AsElementId())
        if not view_phase:
            return False, "Could not retrieve view phase"
        
        # Build phase order map from doc.Phases (chronological order: Past→Future)
        # This is the correct way to compare phases, NOT by ElementId
        phase_order = {}
        try:
            phases = doc.Phases
            for i, phase in enumerate(phases):
                phase_order[phase.Id.IntegerValue] = i
        except:
            # Fallback if doc.Phases not available
            pass
        
        view_phase_id = view_phase.Id.IntegerValue
        view_phase_order = phase_order.get(view_phase_id, -1)
        
        # Get element's created and demolished phases
        cp_param = el.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        dp_param = el.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
        
        created_phase_id = None
        demolished_phase_id = None
        created_phase_order = -1
        demolished_phase_order = -1
        
        if cp_param and cp_param.HasValue:
            cp_id = cp_param.AsElementId()
            if cp_id != DB.ElementId.InvalidElementId:
                created_phase_id = cp_id.IntegerValue
                created_phase_order = phase_order.get(created_phase_id, -1)
        
        if dp_param and dp_param.HasValue:
            dp_id = dp_param.AsElementId()
            if dp_id != DB.ElementId.InvalidElementId:
                demolished_phase_id = dp_id.IntegerValue
                demolished_phase_order = phase_order.get(demolished_phase_id, -1)
        
        # Element has no phase info - not affected by phase filter
        if created_phase_id is None:
            return False, "Element has no phase - not affected by filter"
        
        # Determine element's phase status relative to view phase
        # PhaseStatus: New=0, Existing=1, Demolished=2, Temporary=3
        # Use phase_order for comparisons (NOT ElementId values!)
        phase_status = None
        status_name = ""
        
        # Get phase names for messages
        created_phase = doc.GetElement(DB.ElementId(created_phase_id))
        created_phase_name = created_phase.Name if created_phase else "Unknown"
        
        # PhaseStatus enum values (use integers since DB.PhaseStatus may not be accessible):
        # New=0, Existing=1, Demolished=2, Temporary=3
        PHASE_STATUS_NEW = 0
        PHASE_STATUS_EXISTING = 1
        PHASE_STATUS_DEMOLISHED = 2
        PHASE_STATUS_TEMPORARY = 3
        
        if created_phase_id == view_phase_id:
            if demolished_phase_id is not None and demolished_phase_id == view_phase_id:
                # Created and demolished in same phase = Temporary
                phase_status = PHASE_STATUS_TEMPORARY
                status_name = "Temporary"
            else:
                # Created in view phase, not demolished = New
                phase_status = PHASE_STATUS_NEW
                status_name = "New"
        elif created_phase_order < view_phase_order:
            # Element created in an EARLIER phase than view phase
            if demolished_phase_id is not None and demolished_phase_id == view_phase_id:
                # Created before, demolished in view phase = Demolished
                phase_status = PHASE_STATUS_DEMOLISHED
                status_name = "Demolished"
            elif demolished_phase_id is None or demolished_phase_order > view_phase_order:
                # Created before, still exists = Existing
                phase_status = PHASE_STATUS_EXISTING
                status_name = "Existing"
            else:
                # Demolished before view phase - already caught by check_phase_hint
                return False, "OK"
        else:
            # Created after view phase - already caught by check_phase_hint
            return False, "OK"
        
        if phase_status is None:
            return False, "Could not determine phase status"
        
        # Check what the phase filter does with this status
        # PhaseStatusPresentation: ShowByCategory=0, ShowOverridden=1, DontShow=2
        PRESENTATION_SHOW_BY_CATEGORY = 0
        PRESENTATION_SHOW_OVERRIDDEN = 1
        PRESENTATION_DONT_SHOW = 2
        
        try:
            # Get the PhaseStatus enum value for the API call
            # Try multiple import methods since pyRevit's DB alias may not expose all enums
            PhaseStatus = None
            try:
                # Method 1: Direct import
                from Autodesk.Revit.DB import PhaseStatus
            except:
                try:
                    # Method 2: Through clr
                    import clr
                    clr.AddReference('RevitAPI')
                    from Autodesk.Revit.DB import PhaseStatus
                except:
                    try:
                        # Method 3: Access through module
                        import Autodesk.Revit.DB
                        PhaseStatus = Autodesk.Revit.DB.PhaseStatus
                    except:
                        PhaseStatus = None
            
            if PhaseStatus is None:
                raise Exception("Could not access PhaseStatus enum")
            
            phase_status_enum = {
                PHASE_STATUS_NEW: PhaseStatus.New,
                PHASE_STATUS_EXISTING: PhaseStatus.Existing,
                PHASE_STATUS_DEMOLISHED: PhaseStatus.Demolished,
                PHASE_STATUS_TEMPORARY: PhaseStatus.Temporary
            }.get(phase_status)
            
            presentation = phase_filter.GetPhaseStatusPresentation(phase_status_enum)
            presentation_int = int(presentation)
            
            if presentation_int == PRESENTATION_DONT_SHOW:
                return (True, "Phase Filter **{}** hides **{}** elements. Element is {} in view phase **{}**.".format(
                    phase_filter.Name, status_name, status_name.lower(), view_phase.Name))
            elif presentation_int == PRESENTATION_SHOW_OVERRIDDEN:
                return (None, "Phase Filter **{}** shows **{}** elements with overridden graphics (may appear different).".format(
                    phase_filter.Name, status_name))
            else:
                return False, "Phase Filter **{}** shows **{}** elements normally.".format(phase_filter.Name, status_name)
        except Exception as enum_err:
            # Fallback: infer from filter name when we can't access the PhaseStatus enum
            filter_name = phase_filter.Name.lower() if phase_filter.Name else ""
            
            # Common Revit phase filters and what they hide:
            # "Show All" - shows everything
            # "Show Demo + New" - hides Existing, shows New/Demolished/Temporary
            # "Show Previous + Demo" - hides New/Temporary, shows Existing/Demolished
            # "Show Previous + New" - hides Demolished/Temporary, shows Existing/New
            # "Show Previous Phase" - hides New/Demolished/Temporary, shows Existing only
            
            if "show all" in filter_name:
                return False, "Phase Filter **{}** shows all phase statuses.".format(phase_filter.Name)
            
            elif "demo + new" in filter_name or "demo+new" in filter_name:
                # "Show Demo + New" hides Existing
                if status_name == "Existing":
                    return (True, "Phase Filter **{}** hides **Existing** elements. Element was created in an earlier phase.".format(phase_filter.Name))
                else:
                    return False, "Phase Filter **{}** shows **{}** elements.".format(phase_filter.Name, status_name)
            
            elif "previous + new" in filter_name or "previous+new" in filter_name:
                # "Show Previous + New" hides Demolished and Temporary
                if status_name in ["Demolished", "Temporary"]:
                    return (True, "Phase Filter **{}** hides **{}** elements.".format(phase_filter.Name, status_name))
                else:
                    return False, "Phase Filter **{}** shows **{}** elements.".format(phase_filter.Name, status_name)
            
            elif "previous + demo" in filter_name or "previous+demo" in filter_name:
                # "Show Previous + Demo" hides New and Temporary
                if status_name in ["New", "Temporary"]:
                    return (True, "Phase Filter **{}** hides **{}** elements.".format(phase_filter.Name, status_name))
                else:
                    return False, "Phase Filter **{}** shows **{}** elements.".format(phase_filter.Name, status_name)
            
            elif "previous phase" in filter_name:
                # "Show Previous Phase" only shows Existing
                if status_name != "Existing":
                    return (True, "Phase Filter **{}** hides **{}** elements. Only shows Existing.".format(phase_filter.Name, status_name))
                else:
                    return False, "Phase Filter **{}** shows **{}** elements.".format(phase_filter.Name, status_name)
            
            # Generic fallback - warn but can't determine definitively
            return (None, "Phase Filter **{}** applied. Element status: **{}**. Verify filter settings. (Debug: {})".format(
                phase_filter.Name, status_name, str(enum_err)))
    except Exception as e:
        pass
    return False, "OK"


def check_design_option_hint(vw, el):
    try:
        do = el.DesignOption
        if do and not do.IsPrimary:
            return (None, "Element is in secondary Design Option (**{}**). The view may not be set to show it.".format(do.Name))
    except:
        pass
    return False, "OK"


def check_crop_scopebox_hint(vw, el):
    """
    Check if element is outside the view's Crop Region (2D boundary).
    For 3D views with Section Box, checks that separately.
    Far Clip is handled by check_far_clip_offset.
    """
    try:
        # Check Section Box for 3D views
        if isinstance(vw, DB.View3D):
            if vw.IsSectionBoxActive:
                section_box = vw.GetSectionBox()
                if section_box:
                    bb_el = el.get_BoundingBox(None)
                    if not bb_el:
                        bb_el = _bbox_from_geometry(el)
                    
                    if bb_el:
                        # Section box has a transform - we need to transform element coords
                        # into section box local coordinates before comparing
                        T_sb_from_world = section_box.Transform.Inverse
                        
                        # Get element corners in world coordinates
                        def _corners(bb):
                            xs = (bb.Min.X, bb.Max.X)
                            ys = (bb.Min.Y, bb.Max.Y)
                            zs = (bb.Min.Z, bb.Max.Z)
                            for x in xs:
                                for y in ys:
                                    for z in zs:
                                        yield DB.XYZ(x, y, z)
                        
                        # Transform element corners to section box local space
                        pts_local = [T_sb_from_world.OfPoint(p) for p in _corners(bb_el)]
                        
                        el_min_x = min(p.X for p in pts_local)
                        el_max_x = max(p.X for p in pts_local)
                        el_min_y = min(p.Y for p in pts_local)
                        el_max_y = max(p.Y for p in pts_local)
                        el_min_z = min(p.Z for p in pts_local)
                        el_max_z = max(p.Z for p in pts_local)
                        
                        # Now compare against section box bounds (in local space)
                        outside = (el_max_x < section_box.Min.X or el_min_x > section_box.Max.X or
                                   el_max_y < section_box.Min.Y or el_min_y > section_box.Max.Y or
                                   el_max_z < section_box.Min.Z or el_min_z > section_box.Max.Z)
                        
                        if outside:
                            return True, "Element is entirely **outside** the view's **Section Box**. Expand the Section Box to include the element."
            return False, "OK"
        
        # For other views, check Crop Region
        if not vw.CropBoxActive:
            return False, "OK"

        bb_el = el.get_BoundingBox(vw) or el.get_BoundingBox(None)
        if not bb_el:
            return False, "OK"  # Can't check, don't warn

        crop = vw.CropBox
        T_crop_from_world = crop.Transform.Inverse

        def _corners(bb):
            xs = (bb.Min.X, bb.Max.X)
            ys = (bb.Min.Y, bb.Max.Y)
            zs = (bb.Min.Z, bb.Max.Z)
            for x in xs:
                for y in ys:
                    for z in zs:
                        yield DB.XYZ(x, y, z)

        pts_local = [T_crop_from_world.OfPoint(p) for p in _corners(bb_el)]
        minx = min(p.X for p in pts_local); maxx = max(p.X for p in pts_local)
        miny = min(p.Y for p in pts_local); maxy = max(p.Y for p in pts_local)

        # Only check X/Y extents (2D crop boundary)
        # Far clip (Z/depth) is handled separately
        outside = (maxx < crop.Min.X or minx > crop.Max.X or
                   maxy < crop.Min.Y or miny > crop.Max.Y)

        if outside:
            return True, "Element is entirely **outside** the view's **Crop Region**."

    except:
        pass  # Don't warn on errors

    return False, "OK"


def _bbox_from_geometry(el):
    try:
        opt = DB.Options(); opt.DetailLevel = DB.ViewDetailLevel.Fine; opt.IncludeNonVisibleObjects = True
        geom = el.get_Geometry(opt)
        if not geom: return None
        bbox = None
        def merge(b, add):
            if not b: return add
            b.Min = DB.XYZ(min(b.Min.X, add.Min.X), min(b.Min.Y, add.Min.Y), min(b.Min.Z, add.Min.Z))
            b.Max = DB.XYZ(max(b.Max.X, add.Max.X), max(b.Max.Y, add.Max.Y), max(b.Max.Z, add.Max.Z))
            return b
        for g in geom:
            if isinstance(g, DB.GeometryInstance):
                t = g.Transform; sym = g.GetInstanceGeometry()
                for sg in sym:
                    gb = getattr(sg, "BoundingBox", None)
                    if gb:
                        pmin = t.OfPoint(gb.Min); pmax = t.OfPoint(gb.Max); b = DB.BoundingBoxXYZ()
                        b.Min = DB.XYZ(min(pmin.X, pmax.X), min(pmin.Y, pmax.Y), min(pmin.Z, pmax.Z))
                        b.Max = DB.XYZ(max(pmin.X, pmax.X), max(pmin.Y, pmax.Y), max(pmin.Z, pmax.Z))
                        bbox = merge(bbox, b)
            else:
                gb = getattr(g, "BoundingBox", None)
                if gb: bbox = merge(bbox, gb)
        return bbox
    except:
        return None


def _z_extents_for_link(el):
    try:
        ldoc = el.GetLinkDocument()
        if ldoc is None: return None
        levels = list(DB.FilteredElementCollector(ldoc).OfClass(DB.Level))
        if not levels: return None
        zs = [lvl.Elevation for lvl in levels if hasattr(lvl, "Elevation")]
        if not zs: return None
        t = el.GetTotalTransform()
        world_zs = [t.OfPoint(DB.XYZ(0, 0, z)).Z for z in zs]
        return min(world_zs), max(world_zs)
    except:
        return None


def check_view_range_plan(vw, el):
    try:
        if not isinstance(vw, DB.ViewPlan):
            return False, "OK"
        vr = vw.GetViewRange()
        if not vr: return False, "OK"
        def get_plane_z(plane_type):
            level_id = vr.GetLevelId(plane_type)
            if not level_id or level_id == DB.ElementId.InvalidElementId:
                return 1.0e10 if plane_type == DB.PlanViewPlane.TopClipPlane else -1.0e10
            level = doc.GetElement(level_id); offset = vr.GetOffset(plane_type)
            return (level.Elevation if level else 0.0) + offset
        top_z = get_plane_z(DB.PlanViewPlane.TopClipPlane)
        dep_z = get_plane_z(DB.PlanViewPlane.ViewDepthPlane)
        cut_z = get_plane_z(DB.PlanViewPlane.CutPlane)
        bb = el.get_BoundingBox(None); emin = emax = None
        if bb:
            emin, emax = bb.Min.Z, bb.Max.Z
        else:
            if isinstance(el, DB.RevitLinkInstance):
                zpair = _z_extents_for_link(el)
                if zpair: emin, emax = zpair
            elif isinstance(el, DB.ImportInstance):
                bb2 = _bbox_from_geometry(el)
                if bb2: emin, emax = bb2.Min.Z, bb2.Max.Z
            else:
                loc = getattr(el, "Location", None)
                if isinstance(loc, DB.LocationPoint): z = loc.Point.Z; emin, emax = z, z
                elif isinstance(loc, DB.LocationCurve):
                    c = loc.Curve; emin = min(c.GetEndPoint(0).Z, c.GetEndPoint(1).Z); emax = max(c.GetEndPoint(0).Z, c.GetEndPoint(1).Z)
        if emin is None or emax is None: return None, "Could not compute world Z extents for element/link to verify View Range."
        if isinstance(el, DB.ImportInstance):
            if emax < dep_z: return (None, "Import (DWG) elevation **[{:.3f}, {:.3f}]** is entirely **below** the view's **View Depth** ({:.3f}).".format(emin, emax, dep_z))
            if emin > top_z: return (None, "Import (DWG) elevation **[{:.3f}, {:.3f}]** is entirely **above** the view's **Top** plane ({:.3f}).".format(emin, emax, top_z))
            return False, "OK"
        try:
            cat = getattr(el, "Category", None)
            if cat and cat.Id.IntegerValue == -2001320:
                if emin > cut_z: return (None, "Structural Framing is located **above** the view's **Cut Plane** ({:.3f}).".format(cut_z))
        except: pass
        if emax < dep_z: return (None, "Element elevation **[{:.3f}, {:.3f}]** is entirely **below** the view's **View Depth** ({:.3f}).".format(emin, emax, dep_z))
        if emin > top_z:
            if isinstance(el, DB.RevitLinkInstance): return (None, "Linked model elevation **[{:.3f}, {:.3f}]** is entirely **above** the view's **Top** plane ({:.3f}).".format(emin, emax, top_z))
            return (None, "Element elevation **[{:.3f}, {:.3f}]** is entirely **above** the view's **Top** plane ({:.3f}).".format(emin, emax, top_z))
    except:
        return None, "Could not compute world Z extents for element/link to verify View Range."
    return False, "OK"


def check_plan_regions(vw, el):
    """
    Check if element falls within a Plan Region that may have a view range excluding it.
    Plan Regions override the view's main View Range within their boundaries.
    
    Note: The PlanRegion class is not exposed in the public Revit API, so we cannot
    programmatically read the Plan Region's view range values. We detect XY overlap
    and advise the user to check the Plan Region's View Range settings.
    """
    try:
        # Only applies to Plan views
        if not isinstance(vw, DB.ViewPlan):
            return False, "OK"
        
        # Get element's bounding box
        el_bb = el.get_BoundingBox(None)
        if not el_bb:
            loc = getattr(el, "Location", None)
            if isinstance(loc, DB.LocationPoint):
                pt = loc.Point
                el_min_x, el_max_x = pt.X, pt.X
                el_min_y, el_max_y = pt.Y, pt.Y
                el_min_z = pt.Z
            elif isinstance(loc, DB.LocationCurve):
                c = loc.Curve
                p0, p1 = c.GetEndPoint(0), c.GetEndPoint(1)
                el_min_x, el_max_x = min(p0.X, p1.X), max(p0.X, p1.X)
                el_min_y, el_max_y = min(p0.Y, p1.Y), max(p0.Y, p1.Y)
                el_min_z = min(p0.Z, p1.Z)
            else:
                return False, "OK"
        else:
            el_min_x, el_max_x = el_bb.Min.X, el_bb.Max.X
            el_min_y, el_max_y = el_bb.Min.Y, el_bb.Max.Y
            el_min_z = el_bb.Min.Z
        
        # Collect all Plan Regions in this view
        plan_regions = list(DB.FilteredElementCollector(doc, vw.Id)
                          .OfCategory(DB.BuiltInCategory.OST_PlanRegion)
                          .WhereElementIsNotElementType())
        
        if not plan_regions:
            return False, "OK"
        
        # Check each plan region for XY overlap
        for pr in plan_regions:
            try:
                # Get plan region bounding box
                pr_bb = pr.get_BoundingBox(vw)
                if not pr_bb:
                    pr_bb = pr.get_BoundingBox(None)
                if not pr_bb:
                    continue
                
                # Check for XY overlap (bounding box intersection)
                x_overlap = (el_min_x <= pr_bb.Max.X) and (el_max_x >= pr_bb.Min.X)
                y_overlap = (el_min_y <= pr_bb.Max.Y) and (el_max_y >= pr_bb.Min.Y)
                
                if x_overlap and y_overlap:
                    # Element's XY extent overlaps with this Plan Region
                    return (True, 
                        "Element overlaps with a **Plan Region** in this view. "
                        "The Plan Region has its own View Range that overrides the main view's settings. "
                        "Element is at elevation **{:.1f}'** - check the Plan Region's View Range "
                        "(especially Top clip plane) to ensure it includes this elevation.".format(el_min_z))
                
            except:
                continue
        
        return False, "OK"
        
    except:
        pass
    return False, "OK"


def check_far_clip_offset(vw, el):
    """
    Check if element is beyond the Far Clip Offset in Section/Elevation/Detail views.
    For these views, the crop box Z dimension represents the depth (far clip).
    """
    try:
        # Only applies to Section, Elevation, and Detail views
        if not hasattr(vw, 'ViewType'):
            return False, "OK"
        
        if vw.ViewType not in [DB.ViewType.Section, DB.ViewType.Elevation, DB.ViewType.Detail]:
            return False, "OK"
        
        # Get element bounding box
        bb_el = el.get_BoundingBox(vw) or el.get_BoundingBox(None)
        if not bb_el:
            bb_el = _bbox_from_geometry(el)
        if not bb_el:
            return False, "OK"
        
        # Get crop box for view coordinate transform
        crop = vw.CropBox
        if not crop:
            return False, "OK"
        
        # Transform element to view coordinates
        T_crop_from_world = crop.Transform.Inverse
        
        def _corners(bb):
            xs = (bb.Min.X, bb.Max.X)
            ys = (bb.Min.Y, bb.Max.Y)
            zs = (bb.Min.Z, bb.Max.Z)
            for x in xs:
                for y in ys:
                    for z in zs:
                        yield DB.XYZ(x, y, z)
        
        pts_local = [T_crop_from_world.OfPoint(p) for p in _corners(bb_el)]
        minz = min(p.Z for p in pts_local)
        maxz = max(p.Z for p in pts_local)
        
        # In view coordinates, Z is depth
        # crop.Min.Z is near clip, crop.Max.Z is far clip
        # Element is beyond far clip if its nearest point (minz) is past crop.Max.Z
        if minz > crop.Max.Z:
            return (True, "Element is entirely beyond the **Far Clip Offset**. Increase Far Clip Offset to include the element.")
        
        # Element is in front of near clip (behind the viewer)
        if maxz < crop.Min.Z:
            return (True, "Element is entirely in front of the view's cut plane.")
        
    except:
        pass
    
    return False, "OK"


def check_element_hidden_in_view(vw, el):
    if isinstance(el, DB.RevitLinkInstance): 
        return False, "OK"
    if not el.CanBeHidden(vw): 
        return False, "OK"
    if el.IsHidden(vw):
        category_is_hidden = False
        try:
            cat = el.Category
            if cat and vw.GetCategoryHidden(cat.Id): 
                category_is_hidden = True
        except: 
            pass
        if not category_is_hidden: 
            return True, "Hidden by **Hide in View -> Elements**."
    return False, "OK"


def check_temporary_hide_isolate(vw, el):
    """
    Check if element is hidden due to Temporary Hide/Isolate mode.
    This is when user uses the sunglasses icon to temporarily hide or isolate elements.
    """
    try:
        # Check if temporary hide/isolate is active in the view
        if not vw.IsTemporaryHideIsolateActive():
            return False, "OK"
        
        # Temporary mode is active - this could be hiding the element
        # Options: Isolate Element, Isolate Category, Hide Element, Hide Category
        return (True, "**Temporary Hide/Isolate** is active in this view. Click the sunglasses icon or **Reset Temporary Hide/Isolate** (type 'HR') to restore normal visibility.")
        
    except:
        pass
    
    return False, "OK"


def get_detail_level_name(detail_level):
    """Convert ViewDetailLevel enum to readable string."""
    try:
        level_map = {
            DB.ViewDetailLevel.Coarse: "Coarse",
            DB.ViewDetailLevel.Medium: "Medium",
            DB.ViewDetailLevel.Fine: "Fine"
        }
        return level_map.get(detail_level, str(detail_level))
    except:
        return str(detail_level)


def _get_element_level(el):
    """
    Return the Level associated with an element, or None.
    Tries typed BuiltInParameter BIPs first (most reliable), then falls back
    to finding the nearest level at or below the element's Z midpoint.
    This fallback is essential for fabrication parts which have no standard level BIP.
    Returns: DB.Level or None
    """
    level_bips = [
        DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
        DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
        DB.BuiltInParameter.RBS_START_LEVEL_PARAM,
        DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM,
    ]
    for bip in level_bips:
        try:
            p = el.get_Parameter(bip)
            if p and p.HasValue:
                lid = p.AsElementId()
                if lid and lid != DB.ElementId.InvalidElementId:
                    lvl = doc.GetElement(lid)
                    if isinstance(lvl, DB.Level):
                        return lvl
        except:
            pass

    # Fallback: nearest level at or below element Z midpoint
    try:
        bb = el.get_BoundingBox(None)
        if bb:
            el_z = (bb.Min.Z + bb.Max.Z) / 2.0
        else:
            loc = getattr(el, "Location", None)
            if isinstance(loc, DB.LocationPoint):
                el_z = loc.Point.Z
            elif isinstance(loc, DB.LocationCurve):
                el_z = loc.Curve.GetEndPoint(0).Z
            else:
                return None

        levels = sorted(
            [l for l in DB.FilteredElementCollector(doc).OfClass(DB.Level)],
            key=lambda l: l.Elevation
        )
        nearest = None
        for l in levels:
            if l.Elevation <= el_z + 0.01:  # small tolerance
                nearest = l
        return nearest
    except:
        pass

    return None


def check_level_category(vw, el):
    """
    Only fires when the selected element IS a Level datum.
    Checks three things:
      1. V/G Levels category hidden (plan / section / elevation / 3D)
      2. No associated floor plan view exists for this level (3D views only).
         Revit greys out level crosshairs in 3D when the level has no associated
         plan view -- the crosshair renders in a dim/grey state regardless of
         V/G settings. Creating a Floor Plan for the level restores full display.
      3. Level elevation outside the active Section Box Z range (3D views only).
         This hides the datum entirely rather than greying it out.
    """
    try:
        if not isinstance(el, DB.Level):
            return False, "OK"

        el_name = el.Name if el.Name else "<Unnamed Level>"
        el_elev = el.Elevation

        # --- Check 1: Levels category hidden in V/G ---
        try:
            levels_cat_id = DB.ElementId(DB.BuiltInCategory.OST_Levels)
            if vw.GetCategoryHidden(levels_cat_id):
                return (True,
                    "The selected element is a **Level** datum. The **Levels** category is turned "
                    "**OFF** in Visibility/Graphics for this view. Turn it back on under "
                    "Model Categories -> Levels.")
        except:
            pass

        # --- Check 2: No associated plan view (3D grey-out) ---
        if isinstance(vw, DB.View3D):
            try:
                has_plan = False
                for v in DB.FilteredElementCollector(doc).OfClass(DB.ViewPlan):
                    try:
                        gen = v.GenLevel
                        if gen and gen.Id == el.Id:
                            has_plan = True
                            break
                    except:
                        continue
                if not has_plan:
                    return (True,
                        "Level **'{}'** has no associated Floor Plan view. Revit greys out level "
                        "crosshairs in 3D views when no plan view exists for that level. "
                        "Go to View tab -> Plan Views -> Floor Plan and create a view for "
                        "**'{}'** to restore the full-color crosshair.".format(el_name, el_name))
            except:
                pass

        # --- Check 3: Section box Z clipping (3D views only, hides datum entirely) ---
        if isinstance(vw, DB.View3D) and vw.IsSectionBoxActive:
            try:
                section_box = vw.GetSectionBox()
                if section_box:
                    sb_center_world = section_box.Transform.OfPoint(
                        DB.XYZ(
                            (section_box.Min.X + section_box.Max.X) / 2.0,
                            (section_box.Min.Y + section_box.Max.Y) / 2.0,
                            0.0
                        )
                    )
                    probe_pt_world = DB.XYZ(sb_center_world.X, sb_center_world.Y, el_elev)
                    T_inv = section_box.Transform.Inverse
                    probe_local = T_inv.OfPoint(probe_pt_world)

                    if probe_local.Z < section_box.Min.Z or probe_local.Z > section_box.Max.Z:
                        return (True,
                            "Level **'{}'** (elev {:.2f}') is **outside the Section Box Z range**. "
                            "The datum will be hidden entirely. Expand the section box vertically "
                            "to include this elevation.".format(el_name, el_elev))
            except:
                pass

        return False, "OK"

    except:
        return False, "OK"


def check_level_association(vw, el):
    """
    Only fires for plan views. Compares the element's associated level to the
    view's associated level (vw.GenLevel). A mismatch is a strong hint that the
    element falls outside this view's View Range, and explains WHY check_view_range_plan
    may also fire. Placed just before that check in the pipeline.

    Uses _get_element_level() which handles fabrication parts via Z-elevation fallback.
    Returns None (yellow warning) on mismatch so both checks can appear together.
    """
    try:
        if not isinstance(vw, DB.ViewPlan):
            return False, "OK"

        # Don't fire for Level datums -- check_level_category handles those
        if isinstance(el, DB.Level):
            return False, "OK"

        # Get the view's associated level
        view_level = None
        try:
            if hasattr(vw, "GenLevel") and vw.GenLevel:
                view_level = vw.GenLevel
        except:
            pass

        if not view_level:
            return False, "OK"

        # Get the element's associated level
        el_level = _get_element_level(el)
        if not el_level:
            return False, "OK"

        # Same level -- no issue
        if el_level.Id == view_level.Id:
            return False, "OK"

        # Mismatch -- report with elevations for context
        el_elev  = el_level.Elevation
        vw_elev  = view_level.Elevation
        el_lname = el_level.Name if el_level.Name else "<Unnamed>"
        vw_lname = view_level.Name if view_level.Name else "<Unnamed>"

        return (None,
            "Element is associated with level **'{}'** (elev {:.2f}'), but this view is "
            "associated with **'{}'** (elev {:.2f}'). If the element's geometry does not "
            "fall within this view's View Range, it will not appear. Check View Range or "
            "verify the element was placed on the correct level.".format(
                el_lname, el_elev, vw_lname, vw_elev))

    except:
        return False, "OK"


def check_detail_level(vw, el, source_view=None):
    """Check if view detail level might affect element visibility.
    
    Only warns when:
    - Source and target have DIFFERENT detail levels
    - Target is Coarse and source is not
    - Element is a detail-sensitive category
    """
    try:
        if not hasattr(vw, 'DetailLevel'):
            return False, "OK"
            
        target_level = vw.DetailLevel
        
        # Only check if source view is provided and has different detail level
        if not source_view or not hasattr(source_view, 'DetailLevel'):
            return False, "OK"
        
        source_level = source_view.DetailLevel
        
        # If detail levels are the same, no issue
        if source_level == target_level:
            return False, "OK"
        
        # Only flag as potential issue if target is LOWER detail than source
        # (Coarse < Medium < Fine)
        level_order = {
            DB.ViewDetailLevel.Coarse: 0,
            DB.ViewDetailLevel.Medium: 1,
            DB.ViewDetailLevel.Fine: 2
        }
        
        source_val = level_order.get(source_level, 1)
        target_val = level_order.get(target_level, 1)
        
        # If target has same or higher detail than source, not likely the cause
        if target_val >= source_val:
            return False, "OK"
        
        # Target has LOWER detail - check if element category is detail-sensitive
        cat = getattr(el, "Category", None)
        cat_id = cat.Id.IntegerValue if cat else 0
        
        # Categories known to have detail-level-dependent geometry
        detail_sensitive_cats = {
            -2000170: "Detail Items",
            -2000095: "Repeating Detail Component",
            -2008126: "Fabrication Parts",
            -2008128: "Fabrication Ductwork",
            -2008129: "Fabrication Pipework",
            -2008130: "Fabrication Containment",
        }
        
        source_name = get_detail_level_name(source_level)
        target_name = get_detail_level_name(target_level)
        
        if cat_id in detail_sensitive_cats:
            cat_name = detail_sensitive_cats[cat_id]
            return (True, "**{}** elements require higher detail levels. Source is **{}**, Target is **{}**.".format(cat_name, source_name, target_name))
        
        # For other categories, only warn if target is Coarse
        if target_level == DB.ViewDetailLevel.Coarse and source_level != DB.ViewDetailLevel.Coarse:
            return (None, "Target view is **Coarse**, source is **{}**. Some family geometry may be hidden at Coarse detail level.".format(source_name))
        
        return False, "OK"
        
    except:
        return False, "OK"


# -------------------------
# Main - Handle WPF Window + Pick Element Flow
# -------------------------
def run():
    window = ElementVisibilityWindow()
    result = window.ShowDialog()
    
    # Handle Pick Element button (result = False)
    if result == False:
        try:
            uidoc.RefreshActiveView()
        except:
            pass
        
        el = None
        try:
            el = revit.pick_element(message="Pick the INSTANCE that is VISIBLE in the SOURCE view.")
        except:
            return
        
        if not el:
            return
        
        if isinstance(el, DB.ElementType):
            forms.alert("You picked a Type. Please pick an Instance.")
            return
        
        # Save and reopen window
        cfg.set_option('last_element_id', str(el.Id.IntegerValue))
        script.save_config()
        
        window2 = ElementVisibilityWindow()
        window2._set_element(el)
        result2 = window2.ShowDialog()
        
        if result2 != True:
            return
        
        el = window2.selected_element
        vsrc = window2.selected_source
        vtgt = window2.selected_target
    
    elif result == True:
        # User clicked Analyze
        el = window.selected_element
        vsrc = window.selected_source
        vtgt = window.selected_target
    
    else:
        # Cancelled
        return
    
    if not el or not vsrc or not vtgt:
        return

    # Build results data for ResultsWindow
    results_data = []
    results_data.append(('header', 'Element Visibility Analysis'))
    
    # Gather detailed element information
    el_name = _name(el)
    el_id = el.Id.IntegerValue
    el_category = ""
    el_type = ""
    el_family = ""
    
    try:
        cat = getattr(el, "Category", None)
        if cat:
            el_category = cat.Name
    except:
        pass
    
    try:
        if isinstance(el, DB.FamilyInstance):
            if el.Symbol:
                el_type = el.Symbol.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if el_type:
                    el_type = el_type.AsString() or ""
                else:
                    el_type = el.Symbol.Name if hasattr(el.Symbol, 'Name') else ""
                if el.Symbol.Family:
                    el_family = el.Symbol.Family.Name
        elif isinstance(el, DB.ImportInstance):
            # For imports, try to get the filename
            try:
                import_param = el.get_Parameter(DB.BuiltInParameter.IMPORT_SYMBOL_NAME)
                if import_param and import_param.HasValue:
                    el_type = import_param.AsString()
            except:
                pass
    except:
        pass
    
    # Build consolidated element info line: Category - Family - Type - Element (Id)
    info_parts = []
    if el_category:
        info_parts.append(el_category)
    if el_family:
        info_parts.append(el_family)
    if el_type:
        info_parts.append(el_type)
    info_parts.append("{} (Id {})".format(el_name, el_id))
    
    results_data.append(('info', " - ".join(info_parts)))
    results_data.append(('info', 'Source View (visible): {}'.format(_view_label(vsrc))))
    results_data.append(('info', 'Target View (missing): {}'.format(_view_label(vtgt))))
    
    # Show view template if applied - check if it controls V/G settings
    try:
        template_id = vtgt.ViewTemplateId
        if template_id and template_id.IntegerValue > 0:
            template = doc.GetElement(template_id)
            if template:
                # Check if V/G is locked by testing CanCategoryBeHidden on element's category
                vg_locked = False
                try:
                    el_cat = getattr(el, "Category", None)
                    if el_cat and not vtgt.CanCategoryBeHidden(el_cat.Id):
                        vg_locked = True
                except:
                    pass
                
                if vg_locked:
                    results_data.append(('info_alert', 'View Template: {} (V/G settings locked by template)'.format(template.Name)))
                else:
                    results_data.append(('info', 'View Template: {}'.format(template.Name)))
    except:
        pass

    results_data.append(('subheader', 'Checks on Target View'))

    # Create wrapper for detail level check that includes source view
    def check_detail_level_wrapper(vw, el):
        return check_detail_level(vw, el, source_view=vsrc)
    
    # Create wrapper for subcategory check that includes source view
    def check_subcategory_hidden_wrapper(vw, el):
        return check_subcategory_hidden(vw, el, source_view=vsrc)

    # Note: Imported Categories check is handled separately at the end
    # because we only want to show it if no other cause is found
    checks = [
        ("Revit Link Visibility",       check_revit_link_visibility),
        ("Workset Visibility",          check_workset_visibility),
        ("Category Hidden",             check_category_hidden),
        ("Level Datum Visibility",      check_level_category),
        ("Subcategory Hidden",          check_subcategory_hidden_wrapper),
        ("View Filters",                check_filters),
        ("Element Hidden in View",      check_element_hidden_in_view),
        ("Temporary Hide/Isolate",      check_temporary_hide_isolate),
        ("Detail Level",                check_detail_level_wrapper),
        ("Discipline",                  check_discipline_hint),
        ("Phase",                       check_phase_hint),
        ("Phase Filter",                check_phase_filter_hint),
        ("Design Option",               check_design_option_hint),
        ("Crop Region / Section Box",   check_crop_scopebox_hint),
        ("Level Association (Plan)",    check_level_association),
        ("View Range (Plan)",           check_view_range_plan),
        ("Plan Regions",                check_plan_regions),
        ("Far Clip Offset",             check_far_clip_offset),
    ]

    likely = []
    culprits = 0

    for name, fn in checks:
        try:
            state, msg = fn(vtgt, el)
            if state is True:
                culprits += 1
                likely.append(msg)
                results_data.append(('error', (name, msg)))
            elif state is None:
                likely.append(msg)
                results_data.append(('warning', (name, msg)))
            elif state is False:
                results_data.append(('ok', (name, '')))
            elif state == "error":
                results_data.append(('error', (name, msg)))
            else:
                results_data.append(('ok', (name, '')))
        except Exception as ex:
            results_data.append(('error', (name, str(ex))))

    # Check Imported Categories LAST - only show warning if:
    # 1. No other definitive cause was found (culprits == 0)
    # 2. Element is a FamilyInstance or ImportInstance
    # 3. For FamilyInstance: family contains imported CAD AND there are hidden import categories
    import_check_result = check_imported_categories(vtgt, el, culprits_found=(culprits > 0), source_view=vsrc)
    import_state, import_msg = import_check_result
    
    if import_state is True:
        culprits += 1
        likely.append(import_msg)
        results_data.append(('error', ("Imported Categories", import_msg)))
    elif import_state is None:
        likely.append(import_msg)
        results_data.append(('warning', ("Imported Categories", import_msg)))
    else:
        results_data.append(('ok', ("Imported Categories", '')))

    if culprits == 0 and isinstance(el, DB.ImportInstance):
        likely.append(
            "When selecting Reveal Hidden Elements and the element appears, right click on element -> Unhide in View; "
            "if element is available, click on Elements to unhide. "
            "If Elements is not available, go to Imported Categories under Visibility/Graphics Override "
            "and ensure that relevant checkboxes are checked."
        )

    results_data.append(('subheader', 'Likely Causes'))
    
    if likely:
        seen = set()
        for m in likely:
            if not m or m in seen: 
                continue
            seen.add(m)
            # Clean up markdown formatting for display
            clean_msg = m.replace('**', '').replace('`', '')
            results_data.append(('likely', clean_msg))
    else:
        results_data.append(('likely', "No obvious blockers found. Check these potential causes: Display Settings (transparency/halftone), Analytical Model visibility, View-specific elements (detail items only exist in one view), Underlay settings, or linked model visibility overrides."))

    # Show results window
    try:
        results_win = ResultsWindow("Element Visibility Results", results_data)
        results_win.ShowDialog()
    except Exception as e:
        # Fallback to simple alert
        forms.alert("Error showing results: {}\n\nCheck completed with {} likely causes found.".format(str(e), len(likely)))

run()