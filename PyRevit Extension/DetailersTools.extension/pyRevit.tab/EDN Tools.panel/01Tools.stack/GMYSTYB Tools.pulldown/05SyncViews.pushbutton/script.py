# -*- coding: utf-8 -*-
"""Sync Views - Dual-mode tool for 2D sync and 3D view matching.

Uses MWE_PALETTE_CONTEXT environment variable (set by startup.py) to determine mode:
- From Palette: MWE_PALETTE_CONTEXT=TRUE -> CONTINUOUS AUTO-SYNC
- From Dropdown: MWE_PALETTE_CONTEXT not set -> ONE-SHOT MODAL (safe fallback)

CRITICAL: All Revit API objects are fetched FRESH at sync time only.
NO API objects (View, UIView, ElementId, ViewType) stored in UI controls.

NOTE: Uses __context__ = 'zero-doc' instead of 'doc-open' because pyRevit 6.1.0's
C# loader can incorrectly grey out this button with 'doc-open'. Runtime doc check
replaces the loader-level context check.

Author: Jeremiah Griffith
Version: 3.1.0 (pyRevit 6.1.0 compatible, Owner radio on 3D tab)
"""

__title__ = "Sync\nViews"
__author__ = "Jeremiah Griffith"
__doc__ = "Sync pan and zoom across multiple open views with matching type and scale."
__context__ = 'zero-doc'

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow
from pyrevit.coreutils import envvars
from System.Windows import Visibility, Thickness, HorizontalAlignment, VerticalAlignment, FontWeights
from System.Windows.Media import SolidColorBrush
from System.Windows.Controls import (
    StackPanel, TextBlock, CheckBox, RadioButton, Expander, Orientation
)
from System.Windows.Threading import DispatcherTimer
from System import TimeSpan
import System.Windows.Media
import System.Windows

doc = revit.doc
uidoc = revit.uidoc

if not doc:
    forms.alert("No document open.\nOpen a Revit project first.", title="Sync Views")
    import sys
    sys.exit()


# =============================================================================
# CONTEXT DETECTION
# =============================================================================
def is_palette_context():
    """Check if running from MWE Palette's safe context."""
    try:
        val = envvars.get_pyrevit_env_var('MWE_PALETTE_CONTEXT')
        return val == 'TRUE'
    except:
        return False


# =============================================================================
# VIEW TYPE / SCALE DISPLAY HELPERS
# =============================================================================
VIEW_TYPE_NAMES = {
    1: "Floor Plan",
    2: "Ceiling Plan",
    3: "Section",
    4: "Elevation",
    10: "Drafting View",
    11: "3D View",
    17: "Detail",
    114: "Legend",
    115: "Structural Plan",
    116: "Area Plan",
}

SCALE_NAMES = {
    96: '1/8"', 48: '1/4"', 32: '3/8"', 24: '1/2"',
    16: '3/4"', 12: '1"', 8: '1-1/2"', 4: '3"',
}

SKIP_VIEW_TYPES = {12, 20}  # Schedule, DrawingSheet


def get_view_type_name(view_type_int):
    """Human-readable view type from integer."""
    return VIEW_TYPE_NAMES.get(view_type_int, "View")


def get_scale_string(scale_value):
    """Human-readable scale string from integer."""
    if scale_value == 1:
        return "1:1"
    return SCALE_NAMES.get(scale_value, "1:{}".format(scale_value))


# =============================================================================
# VIEW DATA COLLECTION - Returns ONLY serializable data (ints and strings)
# =============================================================================
def get_open_views_info():
    """Collect info about all open views. No API objects stored."""
    ui_views = uidoc.GetOpenUIViews()
    all_views = {}
    views_by_compat = {}

    for ui_view in ui_views:
        view_id = ui_view.ViewId
        view = doc.GetElement(view_id)
        if view is None:
            continue

        view_type_int = int(view.ViewType)
        if view_type_int in SKIP_VIEW_TYPES:
            continue

        vid_int = view_id.IntegerValue
        is_3d = isinstance(view, DB.View3D)

        scale_value = 0
        try:
            scale_param = view.get_Parameter(DB.BuiltInParameter.VIEW_SCALE)
            if scale_param and scale_param.HasValue:
                scale_value = scale_param.AsInteger()
        except:
            pass

        compat_key = (view_type_int, scale_value)

        view_info = {
            'vid_int': vid_int,
            'name': str(view.Name),
            'type_name': "3D View" if is_3d else get_view_type_name(view_type_int),
            'scale_string': get_scale_string(scale_value),
            'compat_key': compat_key,
            'is_3d': is_3d,
        }

        all_views[vid_int] = view_info
        if compat_key not in views_by_compat:
            views_by_compat[compat_key] = []
        views_by_compat[compat_key].append(view_info)

    return all_views, views_by_compat


# =============================================================================
# ONE-SHOT SYNC (2D zoom corners)
# =============================================================================
def do_one_shot_sync(source_vid_int, target_vid_ints):
    """Sync zoom rectangle from source to targets. Fetches fresh UIViews."""
    ui_views = uidoc.GetOpenUIViews()
    uv_dict = {uv.ViewId.IntegerValue: uv for uv in ui_views}

    source_uv = uv_dict.get(source_vid_int)
    if not source_uv:
        return 0

    try:
        corners = source_uv.GetZoomCorners()
        if len(corners) < 2:
            return 0
        c1, c2 = corners[0], corners[1]
    except:
        return 0

    synced = 0
    for vid in target_vid_ints:
        if vid == source_vid_int:
            continue
        uv = uv_dict.get(vid)
        if uv:
            try:
                uv.ZoomAndCenterRectangle(c1, c2)
                synced += 1
            except:
                pass

    return synced


# =============================================================================
# FAILURE SUPPRESSION FOR 3D PROPERTY SYNC
# =============================================================================
class _SuppressFailures(DB.IFailuresPreprocessor):
    """Suppress all Revit failure dialogs.

    Warnings get deleted. Hard errors (like worksharing ownership) that
    can't be deleted trigger a silent rollback instead of showing a dialog.
    """
    def PreprocessFailures(self, failuresAccessor):
        has_errors = False
        try:
            for msg in failuresAccessor.GetFailureMessages():
                try:
                    sev = msg.GetSeverity()
                    if sev == DB.FailureSeverity.Warning:
                        failuresAccessor.DeleteWarning(msg)
                    else:
                        has_errors = True
                except:
                    has_errors = True
        except:
            pass
        if has_errors:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# =============================================================================
# 3D VIEW PROPERTY SYNC (camera orientation + optional visual props)
# =============================================================================
def sync_view_properties(source_vid_int, target_vid_ints, options, active_doc=None):
    """Sync 3D camera orientation and optional visual properties via Transaction.

    Camera orientation is always synced (SetOrientation).
    Visual style and detail level are optional checkboxes.
    Worksharing errors silently roll back (no dialog).
    Returns number of views synced.
    """
    use_doc = active_doc or doc

    # Read source
    try:
        source_view = use_doc.GetElement(DB.ElementId(source_vid_int))
    except:
        return 0
    if source_view is None:
        return 0

    # Camera orientation (always for 3D)
    src_orient = None
    if isinstance(source_view, DB.View3D):
        try:
            src_orient = source_view.GetOrientation()
        except:
            pass

    # Optional properties
    src_style = None
    src_detail = None
    if options.get('visual_style', False):
        try:
            src_style = source_view.DisplayStyle
        except:
            pass
    if options.get('detail_level', False):
        try:
            src_detail = source_view.DetailLevel
        except:
            pass

    if src_orient is None and src_style is None and src_detail is None:
        return 0

    # Collect targets
    targets = []
    for vid in target_vid_ints:
        if vid == source_vid_int:
            continue
        try:
            tv = use_doc.GetElement(DB.ElementId(vid))
            if tv:
                targets.append(tv)
        except:
            continue

    if not targets:
        return 0

    # Apply in failure-suppressed Transaction
    synced = 0
    try:
        t = DB.Transaction(use_doc, "Sync View Properties")
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(_SuppressFailures())
        opts.SetClearAfterRollback(True)
        t.SetFailureHandlingOptions(opts)

        t.Start()
        try:
            for tv in targets:
                changed = False
                try:
                    if src_orient is not None and isinstance(tv, DB.View3D):
                        tv.SetOrientation(src_orient)
                        changed = True
                except:
                    pass
                try:
                    if src_style is not None:
                        tv.DisplayStyle = src_style
                        changed = True
                except:
                    pass
                try:
                    if src_detail is not None:
                        tv.DetailLevel = src_detail
                        changed = True
                except:
                    pass
                if changed:
                    synced += 1

            if synced > 0:
                t.Commit()
            else:
                t.RollBack()
        except:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
    except:
        pass

    return synced


# =============================================================================
# CONTINUOUS SYNC COMPONENTS (palette context only)
# =============================================================================
sync_state = None
sync_event = None
sync_handler = None


class SyncState:
    """Tracks which views are syncing and their last known zoom/camera state."""
    def __init__(self):
        self.synced_view_ids = set()
        self.last_zoom_states = {}
        self.is_syncing = False
        self.match_props = None  # None = 2D mode, dict = 3D mode

    def clear(self):
        self.synced_view_ids = set()
        self.last_zoom_states = {}
        self.is_syncing = False
        self.match_props = None


def _get_view_state(ui_view, active_doc, vid, include_3d):
    """Get serializable view state tuple. For 3D, includes camera orientation."""
    try:
        corners = ui_view.GetZoomCorners()
        if len(corners) < 2:
            return None
        c1, c2 = corners[0], corners[1]
        parts = [
            round(c1.X, 2), round(c1.Y, 2), round(c1.Z, 2),
            round(c2.X, 2), round(c2.Y, 2), round(c2.Z, 2),
        ]
        if include_3d:
            try:
                v = active_doc.GetElement(DB.ElementId(vid))
                if v and isinstance(v, DB.View3D):
                    orient = v.GetOrientation()
                    parts.extend([
                        round(orient.EyePosition.X, 2),
                        round(orient.EyePosition.Y, 2),
                        round(orient.EyePosition.Z, 2),
                        round(orient.ForwardDirection.X, 3),
                        round(orient.ForwardDirection.Y, 3),
                        round(orient.ForwardDirection.Z, 3),
                        round(orient.UpDirection.X, 3),
                        round(orient.UpDirection.Y, 3),
                        round(orient.UpDirection.Z, 3),
                    ])
            except:
                pass
        return tuple(parts)
    except:
        return None


def setup_continuous_sync():
    """Set up ExternalEvent for continuous sync. Only call from palette context."""
    global sync_state, sync_event, sync_handler

    try:
        from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

        class SyncViewsHandler(IExternalEventHandler):
            def Execute(self, app):
                global sync_state
                if sync_state is None or sync_state.is_syncing:
                    return
                if not sync_state.synced_view_ids:
                    return
                try:
                    current_uidoc = app.ActiveUIDocument
                    if not current_uidoc:
                        return
                    current_doc = current_uidoc.Document

                    ui_views = current_uidoc.GetOpenUIViews()
                    uv_dict = {uv.ViewId.IntegerValue: uv for uv in ui_views}

                    is_3d = sync_state.match_props is not None

                    # Detect which view changed
                    changed_vid = None
                    new_state = None
                    for vid in sync_state.synced_view_ids:
                        uv = uv_dict.get(vid)
                        if not uv:
                            continue
                        current = _get_view_state(uv, current_doc, vid, is_3d)
                        if current and sync_state.last_zoom_states.get(vid) != current:
                            changed_vid = vid
                            new_state = current
                            break

                    if changed_vid and new_state:
                        sync_state.is_syncing = True
                        try:
                            # 3D: SetOrientation first (camera angle),
                            # then ZoomAndCenterRectangle (zoom level)
                            if is_3d:
                                try:
                                    target_vids = [
                                        v for v in sync_state.synced_view_ids
                                        if v != changed_vid
                                    ]
                                    sync_view_properties(
                                        changed_vid, target_vids,
                                        sync_state.match_props,
                                        active_doc=current_doc
                                    )
                                except:
                                    pass

                            # Zoom sync via ZoomAndCenterRectangle (both modes)
                            c1 = DB.XYZ(new_state[0], new_state[1], new_state[2])
                            c2 = DB.XYZ(new_state[3], new_state[4], new_state[5])
                            for vid in sync_state.synced_view_ids:
                                if vid != changed_vid:
                                    uv = uv_dict.get(vid)
                                    if uv:
                                        try:
                                            uv.ZoomAndCenterRectangle(c1, c2)
                                        except:
                                            pass

                            # Update all states after sync
                            for vid in sync_state.synced_view_ids:
                                uv = uv_dict.get(vid)
                                if uv:
                                    state = _get_view_state(
                                        uv, current_doc, vid, is_3d
                                    )
                                    if state:
                                        sync_state.last_zoom_states[vid] = state
                        finally:
                            sync_state.is_syncing = False
                except:
                    pass

            def GetName(self):
                return "SyncViewsHandler"

        sync_state = SyncState()
        sync_handler = SyncViewsHandler()
        sync_event = ExternalEvent.Create(sync_handler)
        return True

    except Exception:
        return False


def initialize_zoom_states(view_ids, include_3d=False):
    """Snapshot current zoom/camera states for change detection."""
    global sync_state
    if not sync_state:
        return

    ui_views = uidoc.GetOpenUIViews()
    uv_dict = {uv.ViewId.IntegerValue: uv for uv in ui_views}

    sync_state.last_zoom_states = {}
    for vid in view_ids:
        uv = uv_dict.get(vid)
        if uv:
            state = _get_view_state(uv, doc, vid, include_3d)
            if state:
                sync_state.last_zoom_states[vid] = state


# =============================================================================
# WPF COLOR HELPERS
# =============================================================================
def _rgb(r, g, b):
    """Shorthand for SolidColorBrush from RGB."""
    return SolidColorBrush(System.Windows.Media.Color.FromRgb(r, g, b))


CLR_BLUE = _rgb(37, 99, 235)
CLR_GREEN = _rgb(22, 163, 74)
CLR_MUTED = _rgb(107, 114, 128)
CLR_LIGHT_MUTED = _rgb(156, 163, 175)
CLR_TEXT = _rgb(55, 65, 81)
CLR_HEADER = _rgb(55, 65, 81)
CLR_WHITE = _rgb(255, 255, 255)
CLR_BG = _rgb(240, 240, 240)


# =============================================================================
# XAML - Unified UI that adapts based on mode
# =============================================================================
XAML_STRING = '''
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sync Views"
        Height="500" Width="380"
        WindowStartupLocation="CenterScreen"
        Background="#F5F7FB"
        FontFamily="Segoe UI"
        FontSize="12"
        ShowInTaskbar="True"
        Topmost="True"
        ResizeMode="CanResizeWithGrip"
        MinHeight="300" MinWidth="300">

    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Main Content Panel -->
        <StackPanel x:Name="main_content" Grid.Row="0" Grid.RowSpan="4">
            <!-- Header -->
            <StackPanel Margin="0,0,0,8">
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Sync Views" FontSize="18" FontWeight="SemiBold" Foreground="#1F2937"/>
                    <TextBlock x:Name="mode_label" Text="" FontSize="10" Foreground="#6B7280"
                               VerticalAlignment="Center" Margin="10,0,0,0"/>
                </StackPanel>
            </StackPanel>

            <!-- Tab Toggle -->
            <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="tab_2d_btn" Grid.Column="0" Content="Sync 2D" Height="32"
                        FontWeight="SemiBold" Cursor="Hand"
                        Background="#FFFFFF" Foreground="#6B7280"
                        BorderBrush="#E5E8F0" BorderThickness="1,1,0,1"/>
                <Button x:Name="tab_3d_btn" Grid.Column="1" Content="Match 3D" Height="32"
                        FontWeight="SemiBold" Cursor="Hand"
                        Background="#FFFFFF" Foreground="#6B7280"
                        BorderBrush="#E5E8F0" BorderThickness="1"/>
            </Grid>

            <!-- Description -->
            <TextBlock x:Name="header_desc" Text="" FontSize="11" Foreground="#6B7280"
                       TextWrapping="Wrap" Margin="0,0,0,8"/>

            <!-- Views List -->
            <Border Background="White" BorderBrush="#E5E8F0" BorderThickness="1" CornerRadius="6"
                    Height="200">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Border Grid.Row="0" Background="#F9FAFB" BorderBrush="#E5E8F0"
                            BorderThickness="0,0,0,1" CornerRadius="6,6,0,0" Padding="12,8">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Open Views" FontWeight="SemiBold" Foreground="#374151"/>
                            <TextBlock x:Name="active_view_label" Text="" FontSize="10"
                                       Foreground="#16A34A" VerticalAlignment="Center" Margin="10,0,0,0"/>
                        </StackPanel>
                    </Border>

                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="12">
                        <StackPanel x:Name="views_container"/>
                    </ScrollViewer>
                </Grid>
            </Border>

            <!-- Match Options (3D only) -->
            <StackPanel x:Name="match_options_panel" Visibility="Collapsed" Margin="0,8,0,0">
                <Expander IsExpanded="True">
                    <Expander.Header>
                        <TextBlock Text="Match Options" FontWeight="SemiBold"
                                   Foreground="#374151" FontSize="12"/>
                    </Expander.Header>
                    <StackPanel Margin="20,5,0,5">
                        <CheckBox x:Name="opt_visual_style"
                                  Content="Visual Style (Wireframe, Shaded, etc.)"
                                  IsChecked="True"/>
                        <CheckBox x:Name="opt_detail_level"
                                  Content="Detail Level (Coarse, Medium, Fine)"
                                  IsChecked="True"/>
                    </StackPanel>
                </Expander>
            </StackPanel>

            <!-- Status -->
            <TextBlock x:Name="status_text" Text="Select 2 or more views" FontSize="11"
                       Foreground="#6B7280" HorizontalAlignment="Center" Margin="0,12,0,0"/>

            <!-- Buttons -->
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
                <Button x:Name="select_all_btn" Content="Select All"
                        Width="75" Margin="0,0,8,0" Padding="8,6"/>
                <Button x:Name="clear_btn" Content="Clear"
                        Width="60" Margin="0,0,8,0" Padding="8,6"/>
                <Button x:Name="close_btn" Content="Close"
                        Width="60" Margin="0,0,8,0" Padding="8,6"/>
                <Button x:Name="action_btn" Content="Start Sync" Width="90"
                        Background="#2563EB" Foreground="White" BorderThickness="0"
                        FontWeight="SemiBold" Padding="8,6" IsEnabled="False"/>
            </StackPanel>
        </StackPanel>

        <!-- Stop Button (shown when continuous syncing) -->
        <Button x:Name="stop_btn" Grid.Row="0" Grid.RowSpan="4" Visibility="Collapsed"
                Background="#DC2626" BorderThickness="0" Margin="4,2"/>
    </Grid>
</Window>
'''


# =============================================================================
# UNIFIED WINDOW CLASS
# =============================================================================
class SyncViewsWindow(WPFWindow):
    """Unified window for both continuous and one-shot sync modes."""

    def __init__(self, continuous_mode=False):
        WPFWindow.__init__(self, XAML_STRING, literal_string=True)

        self.continuous_mode = continuous_mode
        self.current_tab = '2d'
        self.selected_view_ids = set()
        self.view_checkboxes = {}
        self.view_info_map = {}
        self.active_vid_int = None
        self.sync_timer = None
        self.source_vid_int = None
        self.source_radios = {}
        self.owner_vid_int = None
        self.owner_radios = {}

        # Get active view ID as integer
        try:
            if uidoc.ActiveView:
                self.active_vid_int = uidoc.ActiveView.Id.IntegerValue
        except:
            pass

        # Configure UI based on mode
        if continuous_mode:
            self.mode_label.Text = "(Auto-Sync)"
            self.action_btn.Content = "Start Sync"
            self.sync_timer = DispatcherTimer()
            self.sync_timer.Interval = TimeSpan.FromMilliseconds(100)
            self.sync_timer.Tick += self.on_timer_tick
        else:
            self.mode_label.Text = "(One-Shot)"
            self.action_btn.Content = "Sync Now"

        # Wire up buttons
        self.action_btn.Click += self.on_action
        self.stop_btn.Click += self.on_stop
        self.select_all_btn.Click += self.on_select_all
        self.clear_btn.Click += self.on_clear
        self.close_btn.Click += self.on_close
        self.tab_2d_btn.Click += self.on_tab_2d
        self.tab_3d_btn.Click += self.on_tab_3d
        self.Closing += self.on_window_closing

        # Auto-detect: open on Match 3D tab if active view is 3D
        start_tab = '2d'
        try:
            if uidoc.ActiveView and isinstance(uidoc.ActiveView, DB.View3D):
                start_tab = '3d'
        except:
            pass
        self._switch_tab(start_tab)

    # -----------------------------------------------------------------
    # Timer (continuous mode only)
    # -----------------------------------------------------------------
    def on_timer_tick(self, sender, args):
        """Raise the external event for Revit to process on next idle."""
        if sync_event:
            sync_event.Raise()

    # -----------------------------------------------------------------
    # Tab switching
    # -----------------------------------------------------------------
    def on_tab_2d(self, sender, args):
        self._switch_tab('2d')

    def on_tab_3d(self, sender, args):
        self._switch_tab('3d')

    def _switch_tab(self, tab):
        self.current_tab = tab

        if tab == '2d':
            self.tab_2d_btn.Background = CLR_BLUE
            self.tab_2d_btn.Foreground = CLR_WHITE
            self.tab_3d_btn.Background = CLR_WHITE
            self.tab_3d_btn.Foreground = CLR_MUTED
            self.match_options_panel.Visibility = Visibility.Collapsed
            if self.continuous_mode:
                self.header_desc.Text = (
                    "Select views to sync. Pan/zoom any view, others follow automatically."
                )
                self.action_btn.Content = "Start Sync"
            else:
                self.header_desc.Text = (
                    "Check views to sync, select source (radio), then Sync Now."
                )
                self.action_btn.Content = "Sync Now"
        else:
            self.tab_3d_btn.Background = CLR_BLUE
            self.tab_3d_btn.Foreground = CLR_WHITE
            self.tab_2d_btn.Background = CLR_WHITE
            self.tab_2d_btn.Foreground = CLR_MUTED
            self.match_options_panel.Visibility = Visibility.Visible
            if self.continuous_mode:
                self.header_desc.Text = (
                    "Select 3D views to match. Orbit/zoom any view, others follow."
                )
                self.action_btn.Content = "Start Match"
            else:
                self.header_desc.Text = (
                    "Check 3D views to match, then click Match Now."
                )
                self.action_btn.Content = "Match Now"

        self.load_views()

    # -----------------------------------------------------------------
    # View loading
    # -----------------------------------------------------------------
    def load_views(self):
        self.views_container.Children.Clear()
        self.view_checkboxes.clear()
        self.view_info_map.clear()
        self.selected_view_ids.clear()
        self.source_radios.clear()
        self.source_vid_int = None
        self.owner_vid_int = None
        self.owner_radios.clear()
        self.active_view_label.Text = ""

        all_views, views_by_compat = get_open_views_info()
        self.view_info_map = all_views

        if self.current_tab == '3d':
            self._load_3d_views(all_views)
        else:
            self._load_2d_views(all_views, views_by_compat)

        self.update_status()

    def _load_3d_views(self, all_views):
        """Populate 3D tab with all open 3D views."""
        views_3d = [v for v in all_views.values() if v.get('is_3d', False)]

        if len(views_3d) < 2:
            msg = TextBlock()
            msg.Text = "Need 2+ open 3D views to match."
            msg.Foreground = CLR_MUTED
            msg.TextWrapping = System.Windows.TextWrapping.Wrap
            self.views_container.Children.Add(msg)
            return

        for view_info in sorted(views_3d, key=lambda x: x['name']):
            vid_int = view_info['vid_int']

            view_panel = StackPanel()
            view_panel.Orientation = Orientation.Horizontal
            view_panel.Margin = Thickness(0, 3, 0, 3)

            # Owner radio button
            rb = RadioButton()
            rb.Tag = vid_int
            rb.GroupName = "OwnerView"
            rb.Margin = Thickness(0, 0, 8, 0)
            rb.VerticalAlignment = VerticalAlignment.Center
            rb.Checked += self.on_owner_selected
            view_panel.Children.Add(rb)
            self.owner_radios[vid_int] = rb

            cb = CheckBox()
            cb.Tag = vid_int
            cb.Margin = Thickness(0, 0, 8, 0)
            cb.Checked += self.on_view_checked
            cb.Unchecked += self.on_view_unchecked
            view_panel.Children.Add(cb)
            self.view_checkboxes[vid_int] = cb

            name_text = TextBlock()
            name_text.Text = view_info['name']
            name_text.Foreground = CLR_TEXT
            name_text.FontSize = 11
            name_text.VerticalAlignment = VerticalAlignment.Center
            view_panel.Children.Add(name_text)

            if self.active_vid_int and vid_int == self.active_vid_int:
                self._add_active_label(view_panel)
                rb.IsChecked = True
                cb.IsChecked = True

            self.views_container.Children.Add(view_panel)

    def _load_2d_views(self, all_views, views_by_compat):
        """Populate 2D tab with views grouped by type+scale compatibility."""
        if not views_by_compat:
            msg = TextBlock()
            msg.Text = "No views open.\nOpen 2+ views of the same type/scale."
            msg.Foreground = CLR_MUTED
            msg.TextWrapping = System.Windows.TextWrapping.Wrap
            self.views_container.Children.Add(msg)
            return

        sorted_groups = sorted(
            views_by_compat.items(),
            key=lambda x: (get_view_type_name(x[0][0]), x[0][1])
        )
        has_syncable = False

        for compat_key, views in sorted_groups:
            # Filter out 3D views from the 2D tab
            views = [v for v in views if not v.get('is_3d', False)]
            if len(views) < 2:
                continue
            has_syncable = True

            # Group header
            header = StackPanel()
            header.Orientation = Orientation.Horizontal

            type_text = TextBlock()
            type_text.Text = views[0]['type_name']
            type_text.FontWeight = FontWeights.Bold
            type_text.Foreground = CLR_BLUE
            type_text.FontSize = 12
            header.Children.Add(type_text)

            scale_text = TextBlock()
            scale_text.Text = "  @  " + views[0]['scale_string']
            scale_text.Foreground = CLR_MUTED
            scale_text.FontSize = 11
            header.Children.Add(scale_text)

            count_text = TextBlock()
            count_text.Text = "  ({} views)".format(len(views))
            count_text.Foreground = CLR_LIGHT_MUTED
            count_text.FontSize = 10
            header.Children.Add(count_text)

            expander = Expander()
            expander.Header = header
            expander.IsExpanded = True
            expander.Margin = Thickness(0, 0, 0, 10)

            views_stack = StackPanel()
            views_stack.Margin = Thickness(20, 5, 0, 5)

            for view_info in sorted(views, key=lambda x: x['name']):
                vid_int = view_info['vid_int']

                view_panel = StackPanel()
                view_panel.Orientation = Orientation.Horizontal
                view_panel.Margin = Thickness(0, 2, 0, 2)

                cb = CheckBox()
                cb.Tag = vid_int
                cb.Margin = Thickness(0, 0, 8, 0)
                cb.Checked += self.on_view_checked
                cb.Unchecked += self.on_view_unchecked
                view_panel.Children.Add(cb)
                self.view_checkboxes[vid_int] = cb

                # Source radio button (one-shot mode only)
                if not self.continuous_mode:
                    rb = RadioButton()
                    rb.Tag = vid_int
                    rb.GroupName = "SourceView"
                    rb.Margin = Thickness(0, 0, 8, 0)
                    rb.VerticalAlignment = VerticalAlignment.Center
                    rb.Checked += self.on_source_selected
                    view_panel.Children.Add(rb)
                    self.source_radios[vid_int] = rb

                name_text = TextBlock()
                name_text.Text = view_info['name']
                name_text.Foreground = CLR_TEXT
                name_text.FontSize = 11
                name_text.VerticalAlignment = VerticalAlignment.Center
                view_panel.Children.Add(name_text)

                is_active = self.active_vid_int and vid_int == self.active_vid_int
                if is_active:
                    self._add_active_label(view_panel)
                    cb.IsChecked = True
                    if not self.continuous_mode:
                        rb.IsChecked = True
                        self.source_vid_int = vid_int
                        self.active_view_label.Text = "Source: " + view_info['name']

                views_stack.Children.Add(view_panel)

            expander.Content = views_stack
            self.views_container.Children.Add(expander)

        if not has_syncable:
            msg = TextBlock()
            msg.Text = "No syncable groups.\nNeed 2+ views with same type AND scale."
            msg.Foreground = CLR_MUTED
            msg.TextWrapping = System.Windows.TextWrapping.Wrap
            self.views_container.Children.Add(msg)

    def _add_active_label(self, panel):
        """Append green (Active) label to a view row."""
        label = TextBlock()
        label.Text = "  (Active)"
        label.Foreground = CLR_GREEN
        label.FontSize = 10
        label.FontWeight = FontWeights.SemiBold
        label.VerticalAlignment = VerticalAlignment.Center
        panel.Children.Add(label)

    # -----------------------------------------------------------------
    # Checkbox / radio handlers
    # -----------------------------------------------------------------
    def on_view_checked(self, sender, args):
        self.selected_view_ids.add(sender.Tag)
        self.update_status()

    def on_view_unchecked(self, sender, args):
        self.selected_view_ids.discard(sender.Tag)
        self.update_status()

    def on_source_selected(self, sender, args):
        """Handle source view radio button selection."""
        vid_int = sender.Tag
        self.source_vid_int = vid_int

        # Auto-check the checkbox for this view
        if vid_int in self.view_checkboxes:
            self.view_checkboxes[vid_int].IsChecked = True

        # Update the source label
        if vid_int in self.view_info_map:
            self.active_view_label.Text = "Source: " + self.view_info_map[vid_int]['name']

        self.update_status()

    def on_owner_selected(self, sender, args):
        """Handle 3D owner radio button selection."""
        vid_int = sender.Tag
        self.owner_vid_int = vid_int

        # Auto-check the checkbox for this view
        if vid_int in self.view_checkboxes:
            self.view_checkboxes[vid_int].IsChecked = True

        # Update the owner label
        if vid_int in self.view_info_map:
            self.active_view_label.Text = "Owner: " + self.view_info_map[vid_int]['name']

        self.update_status()

    # -----------------------------------------------------------------
    # Status bar
    # -----------------------------------------------------------------
    def update_status(self):
        count = len(self.selected_view_ids)

        if count < 2:
            self.status_text.Text = "Select {} more view(s)".format(2 - count)
            self.status_text.Foreground = CLR_MUTED
            self.action_btn.IsEnabled = False
        else:
            verb = "match" if self.current_tab == '3d' else "sync"
            if self.continuous_mode:
                self.status_text.Text = "{} views ready to {}".format(count, verb)
            else:
                if self.current_tab == '3d':
                    self.status_text.Text = "{} views ready to match".format(count)
                elif self.source_vid_int and self.source_vid_int in self.selected_view_ids:
                    self.status_text.Text = "{} views ready".format(count)
                else:
                    self.status_text.Text = "{} views ready - select a source".format(count)
            self.status_text.Foreground = CLR_GREEN
            self.action_btn.IsEnabled = True

    # -----------------------------------------------------------------
    # Action handlers
    # -----------------------------------------------------------------
    def on_action(self, sender, args):
        """Handle main action button."""
        if len(self.selected_view_ids) < 2:
            return

        if self.continuous_mode:
            self._start_continuous_sync()
        else:
            self._do_one_shot()

    def _start_continuous_sync(self):
        """Start continuous sync mode (palette only)."""
        global sync_state

        if not sync_state:
            return

        # Configure sync state
        sync_state.synced_view_ids = set(self.selected_view_ids)

        is_3d = self.current_tab == '3d'
        if is_3d:
            sync_state.match_props = {
                'visual_style': self.opt_visual_style.IsChecked,
                'detail_level': self.opt_detail_level.IsChecked,
            }
        else:
            sync_state.match_props = None

        initialize_zoom_states(sync_state.synced_view_ids, include_3d=is_3d)

        # Collapse UI into minimal stop button
        self.main_content.Visibility = Visibility.Collapsed
        self.stop_btn.Visibility = Visibility.Visible

        self.MinWidth = 0
        self.MinHeight = 0
        self.Width = 55
        self.Height = 80
        self.Title = "___STOP___"
        self.Background = CLR_BG
        self.ResizeMode = System.Windows.ResizeMode.NoResize
        self.Content.Margin = Thickness(2)

        if self.sync_timer:
            self.sync_timer.Start()

    def _do_one_shot(self):
        """Perform one-shot sync/match."""
        # Determine source: 3D owner > user-selected > active > first selected
        source_vid = None
        if self.current_tab == '3d' and self.owner_vid_int and self.owner_vid_int in self.selected_view_ids:
            source_vid = self.owner_vid_int
        elif self.source_vid_int and self.source_vid_int in self.selected_view_ids:
            source_vid = self.source_vid_int
        elif self.active_vid_int and self.active_vid_int in self.selected_view_ids:
            source_vid = self.active_vid_int
        else:
            source_vid = sorted(self.selected_view_ids)[0]

        target_vids = [v for v in self.selected_view_ids if v != source_vid]

        # 3D: SetOrientation + properties FIRST
        if self.current_tab == '3d':
            props = {
                'visual_style': self.opt_visual_style.IsChecked,
                'detail_level': self.opt_detail_level.IsChecked,
            }
            sync_view_properties(source_vid, target_vids, props)

        # ZoomAndCenterRectangle for zoom level (both tabs)
        synced = do_one_shot_sync(source_vid, list(self.selected_view_ids))

        verb = "Matched" if self.current_tab == '3d' else "Synced"
        if synced > 0:
            self.status_text.Text = "{} {} views!".format(verb, synced)
            self.status_text.Foreground = CLR_GREEN

    # -----------------------------------------------------------------
    # Stop / close / cleanup
    # -----------------------------------------------------------------
    def on_stop(self, sender, args):
        """Stop sync and close window."""
        self._stop_sync()
        self.Close()

    def _stop_sync(self):
        """Stop the sync operation and clean up."""
        global sync_state
        if self.sync_timer:
            self.sync_timer.Stop()
        if sync_state:
            sync_state.clear()

    def on_window_closing(self, sender, args):
        """Clean up when window closes."""
        self._stop_sync()

    def on_select_all(self, sender, args):
        for cb in self.view_checkboxes.values():
            cb.IsChecked = True

    def on_clear(self, sender, args):
        for cb in self.view_checkboxes.values():
            cb.IsChecked = False
        self.selected_view_ids.clear()
        self.update_status()

    def on_close(self, sender, args):
        self._stop_sync()
        self.Close()


# =============================================================================
# MAIN EXECUTION
# =============================================================================
try:
    use_continuous = is_palette_context()

    if use_continuous:
        setup_continuous_sync()
        window = SyncViewsWindow(continuous_mode=True)
        window.Show()  # Modeless
    else:
        window = SyncViewsWindow(continuous_mode=False)
        window.ShowDialog()  # Modal

except Exception as e:
    forms.alert("Error: {}".format(str(e)), title="Sync Views Error")
