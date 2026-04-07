# -*- coding: utf-8 -*-
"""NW Weight Tool - Model Setup (UI-only)

Behavior (v5):
- Shell Year defaults to **Auto (Revit YYYY)**.
- User may override with an explicit year (2022-2026).
- At runtime, if Auto is selected, the running Revit year is detected and used.
- UI hint text never says "Default shell for 20XX".

Expected injected context:
MODEL_SETUP_CONTEXT = {
    "ext_event": <ExternalEvent>,
    "req": <ModelSetupRequest>,
    "show_main": <callable to re-open main window>
}
"""

import clr

from pyrevit import revit

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows.Markup import XamlReader

from System.Windows.Interop import WindowInteropHelper
from System.Windows import MessageBox
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import OpenFileDialog, DialogResult


# ----------------------------- Network defaults -----------------------------

DEFAULT_SHARED_PARAMETER_FILE = r"\\si.net\si\TCM\BIM\Shared Parameters\NW_Shared Parameters.txt"

SCHEDULE_ASSET_DIR = r"\\si.net\si\TCM\BIM\PyRevit\PyRevit_Content_Library\Schedule_Assets"

DEFAULT_SHELL_BY_YEAR = {
    2022: SCHEDULE_ASSET_DIR + r"\PyRevit_Schedule_Asset_R_22.rvt",
    2023: SCHEDULE_ASSET_DIR + r"\PyRevit_Schedule_Asset_R_23.rvt",
    2024: SCHEDULE_ASSET_DIR + r"\PyRevit_Schedule_Asset_R_24.rvt",
    2025: SCHEDULE_ASSET_DIR + r"\PyRevit_Schedule_Asset_R_25.rvt",
    2026: SCHEDULE_ASSET_DIR + r"\PyRevit_Schedule_Asset_R_26.rvt",
}


# ----------------------------- Helpers -----------------------------

def _strip_quotes(s):
    s = (s or "").strip()
    if len(s) >= 2 and ((s[0] == '"' and s[-1] == '"') or (s[0] == "'" and s[-1] == "'")):
        s = s[1:-1].strip()
    return s


def _revit_year_default():
    try:
        return int(revit.app.VersionNumber)
    except:
        return 2025


def _year_to_index(year):
    # ComboBox indices: 0=Auto, 1=2022, 2=2023, ...
    mapping = {2022: 1, 2023: 2, 2024: 3, 2025: 4, 2026: 5}
    try:
        return mapping.get(int(year), 0)
    except:
        return 0


# ----------------------------- Injected context -----------------------------

try:
    _CTX = MODEL_SETUP_CONTEXT
    _EXT_EVENT = _CTX["ext_event"]
    _REQ = _CTX["req"]
    _SHOW_MAIN = _CTX["show_main"]
except:
    _CTX = None
    _EXT_EVENT = None
    _REQ = None
    _SHOW_MAIN = None


# ----------------------------- WPF -----------------------------

XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Model Setup"
        Width="750" Height="610"
        WindowStartupLocation="CenterScreen"
        Topmost="False"
        ResizeMode="CanResizeWithGrip" Background="#F5F6F8" FontFamily="Segoe UI">
<Window.Resources>
  <Style TargetType="TextBlock" x:Key="SectionTitle">
    <Setter Property="FontSize" Value="15"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
  </Style>

  <Style TargetType="Border" x:Key="Card">
    <Setter Property="CornerRadius" Value="12"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush" Value="#D9D9D9"/>
    <Setter Property="Background" Value="White"/>
    <Setter Property="Padding" Value="14"/>
    <Setter Property="Margin" Value="0,0,0,12"/>
  </Style>

  <Style TargetType="Button" x:Key="PrimaryBtn">
    <Setter Property="Padding" Value="14,8"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="BorderThickness" Value="0"/>
    <Setter Property="Background" Value="#FF6A00"/>
    <Setter Property="Foreground" Value="White"/>
  </Style>

  <Style TargetType="Button" x:Key="SecondaryBtn">
    <Setter Property="Padding" Value="14,8"/>
    <Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush" Value="#D0D0D0"/>
    <Setter Property="Background" Value="White"/>
    <Setter Property="Foreground" Value="#333333"/>
  </Style>

  <Style TargetType="Button" x:Key="HelpBtn">
    <Setter Property="Width" Value="20"/>
    <Setter Property="Height" Value="20"/>
    <Setter Property="FontWeight" Value="Bold"/>
    <Setter Property="Padding" Value="0"/>
    <Setter Property="BorderThickness" Value="1"/>
    <Setter Property="BorderBrush" Value="#CFCFCF"/>
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Foreground" Value="#333333"/>
  </Style>
</Window.Resources>

  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="NW Weight Tool - Model Setup" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,10"/>

    <StackPanel Grid.Row="1" Orientation="Vertical" Margin="0,0,0,10">

      <!-- Schedules -->
      <GroupBox BorderThickness="0" Background="Transparent" Padding="10" Margin="0,0,0,10">
<GroupBox.Header>
          <DockPanel LastChildFill="False">
            <TextBlock Text="Schedules to Import" FontWeight="SemiBold"  Style="{StaticResource SectionTitle}"/>
            <Button x:Name="BtnHelpSchedules" Content="?" Width="18" Height="18"
                    Margin="8,0,0,0" Padding="0" FontWeight="Bold"
                    ToolTip="Help: Schedules to Import" Style="{StaticResource HelpBtn}" />
          </DockPanel>
        </GroupBox.Header>
        <StackPanel>
          <CheckBox x:Name="ChkWeight" Content="Import NW__Weight Schedule" Margin="0,2,0,2"/>          <CheckBox x:Name="ChkQC"     Content="Import NW__Weight Schedule QC report" Margin="0,2,0,2"/>

          <CheckBox x:Name="ChkCat" Content="Import NW__Weight Schedule with Category" Margin="0,2,0,2"/>
          <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
            <TextBlock Text="Shell Year:" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <ComboBox x:Name="CboYear" Width="170">
              <ComboBoxItem Content="Auto"/>
              <ComboBoxItem Content="2022"/>
              <ComboBoxItem Content="2023"/>
              <ComboBoxItem Content="2024"/>
              <ComboBoxItem Content="2025"/>
              <ComboBoxItem Content="2026"/>
            </ComboBox>

            <TextBlock Text="   Collision:" VerticalAlignment="Center" Margin="16,0,8,0"/>
            <ComboBox x:Name="CboCollision" Width="220" SelectedIndex="1">
              <ComboBoxItem Content="Skip if exists"/>
              <ComboBoxItem Content="Rename existing then import"/>
              <ComboBoxItem Content="Overwrite existing"/>
            </ComboBox>
          </StackPanel>

          <TextBlock x:Name="TxtShellHint" Margin="0,10,0,0" Opacity="0.85"/>
        </StackPanel>
      </GroupBox>
      <!-- Parameters -->
      <GroupBox BorderThickness="0" Background="Transparent" Padding="10">
<GroupBox.Header>
          <DockPanel LastChildFill="False">
            <TextBlock Text="Required Shared Parameters" FontWeight="SemiBold"  Style="{StaticResource SectionTitle}"/>
            <Button x:Name="BtnHelpParams" Content="?" Width="18" Height="18"
                    Margin="8,0,0,0" Padding="0" FontWeight="Bold"
                    ToolTip="Help: Required Shared Parameters" Style="{StaticResource HelpBtn}" />
          </DockPanel>
        </GroupBox.Header>
        <StackPanel>
          <CheckBox x:Name="ChkParams"
                    Content="Install/Bind CPBOM_Category, CP_Weight, CP_Volume (Construction)"
                    Margin="0,2,0,10"/>

          <TextBlock Text="Shared Parameter TXT:"/>
          <StackPanel Orientation="Horizontal" Margin="0,4,0,0">
            <TextBox x:Name="TxtSpFile" Width="540" Height="24"/>
            <Button x:Name="BtnBrowseSp" Content="Browse" Width="110" Height="24" Margin="10,0,0,0"/>
          </StackPanel>

          <TextBlock Text="Note: Setup will temporarily change Revit's SharedParametersFilename and restore it afterward."
                     Margin="0,6,0,0" Opacity="0.75"/>
        </StackPanel>
      </GroupBox>
    </StackPanel>

    <!-- Status -->
    <GroupBox BorderThickness="0" Background="Transparent" Grid.Row="2" Padding="10">
<GroupBox.Header>
        <DockPanel LastChildFill="False">
          <TextBlock Text="Status" FontWeight="SemiBold"  Style="{StaticResource SectionTitle}"/>
          <Button x:Name="BtnHelpStatus" Content="?" Width="18" Height="18"
                  Margin="8,0,0,0" Padding="0" FontWeight="Bold"
                  ToolTip="Help: Status" Style="{StaticResource HelpBtn}" />
        </DockPanel>
      </GroupBox.Header>
      <TextBox x:Name="TxtStatus"
               FontFamily="Consolas"
               FontSize="12"
               TextWrapping="Wrap"
               VerticalScrollBarVisibility="Auto"
               IsReadOnly="True"
               Background="Transparent"
               BorderThickness="0"/>
    </GroupBox>
    <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="BtnBack" Content="Back" Width="120" Height="32" Margin="0,0,10,0"/>
      <Button x:Name="BtnRun" Content="Run Setup" Width="120" Height="32" Margin="0,0,10,0"/>
      <Button x:Name="BtnClose" Content="Close" Width="120" Height="32"/>
    </StackPanel>
  </Grid>
</Window>
"""


class ModelSetupWindow(object):
    def __init__(self):
        self._w = XamlReader.Parse(XAML)

        self.ChkWeight = self._w.FindName("ChkWeight")
        self.ChkQC = self._w.FindName("ChkQC")

        self.ChkCat = self._w.FindName("ChkCat")
        self.CboYear = self._w.FindName("CboYear")
        self.CboCollision = self._w.FindName("CboCollision")

        self.TxtShellHint = self._w.FindName("TxtShellHint")

        self.ChkParams = self._w.FindName("ChkParams")
        self.TxtSpFile = self._w.FindName("TxtSpFile")
        self.BtnBrowseSp = self._w.FindName("BtnBrowseSp")

        self.TxtStatus = self._w.FindName("TxtStatus")
        self.BtnBack = self._w.FindName("BtnBack")
        self.BtnRun = self._w.FindName("BtnRun")
        self.BtnClose = self._w.FindName("BtnClose")

        # Help buttons
        self.BtnHelpSchedules = self._w.FindName("BtnHelpSchedules")
        self.BtnHelpParams    = self._w.FindName("BtnHelpParams")
        self.BtnHelpStatus    = self._w.FindName("BtnHelpStatus")
        self.TxtSpFile.Text = DEFAULT_SHARED_PARAMETER_FILE

        # Default selection
        # - If user previously ran with an explicit year, preserve it
        # - Otherwise, keep Auto selected
        try:
            prev_year = getattr(_REQ, "shell_year", None) if _REQ is not None else None
        except:
            prev_year = None

        if prev_year is not None:
            self.CboYear.SelectedIndex = _year_to_index(prev_year)
        else:
            self.CboYear.SelectedIndex = 0  # Auto

        self._update_shell_hint()

        self.BtnBrowseSp.Click += self._browse_sp
        self.CboYear.SelectionChanged += self._on_year_changed

        self.BtnBack.Click += self._back
        self.BtnRun.Click += self._run
        self.BtnClose.Click += self._close

        # Help button clicks
        self.BtnHelpSchedules.Click += self._help_schedules
        self.BtnHelpParams.Click += self._help_params
        self.BtnHelpStatus.Click += self._help_status
        if _EXT_EVENT is None or _REQ is None:
            self._set_status("ERROR: Open Model Setup from the main Weight tool 'Model Setup' button.")
        else:
            _REQ._ui_set_status = self._set_status
            self._set_status("Ready.\n\nSelect options and click Run Setup.")

        # ----------------------------- Help popups -----------------------------
    def _help_schedules(self, sender, args):
            MessageBox.Show(
                "Schedules to Import:\n"
                "\n"
                "NW_Weight Schedule:\n"
                "- Total dry weight, volume, and wet weight.\n"
                "\n"
                "NW_Weight Schedule QC report:\n"
                "- Itemized QC schedule for review and validation.\n"
                "- Includes Revit Category, CP_BOM_Category, Family and Type, CP_Weight, CP_Volume, and Wet Weight.\n"
                "- Use it to identify missing or incorrect data.\n"
                "\n"
                "NW_Weight Schedule with Category:\n"
                "- Totals by CP_BOM_Category for reporting.\n"
                "- Shows dry weight, volume, and wet weight.\n"
                "\n"
                "Shell Year:\n"
                "- Controls which schedule template file is used for import.\n"
                "\n"
                "Collision:\n"
                "- Rename existing then import keeps old schedules and brings in fresh copies.",
                "NW Weight Tool - Schedule Import"
            )

    def _help_params(self, sender, args):
            MessageBox.Show(
            """Required Shared Parameters:
            Installs/binds the shared parameters the tool writes to.
            
            - CP_Weight: Stores calculated/derived weight.
            - CP_Volume: Stores calculated/derived volume.
            - CP_BOM_Category: Stores a simplified category label for BOM/QC use.
            
            Shared Parameter TXT:
            Browse to the NW shared parameter text file. Setup temporarily points Revit to this file to bind the parameters, then restores your original setting afterward.""",
                        "NW Weight Tool - Help (Parameters)"
            )

    def _help_status(self, sender, args):
            MessageBox.Show(
            """Status:
            Shows a running log of what Model Setup is doing (imports/installs and any warnings).
            
            Tip: If something fails, copy the Status text and send it to BIM/VDC.""",
                        "NW Weight Tool - Help (Status)"
            )

    def Show(self):
        self._w.Show()

    def _set_status(self, text):
        try:
            self.TxtStatus.Text = text or ""
            try:
                self.TxtStatus.ScrollToEnd()
            except:
                pass
        except:
            pass

    def _year_value(self):
        try:
            c = (self.CboYear.SelectedItem.Content or "").strip()
            if not c or c.lower().startswith("auto"):
                return None
            return int(c)
        except:
            return None

    def _default_shell_for_year(self, yr):
        try:
            if yr is None:
                return ""
            return DEFAULT_SHELL_BY_YEAR.get(int(yr), "")
        except:
            return ""

    def _update_shell_hint(self):
        detected = _revit_year_default()
        yr = self._year_value()  # None means Auto
        effective = yr if yr is not None else detected
        shell_path = self._default_shell_for_year(effective)

        try:
            if yr is None:
                self.TxtShellHint.Text = "Auto:\n{}".format(op.dirname(shell_path))
            else:
                self.TxtShellHint.Text = "{}:\n{}".format(yr, shell_path)
        except:
            pass

    def _on_year_changed(self, sender, args):
        self._update_shell_hint()

    def _browse_sp(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Title = "Select Shared Parameter TXT"
        dlg.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        if dlg.ShowDialog() == DialogResult.OK:
            self.TxtSpFile.Text = dlg.FileName

    def _collision_mode(self):
        try:
            ctext = self.CboCollision.SelectedItem.Content or ""
            if "Skip" in ctext:
                return "skip"
            if "Overwrite" in ctext:
                return "overwrite"
            return "rename"
        except:
            return "rename"

    def _run(self, sender, args):
        if _EXT_EVENT is None or _REQ is None:
            self._set_status("Cannot run: missing context.")
            return

        selected_year = self._year_value()  # None = Auto

        _REQ.shell_year = selected_year
        _REQ.collision_mode = self._collision_mode()

        _REQ.import_weight = bool(self.ChkWeight.IsChecked)
        _REQ.import_qc = bool(self.ChkQC.IsChecked)

        _REQ.import_cat = bool(self.ChkCat.IsChecked)
        _REQ.install_params = bool(self.ChkParams.IsChecked)
        _REQ.shared_param_file = _strip_quotes(self.TxtSpFile.Text)

        detected = _revit_year_default()
        effective_year = selected_year if selected_year is not None else detected
        computed_shell = self._default_shell_for_year(effective_year)

        year_label = str(selected_year) if selected_year is not None else "Auto"

        self._set_status(
            "Queued...\n\n"
            "Shell Year: {}\n"
            "{}:\n"
            "Collision: {}\n\n"
            "Running via ExternalEvent...".format(year_label, computed_shell, _REQ.collision_mode)
        )

        _EXT_EVENT.Raise()

    def _back(self, sender, args):
        try:
            if _REQ is not None:
                _REQ._ui_set_status = None
        except:
            pass
        try:
            if _SHOW_MAIN:
                _SHOW_MAIN()
        except:
            pass
        self._close(sender, args)

    def _close(self, sender, args):
        try:
            if _REQ is not None:
                _REQ._ui_set_status = None
        except:
            pass
        try:
            self._w.Close()
        except:
            pass


win = ModelSetupWindow()
win.Show()
