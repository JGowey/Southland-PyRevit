# -*- coding: utf-8 -*-
"""
SP Loader - Entry Point (script.py)
====================================
pyRevit button entry point for the SP Loader (Param Binder) tool.

Sets the config directory paths, then opens the Launcher form which lists
available parameter configs and lets the user install, create, edit,
duplicate, or delete them.

File layout:
    script.py         - This file (entry point + Launcher UI)
    binder_core.py    - Backend: paths, JSON I/O, SP files, installer
    ui_builder.py     - Config editor UI (lazy-imported on Create/Edit)
    ui_param_editor.py - (Removed — functionality is in ui_builder.py)
    bundle.yaml       - pyRevit button metadata
"""

import os
import json
import time
import getpass
import clr

clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
from System import Array, String
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Application, Form, Button, ListView, ListViewItem, ColumnHeader,
    AnchorStyles, View, DialogResult, Label, TextBox, ComboBox, ComboBoxStyle,
    FormStartPosition, MessageBox, MessageBoxButtons
)

from Autodesk.Revit.UI import TaskDialog

import binder_core as core


# =============================================================================
# LAUNCHER HELPERS
# =============================================================================

def _fmt_mtime(path):
    """Formats a file's modification time as 'YYYY-MM-DD HH:MM'."""
    try:
        ts = os.path.getmtime(path)
        return time.strftime("%Y-%m-%d %H:%M", time.localtime(ts))
    except:
        return ""

def _read_display_name(path):
    """Reads the display_name field from a JSON config file."""
    try:
        data = core.load_json_file(path)
        return data.get("display_name", "") or os.path.basename(path)
    except:
        return os.path.basename(path)


# =============================================================================
# DUPLICATE CONFIG DIALOG
# =============================================================================

class DuplicateDialog(Form):
    """Small dialog for duplicating a config: enter new name and choose scope."""

    def __init__(self, default_name, can_shared):
        Form.__init__(self)
        self.Text = "Duplicate Config"
        self.Size = Size(430, 170)
        self.StartPosition = FormStartPosition.CenterParent

        lbl = Label(); lbl.Text = "New display name:"; lbl.Location = Point(12, 14); lbl.Size = Size(120, 18)
        self.txt = TextBox(); self.txt.Location = Point(12, 34); self.txt.Size = Size(380, 22); self.txt.Text = default_name

        lbl2 = Label(); lbl2.Text = "Save to:"; lbl2.Location = Point(12, 64); lbl2.Size = Size(50, 18)
        self.cbo = ComboBox(); self.cbo.Location = Point(72, 62); self.cbo.Size = Size(140, 22)
        self.cbo.DropDownStyle = ComboBoxStyle.DropDownList
        if can_shared: self.cbo.Items.Add("shared")
        self.cbo.Items.Add("user")
        self.cbo.SelectedIndex = 0

        btnOK = Button(); btnOK.Text = "OK"; btnOK.Location = Point(220, 96); btnOK.Size = Size(80, 26)
        btnCancel = Button(); btnCancel.Text = "Cancel"; btnCancel.Location = Point(312, 96); btnCancel.Size = Size(80, 26)

        def _ok(_s,_a): self.DialogResult = DialogResult.OK; self.Close()
        def _cancel(_s,_a): self.DialogResult = DialogResult.Cancel; self.Close()
        btnOK.Click += _ok; btnCancel.Click += _cancel

        for w in [lbl, self.txt, lbl2, self.cbo, btnOK, btnCancel]: self.Controls.Add(w)


# =============================================================================
# LAUNCHER FORM
# =============================================================================

class Launcher(Form):
    """
    Main launcher window. Lists all available parameter configs (shared + user)
    and provides buttons to install, create, edit, duplicate, or delete them.
    """

    def __init__(self):
        Form.__init__(self)
        self.Text = "Param Binder"
        self.Size = Size(900, 580)
        self.MinimumSize = Size(720, 440)
        self.StartPosition = FormStartPosition.CenterScreen

        # Path info label
        self.lblPaths = Label()
        self.lblPaths.Text = "Shared: {0}\nMine:  {1}".format(core.shared_dir() or "(none)", core.user_dir() or "(none)")
        self.lblPaths.Location = Point(12, 8); self.lblPaths.Size = Size(860, 28)

        # Config list
        self.lv = ListView(); self.lv.View = View.Details; self.lv.FullRowSelect = True; self.lv.MultiSelect = False
        self.lv.Location = Point(12, 44); self.lv.Size = Size(860, 430)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        for t, w in [("Display Name", 300), ("Scope", 80), ("File Name", 300), ("Modified", 150)]:
            ch = ColumnHeader(); ch.Text = t; ch.Width = w; self.lv.Columns.Add(ch)

        # Fix: ensure clicking anywhere on a row selects it
        def _lv_mouse_down(_s, e):
            hit = self.lv.HitTest(e.Location)
            if hit.Item is not None:
                hit.Item.Selected = True
        self.lv.MouseDown += _lv_mouse_down

        # Action buttons
        self.btnInstall = Button(); self.btnInstall.Text = "Install Selected"; self.btnInstall.Size = Size(120, 30)
        self.btnCreate = Button(); self.btnCreate.Text = "Create New"; self.btnCreate.Size = Size(120, 30)
        self.btnEdit = Button(); self.btnEdit.Text = "Edit Selected"; self.btnEdit.Size = Size(120, 30)
        self.btnDup = Button(); self.btnDup.Text = "Duplicate Selected"; self.btnDup.Size = Size(150, 30)
        self.btnDelete = Button(); self.btnDelete.Text = "Delete Selected"; self.btnDelete.Size = Size(140, 30)
        self.btnRefresh = Button(); self.btnRefresh.Text = "Refresh"; self.btnRefresh.Size = Size(85, 30)
        self.btnClose = Button(); self.btnClose.Text = "Close"; self.btnClose.Size = Size(90, 30)

        self._btns = [self.btnInstall, self.btnCreate, self.btnEdit, self.btnDup, self.btnDelete, self.btnRefresh, self.btnClose]

        # Dynamic button layout: action buttons left-aligned, Refresh+Close right-aligned
        def _layout(_s=None,_a=None):
            pad = 8
            y = self.ClientSize.Height - self.btnInstall.Height - 12
            x = 12
            for b in self._btns[:-2]:
                b.Location = Point(x, y); b.Anchor = AnchorStyles.Left | AnchorStyles.Bottom
                x += b.Width + pad
            self.btnRefresh.Location = Point(self.ClientSize.Width - self.btnRefresh.Width - 5 - self.btnClose.Width - pad, y)
            self.btnRefresh.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            self.btnClose.Location = Point(self.ClientSize.Width - self.btnClose.Width - 10, y)
            self.btnClose.Anchor = AnchorStyles.Right | AnchorStyles.Bottom
            self.lv.Height = y - self.lv.Top - 12
        self.Resize += _layout
        _layout()

        for w in [self.lblPaths, self.lv] + self._btns: self.Controls.Add(w)

        # Wire button handlers
        self.btnRefresh.Click += self._reload
        self.btnCreate.Click += self._create
        self.btnEdit.Click += self._edit
        self.btnInstall.Click += self._install
        self.btnDup.Click += self._duplicate
        self.btnDelete.Click += self._delete
        self.btnClose.Click += lambda s, a: self.Close()

        self._reload()

    # --- Helpers ---

    def _sel_path(self):
        """Returns the file path of the selected config, or None."""
        if self.lv.SelectedItems.Count == 0: return None
        tag = self.lv.SelectedItems[0].Tag
        return tag if tag and os.path.exists(tag) else None

    def _reload(self, *_):
        """Scans shared and user config directories and repopulates the list."""
        self.lv.Items.Clear()
        rows = []
        for scope, folder in [("shared", core.shared_dir()), ("user", core.user_dir())]:
            if not folder: continue
            try:
                for fn in os.listdir(folder):
                    if not fn.lower().endswith(".json"): continue
                    p = os.path.join(folder, fn)
                    rows.append((scope, p))
            except: pass

        for scope, p in rows:
            sub = Array[String]([
                _read_display_name(p),
                scope,
                os.path.basename(p),
                _fmt_mtime(p)
            ])
            it = ListViewItem(sub); it.Tag = p
            self.lv.Items.Add(it)

    # --- Actions ---

    def _create(self, *_):
        """Opens the Builder UI to create a new config (lazy import)."""
        import ui_builder as builder
        builder.show_builder(None)
        self._reload()

    def _edit(self, *_):
        """Opens the Builder UI to edit the selected config (lazy import)."""
        p = self._sel_path()
        if not p:
            TaskDialog.Show("Param Binder", "Select a config to edit.")
            return
        import ui_builder as builder
        builder.show_builder(p)
        self._reload()

    def _install(self, *_):
        """Installs the selected config's parameters into the active document."""
        p = self._sel_path()
        if not p:
            TaskDialog.Show("Param Binder", "Select a config to install.")
            return
        try:
            core.install_from_json(p)
        except Exception as e:
            TaskDialog.Show("Param Binder - Install Error", str(e))

    def _duplicate(self, *_):
        """Duplicates the selected config with a new name and scope."""
        p = self._sel_path()
        if not p:
            TaskDialog.Show("Param Binder", "Select a config to duplicate.")
            return
        data = core.load_json_file(p)
        old_name = data.get("display_name", os.path.basename(p))
        dlg = DuplicateDialog(old_name + " copy", core.can_edit_shared())
        if dlg.ShowDialog() != DialogResult.OK:
            return
        new_disp = dlg.txt.Text.strip() or (old_name + " copy")
        scope = dlg.cbo.SelectedItem or "user"
        folder = core.shared_dir() if scope == "shared" else core.user_dir()
        if scope == "shared" and (not os.path.isdir(folder) or not os.access(folder, os.W_OK)):
            folder = core.user_dir()
            scope = "user"
            TaskDialog.Show("Param Binder", "No write access to shared; saving to your user folder.")

        data["display_name"] = new_disp
        try:
            data["last_modified_by"] = getpass.getuser()
        except:
            pass
        name = core.slugify(new_disp) + ".json"
        target = os.path.join(folder, name)
        # Write to temp file first, then rename (atomic save)
        tmp = target + ".tmp"
        try:
            with open(tmp, "w") as fp: json.dump(data, fp, indent=2)
            if os.path.exists(target):
                try: os.remove(target)
                except: pass
            os.rename(tmp, target)
            TaskDialog.Show("Param Binder", "Duplicated to:\n{}".format(target))
        except Exception as e:
            TaskDialog.Show("Param Binder - Duplicate Error", str(e))
        self._reload()

    def _delete(self, *_):
        """Deletes the selected config file after confirmation."""
        p = self._sel_path()
        if not p:
            TaskDialog.Show("Param Binder", "Select a config to delete.")
            return
        res = MessageBox.Show("Permanently delete this file?\n\n{}".format(p), "Param Binder",
                                 MessageBoxButtons.OKCancel)
        if res.ToString() != "OK":
            return
        try:
            os.remove(p)
            self._reload()
        except Exception as e:
            TaskDialog.Show("Param Binder - Delete Error", str(e))


def show_launcher():
    """Creates and shows the Launcher form."""
    Application.EnableVisualStyles()
    frm = Launcher()
    frm.ShowDialog()


# =============================================================================
# ENTRY POINT
# =============================================================================

try:
    BUTTON_DIR = os.path.dirname(__file__)
    SHARED_CFG_DIR = os.path.join(BUTTON_DIR, "Configs")

    core.set_paths(shared_cfg_dir=SHARED_CFG_DIR)
    show_launcher()
except Exception as e:
    TaskDialog.Show("Param Binder - Error", str(e))
