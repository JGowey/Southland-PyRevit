# -*- coding: utf-8 -*-
"""
ScriptSandbox.py  (modeless — GitHub edition)
==============================
Modeless pyRevit sandbox with GitHub version control.
One repo per tool. Save .py commits to GitHub automatically.

HOW TO INSTALL
--------------
Drop into a pyRevit pushbutton folder as  script.py
  MyTools.extension/MyTools.tab/Dev Tools.panel/Sandbox.pushbutton/

FIRST TIME SETUP (GitHub token)
--------------------------------
1. Go to https://github.com/settings/tokens
2. Click "Generate new token (classic)"
3. Name it anything (e.g. "pyrevit-sandbox")
4. Check scope:  repo  (full control of private repositories)
5. Click "Generate token" — copy it immediately
6. Paste it into the Settings panel in the sandbox UI
7. Enter your GitHub username
8. Click Save Settings — stored locally, never entered again

USAGE
-----
1. Type a Tool Name (top of left panel)
2. Paste or write your script in the editor
3. Click  Save .py  — creates the repo if needed, commits the script
4. History panel shows all commits for this tool
5. Click any commit → Restore to Editor to go back to that version

SHORTCUTS
---------
  Ctrl+Enter  — Run
  Ctrl+S      — Save .py (commit)
  Ctrl+O      — Load from file
"""

import traceback
import threading
import os
import datetime
import json
import base64

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System.Net")

from System.Threading import Thread, ThreadStart, ApartmentState
from pyrevit import revit, DB, script as pyscript, forms

output = pyscript.get_output()

_PLACEHOLDER  = "# Paste your script here or load a file\n# then click Run or Ctrl+Enter\n"
_NOTES_HINT   = "Notes — paste Claude responses, param names, column maps...\n"
_FONT_MIN, _FONT_MAX, _FONT_DEF = 9, 22, 12
_MAX_HIST = 50

# Config file lives next to this script
_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sandbox_config.json")


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------
def load_config():
    try:
        if os.path.exists(_CONFIG_PATH):
            with open(_CONFIG_PATH, "r") as f:
                return json.load(f)
    except Exception:
        pass
    return {"github_token": "", "github_user": ""}

def save_config(cfg):
    try:
        with open(_CONFIG_PATH, "w") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception as ex:
        return str(ex)


# ---------------------------------------------------------------------------
# GitHub API  (pure HTTP — no git binary needed)
# ---------------------------------------------------------------------------
def _gh_request(method, path, token, body=None):
    """Raw GitHub API call. Returns (status_code, dict)."""
    from System.Net import WebClient, WebException
    from System import Uri
    from System.Text import Encoding
    url = "https://api.github.com{}".format(path)
    client = WebClient()
    client.Headers.Add("Authorization", "token {}".format(token))
    client.Headers.Add("User-Agent",    "pyrevit-sandbox")
    client.Headers.Add("Accept",        "application/vnd.github.v3+json")
    client.Headers.Add("Content-Type",  "application/json")
    try:
        if method == "GET":
            resp = client.DownloadString(Uri(url))
            return 200, json.loads(resp)
        elif method in ("PUT", "POST", "PATCH"):
            payload = Encoding.UTF8.GetBytes(json.dumps(body) if body else "{}")
            resp = client.UploadString(Uri(url), method, Encoding.UTF8.GetString(payload))
            return 200, json.loads(resp) if resp else {}
    except WebException as ex:
        try:
            import System.IO
            body_txt = System.IO.StreamReader(ex.Response.GetResponseStream()).ReadToEnd()
            code = int(ex.Response.StatusCode)
            return code, json.loads(body_txt) if body_txt else {}
        except Exception:
            return 0, {"message": str(ex)}
    except Exception as ex:
        return 0, {"message": str(ex)}


def gh_ensure_repo(token, user, repo_name):
    """Create repo if it doesn't exist. Returns (ok, message)."""
    code, data = _gh_request("GET", "/repos/{}/{}".format(user, repo_name), token)
    if code == 200:
        return True, "exists"
    if code == 404:
        code2, data2 = _gh_request("POST", "/user/repos", token, {
            "name": repo_name,
            "description": "pyRevit sandbox — {}".format(repo_name),
            "private": True,
            "auto_init": True
        })
        if code2 == 201:
            return True, "created"
        return False, "Could not create repo: {}".format(data2.get("message", code2))
    return False, "GitHub error {}: {}".format(code, data.get("message", ""))


def gh_commit_file(token, user, repo_name, filename, content_str, message):
    """
    Upsert filename in repo root with content_str.
    Returns (ok, commit_sha_or_error).
    """
    b64 = base64.b64encode(content_str.encode("utf-8")).decode("ascii")
    path = "/repos/{}/{}/contents/{}".format(user, repo_name, filename)

    # Get current SHA if file exists (needed for update)
    sha = None
    code, data = _gh_request("GET", path, token)
    if code == 200:
        sha = data.get("sha")

    body = {"message": message, "content": b64}
    if sha:
        body["sha"] = sha

    code2, data2 = _gh_request("PUT", path, token, body)
    if code2 in (200, 201):
        try:
            commit_sha = data2["commit"]["sha"][:7]
        except Exception:
            commit_sha = "ok"
        return True, commit_sha
    return False, "Commit failed {}: {}".format(code2, data2.get("message", ""))


def gh_list_commits(token, user, repo_name, filename):
    """
    Returns list of dicts {sha, message, date} for filename, newest first.
    """
    path = "/repos/{}/{}/commits?path={}&per_page=30".format(user, repo_name, filename)
    code, data = _gh_request("GET", path, token)
    if code != 200 or not isinstance(data, list):
        return []
    result = []
    for c in data:
        try:
            result.append({
                "sha":     c["sha"][:7],
                "sha_full": c["sha"],
                "message": c["commit"]["message"],
                "date":    c["commit"]["author"]["date"][:16].replace("T", " ")
            })
        except Exception:
            pass
    return result


def gh_get_file_at_commit(token, user, repo_name, filename, sha_full):
    """Fetch file content at a specific commit SHA."""
    path = "/repos/{}/{}/contents/{}?ref={}".format(user, repo_name, filename, sha_full)
    code, data = _gh_request("GET", path, token)
    if code == 200:
        try:
            return True, base64.b64decode(data["content"]).decode("utf-8")
        except Exception as ex:
            return False, str(ex)
    return False, "Could not fetch: {}".format(data.get("message", code))


# ---------------------------------------------------------------------------
# Script execution
# ---------------------------------------------------------------------------
def run_script(code, status_cb):
    import pyrevit as _pyr
    ns = {
        "__name__": "__sandbox__", "__file__": "<sandbox>",
        "revit": revit, "doc": revit.doc, "uidoc": revit.uidoc,
        "DB": DB, "output": output, "forms": forms,
        "script": pyscript, "pyrevit": _pyr,
    }
    try:
        exec(compile(code, "<sandbox>", "exec"), ns)
        status_cb("Done", True, None)
    except SystemExit:
        status_cb("Exited cleanly", True, None)
    except SyntaxError as ex:
        output.print_md("## Error\n```\n{}\n```".format(traceback.format_exc()))
        status_cb("SyntaxError line {} — {}".format(ex.lineno, ex.msg), False, ex.lineno)
    except Exception:
        tb = traceback.format_exc()
        output.print_md("## Error\n```\n{}\n```".format(tb))
        lineno = None
        for line in reversed(tb.splitlines()):
            if "<sandbox>" in line and "line" in line.lower():
                try:
                    lineno = int(line.strip().split("line")[1].strip().split(",")[0])
                except Exception:
                    pass
                break
        msg = "Error line {} — see output".format(lineno) if lineno else "Error — see output"
        status_cb(msg, False, lineno)


# ---------------------------------------------------------------------------
# Line numbers editor
# ---------------------------------------------------------------------------
def make_editor_with_linenos(bg_editor, bg_lineno, fg_green, fg_gray, fg_white, font_size):
    from System.Windows import GridLength, GridUnitType, Thickness, TextWrapping
    from System.Windows.Controls import (Grid, TextBox, ScrollBarVisibility, ColumnDefinition)
    from System.Windows.Media import FontFamily
    from System.Windows.Threading import DispatcherTimer, DispatcherPriority
    from System import TimeSpan

    container = Grid()
    c0 = ColumnDefinition(); c0.Width = GridLength(42)
    c1 = ColumnDefinition(); c1.Width = GridLength(1, GridUnitType.Star)
    container.ColumnDefinitions.Add(c0)
    container.ColumnDefinitions.Add(c1)

    txt_lines = TextBox()
    txt_lines.IsReadOnly          = True
    txt_lines.FontFamily          = FontFamily("Consolas")
    txt_lines.FontSize            = font_size
    txt_lines.Background          = bg_lineno
    txt_lines.Foreground          = fg_gray
    txt_lines.BorderThickness     = Thickness(0)
    txt_lines.Padding             = Thickness(4, 6, 4, 6)
    txt_lines.TextWrapping        = TextWrapping.NoWrap
    txt_lines.VerticalScrollBarVisibility   = ScrollBarVisibility.Hidden
    txt_lines.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden
    txt_lines.Text = "1"
    Grid.SetColumn(txt_lines, 0)
    container.Children.Add(txt_lines)

    txt_code = TextBox()
    txt_code.AcceptsReturn    = True
    txt_code.AcceptsTab       = True
    txt_code.FontFamily       = FontFamily("Consolas")
    txt_code.FontSize         = font_size
    txt_code.Background       = bg_editor
    txt_code.Foreground       = fg_green
    txt_code.CaretBrush       = fg_white
    txt_code.TextWrapping     = TextWrapping.NoWrap
    txt_code.BorderThickness  = Thickness(0)
    txt_code.Padding          = Thickness(6, 6, 6, 6)
    txt_code.Text             = _PLACEHOLDER
    txt_code.VerticalScrollBarVisibility   = ScrollBarVisibility.Auto
    txt_code.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
    Grid.SetColumn(txt_code, 1)
    container.Children.Add(txt_code)

    def _update_lines(s, e):
        n = txt_code.Text.count("\n") + 1
        txt_lines.Text = "\n".join(str(i) for i in range(1, n + 1))
    txt_code.TextChanged += _update_lines

    _last = [0.0]
    def _tick(s, e):
        try:
            off = txt_code.VerticalOffset
            if abs(off - _last[0]) > 0.1:
                _last[0] = off
                txt_lines.ScrollToVerticalOffset(off)
        except Exception:
            pass

    timer = DispatcherTimer(DispatcherPriority.Background)
    timer.Interval = TimeSpan.FromMilliseconds(50)
    timer.Tick    += _tick
    timer.Start()

    def _unloaded(s, e):
        try: timer.Stop()
        except Exception: pass
    container.Unloaded += _unloaded

    return container, txt_code, txt_lines


# ---------------------------------------------------------------------------
# Main WPF launch — all on STA thread
# ---------------------------------------------------------------------------
def _launch():
    from System.Windows import (Window, Thickness, HorizontalAlignment,
                                 VerticalAlignment, TextWrapping,
                                 GridLength, GridUnitType, FontWeights)
    from System.Windows.Controls import (
        Grid, GridSplitter, StackPanel, Label, TextBox, Button,
        ScrollBarVisibility, Orientation, RowDefinition, ColumnDefinition,
        ListBox, ListBoxItem, TabControl, TabItem
    )
    from System.Windows.Media import Brushes, FontFamily, SolidColorBrush, Color
    from System.Windows.Input import Key, ModifierKeys
    from System.Windows.Threading import Dispatcher

    def mkbrush(r, g, b):
        br = SolidColorBrush(Color.FromRgb(r, g, b))
        br.Freeze()
        return br

    BG_DARK   = mkbrush(18,  18,  18)
    BG_PANEL  = mkbrush(28,  28,  36)
    BG_INPUT  = mkbrush(22,  22,  30)
    BG_EDITOR = mkbrush(12,  12,  12)
    BG_LINENO = mkbrush(22,  22,  28)
    BG_HIST   = mkbrush(18,  18,  26)
    BG_TOOL   = mkbrush(24,  24,  32)
    BG_SETT   = mkbrush(20,  20,  28)
    FG_GREEN  = mkbrush(80,  220, 120)
    FG_BLUE   = mkbrush(100, 180, 255)
    FG_YELLOW = mkbrush(255, 210,  70)
    FG_GRAY   = mkbrush(100, 100, 120)
    FG_WHITE  = Brushes.White
    FG_RED    = mkbrush(255,  80,  80)
    FG_ORANGE = mkbrush(255, 160,  40)
    BTN_RUN   = mkbrush(34,  160,  74)
    BTN_SAVE  = mkbrush(30,  120, 200)
    BTN_LOAD  = mkbrush(80,   80,  95)
    BTN_GRAY  = mkbrush(55,   55,  68)
    BTN_PIN   = mkbrush(90,   50, 170)
    BTN_PINOFF= mkbrush(55,   55,  68)
    BTN_GH    = mkbrush(30,   30,  30)

    cfg   = load_config()
    state = {"font_size": _FONT_DEF, "pinned": True}

    # refs shared across closures
    code_ref   = [None]
    ln_ref     = [None]
    status_ref = [None]
    tool_ref   = [None]   # tool name TextBox

    # ── helpers ──────────────────────────────────────────────────────────
    def mklbl(text, size, fg, bold=False):
        lbl = Label()
        lbl.Content    = text
        lbl.FontSize   = size
        lbl.Foreground = fg
        lbl.Padding    = Thickness(4, 2, 4, 2)
        if bold: lbl.FontWeight = FontWeights.Bold
        lbl.VerticalContentAlignment = VerticalAlignment.Center
        return lbl

    def mkbtn(text, handler, bg, fg, w, h=28):
        b = Button()
        b.Content         = text
        b.Width           = w
        b.Height          = h
        b.FontSize        = 11
        b.Background      = bg
        b.Foreground      = fg
        b.BorderThickness = Thickness(0)
        b.Margin          = Thickness(0, 0, 4, 0)
        b.Click          += handler
        return b

    def mktxt(height=24, multiline=False, fg=None, bg=None, fontsize=11):
        t = TextBox()
        t.FontFamily      = FontFamily("Consolas")
        t.FontSize        = fontsize
        t.Height          = height
        t.Foreground      = fg or FG_WHITE
        t.Background      = bg or BG_INPUT
        t.CaretBrush      = FG_WHITE
        t.BorderThickness = Thickness(1)
        t.Padding         = Thickness(4, 2, 4, 2)
        if multiline:
            t.AcceptsReturn = True
            t.AcceptsTab    = True
            t.Height        = height
            t.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        return t

    def set_status(msg, ok=True, lineno=None):
        if lineno: msg = "Line {} — {}".format(lineno, msg)
        if status_ref[0]:
            status_ref[0].Content    = msg
            status_ref[0].Foreground = FG_GREEN if ok else FG_RED

    def get_tool_name():
        t = tool_ref[0]
        if t is None: return "untitled"
        raw = t.Text.strip()
        if not raw or raw == "Enter tool name...": return "untitled"
        # sanitize for repo name
        import re
        return re.sub(r"[^a-zA-Z0-9_\-]", "-", raw).lower().strip("-") or "untitled"

    def get_filename():
        return "{}.py".format(get_tool_name())

    # ── GitHub operations ────────────────────────────────────────────────
    def gh_save(s, e):
        code = code_ref[0].Text if code_ref[0] else ""
        if not code or code.strip() == _PLACEHOLDER.strip():
            set_status("Nothing to save.", False); return

        token = cfg.get("github_token", "").strip()
        user  = cfg.get("github_user",  "").strip()
        if not token or not user:
            set_status("Set GitHub token & username in Settings tab first.", False)
            return

        tool   = get_tool_name()
        fname  = get_filename()
        ts     = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        msg    = "save {} {}".format(tool, ts)

        set_status("Committing to GitHub...", True)
        btn_gh_save.IsEnabled = False

        def _work():
            ok, result = gh_ensure_repo(token, user, tool)
            if not ok:
                win.Dispatcher.Invoke(lambda: _gh_done(False, result))
                return
            ok2, result2 = gh_commit_file(token, user, tool, fname, code, msg)
            win.Dispatcher.Invoke(lambda: _gh_done(ok2, result2, tool, token, user, fname))

        t = threading.Thread(target=_work); t.IsBackground = True; t.Start()

    def _gh_done(ok, result, tool=None, token=None, user=None, fname=None):
        btn_gh_save.IsEnabled = True
        if ok:
            set_status("Committed {} — sha {}".format(fname or "", result), True)
            if tool and token and user and fname:
                _refresh_history(token, user, tool, fname)
        else:
            set_status("GitHub error: {}".format(result), False)

    def _refresh_history(token, user, tool, fname):
        def _work():
            commits = gh_list_commits(token, user, tool, fname)
            win.Dispatcher.Invoke(lambda: _populate_history(commits, token, user, tool, fname))
        t = threading.Thread(target=_work); t.IsBackground = True; t.Start()

    def _populate_history(commits, token, user, tool, fname):
        lst_hist.Items.Clear()
        if not commits:
            item = ListBoxItem()
            item.Content    = "No commits yet"
            item.Foreground = FG_GRAY
            item.Background = BG_HIST
            lst_hist.Items.Add(item)
            return
        for c in commits:
            item = ListBoxItem()
            item.Content    = "{}  {}  {}".format(c["sha"], c["date"], c["message"][:30])
            item.Foreground = FG_GREEN
            item.Background = BG_HIST
            item.FontFamily = FontFamily("Consolas")
            item.FontSize   = 10
            item.Tag        = (c["sha_full"], token, user, tool, fname)
            lst_hist.Items.Add(item)

    def on_load_commit(s, e):
        item = lst_hist.SelectedItem
        if item is None:
            set_status("Select a commit first.", False); return
        try:
            sha_full, token, user, tool, fname = item.Tag
        except Exception:
            set_status("Can't restore this entry.", False); return

        set_status("Fetching commit...", True)
        def _work():
            ok, content = gh_get_file_at_commit(token, user, tool, fname, sha_full)
            win.Dispatcher.Invoke(lambda: _restore_commit(ok, content))
        t = threading.Thread(target=_work); t.IsBackground = True; t.Start()

    def _restore_commit(ok, content):
        if ok:
            code_ref[0].Text       = content
            code_ref[0].Foreground = FG_GREEN
            set_status("Restored from commit.", True)
        else:
            set_status("Restore failed: {}".format(content), False)

    def on_refresh_hist(s, e):
        token = cfg.get("github_token", "").strip()
        user  = cfg.get("github_user",  "").strip()
        if not token or not user:
            set_status("Set GitHub token & username in Settings first.", False); return
        tool  = get_tool_name()
        fname = get_filename()
        set_status("Loading history...", True)
        _refresh_history(token, user, tool, fname)

    # ── other handlers ───────────────────────────────────────────────────
    def on_run(s, e):
        txt = code_ref[0]
        if txt is None: return
        code = txt.Text.strip()
        if not code or code == _PLACEHOLDER.strip():
            set_status("Nothing to run.", False); return

        btn_run.IsEnabled    = False
        btn_gh_save.IsEnabled = False
        set_status("Running...  (UI will be unresponsive until script finishes)", True)
        # Force UI repaint before blocking
        from System.Windows.Threading import Dispatcher, DispatcherPriority
        Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Render, lambda: None)

        output.print_md("---\n## Sandbox Run  [{}]".format(get_tool_name()))

        # Run on same thread (Revit API requirement) but keep status visible
        run_script(code, lambda msg, ok, ln: _done_run(msg, ok, ln))

    def _done_run(msg, ok, ln):
        btn_run.IsEnabled     = True
        btn_gh_save.IsEnabled = True
        set_status(msg, ok, ln)
        # Force repaint so status updates are visible immediately
        from System.Windows.Threading import Dispatcher, DispatcherPriority
        try:
            Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Render, lambda: None)
        except Exception:
            pass

    def on_load_file(s, e):
        path = forms.pick_file(file_ext="py", title="Load Script")
        if path:
            try:
                from System.IO import File
                from System.Text import Encoding as E
                code_ref[0].Text       = File.ReadAllText(path, E.UTF8)
                code_ref[0].Foreground = FG_GREEN
                set_status("Loaded: {}".format(os.path.basename(path)))
            except Exception as ex:
                set_status("Load failed: {}".format(ex), False)

    def on_clear(s, e):
        code_ref[0].Text       = _PLACEHOLDER
        code_ref[0].Foreground = FG_GREEN
        set_status("Cleared.")

    def on_font_up(s, e):
        if state["font_size"] < _FONT_MAX:
            state["font_size"] += 1
            code_ref[0].FontSize = state["font_size"]
            ln_ref[0].FontSize   = state["font_size"]
            set_status("Font {}pt".format(state["font_size"]))

    def on_font_dn(s, e):
        if state["font_size"] > _FONT_MIN:
            state["font_size"] -= 1
            code_ref[0].FontSize = state["font_size"]
            ln_ref[0].FontSize   = state["font_size"]
            set_status("Font {}pt".format(state["font_size"]))

    def on_pin(s, e):
        state["pinned"] = not state["pinned"]
        win.Topmost        = state["pinned"]
        btn_pin.Content    = "📌 On Top"  if state["pinned"] else "📌 Off"
        btn_pin.Background = BTN_PIN      if state["pinned"] else BTN_PINOFF

    def on_save_settings(s, e):
        cfg["github_token"] = txt_token.Password.strip()
        cfg["github_user"]  = txt_user.Text.strip()
        result = save_config(cfg)
        if result is True:
            set_status("Settings saved.", True)
        else:
            set_status("Save failed: {}".format(result), False)

    def on_key(s, e):
        ctrl = e.KeyboardDevice.Modifiers == ModifierKeys.Control
        if   ctrl and e.Key == Key.Return: on_run(None,None);       e.Handled = True
        elif ctrl and e.Key == Key.S:      gh_save(None,None);      e.Handled = True
        elif ctrl and e.Key == Key.O:      on_load_file(None,None); e.Handled = True

    # ── window ───────────────────────────────────────────────────────────
    win = Window()
    win.Title      = "Script Sandbox"
    win.Width      = 650
    win.Height     = 950
    win.MinWidth   = 500
    win.MinHeight  = 600
    win.Topmost    = True
    win.Background = BG_DARK
    win.KeyDown   += on_key

    root = Grid()
    root.Margin = Thickness(0)
    for w in [GridLength(185),
              GridLength(10),
              GridLength(1, GridUnitType.Star)]:
        cd = ColumnDefinition(); cd.Width = w
        root.ColumnDefinitions.Add(cd)

    # ════════════════════════════════════════════════════════════════════
    # LEFT PANEL  — tabs: Notes | Settings
    # ════════════════════════════════════════════════════════════════════
    tabs = TabControl()
    tabs.Background   = BG_PANEL
    tabs.BorderThickness = Thickness(0)

    # ── TAB: Notes + History ─────────────────────────────────────────────
    tab_notes = TabItem()
    tab_notes.Header     = "Notes / History"
    tab_notes.Foreground = FG_WHITE
    tab_notes.Background = BG_PANEL

    notes_grid = Grid()
    notes_grid.Background = BG_PANEL
    for h in [GridLength(36),                        # Tool name row
              GridLength(26),                        # Notes header
              GridLength(2, GridUnitType.Star),      # Notes body
              GridLength(26),                        # History header + refresh
              GridLength(3, GridUnitType.Star),      # History list
              GridLength(66)]:                       # History buttons
        rd = RowDefinition(); rd.Height = h
        notes_grid.RowDefinitions.Add(rd)

    # Tool name input
    tool_row = StackPanel()
    tool_row.Orientation    = Orientation.Vertical
    tool_row.Background     = BG_TOOL
    tool_row.Margin         = Thickness(4, 4, 4, 2)

    txt_toolname = mktxt(height=26, fg=FG_YELLOW, bg=BG_INPUT, fontsize=12)
    txt_toolname.Text = "Enter tool name..."
    txt_toolname.Foreground = FG_GRAY
    def _tool_focus(s, e):
        if s.Foreground == FG_GRAY:
            s.Text = ""; s.Foreground = FG_YELLOW
    txt_toolname.GotFocus += _tool_focus
    tool_ref[0] = txt_toolname
    tool_row.Children.Add(txt_toolname)
    Grid.SetRow(tool_row, 0); notes_grid.Children.Add(tool_row)

    # Notes header
    nhdr = mklbl("NOTES", 10, FG_YELLOW, bold=True)
    nhdr.Background = BG_TOOL
    Grid.SetRow(nhdr, 1); notes_grid.Children.Add(nhdr)

    # Notes body
    txt_notes = TextBox()
    txt_notes.AcceptsReturn   = True; txt_notes.AcceptsTab = True
    txt_notes.TextWrapping    = TextWrapping.Wrap
    txt_notes.FontFamily      = FontFamily("Consolas")
    txt_notes.FontSize        = 10
    txt_notes.Background      = BG_INPUT
    txt_notes.Foreground      = FG_GRAY
    txt_notes.CaretBrush      = FG_WHITE
    txt_notes.BorderThickness = Thickness(0)
    txt_notes.Padding         = Thickness(6)
    txt_notes.Text            = _NOTES_HINT
    txt_notes.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
    def _nf(s, e):
        if s.Foreground == FG_GRAY: s.Text = ""; s.Foreground = FG_YELLOW
    txt_notes.GotFocus += _nf
    Grid.SetRow(txt_notes, 2); notes_grid.Children.Add(txt_notes)

    # History header row with refresh button
    hist_hdr_row = StackPanel()
    hist_hdr_row.Orientation = Orientation.Horizontal
    hist_hdr_row.Background  = BG_TOOL
    hhdr = mklbl("HISTORY", 10, FG_BLUE, bold=True)
    btn_refresh = mkbtn("↻", on_refresh_hist, BTN_GRAY, FG_WHITE, 26, 24)
    hist_hdr_row.Children.Add(hhdr)
    hist_hdr_row.Children.Add(btn_refresh)
    Grid.SetRow(hist_hdr_row, 3); notes_grid.Children.Add(hist_hdr_row)

    # History list
    lst_hist = ListBox()
    lst_hist.Background      = BG_HIST
    lst_hist.BorderThickness = Thickness(0)
    lst_hist.FontFamily      = FontFamily("Consolas")
    lst_hist.FontSize        = 10
    Grid.SetRow(lst_hist, 4); notes_grid.Children.Add(lst_hist)

    # History buttons
    hbtn = StackPanel()
    hbtn.Orientation = Orientation.Vertical
    hbtn.Background  = BG_TOOL
    hbtn.Margin      = Thickness(4, 4, 4, 4)
    btn_restore  = mkbtn("Restore to Editor", on_load_commit,  BTN_SAVE, FG_WHITE, 168, 28)
    btn_clr_hist = mkbtn("Clear List",        lambda s,e: lst_hist.Items.Clear(),
                         BTN_GRAY, FG_WHITE, 168, 28)
    hbtn.Children.Add(btn_restore)
    hbtn.Children.Add(btn_clr_hist)
    Grid.SetRow(hbtn, 5); notes_grid.Children.Add(hbtn)

    tab_notes.Content = notes_grid
    tabs.Items.Add(tab_notes)
    tabs.SelectedIndex = 0

    # ── TAB: Settings ────────────────────────────────────────────────────
    tab_sett = TabItem()
    tab_sett.Header     = "⚙ GitHub"
    tab_sett.Foreground = FG_WHITE
    tab_sett.Background = BG_PANEL

    sett_grid = Grid()
    sett_grid.Background = BG_SETT
    sett_grid.Margin     = Thickness(8)
    for h in [GridLength(24), GridLength(30),   # token label + input
              GridLength(12),                    # spacer
              GridLength(24), GridLength(30),   # user label + input
              GridLength(16),                    # spacer
              GridLength(32),                    # save btn
              GridLength(1, GridUnitType.Star),  # instructions
              ]:
        rd = RowDefinition(); rd.Height = h
        sett_grid.RowDefinitions.Add(rd)

    Grid.SetRow(mklbl("GitHub Personal Access Token:", 10, FG_GRAY), 0)
    sett_grid.Children.Add(mklbl("GitHub Personal Access Token:", 10, FG_GRAY))

    from System.Windows.Controls import PasswordBox
    txt_token = PasswordBox()
    txt_token.FontFamily      = FontFamily("Consolas")
    txt_token.FontSize        = 11
    txt_token.Height          = 28
    txt_token.Background      = BG_INPUT
    txt_token.Foreground      = FG_WHITE
    txt_token.BorderThickness = Thickness(1)
    txt_token.Padding         = Thickness(4, 2, 4, 2)
    txt_token.Password        = cfg.get("github_token", "")
    Grid.SetRow(txt_token, 1); sett_grid.Children.Add(txt_token)

    Grid.SetRow(mklbl("GitHub Username:", 10, FG_GRAY), 3)
    sett_grid.Children.Add(mklbl("GitHub Username:", 10, FG_GRAY))

    txt_user = mktxt(height=28, fg=FG_WHITE, bg=BG_INPUT)
    txt_user.Text = cfg.get("github_user", "")
    Grid.SetRow(txt_user, 4); sett_grid.Children.Add(txt_user)

    btn_sett_save = mkbtn("Save Settings", on_save_settings, BTN_RUN, FG_WHITE, 160, 30)
    Grid.SetRow(btn_sett_save, 6); sett_grid.Children.Add(btn_sett_save)

    instructions = TextBox()
    instructions.IsReadOnly      = True
    instructions.TextWrapping    = TextWrapping.Wrap
    instructions.Background      = BG_SETT
    instructions.Foreground      = FG_GRAY
    instructions.BorderThickness = Thickness(0)
    instructions.FontSize        = 10
    instructions.Padding         = Thickness(0, 8, 0, 0)
    instructions.Text = (
        "HOW TO GET A TOKEN\n"
        "──────────────────\n"
        "1. Go to github.com/settings/tokens\n"
        "2. Click 'Generate new token (classic)'\n"
        "3. Name it: pyrevit-sandbox\n"
        "4. Check scope: repo\n"
        "5. Click Generate token\n"
        "6. Copy it — paste above\n"
        "7. Enter your GitHub username\n"
        "8. Click Save Settings\n\n"
        "Each tool name becomes its own\n"
        "private repo on your GitHub.\n"
        "Save .py commits the script with\n"
        "a timestamp message."
    )
    Grid.SetRow(instructions, 7); sett_grid.Children.Add(instructions)

    tab_sett.Content = sett_grid
    tabs.Items.Add(tab_sett)

    Grid.SetColumn(tabs, 0); root.Children.Add(tabs)

    # ── SPLITTER ─────────────────────────────────────────────────────────
    spl = GridSplitter()
    spl.Width               = 10
    spl.HorizontalAlignment = HorizontalAlignment.Stretch
    spl.Background          = BG_PANEL
    spl.ShowsPreview        = True
    Grid.SetColumn(spl, 1); root.Children.Add(spl)

    # ════════════════════════════════════════════════════════════════════
    # RIGHT PANEL  — toolbar + editor + status
    # ════════════════════════════════════════════════════════════════════
    right = Grid()
    right.Background = BG_DARK
    for h in [GridLength(36),
              GridLength(30),
              GridLength(1, GridUnitType.Star),
              GridLength(24)]:
        rd = RowDefinition(); rd.Height = h
        right.RowDefinitions.Add(rd)

    # Toolbar row 1 — run / save to GH / load / clear
    tb1 = StackPanel()
    tb1.Orientation       = Orientation.Horizontal
    tb1.VerticalAlignment = VerticalAlignment.Center
    tb1.Background        = BG_TOOL
    tb1.Margin            = Thickness(4, 3, 4, 0)

    btn_run     = mkbtn("▶  Run",      on_run,       BTN_RUN,  FG_WHITE, 80)
    btn_gh_save = mkbtn("⬆ Save .py",  gh_save,      BTN_SAVE, FG_WHITE, 90)
    btn_load    = mkbtn("📂 Load .py", on_load_file, BTN_LOAD, FG_WHITE, 90)
    btn_clr     = mkbtn("Clear",       on_clear,     BTN_GRAY, FG_WHITE, 55)
    for b in [btn_run, btn_gh_save, btn_load, btn_clr]:
        tb1.Children.Add(b)
    Grid.SetRow(tb1, 0); right.Children.Add(tb1)

    # Toolbar row 2 — pin + font
    tb2 = StackPanel()
    tb2.Orientation       = Orientation.Horizontal
    tb2.VerticalAlignment = VerticalAlignment.Center
    tb2.Background        = BG_TOOL
    tb2.Margin            = Thickness(4, 0, 4, 3)

    btn_pin = mkbtn("📌 On Top", on_pin,     BTN_PIN,  FG_WHITE, 90, 24)
    btn_fdn = mkbtn("A−",        on_font_dn, BTN_GRAY, FG_WHITE, 34, 24)
    btn_fup = mkbtn("A+",        on_font_up, BTN_GRAY, FG_WHITE, 34, 24)
    sz_lbl  = mklbl("Size", 10, FG_GRAY)
    for w in [btn_pin, sz_lbl, btn_fdn, btn_fup]:
        tb2.Children.Add(w)
    Grid.SetRow(tb2, 1); right.Children.Add(tb2)

    # Editor with line numbers
    editor_grid, txt_code, txt_lines = make_editor_with_linenos(
        BG_EDITOR, BG_LINENO, FG_GREEN, FG_GRAY, FG_WHITE, _FONT_DEF)
    code_ref[0] = txt_code
    ln_ref[0]   = txt_lines
    Grid.SetRow(editor_grid, 2); right.Children.Add(editor_grid)

    # Status bar
    lbl_status = mklbl("Ready  |  Ctrl+Enter run  |  Ctrl+S commit  |  Ctrl+O load", 9, FG_GRAY)
    lbl_status.Background = BG_TOOL
    status_ref[0] = lbl_status
    Grid.SetRow(lbl_status, 3); right.Children.Add(lbl_status)

    Grid.SetColumn(right, 2); root.Children.Add(right)

    win.Content = root
    win.Show()
    Dispatcher.Run()


# ---------------------------------------------------------------------------
# STA thread
# ---------------------------------------------------------------------------
t = Thread(ThreadStart(_launch))
t.SetApartmentState(ApartmentState.STA)
t.IsBackground = True
t.Start()