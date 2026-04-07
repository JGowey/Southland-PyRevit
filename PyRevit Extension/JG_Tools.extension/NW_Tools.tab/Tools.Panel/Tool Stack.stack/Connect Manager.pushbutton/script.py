# -*- coding: utf-8 -*-
"""
NW Connect Manager
pyRevit Script  |  Revit 2022+  |  IronPython 2.7

Disconnect  -- Severs connector relationships between pipe, duct,
               fittings, and accessories in the selected scope.
Reconnect   -- Restores connector relationships between elements
               that are physically touching but logically disconnected.

Elements are temporarily pinned and unpinned during each transaction
to prevent loss of any elements. Original pin state is always restored.

Window is modeless (stays open while you work in Revit).
doc/uidoc resolved fresh on every Run.
"""

from __future__ import print_function
import datetime
import threading   # threading.Event for cancel

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
)
from Autodesk.Revit.UI import (
    IExternalEventHandler,
    ExternalEvent,
)

import System
from System.Windows.Forms import (
    Form, Label, Button, RadioButton,
    TextBox, ScrollBars, ProgressBar, ProgressBarStyle,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    DialogResult, Application,
    FormStartPosition, FormBorderStyle, BorderStyle,
    FlatStyle, Panel,
)
from System.Drawing import Size, Point, Color, Font, FontStyle

from pyrevit import script
logger = script.get_logger()

_uiapp = __revit__  # noqa: F821


# ===========================================================================
# ELEMENT COLLECTION
# ===========================================================================

_TARGET_CATEGORIES = [
    BuiltInCategory.OST_FabricationPipework,
    BuiltInCategory.OST_FabricationDuctwork,
    BuiltInCategory.OST_FabricationHangers,
    BuiltInCategory.OST_FabricationContainment,
    BuiltInCategory.OST_PipeCurves,
    BuiltInCategory.OST_DuctCurves,
    BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_PipeAccessory,
    BuiltInCategory.OST_DuctAccessory,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures,
    BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_CableTrayFitting,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_Assemblies,
]

_TARGET_CAT_IDS = frozenset(int(bic) for bic in _TARGET_CATEGORIES)

def _is_target_element(el):
    try:
        cat = el.Category
        if cat is None:
            return False
        return cat.Id.IntegerValue in _TARGET_CAT_IDS
    except Exception:
        return False


def collect_elements(doc, uidoc, scope):
    """
    Return a deduplicated list of target MEP/fabrication elements.
    scope: "selected" | "view" | "model"
    Selected scope returns exactly the selected elements -- no assembly expansion.
    """
    seen = set()
    result = []

    def _add(el):
        eid = el.Id.IntegerValue
        if eid in seen:
            return
        seen.add(eid)
        if _is_target_element(el):
            result.append(el)

    if scope == "selected":
        sel = uidoc.Selection.GetElementIds()
        for eid in sel:
            el = doc.GetElement(eid)
            if el is not None:
                _add(el)

    elif scope == "view":
        view = doc.ActiveView
        col = FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType()
        for el in col:
            if not el.IsHidden(view):
                _add(el)

    else:  # model
        col = FilteredElementCollector(doc).WhereElementIsNotElementType()
        for el in col:
            _add(el)

    return result


# ===========================================================================
# CONNECTOR HELPERS
# ===========================================================================

def _get_connectors(el):
    try:
        mgr = None
        if hasattr(el, "ConnectorManager"):
            mgr = el.ConnectorManager
        if mgr is None and hasattr(el, "MEPModel") and el.MEPModel is not None:
            mgr = el.MEPModel.ConnectorManager
        if mgr is None:
            return []
        return list(mgr.Connectors)
    except Exception:
        return []


def _disconnect_one(el):
    """Sever all connections on el. Returns list of error strings."""
    errors = []
    for conn in _get_connectors(el):
        try:
            if conn.IsConnected:
                refs = list(conn.AllRefs)
                for ref in refs:
                    try:
                        conn.DisconnectFrom(ref)
                    except Exception as ex:
                        errors.append(str(ex))
        except Exception as ex:
            errors.append(str(ex))
    return errors


def _build_connector_index(elements_by_id, cell_size=0.02):
    """
    Build a spatial hash of all unconnected connectors keyed by grid cell.
    cell_size in feet -- must be >= 2x THRESHOLD so adjacent connectors
    are always in the same or neighboring cell.

    Returns dict: (ix, iy, iz) -> list of (connector, element_id)
    """
    index = {}
    for eid, el in elements_by_id.items():
        for conn in _get_connectors(el):
            try:
                if conn.IsConnected:
                    continue
                pt = conn.Origin
                key = (int(pt.X / cell_size),
                       int(pt.Y / cell_size),
                       int(pt.Z / cell_size))
                if key not in index:
                    index[key] = []
                index[key].append((conn, eid))
            except Exception:
                pass
    return index


def _reconnect_from_index(el, connector_index, cell_size=0.02):
    """
    For each unconnected connector on el, look up only the connectors
    in neighboring grid cells (3x3x3 neighborhood = 27 cells max).
    Connect the first match within THRESHOLD distance.

    O(1) per connector lookup vs O(n) full scan.
    Returns list of error strings.
    """
    THRESHOLD = 0.01
    el_id     = el.Id.IntegerValue
    errors    = []

    for conn in _get_connectors(el):
        try:
            if conn.IsConnected:
                continue
            pt  = conn.Origin
            ix  = int(pt.X / cell_size)
            iy  = int(pt.Y / cell_size)
            iz  = int(pt.Z / cell_size)

            matched = False
            for dx in (-1, 0, 1):
                if matched: break
                for dy in (-1, 0, 1):
                    if matched: break
                    for dz in (-1, 0, 1):
                        if matched: break
                        neighbors = connector_index.get((ix+dx, iy+dy, iz+dz), [])
                        for other_conn, other_eid in neighbors:
                            if other_eid == el_id:
                                continue
                            try:
                                if other_conn.IsConnected:
                                    continue
                                if pt.DistanceTo(other_conn.Origin) <= THRESHOLD:
                                    conn.ConnectTo(other_conn)
                                    matched = True
                                    break
                            except Exception:
                                pass
        except Exception as ex:
            errors.append(str(ex))
    return errors


# ===========================================================================
# PROGRESS HELPERS
# ===========================================================================
#
# Everything runs on the Revit API thread inside Execute().
# Application.DoEvents() flushes WinForms paint queue mid-loop so the bar
# repaints and Cancel button clicks register while the loop runs.
# ---------------------------------------------------------------------------

_PROGRESS_STEP = 250
_PROGRESS_STEP_RECONNECT = 50  # reconnect is slower per element - pump UI more often


def _el_label(action, el, done, total):
    try:
        cat = el.Category.Name if el.Category else "Element"
    except Exception:
        cat = "Element"
    return "{}  {}/{}  --  {} [{}]".format(action, done + 1, total,
                                            cat, el.Id.IntegerValue)


def _make_progress(form, total):
    """
    Return a progress callback. Accepts an optional pct_offset and pct_scale
    so callers can map each pass to a sub-range of the bar:

      pass 1 (pin):     offset=0,  scale=0.2  → 0-20%
      pass 2 (work):    offset=20, scale=0.6  → 20-80%
      pass 3 (restore): offset=80, scale=0.2  → 80-100%

    Default (offset=0, scale=1.0) maps the full bar as before.
    """
    if form is None or form.IsDisposed:
        return None

    def callback(done, total, label, pct_offset=0, pct_scale=1.0):
        # Callers pre-throttle with their own step (% _PROGRESS_STEP or
        # % _PROGRESS_STEP_RECONNECT) before calling here, so no second
        # check needed -- just update and pump.
        try:
            if form.IsDisposed:
                return
            raw = (done + 1) / float(total) if total > 0 else 1.0
            pct = int(pct_offset + raw * 100 * pct_scale)
            form._progress_bar.Value   = min(max(pct, 0), 100)
            form._progress_lbl.Text    = label
            form._progress_bar.Visible = True
            form._progress_lbl.Visible = True
            Application.DoEvents()
        except Exception:
            pass

    return callback


def _update_progress_phase(form, pct, label):
    """Direct single-update for phase transitions (between passes)."""
    try:
        if form is None or form.IsDisposed:
            return
        form._progress_bar.Value   = min(max(pct, 0), 100)
        form._progress_lbl.Text    = label
        form._progress_bar.Visible = True
        form._progress_lbl.Visible = True
        Application.DoEvents()
    except Exception:
        pass


# ===========================================================================
# CORE ACTIONS
# ===========================================================================

def do_disconnect(doc, elements, form=None, progress=None, cancel_event=None):
    """
    Transaction 1: Snapshot pin states, pin all elements, COMMIT.  (bar 0-20%)
    Transaction 2: Disconnect all connectors, COMMIT.              (bar 20-80%)
    Transaction 3: Restore original pin states, COMMIT.            (bar 80-100%)

    Three separate committed transactions are required. Revit's auto-delete
    logic fires at transaction commit boundaries. Committing the pin pass first
    ensures all elements are locked before Revit evaluates any connector change.
    Returns: counts dict, errors list
    """
    counts = {"processed": 0, "cancelled": 0, "errors": 0}
    errors = []
    total  = len(elements)

    # --- Transaction 1: snapshot + pin all  (bar 0-20%) ---
    pin_states = {}
    _update_progress_phase(form, 0, "Pass 1/3 -- Pinning {:,} elements...".format(total))
    with Transaction(doc, "NW: Disconnect (pin)") as t:
        t.Start()
        try:
            for i, el in enumerate(elements):
                try:
                    pin_states[el.Id.IntegerValue] = el.Pinned
                    if not el.Pinned:
                        el.Pinned = True
                except Exception:
                    pin_states[el.Id.IntegerValue] = True
                if i % _PROGRESS_STEP == 0:
                    # DoEvents here lets cancel button clicks register during pin pass
                    if progress:
                        progress(i, total, "Pass 1/3 -- Pinning  {}/{}".format(i + 1, total),
                                 pct_offset=0, pct_scale=0.2)
                    else:
                        Application.DoEvents()
            t.Commit()
        except Exception as ex:
            t.RollBack()
            errors.append("Pin pass failed: {}".format(ex))
            return counts, errors

    # Cancel check between T1 and T2 -- T1 already committed so pins are set.
    # We must still run T3 to restore them, so jump straight to restore.
    if cancel_event and cancel_event.is_set():
        counts["cancelled"] = total
        _update_progress_phase(form, 80, "Cancelled -- restoring pin states...")
    else:
        # --- Transaction 2: disconnect all  (bar 20-80%) ---
        _update_progress_phase(form, 20, "Pass 2/3 -- Disconnecting...")
        with Transaction(doc, "NW: Disconnect") as t:
            t.Start()
            try:
                for i, el in enumerate(elements):
                    if cancel_event and cancel_event.is_set():
                        counts["cancelled"] = total - i
                        break
                    if progress:
                        progress(i, total, _el_label("DISCONNECT", el, i, total),
                                 pct_offset=20, pct_scale=0.6)
                    try:
                        live = doc.GetElement(el.Id)
                        if live is None:
                            errors.append("[{}] missing after pin commit".format(el.Id.IntegerValue))
                            counts["errors"] += 1
                            continue
                        el_errors = _disconnect_one(live)
                        errors.extend(el_errors)
                        counts["errors"] += len(el_errors)
                        counts["processed"] += 1
                    except Exception as ex:
                        errors.append("[{}] {}".format(el.Id.IntegerValue, ex))
                        counts["errors"] += 1
                t.Commit()
            except Exception as ex:
                t.RollBack()
                errors.append("Disconnect pass failed: {}".format(ex))

    # --- Transaction 3: restore pin states  (bar 80-100%) ---
    # Always runs -- even on cancel, so pins are never left in a modified state.
    _update_progress_phase(form, 80, "Pass 3/3 -- Restoring pin states...")
    with Transaction(doc, "NW: Disconnect (restore pins)") as t:
        t.Start()
        try:
            for i, el in enumerate(elements):
                try:
                    live = doc.GetElement(el.Id)
                    if live is None:
                        continue
                    if not pin_states.get(el.Id.IntegerValue, True):
                        live.Pinned = False
                except Exception:
                    pass
                if i % _PROGRESS_STEP == 0:
                    if progress:
                        progress(i, total, "Pass 3/3 -- Restoring  {}/{}".format(i + 1, total),
                                 pct_offset=80, pct_scale=0.2)
                    else:
                        Application.DoEvents()
            t.Commit()
        except Exception as ex:
            t.RollBack()
            errors.append("Restore pins pass failed: {}".format(ex))

    return counts, errors


def do_reconnect(doc, elements, form=None, progress=None, cancel_event=None):
    """
    Transaction 1: Snapshot pin states, pin all elements, COMMIT.  (bar 0-20%)
    Transaction 2: Reconnect physically touching connectors, COMMIT.(bar 20-80%)
    Transaction 3: Restore original pin states, COMMIT.            (bar 80-100%)

    Same rationale as do_disconnect -- pin commit must fully close before
    any connector work begins.
    Returns: counts dict, errors list
    """
    counts = {"processed": 0, "cancelled": 0, "errors": 0}
    errors = []
    total  = len(elements)

    # --- Transaction 1: snapshot + pin all  (bar 0-20%) ---
    pin_states = {}
    _update_progress_phase(form, 0, "Pass 1/3 -- Pinning {:,} elements...".format(total))
    with Transaction(doc, "NW: Reconnect (pin)") as t:
        t.Start()
        try:
            for i, el in enumerate(elements):
                try:
                    pin_states[el.Id.IntegerValue] = el.Pinned
                    if not el.Pinned:
                        el.Pinned = True
                except Exception:
                    pin_states[el.Id.IntegerValue] = True
                if i % _PROGRESS_STEP == 0:
                    if progress:
                        progress(i, total, "Pass 1/3 -- Pinning  {}/{}".format(i + 1, total),
                                 pct_offset=0, pct_scale=0.2)
                    else:
                        Application.DoEvents()
            t.Commit()
        except Exception as ex:
            t.RollBack()
            errors.append("Pin pass failed: {}".format(ex))
            return counts, errors

    # Cancel check between T1 and T2 -- must still run T3 to restore pins.
    if cancel_event and cancel_event.is_set():
        counts["cancelled"] = total
        _update_progress_phase(form, 80, "Cancelled -- restoring pin states...")
    else:
        # Build live lookup + spatial connector index after pin commit
        elements_by_id = {}
        for el in elements:
            live = doc.GetElement(el.Id)
            if live is not None:
                elements_by_id[el.Id.IntegerValue] = live

        _update_progress_phase(form, 20, "Pass 2/3 -- Building connector index...")
        connector_index = _build_connector_index(elements_by_id)

        # --- Transaction 2: reconnect all  (bar 20-80%) ---
        _update_progress_phase(form, 20, "Pass 2/3 -- Reconnecting...")
        with Transaction(doc, "NW: Reconnect") as t:
            t.Start()
            try:
                for i, el in enumerate(elements):
                    if cancel_event and cancel_event.is_set():
                        counts["cancelled"] = total - i
                        break
                    if i % _PROGRESS_STEP_RECONNECT == 0:
                        if progress:
                            progress(i, total, _el_label("RECONNECT", el, i, total),
                                     pct_offset=20, pct_scale=0.6)
                        else:
                            Application.DoEvents()
                    try:
                        live = elements_by_id.get(el.Id.IntegerValue)
                        if live is None:
                            errors.append("[{}] missing after pin commit".format(el.Id.IntegerValue))
                            counts["errors"] += 1
                            continue
                        el_errors = _reconnect_from_index(live, connector_index)
                        errors.extend(el_errors)
                        counts["errors"] += len(el_errors)
                        counts["processed"] += 1
                    except Exception as ex:
                        errors.append("[{}] {}".format(el.Id.IntegerValue, ex))
                        counts["errors"] += 1
                t.Commit()
            except Exception as ex:
                t.RollBack()
                errors.append("Reconnect pass failed: {}".format(ex))

    # --- Transaction 3: restore pin states  (bar 80-100%) ---
    # Always runs -- even on cancel, pins must be left in their original state.
    _update_progress_phase(form, 80, "Pass 3/3 -- Restoring pin states...")
    with Transaction(doc, "NW: Reconnect (restore pins)") as t:
        t.Start()
        try:
            for i, el in enumerate(elements):
                try:
                    live = doc.GetElement(el.Id)
                    if live is None:
                        continue
                    if not pin_states.get(el.Id.IntegerValue, True):
                        live.Pinned = False
                except Exception:
                    pass
                if i % _PROGRESS_STEP == 0:
                    if progress:
                        progress(i, total, "Pass 3/3 -- Restoring  {}/{}".format(i + 1, total),
                                 pct_offset=80, pct_scale=0.2)
                    else:
                        Application.DoEvents()
            t.Commit()
        except Exception as ex:
            t.RollBack()
            errors.append("Restore pins pass failed: {}".format(ex))

    return counts, errors



# ===========================================================================
# RUN DISPATCHER
# ===========================================================================

def run_action(doc, uidoc, scope, action, form=None, elements=None,
               cancel_event=None):
    if elements is None:
        elements = collect_elements(doc, uidoc, scope)
    if not elements:
        return "No elements found.\r\nScope: {}".format(scope.upper())

    total = len(elements)
    ts    = datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
    SEP   = "-" * 26
    lines = [
        "Scope:    {}".format(scope.upper()),
        "Elements: {:,}".format(total),
        "Time:     {}".format(ts),
        "",
    ]

    prog = _make_progress(form, total)

    if action == "disconnect":
        c, errs = do_disconnect(doc, elements, form=form, progress=prog,
                                cancel_event=cancel_event)
        lines += [
            "DISCONNECT",
            SEP,
            "  Processed  {:,}".format(c["processed"]),
            "  Errors     {}".format(c["errors"]),
        ]
    elif action == "reconnect":
        c, errs = do_reconnect(doc, elements, form=form, progress=prog,
                               cancel_event=cancel_event)
        lines += [
            "RECONNECT",
            SEP,
            "  Processed  {:,}".format(c["processed"]),
            "  Errors     {}".format(c["errors"]),
        ]
    else:
        errs = []
        lines.append("Unknown action: {}".format(action))
        c = {"cancelled": 0}

    if c.get("cancelled"):
        lines.append("  Cancelled  {:,} remaining".format(c["cancelled"]))
    if errs:
        lines.append("")
        lines.append("  -- Errors (first 5) --")
        for err in errs[:5]:
            lines.append("  " + str(err)[:80])

    return "\r\n".join(lines).rstrip()


# ===========================================================================
# HELP TEXT
# ===========================================================================

_HELP_OVERVIEW = (
    "NW Connect Manager\n\n"
    "DISCONNECT\n"
    "  Severs connector relationships between selected pipe,\n"
    "  duct, fittings, and accessories.\n"
    "  Elements are temporarily pinned and unpinned during\n"
    "  the transaction to prevent loss of any elements.\n\n"
    "RECONNECT\n"
    "  Restores connector relationships between pipe, duct,\n"
    "  fittings, and accessories that are physically touching\n"
    "  but logically disconnected.\n"
    "  Elements are temporarily pinned and unpinned during\n"
    "  the transaction to prevent loss of any elements."
)

_HELP_SCOPE = (
    "SCOPE\n\n"
    "Selected Elements\n"
    "  Exactly the elements currently selected in Revit.\n\n"
    "All Elements In Active View\n"
    "  All visible pipe, duct, fittings, and accessories\n"
    "  in the active view. Hidden elements are excluded.\n\n"
    "All Model Elements\n"
    "  Every pipe, duct, fitting, and accessory in the\n"
    "  project, regardless of view or phase."
)


# ===========================================================================
# EXTERNAL EVENT HANDLER
# ===========================================================================

class _RunHandler(IExternalEventHandler):

    def __init__(self):
        self._scope        = "view"
        self._action       = "disconnect"
        self._form         = None
        self._cancel_event = threading.Event()

    def cancel(self):
        self._cancel_event.set()

    def set_request(self, scope, action):
        self._cancel_event.clear()
        self._scope  = scope
        self._action = action

    def Execute(self, uiapp):
        try:
            doc   = uiapp.ActiveUIDocument.Document
            uidoc = uiapp.ActiveUIDocument

            if doc is None:
                result = "No active document."
            elif doc.IsReadOnly:
                result = "Document is read-only."
            else:
                form = self._form
                elements = collect_elements(doc, uidoc, self._scope)
                n = len(elements)
                if form is not None and not form.IsDisposed:
                    try:
                        form._progress_bar.Value   = 0
                        form._progress_lbl.Text    = \
                            "Collected {:,} elements -- starting...".format(n)
                        form._progress_bar.Visible = True
                        form._progress_lbl.Visible = True
                        Application.DoEvents()
                    except Exception:
                        pass

                result = run_action(
                    doc, uidoc, self._scope, self._action,
                    form=form, elements=elements,
                    cancel_event=self._cancel_event
                )

        except Exception as ex:
            result = "Error:\r\n{}".format(ex)

        if self._form is not None and not self._form.IsDisposed:
            self._form.set_result(result)

    def GetName(self):
        return "NW_ConnectManager"


# ===========================================================================
# MODELESS FORM
# ===========================================================================

class ConnectManagerForm(Form):

    W  = 520
    LM = 18
    GW = 484

    C_BG         = Color.FromArgb(245, 245, 245)   # light gray bg
    C_GB         = Color.FromArgb(255, 255, 255)   # white group boxes
    C_TEXT       = Color.FromArgb(30,   30,  30)   # near-black text
    C_SUBTEXT    = Color.FromArgb(110, 110, 110)   # medium gray sublabels
    C_HEAD       = Color.FromArgb(50,   50,  50)   # dark group box titles
    C_WARNING    = Color.FromArgb(180,  80,   0)   # dark orange warning
    C_BTN_DISC   = Color.FromArgb(232, 119,  34)   # NW orange (disconnect)
    C_BTN_RECO   = Color.FromArgb(232, 119,  34)   # NW orange (reconnect)
    C_BTN_CLOSE  = Color.FromArgb(180, 180, 180)   # neutral gray close
    C_BTN_CANCEL = Color.FromArgb(160,  60,  40)   # red-orange cancel
    C_RESULT_BG  = Color.FromArgb(235, 235, 235)   # light result box

    FONT_HEAD  = Font("Segoe UI", 10, FontStyle.Bold)
    FONT_BTN   = Font("Segoe UI",  9, FontStyle.Bold)
    FONT_BODY  = Font("Segoe UI",  9, FontStyle.Regular)
    FONT_SMALL = Font("Segoe UI",  8, FontStyle.Regular)

    def __init__(self, handler, ext_event):
        super(ConnectManagerForm, self).__init__()
        self._handler   = handler
        self._ext_event = ext_event
        self._build_ui()
        self.Text            = "NW  |  Connect Manager"
        self.BackColor       = self.C_BG
        self.FormBorderStyle = FormBorderStyle.FixedToolWindow
        self.StartPosition   = FormStartPosition.CenterScreen
        self.TopMost         = True
        self.ClientSize      = Size(self.W, self._total_h)

    # -------------------------------------------------------------------------
    # Widget helpers
    # -------------------------------------------------------------------------

    def _group_box(self, text, y, h):
        # Label sits above the panel so it never overlaps the border line
        lbl = Label()
        lbl.Text      = text.strip()
        lbl.Font      = self.FONT_SMALL
        lbl.ForeColor = self.C_HEAD
        lbl.BackColor = self.C_BG
        lbl.Location  = Point(self.LM, y)
        lbl.Size      = Size(self.GW, 16)
        self.Controls.Add(lbl)

        pnl = Panel()
        pnl.BackColor   = self.C_GB
        pnl.BorderStyle = BorderStyle.FixedSingle
        pnl.Location    = Point(self.LM, y + 18)
        pnl.Size        = Size(self.GW, h)
        self.Controls.Add(pnl)
        return pnl

    def _radio(self, parent, text, x, y):
        rb = RadioButton()
        rb.Text      = text
        rb.Font      = self.FONT_BODY
        rb.ForeColor = self.C_TEXT
        rb.BackColor = parent.BackColor
        rb.Location  = Point(x, y)
        rb.Size      = Size(self.GW - x - 8, 20)
        parent.Controls.Add(rb)
        return rb

    def _sublabel(self, parent, text, x, y, color=None):
        lbl = Label()
        lbl.Text      = text
        lbl.Font      = self.FONT_SMALL
        lbl.ForeColor = color or self.C_SUBTEXT
        lbl.BackColor = parent.BackColor
        lbl.Location  = Point(x, y)
        lbl.Size      = Size(self.GW - x - 8, 16)
        parent.Controls.Add(lbl)

    def _action_btn(self, text, x, y, w, h, color):
        btn = Button()
        btn.Text      = text
        btn.Font      = self.FONT_BTN
        btn.ForeColor = Color.White
        btn.BackColor = color
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 0
        btn.Location  = Point(x, y)
        btn.Size      = Size(w, h)
        self.Controls.Add(btn)
        return btn

    def _help_btn(self, parent, x, y, help_text):
        btn = Button()
        btn.Text      = "?"
        btn.Font      = self.FONT_SMALL
        btn.ForeColor = Color.FromArgb(30, 30, 30)           # black text
        btn.BackColor = self.C_BTN_CLOSE                     # same dark gray as Close
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize  = 0
        btn.Location  = Point(x, y)
        btn.Size      = Size(22, 18)
        btn.Tag       = help_text
        btn.Click    += self._on_help
        parent.Controls.Add(btn)

    # -------------------------------------------------------------------------
    # Section builders
    # -------------------------------------------------------------------------

    def _build_title_row(self, y):
        lbl = Label()
        lbl.Text      = "Connect  Manager"
        lbl.Font      = Font("Segoe UI", 14, FontStyle.Bold)
        lbl.ForeColor = Color.FromArgb(232, 119, 34)  # NW orange
        lbl.BackColor = self.C_BG
        lbl.Location  = Point(self.LM, y)
        lbl.Size      = Size(self.GW - 30, 28)
        self.Controls.Add(lbl)
        self._help_btn(self, self.W - self.LM - 22, y + 4, _HELP_OVERVIEW)
        return y + 28

    def _build_scope(self, y):
        ROW_H = 40
        PAD_T = 20
        PAD_B = 18
        h = PAD_T + ROW_H * 3 + PAD_B
        gb = self._group_box("Scope", y, h)
        self._help_btn(gb, self.GW - 30, 2, _HELP_SCOPE)

        ry = PAD_T
        self._rb_selected = self._radio(gb, "Selected Elements",              16, ry)
        self._sublabel(gb, "Exactly the elements currently selected",          34, ry + 21)
        ry += ROW_H
        self._rb_view     = self._radio(gb, "All Elements In Active View",    16, ry)
        self._sublabel(gb, "All visible pipe, duct, and fittings in the active view",  34, ry + 21)
        ry += ROW_H
        self._rb_model    = self._radio(gb, "All Model Elements",              16, ry)
        self._sublabel(gb, "Every pipe, duct, and fitting in the project",     34, ry + 21)

        self._rb_selected.Checked = True
        return y + 18 + h  # 18 = label height added by _group_box

    def _build_action_row(self, y):
        BTN_W = (self.GW - 8) // 2
        BTN_H = 52

        self._btn_disconnect = self._action_btn(
            "Disconnect", self.LM, y, BTN_W, BTN_H, self.C_BTN_DISC)
        self._btn_disconnect.Click += self._on_disconnect

        self._btn_reconnect = self._action_btn(
            "Reconnect", self.LM + BTN_W + 8, y, BTN_W, BTN_H, self.C_BTN_RECO)
        self._btn_reconnect.Click += self._on_reconnect

        return y + BTN_H

    def _build_run_row(self, y):
        self._btn_cancel = self._action_btn(
            "Cancel", self.LM, y, 150, 30, self.C_BTN_CANCEL)
        self._btn_cancel.Visible = False
        self._btn_cancel.Click  += self._on_cancel

        btn_close = self._action_btn(
            "Close", self.LM + 158, y, 100, 30, self.C_BTN_CLOSE)
        btn_close.ForeColor = self.C_TEXT
        btn_close.Click    += self._on_close

        return y + 30

    def _build_result(self, y):
        lbl = Label()
        lbl.Text      = "Last run result:"
        lbl.Font      = self.FONT_SMALL
        lbl.ForeColor = self.C_SUBTEXT
        lbl.Location  = Point(self.LM, y)
        lbl.Size      = Size(200, 16)
        self.Controls.Add(lbl)
        y += 18

        self._progress_bar = ProgressBar()
        self._progress_bar.Minimum  = 0
        self._progress_bar.Maximum  = 100
        self._progress_bar.Value    = 0
        self._progress_bar.Style    = ProgressBarStyle.Continuous
        self._progress_bar.Location = Point(self.LM, y)
        self._progress_bar.Size     = Size(self.GW, 14)
        self._progress_bar.Visible  = False
        self.Controls.Add(self._progress_bar)

        self._progress_lbl = Label()
        self._progress_lbl.Text      = ""
        self._progress_lbl.Font      = self.FONT_SMALL
        self._progress_lbl.ForeColor = self.C_SUBTEXT
        self._progress_lbl.BackColor = self.C_BG
        self._progress_lbl.Location  = Point(self.LM, y + 16)
        self._progress_lbl.Size      = Size(self.GW, 15)
        self._progress_lbl.Visible   = False
        self.Controls.Add(self._progress_lbl)
        y += 34

        self._result_box = TextBox()
        self._result_box.Multiline   = True
        self._result_box.ReadOnly    = True
        self._result_box.ScrollBars  = ScrollBars.Vertical
        self._result_box.WordWrap    = False
        self._result_box.BorderStyle = BorderStyle.FixedSingle
        self._result_box.BackColor   = self.C_RESULT_BG
        self._result_box.ForeColor   = self.C_TEXT
        self._result_box.Font        = Font("Consolas", 9, FontStyle.Regular)
        self._result_box.Location    = Point(self.LM, y)
        self._result_box.Size        = Size(self.GW, 180)
        self._result_box.Text        = "No actions run yet."
        self.Controls.Add(self._result_box)

        return y + 180

    # -------------------------------------------------------------------------
    # Master build
    # -------------------------------------------------------------------------

    def _build_ui(self):
        y = 14
        y = self._build_title_row(y);  y += 10
        y = self._build_scope(y);      y += 10
        y = self._build_action_row(y); y += 14
        y = self._build_run_row(y);    y += 12
        y = self._build_result(y);     y += 10
        self._total_h = y

    # -------------------------------------------------------------------------
    # Helpers
    # -------------------------------------------------------------------------

    def _current_scope(self):
        if self._rb_selected.Checked: return "selected"
        if self._rb_view.Checked:     return "view"
        return "model"

    def _start_run(self, action):
        scope = self._current_scope()
        self._result_box.Text = (
            "Running...\r\nScope:  {}\r\nAction: {}".format(
                scope.upper(), action.upper()))
        self._btn_disconnect.Enabled = False
        self._btn_reconnect.Enabled  = False
        self._btn_cancel.Visible     = True
        self._handler.set_request(scope, action)
        self._ext_event.Raise()

    # -------------------------------------------------------------------------
    # Event handlers
    # -------------------------------------------------------------------------

    def _on_help(self, sender, e):
        """Show help text in a topmost Label-based popup -- no scrollbar, no blue highlight."""
        import math
        text  = sender.Tag
        lines = text.split("\n")
        line_h         = 19    # px per visual line at Segoe UI 9
        chars_per_line = 50    # approx chars that fit before wrapping at this width

        # Count visual lines, expanding wrapped lines
        visual_lines = 0
        for ln in lines:
            if len(ln) == 0:
                visual_lines += 1
            else:
                visual_lines += int(math.ceil(len(ln) / float(chars_per_line)))

        dlg_w     = 420
        content_h = visual_lines * line_h + 16
        dlg_h     = content_h + 82  # top pad 12 + OK row + breathing room

        dlg = Form()
        dlg.Text            = "NW  |  Help"
        dlg.BackColor       = self.C_BG
        dlg.FormBorderStyle = FormBorderStyle.FixedToolWindow
        dlg.StartPosition   = FormStartPosition.CenterScreen
        dlg.TopMost         = True
        dlg.ClientSize      = Size(dlg_w, dlg_h)

        lbl = Label()
        lbl.Text      = text
        lbl.Font      = self.FONT_BODY
        lbl.ForeColor = self.C_TEXT
        lbl.BackColor = self.C_BG
        lbl.Location  = Point(14, 12)
        lbl.Size      = Size(dlg_w - 28, content_h)
        lbl.AutoSize  = False
        dlg.Controls.Add(lbl)

        ok = Button()
        ok.Text      = "OK"
        ok.Font      = self.FONT_BTN
        ok.ForeColor = Color.White
        ok.BackColor = self.C_BTN_CLOSE
        ok.FlatStyle = FlatStyle.Flat
        ok.FlatAppearance.BorderSize = 0
        ok.Size      = Size(80, 28)
        ok.Location  = Point((dlg_w - 80) // 2, dlg_h - 44)
        ok.Click    += lambda s, ev: dlg.Close()
        dlg.Controls.Add(ok)
        dlg.AcceptButton = ok

        dlg.ShowDialog(self)

    def _on_disconnect(self, sender, e):
        self._start_run("disconnect")

    def _on_reconnect(self, sender, e):
        self._start_run("reconnect")

    def _on_cancel(self, sender, e):
        self._handler.cancel()
        self._btn_cancel.Text    = "Stopping..."
        self._btn_cancel.Enabled = False

    def _on_close(self, sender, e):
        self.Close()

    def set_result(self, text):
        if self.IsDisposed:
            return
        self._progress_bar.Visible   = False
        self._progress_lbl.Visible   = False
        self._progress_bar.Value     = 0
        self._result_box.Text        = text
        self._btn_cancel.Visible     = False
        self._btn_cancel.Text        = "Cancel"
        self._btn_cancel.Enabled     = True
        self._btn_disconnect.Enabled = True
        self._btn_reconnect.Enabled  = True


# ===========================================================================
# ENTRY POINT
# ===========================================================================

def main():
    handler   = _RunHandler()
    ext_event = ExternalEvent.Create(handler)
    form      = ConnectManagerForm(handler, ext_event)
    handler._form = form
    form.Show()


main()