# -*- coding: utf-8 -*-
""" 
pyRevit - Modeless "Micron Parameter Writer" (ExternalEvent)

Logic (UNCHANGED):
  - Micron Zone:
      * Override wins: if em_CP_Override Parameter Sync == YES and Scope_Const_Work_Pkg_Ext_MTI has value -> return it
      * Else:
          1) Center point in zone (TRUE XYZ)
          2) Else element bbox intersects zone (TRUE XYZ)
          If multiple zones match: smallest zone volume wins

UI (UPDATED ONLY):
  - Refreshed WPF styling/layout to match the newer "Weight and Volume Calculator" look
  - No changes to calculations, defaults, or parameter-writing behavior
"""

from __future__ import division

import math
import clr
import time

from pyrevit import script

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementMulticategoryFilter,
    TransactionGroup,
    Transaction,
    ElementId,
    SynchronizeWithCentralOptions,
    RelinquishOptions,
    TransactWithCentralOptions,
    WorksharingUtils,
    BuiltInParameter,
    DirectShape,
)

from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent
from System.Collections.Generic import List

# WPF
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows.Markup import XamlReader

# Message pump for modeless responsiveness
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import Application

# Fabrication
try:
    from Autodesk.Revit.DB import FabricationPart
except:
    FabricationPart = None

logger = script.get_logger()

# -----------------------------------------------------------------------------
# CONFIG
# -----------------------------------------------------------------------------
WINDOW_TITLE = "Micron Parameter Writer"

CHUNK_SIZE = 10000
UI_UPDATE_INTERVAL = 0.25

# Job Number default (hardcoded)
JOB_NUMBER_VALUE = "B111"

# Micron Zone constants
MSG_NOZONE = "Not in Zone"
GEN_TAG = "#SPACE_SOLID"
GRID_SIZE_FT = 40.0  # tune 30–60 as needed

# Category filter list (UNCHANGED)
CATEGORY_NAMES = set([
    "Air Terminals","Cable Tray Fittings","Cable Trays","Columns","Communication Devices",
    "Conduit Fittings","Conduits","Data Devices","Doors","Duct Accessories","Duct Fittings","Ducts",
    "Electrical Equipment","Electrical Fixtures","Fire Alarm Devices","Fire Protection","Flex Ducts","Flex Pipes",
    "Food Service Equipment","Generic Models","Lighting Devices","Lighting Fixtures","Mechanical Control Devices",
    "Mechanical Equipment","Medical Equipment","MEP Fabrication Ductwork","MEP Fabrication Hangers","MEP Fabrication Pipework",
    "Parts","Pipe Accessories","Pipe Fittings","Pipes","Plumbing Equipment","Plumbing Fixtures","Railings","Ramps",
    "Security Devices","Specialty Equipment","Sprinklers","Structural Columns","Structural Connections",
    "Structural Foundations","Structural Framing","Structural Rebar","Structural Stiffeners","Structural Trusses",
    "Telephone Devices","Walls","Windows","Wires",
])

# -----------------------------------------------------------------------------
# Shared run-state
# -----------------------------------------------------------------------------
class RunState(object):
    def __init__(self):
        self.cancel_requested = False
        self.is_running = False

RUN_STATE = RunState()

def do_events():
    try:
        Application.DoEvents()
    except:
        pass

# -----------------------------------------------------------------------------
# General Helpers
# -----------------------------------------------------------------------------

def is_fab_part(el):
    return (FabricationPart is not None) and isinstance(el, FabricationPart)


def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None


def get_param_string(el, name):
    p = get_param(el, name)
    if not p or not p.HasValue:
        return None
    try:
        s = p.AsString()
        if not s:
            s = p.AsValueString()
        return s
    except:
        return None


def set_param_int(el, name, val):
    p = get_param(el, name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(int(val))
        return True
    except:
        return False


def set_param_string(el, name, val):
    p = get_param(el, name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(val if val is not None else "")
        return True
    except:
        return False


def get_category_ids_by_name(doc_, names_set):
    ids = []
    try:
        cats = doc_.Settings.Categories
        for c in cats:
            try:
                if c and c.Name in names_set:
                    ids.append(c.Id)
            except:
                pass
    except:
        pass
    return ids


def build_category_filter(doc_):
    cat_ids_py = get_category_ids_by_name(doc_, CATEGORY_NAMES)
    cat_ids = List[ElementId]()
    for cid in cat_ids_py:
        cat_ids.Add(cid)
    if cat_ids.Count > 0:
        return ElementMulticategoryFilter(cat_ids)
    return None

# -----------------------------------------------------------------------------
# Sync helper
# -----------------------------------------------------------------------------

def sync_with_central(doc_):
    try:
        if not doc_.IsWorkshared:
            return "Skipped (not workshared)"
    except:
        return "Skipped (unknown workshare state)"

    try:
        ro = RelinquishOptions(True)
        ro.CheckedOutElements = True
        ro.UserWorksets = True

        opts = SynchronizeWithCentralOptions()
        opts.SetRelinquishOptions(ro)
        opts.SaveLocalAfter = True

        doc_.SynchronizeWithCentral(
            TransactWithCentralOptions(),
            opts
        )
        return "SUCCESS"
    except Exception as ex:
        return "FAILED: {}".format(ex)

# -----------------------------------------------------------------------------
# Collection (yields UI while enumerating)
# -----------------------------------------------------------------------------

def collect_elements(doc_, uidoc_, scope_str, cat_filter, only_owned=False, username=None, progress_cb=None, pump_cb=None):

    def is_owned_by_me(e):
        if not only_owned:
            return True
        # If not workshared, ownership isn't meaningful; allow all.
        try:
            if not doc_.IsWorkshared:
                return True
        except:
            return True

        try:
            info = WorksharingUtils.GetWorksharingTooltipInfo(doc_, e.Id)
            owner = (info.Owner or "").strip()
            if not owner:
                return False  # strict owned-only: skip unowned
            return (username is not None) and (owner == username)
        except:
            return True

    def keep(e):
        try:
            if not is_owned_by_me(e):
                return False
            # Always allow Fabrication Parts (they may not have standard categories)
            if is_fab_part(e):
                return True
            if e.Category and e.Category.Name in CATEGORY_NAMES:
                return True
        except:
            pass
        return False

    if scope_str == "Current Selection":
        ids = list(uidoc_.Selection.GetElementIds())
        els = [doc_.GetElement(i) for i in ids]
        els = [e for e in els if e is not None and keep(e)]
        if progress_cb:
            progress_cb("Collected selection", len(els), len(els))
        return els

    collector = FilteredElementCollector(doc_, doc_.ActiveView.Id) if scope_str == "Current View" else FilteredElementCollector(doc_)
    collector = collector.WhereElementIsNotElementType()
    if cat_filter:
        collector = collector.WherePasses(cat_filter)

    results = []
    last_ui = time.time()

    for e in collector:
        if keep(e):
            results.append(e)

        now = time.time()
        if now - last_ui >= UI_UPDATE_INTERVAL:
            last_ui = now
            if progress_cb:
                progress_cb("Collecting elements...", len(results), 0)
            if pump_cb:
                pump_cb()

    if progress_cb:
        progress_cb("Collection complete", len(results), len(results))
    return results

# -----------------------------------------------------------------------------
# Micron Zone logic (DirectShape zones + spatial grid)
# -----------------------------------------------------------------------------
ZONE_CACHE = {}

def _doc_cache_key(doc_):
    """Best-effort stable key for the current document instance."""
    try:
        pn = doc_.PathName
        if pn:
            return pn
    except:
        pass
    try:
        return "{}|{}".format(doc_.Title, doc_.GetHashCode())
    except:
        return str(doc_)


def _floor_div(v, step):
    try:
        return int(math.floor(float(v) / float(step)))
    except:
        return 0


def _cell_key(ix, iy):
    return "{},{}".format(int(ix), int(iy))


def _get_text_bip(elem, bip):
    try:
        p = elem.get_Parameter(bip)
        if p:
            return p.AsString()
    except:
        pass
    return None


def _bbox_contains_xyz(bb, pt):
    if not bb or not pt:
        return False
    try:
        return (pt.X >= bb.Min.X and pt.X <= bb.Max.X and
                pt.Y >= bb.Min.Y and pt.Y <= bb.Max.Y and
                pt.Z >= bb.Min.Z and pt.Z <= bb.Max.Z)
    except:
        return False


def _bbox_intersects_xyz(a, b):
    if not a or not b:
        return False
    try:
        return not (
            a.Max.X < b.Min.X or a.Min.X > b.Max.X or
            a.Max.Y < b.Min.Y or a.Min.Y > b.Max.Y or
            a.Max.Z < b.Min.Z or a.Min.Z > b.Max.Z
        )
    except:
        return False


def _bbox_volume(bb):
    try:
        dx = float(bb.Max.X - bb.Min.X)
        dy = float(bb.Max.Y - bb.Min.Y)
        dz = float(bb.Max.Z - bb.Min.Z)
        if dx <= 0 or dy <= 0 or dz <= 0:
            return float("inf")
        return dx * dy * dz
    except:
        return float("inf")


def _xyz(x, y, z):
    from Autodesk.Revit.DB import XYZ
    return XYZ(x, y, z)


def _get_element_center(el, doc_):
    if not el:
        return None
    try:
        loc = el.Location
        if loc and loc.GetType().FullName == "Autodesk.Revit.DB.LocationPoint":
            pt = loc.Point
            if pt:
                return pt
    except:
        pass
    try:
        bb = el.get_BoundingBox(None)
        if (not bb) and doc_.ActiveView:
            bb = el.get_BoundingBox(doc_.ActiveView)
        if bb:
            return _xyz(
                (bb.Min.X + bb.Max.X) * 0.5,
                (bb.Min.Y + bb.Max.Y) * 0.5,
                (bb.Min.Z + bb.Max.Z) * 0.5
            )
    except:
        pass
    return None


def _get_element_bbox(el, doc_):
    if not el:
        return None
    try:
        bb = el.get_BoundingBox(None)
        if bb:
            return bb
    except:
        pass
    try:
        if doc_ and doc_.ActiveView:
            return el.get_BoundingBox(doc_.ActiveView)
    except:
        pass
    return None


def _count_zone_directshapes(doc_):
    try:
        dss = FilteredElementCollector(doc_).OfClass(DirectShape).WhereElementIsNotElementType()
        n = 0
        for ds in dss:
            mark = _get_text_bip(ds, BuiltInParameter.ALL_MODEL_MARK)
            if mark == GEN_TAG:
                n += 1
        return n
    except:
        return 0


def _build_zone_cache(doc_):
    zones = {}   # idInt -> {"BB":bb, "Name":name, "Vol": vol}
    grid = {}    # "ix,iy" -> [idInt, ...]

    dss = FilteredElementCollector(doc_).OfClass(DirectShape).WhereElementIsNotElementType()
    for ds in dss:
        try:
            mark = _get_text_bip(ds, BuiltInParameter.ALL_MODEL_MARK)
            if mark != GEN_TAG:
                continue

            bb = ds.get_BoundingBox(None)
            if not bb:
                continue

            name = _get_text_bip(ds, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
            if not name:
                try:
                    name = ds.Name
                except:
                    name = None
            if not name:
                continue

            vol = _bbox_volume(bb)
            id_int = ds.Id.IntegerValue
            zones[id_int] = {"BB": bb, "Name": name, "Vol": vol}

            ix0 = _floor_div(bb.Min.X, GRID_SIZE_FT)
            ix1 = _floor_div(bb.Max.X, GRID_SIZE_FT)
            iy0 = _floor_div(bb.Min.Y, GRID_SIZE_FT)
            iy1 = _floor_div(bb.Max.Y, GRID_SIZE_FT)

            for ix in range(ix0, ix1 + 1):
                for iy in range(iy0, iy1 + 1):
                    k = _cell_key(ix, iy)
                    if k not in grid:
                        grid[k] = []
                    grid[k].append(id_int)
        except:
            pass

    ZONE_CACHE.clear()
    ZONE_CACHE.update({
        "DocKey": _doc_cache_key(doc_),
        "GridSize": float(GRID_SIZE_FT),
        "Zones": zones,
        "Grid": grid,
        "ZoneCount": len(zones),
    })


def _ensure_zone_cache(doc_):
    need = False
    dk = _doc_cache_key(doc_)
    try:
        if not ZONE_CACHE.get("DocKey"):
            need = True
        elif ZONE_CACHE.get("DocKey") != dk:
            need = True
        elif float(ZONE_CACHE.get("GridSize", -1.0)) != float(GRID_SIZE_FT):
            need = True
    except:
        need = True

    if not need:
        try:
            current = _count_zone_directshapes(doc_)
            if int(current) != int(ZONE_CACHE.get("ZoneCount", -999)):
                need = True
        except:
            need = True

    if need:
        _build_zone_cache(doc_)


def _param_is_yes(el, param_name):
    p = get_param(el, param_name)
    if not p:
        return False
    try:
        return int(p.AsInteger()) == 1
    except:
        try:
            s = p.AsString() or p.AsValueString() or ""
            t = s.strip().lower()
            return t in ("1", "true", "yes", "y", "checked", "on")
        except:
            return False


def calc_micron_zone(el, doc_):
    if not el or not doc_:
        return MSG_NOZONE

    # OVERRIDE behavior
    override_param = "em_CP_Override Parameter Sync"
    return_param = "Scope_Const_Work_Pkg_Ext_MTI"

    try:
        if _param_is_yes(el, override_param):
            ov = get_param_string(el, return_param)
            if ov and ov.strip():
                return ov
            # if checked but blank, fall through
    except:
        pass

    _ensure_zone_cache(doc_)
    if not ZONE_CACHE or int(ZONE_CACHE.get("ZoneCount", 0)) <= 0:
        return MSG_NOZONE

    # ---- 1) CENTER POINT MATCH (XYZ) ----
    pt = _get_element_center(el, doc_)
    if not pt:
        return MSG_NOZONE

    ix = _floor_div(pt.X, GRID_SIZE_FT)
    iy = _floor_div(pt.Y, GRID_SIZE_FT)

    center_candidates = []
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            k = _cell_key(ix + dx, iy + dy)
            ids = ZONE_CACHE["Grid"].get(k)
            if not ids:
                continue
            for id_int in ids:
                z = ZONE_CACHE["Zones"].get(id_int)
                if not z:
                    continue
                bbz = z["BB"]
                if _bbox_contains_xyz(bbz, pt):
                    center_candidates.append((z.get("Vol", _bbox_volume(bbz)), z["Name"]))

    if center_candidates:
        center_candidates.sort(key=lambda x: x[0])  # smallest vol wins
        return center_candidates[0][1]

    # ---- 2) FALLBACK: ELEMENT BBOX INTERSECTS ZONE (XYZ) ----
    ebb = _get_element_bbox(el, doc_)
    if not ebb:
        return MSG_NOZONE

    ix0 = _floor_div(ebb.Min.X, GRID_SIZE_FT)
    ix1 = _floor_div(ebb.Max.X, GRID_SIZE_FT)
    iy0 = _floor_div(ebb.Min.Y, GRID_SIZE_FT)
    iy1 = _floor_div(ebb.Max.Y, GRID_SIZE_FT)

    tested = set()
    intersect_candidates = []

    for ix2 in range(ix0 - 1, ix1 + 2):
        for iy2 in range(iy0 - 1, iy1 + 2):
            k = _cell_key(ix2, iy2)
            ids = ZONE_CACHE["Grid"].get(k)
            if not ids:
                continue
            for id_int in ids:
                if id_int in tested:
                    continue
                tested.add(id_int)
                z = ZONE_CACHE["Zones"].get(id_int)
                if not z:
                    continue
                bbz = z["BB"]
                if _bbox_intersects_xyz(ebb, bbz):
                    intersect_candidates.append((z.get("Vol", _bbox_volume(bbz)), z["Name"]))

    if intersect_candidates:
        intersect_candidates.sort(key=lambda x: x[0])  # smallest vol wins
        return intersect_candidates[0][1]

    return MSG_NOZONE

# -----------------------------------------------------------------------------
# Construction status logic
# -----------------------------------------------------------------------------

def calc_construction_status_auto(el):
    s = get_param_string(el, "Scope_Coord_Status_External_MTI")
    if not s:
        return "WIP"
    u = s.upper()
    if "IFF" in u:
        return "IFF"
    if "AB" in u:
        return "AP"   # keep your PowerShell mapping exactly
    return "WIP"

# -----------------------------------------------------------------------------
# ExternalEvent
# -----------------------------------------------------------------------------

class BatchWriteRequest(object):
    def __init__(self):
        self.scope = "Current View"
        self.only_owned = False
        self.do_sync = False

        # Outputs
        self.do_jobnum = True
        self.do_scope_new_yes = True
        self.do_future_no = True
        self.do_demo_no = True
        self.do_record_no = True
        self.do_zone = True

        self.do_const_status = True
        self.const_status_mode = "AUTO"  # AUTO, WIP, IFF, AB


class BatchWriteHandler(IExternalEventHandler):
    def __init__(self, req, win):
        self._req = req
        self._win = win
        self._cat_filter = None

    def GetName(self):
        return WINDOW_TITLE

    def Execute(self, uiapp):
        if RUN_STATE.is_running:
            self._win.set_status_lines(["Already running. Click Stop to request stop."])
            return

        RUN_STATE.is_running = True
        RUN_STATE.cancel_requested = False

        start_time = time.time()
        processed = 0
        cancelled = False
        sync_result = "Not requested"

        def pump():
            do_events()

        def prog(phase, value, maximum):
            self._win.set_progress(phase, value, maximum)
            do_events()

        try:
            uidoc_ = uiapp.ActiveUIDocument
            doc_ = uidoc_.Document
            username = None
            try:
                username = uiapp.Application.Username
            except:
                username = None

            if self._cat_filter is None:
                self._cat_filter = build_category_filter(doc_)

            # COLLECTION
            self._win.set_status_lines([
                "Scope: {}".format(self._req.scope),
                "Only owned: {}".format("YES" if self._req.only_owned else "NO"),
                "Collecting..."
            ])
            prog("Collecting elements...", 0, 0)

            elements = collect_elements(
                doc_, uidoc_, self._req.scope, self._cat_filter,
                only_owned=self._req.only_owned, username=username,
                progress_cb=prog, pump_cb=pump
            )

            total = len(elements)
            self._win.set_status_lines([
                "Scope: {}".format(self._req.scope),
                "Only owned: {}".format("YES" if self._req.only_owned else "NO"),
                "Elements collected: {}".format(total)
            ])

            if total == 0:
                self._win.set_progress("Idle", 0, 100)
                return

            if not (
                self._req.do_jobnum or
                self._req.do_scope_new_yes or
                self._req.do_future_no or
                self._req.do_demo_no or
                self._req.do_record_no or
                self._req.do_zone or
                self._req.do_const_status
            ):
                self._win.set_status_lines(["Nothing selected to run (check at least one box)."])
                self._win.set_progress("Idle", 0, 100)
                return

            # WRITING
            self._win.set_progress("Writing parameters...", 0, total)

            # Counters per output
            j_ok = j_skip = 0
            n_ok = n_skip = 0
            f_ok = f_skip = 0
            d_ok = d_skip = 0
            r_ok = r_skip = 0
            z_ok = z_skip = 0
            s_ok = s_skip = 0

            last_ui = time.time()

            tg = TransactionGroup(doc_, "Write MTI Parameters")
            tg.Start()
            try:
                for start in range(0, total, CHUNK_SIZE):
                    if RUN_STATE.cancel_requested:
                        cancelled = True
                        break

                    chunk = elements[start:start + CHUNK_SIZE]
                    t = Transaction(doc_, "Write MTI {}-{}".format(start + 1, min(total, start + CHUNK_SIZE)))
                    t.Start()
                    try:
                        for el in chunk:
                            if RUN_STATE.cancel_requested:
                                cancelled = True
                                break

                            # Job Number
                            if self._req.do_jobnum:
                                if set_param_string(el, "Project Number_MTI", JOB_NUMBER_VALUE):
                                    j_ok += 1
                                else:
                                    j_skip += 1

                            # New scope - Yes
                            if self._req.do_scope_new_yes:
                                if set_param_int(el, "Scope_NEW_MTI", 1):
                                    n_ok += 1
                                else:
                                    n_skip += 1

                            # Future/Demo/Record - No
                            if self._req.do_future_no:
                                if set_param_int(el, "Scope_FUTURE_MTI", 0):
                                    f_ok += 1
                                else:
                                    f_skip += 1

                            if self._req.do_demo_no:
                                if set_param_int(el, "Scope_DEMO_MTI", 0):
                                    d_ok += 1
                                else:
                                    d_skip += 1

                            if self._req.do_record_no:
                                if set_param_int(el, "Scope_RECORD_MTI", 0):
                                    r_ok += 1
                                else:
                                    r_skip += 1

                            # Micron Zone (guarded)
                            if self._req.do_zone:
                                try:
                                    zv = calc_micron_zone(el, doc_)
                                except:
                                    zv = MSG_NOZONE
                                if set_param_string(el, "Scope_Const_Work_Pkg_Ext_MTI", zv):
                                    z_ok += 1
                                else:
                                    z_skip += 1

                            # Construction Status
                            if self._req.do_const_status:
                                mode = (self._req.const_status_mode or "AUTO").upper()
                                if mode == "AUTO":
                                    sv = calc_construction_status_auto(el)
                                elif mode == "WIP":
                                    sv = "WIP"
                                elif mode == "IFF":
                                    sv = "IFF"
                                elif mode == "AB":
                                    sv = "AB"
                                else:
                                    sv = calc_construction_status_auto(el)

                                if set_param_string(el, "Scope_Coord_Status_External_MTI", sv):
                                    s_ok += 1
                                else:
                                    s_skip += 1

                            processed += 1

                            now = time.time()
                            if now - last_ui >= UI_UPDATE_INTERVAL or processed == total:
                                last_ui = now
                                self._win.set_progress("Writing parameters...", processed, total)
                                do_events()

                        t.Commit()
                    except:
                        logger.exception("Chunk failed; rolling back chunk transaction.")
                        t.RollBack()

                    if cancelled:
                        break

                tg.Assimilate()
            except:
                logger.exception("Failed overall; rolling back transaction group.")
                tg.RollBack()
                raise

            # SYNC (only after transactions are closed)
            if (not cancelled) and self._req.do_sync:
                self._win.set_progress("Syncing with Central...", total, total)
                self._win.set_status_lines([
                    "Scope: {}".format(self._req.scope),
                    "Only owned: {}".format("YES" if self._req.only_owned else "NO"),
                    "Processed: {} / {}".format(processed, total),
                    "",
                    "Syncing with Central..."
                ])
                do_events()
                sync_result = sync_with_central(doc_)
            elif cancelled and self._req.do_sync:
                sync_result = "Skipped (cancelled)"

            elapsed = time.time() - start_time

            lines = []
            lines.append("Scope: {}".format(self._req.scope))
            lines.append("Only owned: {}".format("YES" if self._req.only_owned else "NO"))
            lines.append("Elements collected: {}".format(total))
            lines.append("Processed: {} / {}".format(processed, total))
            lines.append("Stopped: {}".format("YES" if cancelled else "NO"))
            lines.append("Elapsed: {:.1f} sec".format(elapsed))
            lines.append("")
            lines.append("Sync with Central: {}".format(sync_result))

            if self._req.do_jobnum:
                lines += ["", "Project Number_MTI", "  wrote:   {}".format(j_ok), "  skipped: {}".format(j_skip)]
            if self._req.do_scope_new_yes:
                lines += ["", "Scope_NEW_MTI", "  wrote:   {}".format(n_ok), "  skipped: {}".format(n_skip)]
            if self._req.do_future_no:
                lines += ["", "Scope_FUTURE_MTI", "  wrote:   {}".format(f_ok), "  skipped: {}".format(f_skip)]
            if self._req.do_demo_no:
                lines += ["", "Scope_DEMO_MTI", "  wrote:   {}".format(d_ok), "  skipped: {}".format(d_skip)]
            if self._req.do_record_no:
                lines += ["", "Scope_RECORD_MTI", "  wrote:   {}".format(r_ok), "  skipped: {}".format(r_skip)]
            if self._req.do_zone:
                lines += ["", "Scope_Const_Work_Pkg_Ext_MTI (Micron Zone)", "  wrote:   {}".format(z_ok), "  skipped: {}".format(z_skip)]
            if self._req.do_const_status:
                lines += ["", "Scope_Coord_Status_External_MTI ({})".format(self._req.const_status_mode), "  wrote:   {}".format(s_ok), "  skipped: {}".format(s_skip)]

            self._win.set_status_lines(lines)
            self._win.set_progress("Done", total, total)

        except Exception as ex:
            logger.exception("Execute failed.")
            self._win.set_status_lines(["Run failed.", "", "ERROR:", str(ex)])
            self._win.set_progress("Error", 0, 100)
        finally:
            RUN_STATE.is_running = False
            do_events()

# -----------------------------------------------------------------------------
# WPF Window (REFRESHED UI ONLY)
# -----------------------------------------------------------------------------
XAML = u"""
<Window xmlns=\"http://schemas.microsoft.com/winfx/2006/xaml/presentation\"
        xmlns:x=\"http://schemas.microsoft.com/winfx/2006/xaml\"
        Title=\"__TITLE__\"
        Width=\"660\" Height=\"940\"
        WindowStartupLocation=\"CenterScreen\"
        Topmost=\"True\"
        ResizeMode=\"CanResizeWithGrip\" Background=\"#F5F6F8\" FontFamily=\"Segoe UI\">

  <Window.Resources>
    <Style TargetType=\"TextBlock\" x:Key=\"SectionTitle\">
      <Setter Property=\"FontSize\" Value=\"15\"/>
      <Setter Property=\"FontWeight\" Value=\"SemiBold\"/>
      <Setter Property=\"Margin\" Value=\"0,0,0,8\"/>
    </Style>

    <Style TargetType=\"Border\" x:Key=\"Card\">
      <Setter Property=\"CornerRadius\" Value=\"12\"/>
      <Setter Property=\"BorderThickness\" Value=\"1\"/>
      <Setter Property=\"BorderBrush\" Value=\"#D9D9D9\"/>
      <Setter Property=\"Background\" Value=\"White\"/>
      <Setter Property=\"Padding\" Value=\"14\"/>
      <Setter Property=\"Margin\" Value=\"0,0,0,12\"/>
    </Style>

    <Style TargetType=\"Button\" x:Key=\"PrimaryBtn\">
      <Setter Property=\"Padding\" Value=\"14,8\"/>
      <Setter Property=\"FontWeight\" Value=\"SemiBold\"/>
      <Setter Property=\"BorderThickness\" Value=\"0\"/>
      <Setter Property=\"Background\" Value=\"#FF6A00\"/>
      <Setter Property=\"Foreground\" Value=\"White\"/>
    </Style>

    <Style TargetType=\"Button\" x:Key=\"SecondaryBtn\">
      <Setter Property=\"Padding\" Value=\"14,8\"/>
      <Setter Property=\"FontWeight\" Value=\"SemiBold\"/>
      <Setter Property=\"BorderThickness\" Value=\"1\"/>
      <Setter Property=\"BorderBrush\" Value=\"#D0D0D0\"/>
      <Setter Property=\"Background\" Value=\"White\"/>
      <Setter Property=\"Foreground\" Value=\"#333333\"/>
    </Style>

    <Style TargetType="Button" x:Key="HelpBtn">
      <Setter Property="Width" Value="18"/>
      <Setter Property="Height" Value="18"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Margin" Value="8,0,0,0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Background" Value="#EEF1F4"/>
      <Setter Property="Foreground" Value="#333333"/>
      <Setter Property="BorderBrush" Value="#D0D0D0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="HorizontalAlignment" Value="Left"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="ToolTipService.ShowDuration" Value="60000"/>
      <Setter Property="Template" >
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border CornerRadius="9" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid Margin=\"14\">
    <Grid.RowDefinitions>
      <RowDefinition Height=\"Auto\"/>
      <RowDefinition Height=\"Auto\"/>
      <RowDefinition Height=\"Auto\"/>
      <RowDefinition Height=\"Auto\"/>
      <RowDefinition Height=\"*\"/>
      <RowDefinition Height=\"Auto\"/>
    </Grid.RowDefinitions>

    <!-- Outputs -->
    <Border Grid.Row=\"0\" Style=\"{StaticResource Card}\">
      <StackPanel>
        <StackPanel Orientation=\"Horizontal\" VerticalAlignment=\"Center\"><TextBlock Text=\"Outputs to write\" Style=\"{StaticResource SectionTitle}\"/><Button x:Name=\"HelpOutputs\" Content=\"?\" Style=\"{StaticResource HelpBtn}\"><Button.ToolTip><TextBlock TextWrapping=\"Wrap\" Width=\"440\" Text=\"Required outputs (recommended):
• Job Number: writes Project Number_MTI = B111.
• New Scope: sets Scope_NEW_MTI = Yes.
• Micron Zone: writes Scope_Const_Work_Pkg_Ext_MTI for Micron construction zones.

Optional outputs (normally leave OFF):
These set Micron scope fields to No (blank/null = faster sync). Turn ON only if you need to force any Yes values back to No.\"/></Button.ToolTip></Button></StackPanel>

        <CheckBox x:Name=\"ChkJob\"   Content=\"Write Job Number (Project Number_MTI = B111)\" IsChecked=\"True\" Margin=\"0,2,0,2\"/>
        <CheckBox x:Name=\"ChkNew\"   Content=\"Write New Scope - Yes (Scope_NEW_MTI = Yes)\" IsChecked=\"True\" Margin=\"0,2,0,2\"/>
        <CheckBox x:Name=\"ChkZone\"  Content=\"Write Micron Zone (Scope_Const_Work_Pkg_Ext_MTI)\" IsChecked=\"True\" Margin=\"0,2,0,8\"/>

        <Separator Margin=\"0,8,0,8\"/>

        <StackPanel Orientation=\"Horizontal\" VerticalAlignment=\"Center\" Margin=\"0,2,0,6\"><CheckBox x:Name=\"ChkStatus\" Content=\"Write Construction Status (Scope_Coord_Status_External_MTI)\" IsChecked=\"True\"/><Button x:Name=\"HelpWriteStatus\" Content=\"?\" Style=\"{StaticResource HelpBtn}\"><Button.ToolTip><TextBlock TextWrapping=\"Wrap\" Width=\"440\" Text=\"Construction Status:
• Auto: keeps existing value; blank becomes WIP. If set to IFF or AB, it stays the same.
• Force WIP: use while modeling (before area sign-off).
• Force IFF: after GC + Micron review/sign-off (Issue for Fabrication).
• Force AB: after laser scan + model updates (As-Built).\"/></Button.ToolTip></Button></StackPanel>

        <StackPanel Orientation=\"Horizontal\" Margin=\"18,0,0,0\">
          <RadioButton x:Name=\"RadAuto\" Content=\"Auto (normalize)\" IsChecked=\"True\" Margin=\"0,0,14,0\"/>
          <RadioButton x:Name=\"RadWIP\"  Content=\"Force WIP\" Margin=\"0,0,14,0\"/>
          <RadioButton x:Name=\"RadIFF\"  Content=\"Force IFF\" Margin=\"0,0,14,0\"/>
          <RadioButton x:Name=\"RadAB\"   Content=\"Force AB\"/>
        </StackPanel>

        <Separator Margin=\"0,12,0,8\"/>

        <TextBlock Text=\"Optional (default OFF)\" FontWeight=\"SemiBold\" Margin=\"0,0,0,4\"/>
        <StackPanel Orientation=\"Vertical\" Margin=\"12,0,0,0\">
          <CheckBox x:Name=\"ChkFuture\" Content=\"Write Future Scope - No (Scope_FUTURE_MTI = No)\" IsChecked=\"False\" Margin=\"0,2,0,2\"/>
          <CheckBox x:Name=\"ChkDemo\"   Content=\"Write Demo Scope - No (Scope_DEMO_MTI = No)\" IsChecked=\"False\" Margin=\"0,2,0,2\"/>
          <CheckBox x:Name=\"ChkRecord\" Content=\"Write Record Scope - No (Scope_RECORD_MTI = No)\" IsChecked=\"False\" Margin=\"0,2,0,2\"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- Options -->
    <Border Grid.Row=\"1\" Style=\"{StaticResource Card}\">
      <StackPanel>
        <StackPanel Orientation=\"Horizontal\" VerticalAlignment=\"Center\"><TextBlock Text=\"Options\" Style=\"{StaticResource SectionTitle}\"/><Button x:Name=\"HelpOptions\" Content=\"?\" Style=\"{StaticResource HelpBtn}\"><Button.ToolTip><TextBlock TextWrapping=\"Wrap\" Width=\"420\" Text=\"Only elements I own: limits writing to elements currently checked out by you (workshared models).
Sync with Central: syncs + relinquishes at the end to help keep the model clean.\"/></Button.ToolTip></Button></StackPanel>
        <CheckBox x:Name=\"ChkOwned\" Content=\"Only elements I own (workshared)\" IsChecked=\"True\" Margin=\"0,2,0,2\"/>
        <CheckBox x:Name=\"ChkSync\"  Content=\"Sync with Central when finished (workshared)\" IsChecked=\"True\" Margin=\"0,2,0,2\"/>
      </StackPanel>
    </Border>

    <!-- Scope -->
    <Border Grid.Row=\"2\" Style=\"{StaticResource Card}\">
      <StackPanel>
        <StackPanel Orientation=\"Horizontal\" VerticalAlignment=\"Center\"><TextBlock Text=\"Scope\" Style=\"{StaticResource SectionTitle}\"/><Button x:Name=\"HelpScope\" Content=\"?\" Style=\"{StaticResource HelpBtn}\"><Button.ToolTip><TextBlock TextWrapping=\"Wrap\" Width=\"420\" Text=\"Choose what to process:
• Current View: elements visible in the active view.
• Entire Model: all matching model elements.
• Current Selection: only the elements you have selected.\"/></Button.ToolTip></Button></StackPanel>
        <StackPanel Orientation=\"Horizontal\" Margin=\"0,0,0,2\">
          <RadioButton x:Name=\"RadView\"  Content=\"Current View\" Margin=\"0,0,14,0\"/>
          <RadioButton x:Name=\"RadModel\" Content=\"Entire Model\" IsChecked=\"True\" Margin=\"0,0,14,0\"/>
          <RadioButton x:Name=\"RadSel\"   Content=\"Current Selection\"/>
        </StackPanel>
      </StackPanel>
    </Border>

    <!-- Status / Progress -->
    <Border Grid.Row=\"4\" Style=\"{StaticResource Card}\">
      <StackPanel>
        <StackPanel Orientation=\"Horizontal\" VerticalAlignment=\"Center\" Margin=\"0,0,0,8\"><TextBlock Text=\"Status\" FontSize=\"15\" FontWeight=\"SemiBold\"/><Button x:Name=\"HelpStatus\" Content=\"?\" Style=\"{StaticResource HelpBtn}\"><Button.ToolTip><TextBlock TextWrapping=\"Wrap\" Width=\"420\" Text=\"Status panel:
Shows progress and live messages while the tool runs.
Idle means nothing is running. Use Stop to safely end a run early.\"/></Button.ToolTip></Button></StackPanel>
        <ProgressBar x:Name=\"Prog\" Height=\"16\" Minimum=\"0\" Maximum=\"100\" Value=\"0\" Margin=\"0,0,0,10\"/>
        <TextBlock x:Name=\"ProgText\" Text=\"Idle\" Margin=\"0,0,0,8\"/>

        <TextBox x:Name=\"StatusText\"
                 Text=\"Ready. Change selection/view then click Run.\"
                 TextWrapping=\"Wrap\"
                 VerticalScrollBarVisibility=\"Auto\"
                 HorizontalScrollBarVisibility=\"Auto\"
                 IsReadOnly=\"True\"
                 AcceptsReturn=\"True\"
                 BorderThickness=\"0\"
                 Background=\"Transparent\"/>
      </StackPanel>
    </Border>

    <!-- Buttons -->
    <StackPanel Grid.Row=\"5\" Orientation=\"Horizontal\" HorizontalAlignment=\"Right\" VerticalAlignment=\"Bottom\" Margin=\"0,12,0,0\">
      <Button x:Name=\"BtnRun\"    Content=\"Run\" Width=\"110\" Height=\"32\" Margin=\"0,0,10,0\" Style=\"{StaticResource PrimaryBtn}\"/>
      <Button x:Name=\"BtnCancel\" Content=\"Stop\" Width=\"110\" Height=\"32\" Margin=\"0,0,10,0\" Style=\"{StaticResource SecondaryBtn}\"/>
      <Button x:Name=\"BtnClose\"  Content=\"Close\" Width=\"110\" Height=\"32\" Style=\"{StaticResource SecondaryBtn}\"/>
    </StackPanel>

  </Grid>
</Window>
""".replace("__TITLE__", WINDOW_TITLE)


class BatchWindow(object):
    def __init__(self):
        self._w = XamlReader.Parse(XAML)
        self._w.Closed += self._on_closed

        # Outputs
        self.ChkJob = self._w.FindName("ChkJob")
        self.ChkNew = self._w.FindName("ChkNew")
        self.ChkFuture = self._w.FindName("ChkFuture")
        self.ChkDemo = self._w.FindName("ChkDemo")
        self.ChkRecord = self._w.FindName("ChkRecord")
        self.ChkZone = self._w.FindName("ChkZone")
        self.ChkStatus = self._w.FindName("ChkStatus")

        self.RadAuto = self._w.FindName("RadAuto")
        self.RadWIP = self._w.FindName("RadWIP")
        self.RadIFF = self._w.FindName("RadIFF")
        self.RadAB = self._w.FindName("RadAB")

        # Options
        self.ChkOwned = self._w.FindName("ChkOwned")
        self.ChkSync  = self._w.FindName("ChkSync")

        # Scope
        self.RadView = self._w.FindName("RadView")
        self.RadModel = self._w.FindName("RadModel")
        self.RadSel = self._w.FindName("RadSel")

        # Buttons
        self.BtnRun = self._w.FindName("BtnRun")
        self.BtnCancel = self._w.FindName("BtnCancel")
        self.BtnClose = self._w.FindName("BtnClose")

        # Status/prog
        self.Prog = self._w.FindName("Prog")
        self.ProgText = self._w.FindName("ProgText")
        self.StatusText = self._w.FindName("StatusText")

        self._req = BatchWriteRequest()
        self._handler = BatchWriteHandler(self._req, self)
        self._ext_event = ExternalEvent.Create(self._handler)

        self.BtnRun.Click += self._on_run
        self.BtnCancel.Click += self._on_cancel
        self.BtnClose.Click += self._on_close

    def Show(self):
        self._w.Show()

    def _get_scope(self):
        if self.RadSel.IsChecked:
            return "Current Selection"
        if self.RadView.IsChecked:
            return "Current View"
        return "Entire Model"

    def _get_status_mode(self):
        if self.RadWIP.IsChecked:
            return "WIP"
        if self.RadIFF.IsChecked:
            return "IFF"
        if self.RadAB.IsChecked:
            return "AB"
        return "AUTO"

    def set_status_lines(self, lines):
        try:
            self.StatusText.Text = "\n".join(lines)
            try:
                self.StatusText.ScrollToEnd()
            except:
                pass
        except:
            pass

    def set_progress(self, phase, value, maximum):
        try:
            self.ProgText.Text = phase
            if maximum == 0:
                self.Prog.IsIndeterminate = True
            else:
                self.Prog.IsIndeterminate = False
                self.Prog.Minimum = 0
                self.Prog.Maximum = float(maximum)
                self.Prog.Value = float(value)
        except:
            pass

    def _on_run(self, sender, args):
        try:
            if RUN_STATE.is_running:
                self.ProgText.Text = "Already running... click Stop."
                return

            # outputs
            self._req.do_jobnum = bool(self.ChkJob.IsChecked)
            self._req.do_scope_new_yes = bool(self.ChkNew.IsChecked)
            self._req.do_future_no = bool(self.ChkFuture.IsChecked)
            self._req.do_demo_no = bool(self.ChkDemo.IsChecked)
            self._req.do_record_no = bool(self.ChkRecord.IsChecked)
            self._req.do_zone = bool(self.ChkZone.IsChecked)

            self._req.do_const_status = bool(self.ChkStatus.IsChecked)
            self._req.const_status_mode = self._get_status_mode()

            # options/scope
            self._req.only_owned = bool(self.ChkOwned.IsChecked)
            self._req.do_sync = bool(self.ChkSync.IsChecked)
            self._req.scope = self._get_scope()

            RUN_STATE.cancel_requested = False
            self.set_progress("Queued...", 0, 100)
            self.set_status_lines(["Queued..."])
            do_events()
            self._ext_event.Raise()
        except Exception as ex:
            self.set_status_lines(["UI error:", str(ex)])

    def _on_cancel(self, sender, args):
        RUN_STATE.cancel_requested = True
        if RUN_STATE.is_running:
            self.ProgText.Text = "Stop requested... stopping soon."
        else:
            self.ProgText.Text = "Stop requested (not running)."
        do_events()

    def _on_close(self, sender, args):
        try:
            self._w.Close()
        except:
            pass

    def _on_closed(self, sender, args):
        try:
            self._ext_event = None
        except:
            pass

# -----------------------------------------------------------------------------
# Launch
# -----------------------------------------------------------------------------
win = BatchWindow()
win.Show()
