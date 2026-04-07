# -*- coding: utf-8 -*-
"""
pyRevit - Modeless "Weight and Volume Calculator" (ExternalEvent)

Writes (internal units):
  - CP_Weight (kg)
  - CP_Volume (ft^3)
  - CP_BOM_Category (string)

UI improvements:
  - In-window ProgressBar + status line
  - UI updates + Cancel responsiveness via time-throttled DoEvents
  - Updates during BOTH collection and writing so it never looks stalled

Performance:
  - No per-element UI update
  - UI pump runs ~4 times/sec by default
  - Chunked Transactions for large models
"""

from __future__ import division

import re
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
WINDOW_TITLE = "Weight and Volume Calculator"

CHUNK_SIZE = 10000               # keep large for speed
UI_UPDATE_INTERVAL = 0.25        # seconds; ~4 updates/sec (minimal overhead but feels alive)

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

KG_PER_LB = 0.45359237
GAL_TO_FT3 = 1.0 / 7.48051948

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
# Fishmouth lookup tables
# -----------------------------------------------------------------------------
WEIGHT_PER_FOOT_STD = {
    "0.125":0.24,"0.25":0.42,"0.375":0.57,"0.5":0.85,"0.75":1.13,"1":1.68,"1.25":2.28,"1.5":2.72,
    "2":3.66,"2.5":5.1,"3":6.6,"3.5":8.17,"4":10.79,"5":14.62,"6":18.97,"8":28.55,"10":40.48,
    "12":53.52,"14":63.52,"16":74.75,"18":92,"20":105,"24":140,"30":182,"36":216,"42":254,"48":280
}
WEIGHT_PER_FOOT_SCH10 = {
    "0.5":0.65,"0.75":0.72,"1":1.27,"1.25":1.63,"1.5":2.03,"2":2.64,"2.5":3.33,"3":4.3,"4":5.62,
    "5":7.77,"6":9.29,"8":14.97,"10":20.55,"12":26.4,"14":30.63,"16":35.43,"18":40.56,"20":45.62,
    "24":56.84,"30":77.45,"36":92.69,"42":109.56,"48":123.93
}
WEIGHT_PER_FOOT_SCH40 = {
    "0.125":0.24,"0.25":0.42,"0.375":0.57,"0.5":0.85,"0.75":1.13,"1":1.68,"1.25":2.28,"1.5":2.72,
    "2":3.66,"2.5":5.8,"3":7.58,"3.5":9.11,"4":10.79,"5":14.62,"6":18.97,"8":28.55,"10":40.48,
    "12":53.52,"14":63.52,"16":74.75,"18":92,"20":105,"24":140,"30":182,"36":216,"42":254,"48":280
}

GPF_STD = {0.125:0.002,0.25:0.005,0.375:0.011,0.5:0.018,0.75:0.041,1.0:0.065,1.25:0.100,1.5:0.140,2.0:0.225,3.0:0.500,4.0:0.900,6.0:2.000,8.0:3.500,10.0:5.500,12.0:8.000,14.0:10.500,16.0:13.500,18.0:17.000,20.0:21.000,24.0:30.000,30.0:46.000,36.0:66.000,42.0:90.000,48.0:117.000}
GPF_SCH10 = {0.125:0.003,0.25:0.006,0.375:0.012,0.5:0.020,0.75:0.045,1.0:0.070,1.25:0.110,1.5:0.150,2.0:0.250,3.0:0.540,4.0:1.000,6.0:2.200,8.0:3.900,10.0:6.100,12.0:8.800,14.0:11.600,16.0:14.800,18.0:18.600,20.0:22.800,24.0:32.800,30.0:50.000,36.0:72.000,42.0:98.000,48.0:128.000}
GPF_SCH40 = {0.125:0.002,0.25:0.005,0.375:0.010,0.5:0.017,0.75:0.038,1.0:0.060,1.25:0.092,1.5:0.130,2.0:0.210,3.0:0.475,4.0:0.850,6.0:1.900,8.0:3.300,10.0:5.200,12.0:7.600,14.0:10.000,16.0:12.900,18.0:16.200,20.0:20.000,24.0:28.800,30.0:44.000,36.0:63.000,42.0:86.000,48.0:112.000}

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def is_fab_part(el):
    return (FabricationPart is not None) and isinstance(el, FabricationPart)

def get_param(el, name):
    try:
        return el.LookupParameter(name)
    except:
        return None

def set_param_double(el, name, val):
    p = get_param(el, name)
    if not p or p.IsReadOnly:
        return False
    try:
        p.Set(float(val))
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

def get_type_elem(el, doc_):
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue != -1:
            return doc_.GetElement(tid)
    except:
        pass
    return None

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

def get_cid_fab(itm):
    try:
        cid = itm.ItemCustomId
        if cid:
            return cid
    except:
        pass
    p = get_param(itm, "ItemCustomId")
    if p and p.HasValue:
        try:
            return int(p.AsString())
        except:
            pass
    return None

def get_spec_string_fab(itm):
    for attr in ("ProductSpecificationDescription", "Specification"):
        try:
            s = getattr(itm, attr)
            if s:
                return s
        except:
            pass
    return ""

def schedule_key(spec_str):
    s = (spec_str or "").upper()
    if "SCH 10" in s: return "SCH10"
    if "SCH 40" in s: return "SCH40"
    if "STD" in s:    return "STD"
    return "STD"

def closest_key_numeric(keys, target):
    best = None
    bestdiff = None
    for k in keys:
        try:
            diff = abs(float(k) - float(target))
        except:
            continue
        if best is None or diff < bestdiff:
            best = k
            bestdiff = diff
    return best

# -----------------------------------------------------------------------------
# Calculators
# -----------------------------------------------------------------------------
def calc_bom_category(el, doc_):
    if is_fab_part(el):
        try:
            if el.Category:
                return el.Category.Name
        except:
            pass
        return "Unknown ITM Category"

    p = get_param(el, "CPL_Family Category")
    if p and p.HasValue:
        try:
            v = (p.AsString() or "").strip()
            if v:
                return v
        except:
            pass

    te = get_type_elem(el, doc_)
    if te:
        tp = get_param(te, "CPL_Family Category")
        if tp and tp.HasValue:
            try:
                v = (tp.AsString() or "").strip()
                if v:
                    return v
            except:
                pass

    try:
        if el.Category:
            return el.Category.Name
    except:
        pass
    return "Unknown Category"

def parse_fishmouth_tap_size(entry_str):
    if not entry_str:
        return None
    matches = re.findall(r"(\d+(?:\.\d+)?|\d+\s*[-/]\s*\d+)", entry_str)
    nums = []
    for m in matches:
        v = re.sub(r"\s", "", m)
        if "-" in v:
            v = v.replace("-", "/")
        if "/" in v:
            try:
                a, b = v.split("/")
                nums.append(float(a) / float(b))
            except:
                pass
        else:
            try:
                nums.append(float(v))
            except:
                pass
    return sorted(nums)[0] if nums else None

def fishmouth_weight_kg(itm):
    entry = None
    for attr in ("ProductEntry", "ProductSizeDescription"):
        try:
            entry = getattr(itm, attr)
            if entry:
                break
        except:
            pass
    if not entry:
        pe = get_param(itm, "Product Entry")
        if pe and pe.HasValue:
            try:
                entry = pe.AsString()
            except:
                entry = None

    tap = parse_fishmouth_tap_size(entry)
    if tap is None:
        return 0.0

    sk = schedule_key(get_spec_string_fab(itm))
    table = WEIGHT_PER_FOOT_STD if sk == "STD" else (WEIGHT_PER_FOOT_SCH10 if sk == "SCH10" else WEIGHT_PER_FOOT_SCH40)

    tap_key = "{0}".format(tap)
    wpf = table.get(tap_key, None)
    if wpf is None:
        ck = closest_key_numeric(table.keys(), tap)
        if ck is not None:
            wpf = table.get(ck, None)
    if wpf is None:
        return 0.0

    try:
        length = float(itm.CenterlineLength or 0.0)
    except:
        length = 0.0
    if length <= 0:
        for alt in ("Tap Length", "Cut Length", "Height"):
            p = get_param(itm, alt)
            if p and p.HasValue:
                try:
                    length = float(p.AsDouble())
                    if length > 0:
                        break
                except:
                    pass
    if length <= 0:
        return 0.0

    return float(wpf) * float(length) * KG_PER_LB

def calc_weight_kg(el, doc_):
    if is_fab_part(el):
        cid = get_cid_fab(el)
        if cid == 2875:
            return fishmouth_weight_kg(el)

        try:
            w = getattr(el, "Weight", None)
            if w and w > 0:
                return float(w)
        except:
            pass
        try:
            w = el.Get_Weight()
            if w and w > 0:
                return float(w)
        except:
            pass

        p = get_param(el, "Weight")
        if p and p.HasValue:
            try:
                v = float(p.AsDouble())
                if v > 0:
                    return v
            except:
                pass

        p = get_param(el, "CP_Weight")
        if p and p.HasValue:
            try:
                return float(p.AsDouble())
            except:
                pass

        return 0.0

    p = get_param(el, "CP_Weight")
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            pass

    te = get_type_elem(el, doc_)
    if te:
        tp = get_param(te, "CP_Weight")
        if tp and tp.HasValue:
            try:
                return float(tp.AsDouble())
            except:
                pass

    return 0.0

def fishmouth_volume_ft3(itm):
    sk = schedule_key(get_spec_string_fab(itm))
    lookup = GPF_STD if sk == "STD" else (GPF_SCH10 if sk == "SCH10" else GPF_SCH40)

    entry = None
    p = get_param(itm, "Product Entry")
    if p and p.HasValue:
        try:
            entry = p.AsString()
        except:
            entry = None
    if not entry:
        return 0.0

    m = re.search(r"\d+[xX](\d+\.?\d*)", entry)
    if not m:
        return 0.0
    try:
        tap_size = float(m.group(1))
    except:
        return 0.0

    ck = None
    bestd = None
    for k in lookup.keys():
        d = abs(float(k) - tap_size)
        if ck is None or d < bestd:
            ck = k
            bestd = d
    if ck is None:
        return 0.0

    try:
        length = float(itm.CenterlineLength or 0.0)
    except:
        length = 0.0
    if length <= 0:
        return 0.0

    gallons = float(length) * float(lookup[ck])
    return gallons * GAL_TO_FT3

def get_double_param(el, name):
    p = get_param(el, name)
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            return None
    return None

def coupler_length_ft(itm):
    try:
        cm = getattr(itm, "ConnectorManager", None)
        if cm:
            conns = cm.Connectors
            pts = []
            for c in conns:
                try:
                    pts.append(c.Origin)
                except:
                    try:
                        pts.append(c.CoordinateSystem.Origin)
                    except:
                        pass
            if len(pts) >= 2:
                p1, p2 = pts[0], pts[1]
                dx, dy, dz = (p2.X - p1.X), (p2.Y - p1.Y), (p2.Z - p1.Z)
                d = math.sqrt(dx*dx + dy*dy + dz*dz)
                if d > 0:
                    return float(d)
    except:
        pass
    try:
        return float(itm.CenterlineLength or 0.0)
    except:
        return 0.0

def calc_volume_ft3(el, doc_):
    if is_fab_part(el):
        cid = get_cid_fab(el)
        desc = ""
        try:
            desc = el.ProductLongDescription or ""
        except:
            desc = ""

        id_ft = get_double_param(el, "Inside Diameter")

        if cid == 2060 and ("Cap" in desc):
            try:
                dims = el.GetDimensions()
                collar = None
                height = None
                for dim in dims:
                    if dim.Name == "Collar":
                        collar = float(el.GetDimensionValue(dim))
                    elif dim.Name == "Height":
                        height = float(el.GetDimensionValue(dim))
                if id_ft and id_ft > 0 and collar and height and collar > 0 and height > 0:
                    r = id_ft / 2.0
                    return float((math.pi * r * r * collar) + ((2.0/3.0) * math.pi * (r ** 3)))
            except:
                pass

        if cid == 2875:
            v = fishmouth_volume_ft3(el)
            if v > 0:
                return v

        if cid == 2522:
            if not id_ft or id_ft <= 0:
                cand = []
                for nm in ("Inside Diameter 1", "Inside Diameter 2"):
                    vv = get_double_param(el, nm)
                    if vv and vv > 0:
                        cand.append(vv)
                if cand:
                    id_ft = min(cand)
            gap = coupler_length_ft(el)
            if id_ft and id_ft > 0 and gap and gap > 0:
                r = id_ft / 2.0
                return float(math.pi * r * r * gap)

        try:
            length = float(el.CenterlineLength or 0.0)
        except:
            length = 0.0
        if id_ft and id_ft > 0 and length > 0:
            r = id_ft / 2.0
            return float(math.pi * r * r * length)

        p = get_param(el, "CP_Volume")
        if p and p.HasValue:
            try:
                return float(p.AsDouble())
            except:
                pass
        return 0.0

    p = get_param(el, "CP_Volume")
    if p and p.HasValue:
        try:
            return float(p.AsDouble())
        except:
            pass
    return 0.0

# -----------------------------------------------------------------------------
# Collection (yields UI while enumerating)
# -----------------------------------------------------------------------------
def collect_elements(doc_, uidoc_, scope_str, cat_filter, progress_cb=None, pump_cb=None):
    def keep(e):
        try:
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
# ExternalEvent
# -----------------------------------------------------------------------------
class BatchWriteRequest(object):
    def __init__(self):
        self.do_weight = True
        self.do_volume = True
        self.do_bomcat = True
        self.scope = "Current View"

class BatchWriteHandler(IExternalEventHandler):
    def __init__(self, req, win):
        self._req = req
        self._win = win
        self._cat_filter = None

    def GetName(self):
        return WINDOW_TITLE

    def Execute(self, uiapp):
        if RUN_STATE.is_running:
            self._win.set_status_lines(["Already running. Click Cancel to request stop."])
            return

        RUN_STATE.is_running = True
        RUN_STATE.cancel_requested = False

        start_time = time.time()
        processed = 0
        cancelled = False

        # helpers to update UI safely
        def pump():
            do_events()

        def prog(phase, value, maximum):
            # maximum==0 means indeterminate collection-style updates
            self._win.set_progress(phase, value, maximum)
            do_events()

        try:
            uidoc_ = uiapp.ActiveUIDocument
            doc_ = uidoc_.Document

            if self._cat_filter is None:
                self._cat_filter = build_category_filter(doc_)

            # COLLECTION (with UI yielding)
            self._win.set_status_lines(["Scope: {}".format(self._req.scope), "Collecting..."])
            prog("Collecting elements...", 0, 0)

            elements = collect_elements(doc_, uidoc_, self._req.scope, self._cat_filter,
                                        progress_cb=prog, pump_cb=pump)

            total = len(elements)
            self._win.set_status_lines(["Scope: {}".format(self._req.scope), "Elements collected: {}".format(total)])

            if total == 0:
                self._win.set_progress("Idle", 0, 100)
                return

            if not (self._req.do_weight or self._req.do_volume or self._req.do_bomcat):
                self._win.set_status_lines(["Nothing selected to run (check at least one box)."])
                self._win.set_progress("Idle", 0, 100)
                return

            # WRITING
            self._win.set_progress("Writing parameters...", 0, total)

            w_ok = w_skip = 0
            v_ok = v_skip = 0
            b_ok = b_skip = 0

            last_ui = time.time()

            tg = TransactionGroup(doc_, "Write CP_*")
            tg.Start()
            try:
                for start in range(0, total, CHUNK_SIZE):
                    if RUN_STATE.cancel_requested:
                        cancelled = True
                        break

                    chunk = elements[start:start + CHUNK_SIZE]
                    t = Transaction(doc_, "Write CP_* {}-{}".format(start + 1, min(total, start + CHUNK_SIZE)))
                    t.Start()
                    try:
                        for el in chunk:
                            if RUN_STATE.cancel_requested:
                                cancelled = True
                                break

                            if self._req.do_weight:
                                val = calc_weight_kg(el, doc_)
                                if set_param_double(el, "CP_Weight", val): w_ok += 1
                                else: w_skip += 1

                            if self._req.do_volume:
                                val = calc_volume_ft3(el, doc_)
                                if set_param_double(el, "CP_Volume", val): v_ok += 1
                                else: v_skip += 1

                            if self._req.do_bomcat:
                                val = calc_bom_category(el, doc_)
                                if set_param_string(el, "CP_BOM_Category", val): b_ok += 1
                                else: b_skip += 1

                            processed += 1

                            # time-based UI update (low overhead, feels alive)
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

            elapsed = time.time() - start_time

            lines = []
            lines.append("Scope: {}".format(self._req.scope))
            lines.append("Elements collected: {}".format(total))
            lines.append("Processed: {} / {}".format(processed, total))
            lines.append("Cancelled: {}".format("YES" if cancelled else "NO"))
            lines.append("Elapsed: {:.1f} sec".format(elapsed))

            if self._req.do_weight:
                lines += ["", "CP_Weight", "  wrote:   {}".format(w_ok), "  skipped: {}".format(w_skip)]
            if self._req.do_volume:
                lines += ["", "CP_Volume", "  wrote:   {}".format(v_ok), "  skipped: {}".format(v_skip)]
            if self._req.do_bomcat:
                lines += ["", "CP_BOM_Category", "  wrote:   {}".format(b_ok), "  skipped: {}".format(b_skip)]

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
# WPF Window (ProgressBar inside)
# -----------------------------------------------------------------------------
XAML = u"""
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="__TITLE__"
        Width="560" Height="600"
        WindowStartupLocation="CenterScreen"
        Topmost="True"
        ResizeMode="CanResizeWithGrip">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="Select outputs to write:" FontWeight="SemiBold" Margin="0,0,0,6"/>

    <StackPanel Grid.Row="1" Orientation="Vertical" Margin="0,0,0,12">
      <CheckBox x:Name="ChkWeight" Content="Write CP_Weight" IsChecked="True" Margin="0,2,0,2"/>
      <CheckBox x:Name="ChkVolume" Content="Write CP_Volume" IsChecked="True" Margin="0,2,0,2"/>
      <CheckBox x:Name="ChkBom"    Content="Write CP_BOM_Category" IsChecked="True" Margin="0,2,0,2"/>
    </StackPanel>

    <TextBlock Grid.Row="2" Text="Scope:" FontWeight="SemiBold" Margin="0,0,0,6"/>

    <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="0,0,0,12">
      <RadioButton x:Name="RadView" Content="Current View" IsChecked="True" Margin="0,0,14,0"/>
      <RadioButton x:Name="RadModel" Content="Entire Model" Margin="0,0,14,0"/>
      <RadioButton x:Name="RadSel" Content="Current Selection"/>
    </StackPanel>

    <Border Grid.Row="4" BorderBrush="#CFCFCF" BorderThickness="1" CornerRadius="6" Padding="10">
      <StackPanel>
        <ProgressBar x:Name="Prog" Height="16" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,8"/>
        <TextBlock x:Name="ProgText" Text="Idle" Margin="0,0,0,8"/>
        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
          <TextBlock x:Name="StatusText"
                     Text="Ready. Change selection/view then click Run."
                     TextWrapping="Wrap"/>
        </ScrollViewer>
      </StackPanel>
    </Border>

    <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
      <Button x:Name="BtnRun" Content="Run" Width="110" Height="32" Margin="0,0,10,0"/>
      <Button x:Name="BtnCancel" Content="Cancel" Width="110" Height="32" Margin="0,0,10,0"/>
      <Button x:Name="BtnClose" Content="Close" Width="110" Height="32"/>
    </StackPanel>

  </Grid>
</Window>
""".replace("__TITLE__", WINDOW_TITLE)

class BatchWindow(object):
    def __init__(self):
        self._w = XamlReader.Parse(XAML)
        self._w.Closed += self._on_closed

        self.ChkWeight = self._w.FindName("ChkWeight")
        self.ChkVolume = self._w.FindName("ChkVolume")
        self.ChkBom = self._w.FindName("ChkBom")

        self.RadView = self._w.FindName("RadView")
        self.RadModel = self._w.FindName("RadModel")
        self.RadSel = self._w.FindName("RadSel")

        self.BtnRun = self._w.FindName("BtnRun")
        self.BtnCancel = self._w.FindName("BtnCancel")
        self.BtnClose = self._w.FindName("BtnClose")

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
        if self.RadModel.IsChecked:
            return "Entire Model"
        return "Current View"

    def set_status_lines(self, lines):
        try:
            self.StatusText.Text = "\n".join(lines)
        except:
            pass

    def set_progress(self, phase, value, maximum):
        try:
            self.ProgText.Text = phase

            # maximum==0 means "collection" (unknown total); just animate-ish by bouncing value
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
                self.ProgText.Text = "Already running... click Cancel."
                return

            self._req.do_weight = bool(self.ChkWeight.IsChecked)
            self._req.do_volume = bool(self.ChkVolume.IsChecked)
            self._req.do_bomcat = bool(self.ChkBom.IsChecked)
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
            self.ProgText.Text = "Cancel requested... stopping soon."
        else:
            self.ProgText.Text = "Cancel requested (not running)."
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
