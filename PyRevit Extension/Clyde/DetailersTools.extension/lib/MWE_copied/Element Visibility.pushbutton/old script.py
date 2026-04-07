# -*- coding: utf-8 -*-
# Element Visibility with WPF UI

from pyrevit import revit, DB, forms, script
from pyrevit.forms import WPFWindow
from System.Collections.Generic import List
from System.Windows import Visibility, WindowState
import os

# --- Core Globals ---
doc   = revit.doc
uidoc = revit.uidoc
out   = script.get_output()
cfg   = script.get_config()

try:
    out.set_title("Element Visibility")
    out.resize(700, 600)
    out.center()
except:
    pass


def _e(msg):
    out.print_md("> " + msg)


def _name(x):
    try:
        n = x.Name
        return n if isinstance(n, basestring) else str(n)
    except:
        return "<no name>"


def _view_label(v):
    vt = getattr(v, "ViewType", None)
    tag = "TEMPLATE • " if getattr(v, "IsTemplate", False) else ""
    nm  = getattr(v, "Name", "<Unnamed>")
    try:
        nm = unicode(nm)
    except:
        pass
    return u"{0}{1} • {2}".format(tag, vt if vt is not None else "View", nm)


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
# WPF Window for UI
# -------------------------
class ElementVisibilityWindow(WPFWindow):
    def __init__(self):
        xaml_path = os.path.join(os.path.dirname(__file__), 'window.xaml')
        WPFWindow.__init__(self, xaml_path)
        
        self.selected_element = None
        self.selected_source = None
        self.selected_target = None
        
        # Get all views
        self.views = _get_views(lambda v: not v.IsTemplate and not isinstance(v, DB.ViewSchedule))
        self.view_labels = [_view_label(v) for v in self.views]
        
        # Populate ComboBoxes
        self.cmbSourceView.Items.Clear()
        self.cmbTargetView.Items.Clear()
        for label in self.view_labels:
            self.cmbSourceView.Items.Add(label)
            self.cmbTargetView.Items.Add(label)
        
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
    
    def OnUseLastElement(self, sender, args):
        last_id_str = cfg.get_option('last_element_id', '')
        if last_id_str:
            try:
                el = doc.GetElement(DB.ElementId(int(last_id_str)))
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
        cfg.last_element_id = str(el.Id.IntegerValue)
        script.save_config()
        
        try:
            self.txtSelectedElement.Text = "{} (Id: {})".format(_name(el), el.Id.IntegerValue)
            self.selectedElementPanel.Visibility = Visibility.Visible
        except:
            pass
        
        self.cmbSourceView.IsEnabled = True
        self.cmbTargetView.IsEnabled = True
        self._check_analyze_button()
    
    def OnSourceViewChanged(self, sender, args):
        if self.cmbSourceView.SelectedIndex >= 0:
            self.selected_source = self.views[self.cmbSourceView.SelectedIndex]
            
            if self.selected_target and self.selected_source.Id == self.selected_target.Id:
                forms.alert("Source and Target views can't be the same.")
                self.cmbSourceView.SelectedIndex = -1
                self.selected_source = None
                return
            
            self._check_analyze_button()
    
    def OnTargetViewChanged(self, sender, args):
        if self.cmbTargetView.SelectedIndex >= 0:
            self.selected_target = self.views[self.cmbTargetView.SelectedIndex]
            
            if self.selected_source and self.selected_target.Id == self.selected_source.Id:
                forms.alert("Source and Target views can't be the same.")
                self.cmbTargetView.SelectedIndex = -1
                self.selected_target = None
                return
            
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


def check_imported_categories(vw, el):
    return False, "OK"


def check_category_hidden(vw, el):
    if isinstance(el, DB.ImportInstance):
        return False, "OK"
    cat = getattr(el, "Category", None)
    if not cat:
        return False, "OK"
    try:
        if vw.CanCategoryBeHidden(cat.Id) and vw.GetCategoryHidden(cat.Id):
            return True, "Model category **'{}'** is turned **OFF** in Visibility/Graphics.".format(cat.Name)
    except:
        pass
    try:
        parent = getattr(cat, "Parent", None)
        if parent and vw.CanCategoryBeHidden(parent.Id) and vw.GetCategoryHidden(parent.Id):
            return True, "Parent model category **'{}'** is turned **OFF** in Visibility/Graphics.".format(parent.Name)
    except:
        pass
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
        if vp and cp and cp.Id.IntegerValue > vp.Id.IntegerValue:
            return (None, "Created in a later phase (**{}**) than the view phase (**{}**).".format(cp.Name, vp.Name))
        if vp and dp and dp.Id.IntegerValue <= vp.Id.IntegerValue:
            return (None, "Demolished in or before the view phase (**{}**). Phase Filter may be hiding it.".format(vp.Name))
    except:
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
    try:
        if not vw.CropBoxActive:
            return False, "OK"

        bb_el = el.get_BoundingBox(vw) or el.get_BoundingBox(None)
        if not bb_el:
            return None, "Could not get element bounding box to check against Crop Region."

        crop = vw.CropBox
        T_world_from_crop = crop.Transform
        T_crop_from_world = T_world_from_crop.Inverse

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
        minz = min(p.Z for p in pts_local); maxz = max(p.Z for p in pts_local)

        if isinstance(vw, DB.ViewPlan):
            outside = (maxx < crop.Min.X or minx > crop.Max.X or
                       maxy < crop.Min.Y or miny > crop.Max.Y)
        else:
            outside = (maxx < crop.Min.X or minx > crop.Max.X or
                       maxy < crop.Min.Y or miny > crop.Max.Y or
                       maxz < crop.Min.Z or minz > crop.Max.Z)

        if outside:
            try:
                if vw.ViewType in [DB.ViewType.Section, DB.ViewType.Elevation]:
                    p_active = vw.LookupParameter("Far Clip Active")
                    if p_active and p_active.AsInteger() == 1:
                        return True, "Element is located entirely beyond the view's **Far Clip Offset**."
            except:
                pass
            return True, "Element is located entirely **outside** the view's active **Crop Region**."

    except Exception as ex:
        return None, "Crop Region check failed: {}".format(ex)

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
        cfg.last_element_id = str(el.Id.IntegerValue)
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

    # ORIGINAL OUTPUT CODE - UNCHANGED
    out.print_md("# Element Visibility")
    out.print_md("- Element: **{}** (Id {})".format(_name(el), el.Id.IntegerValue))
    out.print_md("- Source View (visible): **{}**".format(_view_label(vsrc)))
    out.print_md("- Target View (missing): **{}**".format(_view_label(vtgt)))

    try:
        tid = vtgt.ViewTemplateId
        if tid and tid.IntegerValue > 0:
            vt = doc.GetElement(tid)
            if isinstance(vt, DB.View):
                out.print_md("- Target View Template: `{}`".format(vt.Name))
    except:
        pass

    out.print_md("\n## Checks on Target View")

    checks = [
        ("check_revit_link_visibility",       check_revit_link_visibility),
        ("check_workset_visibility",          check_workset_visibility),
        ("check_imported_categories",         check_imported_categories),
        ("check_category_hidden",             check_category_hidden),
        ("check_filters",                     check_filters),
        ("check_element_hidden_in_view",      check_element_hidden_in_view),
        ("check_discipline_hint",             check_discipline_hint),
        ("check_phase_hint",                  check_phase_hint),
        ("check_design_option_hint",          check_design_option_hint),
        ("check_crop_scopebox_hint",          check_crop_scopebox_hint),
        ("check_view_range_plan",             check_view_range_plan),
    ]

    outputs = []
    likely  = []
    culprits = 0

    for name, fn in checks:
        try:
            state, msg = fn(vtgt, el)
            outputs.append((name, state, msg))
            if state is True:
                culprits += 1
                likely.append(msg)
            elif state is None:
                likely.append(msg)
        except Exception as ex:
            outputs.append((name, "error", str(ex)))

    if culprits == 0 and isinstance(el, DB.ImportInstance):
        for i, (nm, st, ms) in enumerate(outputs):
            if nm == "check_imported_categories":
                outputs[i] = (
                    "check_imported_categories", None,
                    ("**Imported categories** could be deselected in visibility/graphics override "
                     "or one or more **layers** under **location** are turned off. "
                     "Check and expand **Imported Categories** to ensure that relevant checkboxes are checked.")
                )
                break
        likely.append(
            "When selecting **Reveal Hidden Elements** and the element appears, right click on element -> **Unhide in View**; "
            "if element is available, click on Elements to unhide. "
            "If Elements is not available, go to **Imported Categories** under **Visibility/Graphics Override** "
            "and ensure that relevant checkboxes are checked."
        )

    for name, state, msg in outputs:
        if state is True: out.print_md("- X **{}** -> {}".format(name, msg))
        elif state is None: out.print_md("- ! **{}** -> {}".format(name, msg))
        elif state is False: out.print_md("- OK **{}**".format(name))
        elif state == "error": out.print_md("- ERROR **{}** -> {}".format(name, msg))
        else: out.print_md("- OK **{}**".format(name))

    out.print_md("\n## Likely Causes")
    if likely:
        seen = set()
        for m in likely:
            if not m or m in seen: continue
            seen.add(m)
            _e("- " + m)
    else:
        _e("No obvious blockers found. Also check:\n"
           "• Revit Link display settings\n"
           "• Per-element graphic overrides (e.g., transparency, halftone)\n"
           "• View Template rules applied to the target view")

run()