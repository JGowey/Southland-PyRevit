# -*- coding: utf-8 -*-
"""Data - Unified data import/export tool.

Consolidates Export Native Duct Data, Import Native Duct Data,
Export Family Data, and Import Family Data into a single branded
WPF dispatcher window.

Author: Jeremiah Griffith
Version: 1.0.0
"""

from __future__ import print_function
from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
import os, sys, json, csv, math, codecs, uuid, traceback

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
from System.Windows import Window, Visibility
from System.Windows.Markup import XamlReader
from System.IO import StreamReader
import System.Windows
import System.Windows.Input

uidoc = revit.uidoc
doc   = revit.doc
app   = doc.Application
OUT   = script.get_output()
log   = script.get_logger()

REVIT_VERSION = int(app.VersionNumber)
XAML_PATH = os.path.join(os.path.dirname(__file__), "window.xaml")

# Python 2/3 compat
PY3 = sys.version_info[0] >= 3
if PY3:
    string_types = (str,)
else:
    string_types = (str, unicode)


# ================================================================
#  ENGINE 1: EXPORT NATIVE DUCT DATA
# ================================================================

def _xyz_to_tuple_duct(xyz):
    return (round(xyz.X, 9), round(xyz.Y, 9), round(xyz.Z, 9))


def _iter_connectors_duct(elem):
    cm = getattr(elem, 'ConnectorManager', None)
    if cm:
        for c in cm.Connectors:
            yield c
        return
    mep = getattr(elem, 'MEPModel', None)
    if mep and getattr(mep, 'ConnectorManager', None):
        for c in mep.ConnectorManager.Connectors:
            yield c


def _two_ends(cons):
    cons = list(cons)
    n = len(cons)
    if n < 2:
        return (None, None, 0.0)
    best = -1.0
    pair = (cons[0], cons[1])
    for i in range(n):
        for j in range(i + 1, n):
            try:
                d = cons[i].Origin.DistanceTo(cons[j].Origin)
                if d > best:
                    best = d
                    pair = (cons[i], cons[j])
            except:
                pass
    return pair[0], pair[1], best


def _shape_dict(conn):
    d = {'shape': None, 'w': None, 'h': None, 'diameter': None,
         'direction': None, 'coord_system': None}
    try:
        s = conn.Shape
        if s == DB.ConnectorProfileType.Round:
            d['shape'] = 'Round'
            try:
                d['diameter'] = conn.Radius * 2.0
            except:
                pass
        elif s == DB.ConnectorProfileType.Rectangular:
            d['shape'] = 'Rectangular'
            try:
                d['w'] = conn.Width
                d['h'] = conn.Height
            except:
                pass
        elif s == DB.ConnectorProfileType.Oval:
            d['shape'] = 'Oval'
            try:
                d['w'] = conn.Width
                d['h'] = conn.Height
            except:
                pass
        else:
            d['shape'] = 'Unknown'
    except:
        d['shape'] = 'Unknown'

    try:
        coord_sys = conn.CoordinateSystem
        if coord_sys:
            direction_vec = coord_sys.BasisZ
            d['direction'] = _xyz_to_tuple_duct(direction_vec)
            d['coord_system'] = {
                'origin': _xyz_to_tuple_duct(coord_sys.Origin),
                'basis_x': _xyz_to_tuple_duct(coord_sys.BasisX),
                'basis_y': _xyz_to_tuple_duct(coord_sys.BasisY),
                'basis_z': _xyz_to_tuple_duct(coord_sys.BasisZ)
            }
    except:
        try:
            if hasattr(conn, 'Direction'):
                d['direction'] = _xyz_to_tuple_duct(conn.Direction)
        except:
            pass
    return d


def _elem_level_name(elem):
    try:
        if elem.LevelId and elem.LevelId != DB.ElementId.InvalidElementId:
            lvl = doc.GetElement(elem.LevelId)
            return lvl.Name if lvl else None
    except:
        pass
    try:
        av = doc.ActiveView
        return av.GenLevel.Name if av and av.GenLevel else None
    except:
        return None


def _system_name(elem):
    try:
        sys_obj = getattr(elem, 'MEPSystem', None)
        if sys_obj:
            return sys_obj.Name
    except:
        pass
    try:
        p = elem.LookupParameter('System Classification')
        if p:
            return p.AsString()
    except:
        pass
    return None


def _classify_fitting_export(elem, cons):
    n = len(list(cons))
    if n == 2:
        try:
            c0, c1, _ = _two_ends(_iter_connectors_duct(elem))
            s0 = _shape_dict(c0)
            s1 = _shape_dict(c1)
            if s0.get('shape') != s1.get('shape'):
                return 'Transition'
            if s0.get('diameter') and s1.get('diameter') and abs(
                    s0['diameter'] - s1['diameter']) > 1e-6:
                return 'Transition'
            if s0.get('w') and s1.get('w') and abs(
                    s0['w'] - s1['w']) > 1e-6:
                return 'Transition'
            if s0.get('h') and s1.get('h') and abs(
                    s0['h'] - s1['h']) > 1e-6:
                return 'Transition'
            if s0.get('direction') and s1.get('direction'):
                try:
                    from System import Math as SysMath
                    v0 = DB.XYZ(*s0['direction'])
                    v1 = DB.XYZ(*s1['direction'])
                    dot = v0.DotProduct(v1)
                    mag1 = v0.GetLength()
                    mag2 = v1.GetLength()
                    if mag1 > 0 and mag2 > 0:
                        cos_angle = max(-1, min(1, dot / (mag1 * mag2)))
                        angle_deg = SysMath.ToDegrees(SysMath.Acos(cos_angle))
                        if abs(angle_deg - 180) < 5:
                            return 'Coupling'
                        elif 85 < angle_deg < 95:
                            return 'Elbow90'
                        elif 40 < angle_deg < 50:
                            return 'Elbow45'
                        elif 20 < angle_deg < 26:
                            return 'Elbow22'
                        elif 8 < angle_deg < 16:
                            return 'Elbow11'
                        else:
                            return 'ElbowCustom'
                except:
                    pass
            return 'Coupling'
        except:
            return 'Fitting2'
    elif n == 3:
        return 'Tap'
    elif n >= 4:
        return 'Multiway'
    return 'Fitting'


def _export_record(elem):
    rec = {
        'element_id': elem.Id.IntegerValue,
        'category': elem.Category.Name if elem.Category else None,
        'class': elem.__class__.__name__,
        'family_name': (getattr(elem, 'Symbol', None).Family.Name
                        if hasattr(elem, 'Symbol') and elem.Symbol else None),
        'type_name': elem.Name if hasattr(elem, 'Name') else None,
        'level': _elem_level_name(elem),
        'system': _system_name(elem),
    }
    loc = getattr(elem, 'Location', None)
    if isinstance(loc, DB.LocationCurve):
        try:
            c = loc.Curve
            rec['curve_start'] = _xyz_to_tuple_duct(c.GetEndPoint(0))
            rec['curve_end'] = _xyz_to_tuple_duct(c.GetEndPoint(1))
            rec['curve_length_ft'] = c.Length
            start = c.GetEndPoint(0)
            end = c.GetEndPoint(1)
            direction = end - start
            if direction.GetLength() > 0:
                direction = direction.Normalize()
                rec['curve_direction'] = _xyz_to_tuple_duct(direction)
        except:
            pass

    cons = list(_iter_connectors_duct(elem))
    rec['connector_count'] = len(cons)
    c0, c1, span = _two_ends(cons)
    rec['span_ft'] = span
    if c0:
        s0 = _shape_dict(c0)
        s0.update({'origin': _xyz_to_tuple_duct(c0.Origin)})
        rec['conn_start'] = s0
    if c1:
        s1 = _shape_dict(c1)
        s1.update({'origin': _xyz_to_tuple_duct(c1.Origin)})
        rec['conn_end'] = s1
    if len(cons) > 2:
        for i, c in enumerate(cons[2:], start=2):
            s = _shape_dict(c)
            s.update({'origin': _xyz_to_tuple_duct(c.Origin)})
            rec['conn_{}'.format(i)] = s

    if isinstance(elem, DB.Mechanical.Duct):
        rec['kind'] = 'Duct'
        try:
            p = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            if p and p.AsDouble() > 0:
                rec['diameter_ft'] = p.AsDouble()
            else:
                pw = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                ph = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                if pw and ph:
                    rec['width_ft'] = pw.AsDouble()
                    rec['height_ft'] = ph.AsDouble()
        except:
            pass
    elif (elem.Category and
          elem.Category.Id == DB.ElementId(DB.BuiltInCategory.OST_DuctFitting)):
        rec['kind'] = 'DuctFitting'
        rec['fitting_class'] = _classify_fitting_export(elem, cons)
    else:
        rec['kind'] = 'Other'
    return rec


def run_export_native_duct():
    """Engine: Export Native Duct Data."""
    OUT.freeze()
    OUT.print_md('**Export Native Duct Data** (JSONL + CSV with Direction Data)')

    selection_options = [
        "Pick", "Box Select",
        "All In View (all elements in view)",
        "Visible In View (respects visibility/filters)",
        "Use Current Selection"
    ]
    selected_option = forms.CommandSwitchWindow.show(
        selection_options,
        message="Select fittings to export")
    if not selected_option:
        OUT.print_md('Export cancelled by user')
        OUT.unfreeze()
        return

    OUT.print_md('Selection method: {0}'.format(selected_option))
    elems = []

    try:
        if selected_option == "Pick":
            try:
                picked_refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    'Pick duct elements (multiple selection)')
                elems = [doc.GetElement(ref.ElementId) for ref in picked_refs]
            except:
                OUT.print_md('Pick selection cancelled')
                OUT.unfreeze()
                return

        elif selected_option == "Box Select":
            try:
                picked_elems = uidoc.Selection.PickElementsByRectangle(
                    'Drag selection box around duct elements')
                elems = list(picked_elems)
            except:
                OUT.print_md('Box selection cancelled')
                OUT.unfreeze()
                return

        elif selected_option == "All In View (all elements in view)":
            current_view = doc.ActiveView
            catids = List[DB.BuiltInCategory]()
            catids.Add(DB.BuiltInCategory.OST_DuctCurves)
            catids.Add(DB.BuiltInCategory.OST_DuctFitting)
            cat_filter = DB.ElementMulticategoryFilter(catids)
            collector = DB.FilteredElementCollector(doc, current_view.Id)
            elems = list(collector.WherePasses(cat_filter)
                         .WhereElementIsNotElementType())

        elif selected_option == "Visible In View (respects visibility/filters)":
            current_view = doc.ActiveView
            catids = List[DB.BuiltInCategory]()
            catids.Add(DB.BuiltInCategory.OST_DuctCurves)
            catids.Add(DB.BuiltInCategory.OST_DuctFitting)
            cat_filter = DB.ElementMulticategoryFilter(catids)
            vis_filter = DB.VisibleInViewFilter(doc, current_view.Id)
            collector = DB.FilteredElementCollector(doc, current_view.Id)
            elems = list(collector.WherePasses(cat_filter)
                         .WherePasses(vis_filter)
                         .WhereElementIsNotElementType())

        elif selected_option == "Use Current Selection":
            sel_ids = list(uidoc.Selection.GetElementIds())
            if sel_ids:
                elems = [doc.GetElement(i) for i in sel_ids]
            else:
                OUT.print_md('No elements currently selected')
                OUT.unfreeze()
                return
    except Exception as ex:
        OUT.print_md('Selection error: {0}'.format(str(ex)))
        OUT.unfreeze()
        return

    duct_elems = [e for e in elems
                  if isinstance(e, DB.Mechanical.Duct) or
                  (e.Category and e.Category.Id ==
                   DB.ElementId(DB.BuiltInCategory.OST_DuctFitting))]
    elems = duct_elems

    if not elems:
        forms.alert('No duct or duct fittings found in selection.',
                     title='Export Native Duct Data')
        OUT.unfreeze()
        return

    duct_count = sum(1 for e in elems if isinstance(e, DB.Mechanical.Duct))
    fitting_count = len(elems) - duct_count
    OUT.print_md('- Ducts: {0}, Fittings: {1}'.format(duct_count, fitting_count))

    if not forms.alert(
            'Export {0} elements ({1} ducts, {2} fittings)?'.format(
                len(elems), duct_count, fitting_count),
            title='Confirm Export', yes=True, no=True, ok=False):
        OUT.print_md('Export cancelled')
        OUT.unfreeze()
        return

    path = forms.save_file(file_ext='jsonl')
    if not path:
        OUT.unfreeze()
        return

    base = path[:-6] if path.lower().endswith('.jsonl') else path
    jsonl_path = base + '.jsonl'
    csv_path = base + '.csv'

    headers = [
        'element_id', 'kind', 'fitting_class', 'category', 'class',
        'family_name', 'type_name', 'level', 'system',
        'curve_start', 'curve_end', 'curve_length_ft', 'curve_direction',
        'connector_count', 'span_ft',
        'conn_start.shape', 'conn_start.w', 'conn_start.h',
        'conn_start.diameter', 'conn_start.origin', 'conn_start.direction',
        'conn_end.shape', 'conn_end.w', 'conn_end.h',
        'conn_end.diameter', 'conn_end.origin', 'conn_end.direction',
        'diameter_ft', 'width_ft', 'height_ft', 'angle_deg'
    ]

    count = 0
    with open(jsonl_path, 'w') as jf, open(csv_path, 'w') as cf:
        writer = csv.DictWriter(cf, fieldnames=headers)
        writer.writeheader()
        for e in elems:
            try:
                rec = _export_record(e)
                jf.write(json.dumps(rec) + '\n')
                row = dict((k, None) for k in headers)
                for k in ['element_id', 'category', 'class', 'family_name',
                           'type_name', 'level', 'system', 'kind',
                           'fitting_class', 'curve_length_ft',
                           'connector_count', 'span_ft', 'diameter_ft',
                           'width_ft', 'height_ft', 'angle_deg']:
                    if k in rec:
                        row[k] = rec[k]
                row['curve_start'] = rec.get('curve_start')
                row['curve_end'] = rec.get('curve_end')
                row['curve_direction'] = rec.get('curve_direction')
                if 'conn_start' in rec:
                    cs = rec['conn_start']
                    row['conn_start.shape'] = cs.get('shape')
                    row['conn_start.w'] = cs.get('w')
                    row['conn_start.h'] = cs.get('h')
                    row['conn_start.diameter'] = cs.get('diameter')
                    row['conn_start.origin'] = cs.get('origin')
                    row['conn_start.direction'] = cs.get('direction')
                if 'conn_end' in rec:
                    ce = rec['conn_end']
                    row['conn_end.shape'] = ce.get('shape')
                    row['conn_end.w'] = ce.get('w')
                    row['conn_end.h'] = ce.get('h')
                    row['conn_end.diameter'] = ce.get('diameter')
                    row['conn_end.origin'] = ce.get('origin')
                    row['conn_end.direction'] = ce.get('direction')
                writer.writerow(row)
                count += 1
            except Exception as ex:
                OUT.print_md(':x: skip {0}: {1}'.format(
                    e.Id.IntegerValue, ex))

    OUT.print_md('**Done** - Wrote {0} records'.format(count))
    OUT.print_md('- JSONL: `{0}`'.format(jsonl_path))
    OUT.print_md('- CSV: `{0}`'.format(csv_path))
    OUT.unfreeze()

# ================================================================
#  ENGINE 2: IMPORT NATIVE DUCT DATA (inlined)
# ================================================================

def _is_xyz(o): 
    return hasattr(o,'X') and hasattr(o,'Y') and hasattr(o,'Z')

def _to_tuple3(v):
    if v is None: return None
    if _is_xyz(v): return (float(v.X), float(v.Y), float(v.Z))
    if isinstance(v,(list,tuple)) and len(v)==3: return (float(v[0]),float(v[1]),float(v[2]))
    if isinstance(v,dict):
        if set(['X','Y','Z']).issubset(v): return (float(v['X']),float(v['Y']),float(v['Z']))
        if set(['x','y','z']).issubset(v): return (float(v['x']),float(v['y']),float(v['z']))
    return None

def _to_xyz(v):
    t = _to_tuple3(v)
    return DB.XYZ(t[0],t[1],t[2]) if t else None

def _safe_float(x):
    try: 
        return float(x)
    except: 
        return None

def _unit(x):
    try:
        m = (x.X*x.X + x.Y*x.Y + x.Z*x.Z)**0.5
        return DB.XYZ(1,0,0) if m==0 else DB.XYZ(x.X/m, x.Y/m, x.Z/m)
    except: 
        return DB.XYZ(1,0,0)

def _vec(p0,p1): 
    return DB.XYZ(p1[0]-p0[0], p1[1]-p0[1], p1[2]-p0[2])

def _angle_between_vectors(v1, v2):
    """Calculate angle in degrees between two vectors"""
    try:
        dot = v1.DotProduct(v2)
        mag1 = (v1.X*v1.X + v1.Y*v1.Y + v1.Z*v1.Z)**0.5
        mag2 = (v2.X*v2.X + v2.Y*v2.Y + v2.Z*v2.Z)**0.5
        if mag1 == 0 or mag2 == 0: 
            return 0
        cos_angle = dot / (mag1 * mag2)
        cos_angle = max(-1, min(1, cos_angle))  # Clamp to avoid numerical errors
        angle_rad = math.acos(cos_angle)
        return math.degrees(angle_rad)
    except:
        return 0

def _sizes_equal(s1, s2, tolerance=0.01):
    """Check if two size tuples are equal within tolerance"""
    if not s1 or not s2: 
        return False
    w1, h1, d1 = s1
    w2, h2, d2 = s2
    
    # Compare diameters for round
    if d1 and d2:
        return abs(d1 - d2) <= tolerance
    
    # Compare width/height for rectangular
    if w1 and h1 and w2 and h2:
        return abs(w1 - w2) <= tolerance and abs(h1 - h2) <= tolerance
    
    return False

def _is_straight_duct(rec):
    """Check if this record represents a straight duct (same size connectors perpendicular to centerline)"""
    if rec.get('kind') != 'Duct':
        return False
    
    # Must have curve start/end points
    p0 = _to_tuple3(rec.get('curve_start'))
    p1 = _to_tuple3(rec.get('curve_end'))
    if not (p0 and p1):
        return False
    
    # Get connector info if available
    cs = []
    for k in ('conn_start', 'conn_end'):
        v = rec.get(k)
        if isinstance(v, dict) and v.get('origin') is not None:
            cs.append(v)
    
    # If we have connector data, verify they're the same size and perpendicular
    if len(cs) == 2:
        # Check same sizes
        def _get_size_tuple(c):
            shp = (c.get('shape') or '').lower()
            if shp.startswith('rect'):
                return (_safe_float(c.get('w')), _safe_float(c.get('h')), None)
            else:
                return (None, None, _safe_float(c.get('diameter')))
        
        t0 = _get_size_tuple(cs[0])
        t1 = _get_size_tuple(cs[1])
        
        if not _sizes_equal(t0, t1):
            return False
        
        # Check perpendicular to centerline
        centerline = _unit(_to_xyz(_vec(p0, p1)))
        d0 = _to_xyz(_to_tuple3(cs[0].get('direction')))
        d1 = _to_xyz(_to_tuple3(cs[1].get('direction')))
        
        if d0 and d1:
            # Connectors should be roughly opposite and perpendicular to centerline
            angle_to_center_0 = abs(_angle_between_vectors(centerline, d0) - 90)
            angle_to_center_1 = abs(_angle_between_vectors(centerline, d1) - 90)
            opposite_angle = abs(_angle_between_vectors(d0, d1) - 180)
            
            if angle_to_center_0 < 15 and angle_to_center_1 < 15 and opposite_angle < 15:
                return True
    
    # Default: if no connector data, assume it's a straight
    return True

# ---------------- lookups ----------------
def _get_level(name):
    if not name: 
        return None
    for lvl in DB.FilteredElementCollector(doc).OfClass(DB.Level):
        if lvl.Name == name:
            return lvl
    return None

def _get_mep_system_type():
    if not hasattr(_get_mep_system_type,'c'):
        try:
            coll = DB.FilteredElementCollector(doc).OfClass(DB.Mechanical.MechanicalSystemType)
            _get_mep_system_type.c = list(coll)[0] if coll else None
        except: 
            _get_mep_system_type.c = None
    return _get_mep_system_type.c

def _duct_type_for_shape(shape_hint):
    """Return a DuctType whose family name matches the needed shape."""
    need_round = str(shape_hint or '').lower().startswith('round')
    need_rect = str(shape_hint or '').lower().startswith('rect')
    best = None
    round_dt = None
    rect_dt = None
    for dt in DB.FilteredElementCollector(doc).OfClass(DB.Mechanical.DuctType):
        fam = (getattr(dt,'FamilyName','') or '').lower()
        if 'oval' in fam: 
            continue
        if 'round' in fam and not round_dt: 
            round_dt = dt
        if ('rect' in fam or 'rectangular' in fam) and not rect_dt: 
            rect_dt = dt
        if best is None: 
            best = dt
    if need_round and round_dt: 
        return round_dt
    if need_rect and rect_dt:  
        return rect_dt
    return best

# ---------------- sizing ----------------
def _apply_size(elem, shape, w=None, h=None, d=None):
    try:
        if d:
            p = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            if p and not p.IsReadOnly: 
                p.Set(float(d))
                return
        if w and h:
            pw = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            ph = elem.get_Parameter(DB.BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            if pw and not pw.IsReadOnly: 
                pw.Set(float(w))
            if ph and not ph.IsReadOnly: 
                ph.Set(float(h))
    except: 
        pass

# ---------------- creation ----------------
def _create_duct_segment(p0, p1, level, sys_type, shape, size_tuple):
    dt = _duct_type_for_shape(shape)
    if not dt: 
        raise Exception('No duct type for shape {}'.format(shape))
    d = DB.Mechanical.Duct.Create(doc, sys_type.Id, dt.Id, level.Id, _to_xyz(p0), _to_xyz(p1))
    if not d: 
        raise Exception('Duct.Create returned None')
    # size_tuple: (w,h,d) where any may be None
    w, h, dia = size_tuple
    _apply_size(d, shape, w=w, h=h, d=dia)
    return d

def _closest_connector(elem, ref_xyz):
    try:
        cm = getattr(elem,'ConnectorManager',None)
        if cm:
            best = None
            bd = 1e9
            for c in cm.Connectors:
                d = c.Origin.DistanceTo(ref_xyz)
                if d < bd: 
                    best = c
                    bd = d
            return best
        mep = getattr(elem,'MEPModel',None)
        if mep and getattr(mep,'ConnectorManager',None):
            best = None
            bd = 1e9
            for c in mep.ConnectorManager.Connectors:
                d = c.Origin.DistanceTo(ref_xyz)
                if d < bd: 
                    best = c
                    bd = d
            return best
    except: 
        pass
    return None

def _nd_iter_connectors(elem):
    """Helper to iterate over all connectors of an element"""
    try:
        cm = getattr(elem,'ConnectorManager',None)
        if cm:
            for c in cm.Connectors:
                yield c
            return
        mep = getattr(elem,'MEPModel',None)
        if mep and getattr(mep,'ConnectorManager',None):
            for c in mep.ConnectorManager.Connectors:
                yield c
    except: 
        pass

def _try_set_length(elem, length_ft):
    """Enhanced length setting with multiple parameter attempts"""
    if length_ft is None: 
        return False
    
    success = False
    attempts = []
    
    try:
        # Try multiple common length parameter names
        length_params = [
            'Length',
            'length', 
            'Overall Length',
            'Total Length',
            'Transition Length',
            'Distance',
            'Span'
        ]
        
        for param_name in length_params:
            try:
                param = elem.LookupParameter(param_name)
                if param and not param.IsReadOnly:
                    old_val = param.AsDouble() if param.HasValue else 0
                    param.Set(float(length_ft))
                    new_val = param.AsDouble()
                    attempts.append('{0}: {1:.3f} -> {2:.3f}'.format(param_name, old_val, new_val))
                    if abs(new_val - length_ft) < 0.01:
                        success = True
                        OUT.print_md('SUCCESS: Set {0} parameter to {1:.3f} ft'.format(param_name, new_val))
                    else:
                        OUT.print_md('PARTIAL: {0} parameter changed but not to target value'.format(param_name))
            except Exception as param_ex:
                attempts.append('{0}: FAILED ({1})'.format(param_name, str(param_ex)))
        
        # Try built-in parameters
        builtin_params = [
            DB.BuiltInParameter.CURVE_ELEM_LENGTH,
            DB.BuiltInParameter.INSTANCE_LENGTH_PARAM
        ]
        
        for bip in builtin_params:
            try:
                param = elem.get_Parameter(bip)
                if param and not param.IsReadOnly:
                    old_val = param.AsDouble() if param.HasValue else 0
                    param.Set(float(length_ft))
                    new_val = param.AsDouble()
                    param_name = 'BuiltIn_{0}'.format(str(bip).split('.')[-1])
                    attempts.append('{0}: {1:.3f} -> {2:.3f}'.format(param_name, old_val, new_val))
                    if abs(new_val - length_ft) < 0.01:
                        success = True
                        OUT.print_md('SUCCESS: Set built-in parameter to {0:.3f} ft'.format(new_val))
                    else:
                        OUT.print_md('PARTIAL: Built-in parameter changed but not to target value')
            except Exception as bip_ex:
                attempts.append('BuiltIn_{0}: FAILED ({1})'.format(str(bip).split('.')[-1], str(bip_ex)))
        
        # Log all attempts for debugging
        if attempts:
            OUT.print_md('Length parameter attempts: {0}'.format('; '.join(attempts)))
        
    except Exception as ex:
        OUT.print_md('Length parameter setting failed: {0}'.format(ex))
    
    return success

# ---------------- fitting classification ----------------
def _classify_fitting(rec):
    """Classify fitting type based on connector data"""
    # Get connectors
    cs = []
    for k in ('conn_start','conn_end'):
        v = rec.get(k)
        if isinstance(v,dict) and v.get('origin') is not None:
            cs.append(v)
    
    # Add other connectors if available
    if len(cs) < 2:
        for k, v in rec.items():
            if k.startswith('conn_') and isinstance(v,dict) and v.get('origin') is not None and v not in cs:
                cs.append(v)
    
    if len(cs) != 2:
        return 'Unknown', None
    
    # Get shapes and sizes
    def _get_size_shape(c):
        shp = (c.get('shape') or '').capitalize()
        if shp.startswith('Rect'):
            return ('Rectangular', (_safe_float(c.get('w')), _safe_float(c.get('h')), None))
        else:
            return ('Round', (None, None, _safe_float(c.get('diameter'))))
    
    s0, t0 = _get_size_shape(cs[0])
    s1, t1 = _get_size_shape(cs[1])
    
    # DEBUG: Log connector details for troubleshooting
    elem_id = rec.get('element_id', 'n/a')
    OUT.print_md('DEBUG {0}: Conn0 {1} {2}, Conn1 {3} {4}'.format(elem_id, s0, t0, s1, t1))
    
    # Check if it's a transition (different shapes or sizes)
    same_shapes = s0 == s1
    same_sizes = _sizes_equal(t0, t1)
    OUT.print_md('DEBUG {0}: Same shapes: {1}, Same sizes: {2}'.format(elem_id, same_shapes, same_sizes))
    
    if not same_shapes or not same_sizes:
        return 'Transition', (cs, s0, s1, t0, t1)
    
    # Same shape and size - check angle if we have direction data
    p0 = _to_tuple3(cs[0].get('origin'))
    p1 = _to_tuple3(cs[1].get('origin'))
    d0 = _to_tuple3(cs[0].get('direction'))
    d1 = _to_tuple3(cs[1].get('direction'))
    
    if not all([p0, p1, d0, d1]):
        OUT.print_md('DEBUG {0}: Missing position/direction data - assuming ElbowGeneric'.format(elem_id))
        # If we have same shapes/sizes but missing direction data, assume it's an elbow
        # Revit's routing preferences will determine the appropriate angle
        return 'ElbowGeneric', (cs, s0, s1, t0, t1)
    
    # Calculate angle between connector directions
    v0 = _to_xyz(d0)
    v1 = _to_xyz(d1)
    angle = _angle_between_vectors(v0, v1)
    
    OUT.print_md('DEBUG {0}: Angle between connectors: {1:.1f}°'.format(elem_id, angle))
    
    # Classify based on angle (focusing on the 11.25 degree angle you mentioned)
    if abs(angle - 180) < 5:  # Straight-through
        return 'Coupling', (cs, s0, s1, t0, t1)
    elif 85 < angle < 95:  # 90 degree elbow
        return 'Elbow90', (cs, s0, s1, t0, t1)
    elif 40 < angle < 50:   # 45 degree elbow
        return 'Elbow45', (cs, s0, s1, t0, t1)
    elif 20 < angle < 26:   # 22.5 degree elbow (11.25 * 2)
        return 'Elbow22', (cs, s0, s1, t0, t1)
    elif 8 < angle < 16:    # 11.25 degree elbow (your specific requirement)
        return 'Elbow11', (cs, s0, s1, t0, t1)
    else:
        return 'ElbowCustom', (cs, s0, s1, t0, t1)

# ---------------- elbow creation ----------------
def _make_elbow(rec, elbow_type, fitting_data):
    """Create elbow fitting with correct orientation"""
    cs, s0, s1, t0, t1 = fitting_data
    
    # Global context
    level = _get_level(rec.get('level'))
    if not level:
        lvls = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
        level = lvls[0] if lvls else None
    st = _get_mep_system_type()
    if not (level and st):
        return None, 'Missing Level or System Type'
    
    # Get connector positions and directions
    p0 = _to_tuple3(cs[0].get('origin'))
    p1 = _to_tuple3(cs[1].get('origin'))
    d0_data = _to_tuple3(cs[0].get('direction'))
    d1_data = _to_tuple3(cs[1].get('direction'))
    
    if not all([p0, p1, d0_data, d1_data]):
        return None, 'Missing connector position or direction data'
    
    # CRITICAL FIX: Use INWARD directions for stub creation
    # The exported directions are outward from the fitting
    # For proper elbow creation, we need stubs pointing INWARD toward the elbow center
    d0_inward = _to_xyz((-d0_data[0], -d0_data[1], -d0_data[2]))  # Reverse direction
    d1_inward = _to_xyz((-d1_data[0], -d1_data[1], -d1_data[2]))  # Reverse direction
    
    # Calculate angle for logging  
    d0_outward = _to_xyz(d0_data)
    d1_outward = _to_xyz(d1_data)
    angle = _angle_between_vectors(d0_outward, d1_outward)
    OUT.print_md('Elbow {0}: Outward directions angle {1:.1f}°'.format(
        rec.get('element_id', 'n/a'), angle))
    
    # Calculate distance and appropriate stub length
    connector_distance = _to_xyz(p0).DistanceTo(_to_xyz(p1))
    
    # Use conservative stub lengths
    if s0 == 'Rectangular':
        stub_len = min(0.2, connector_distance * 0.15)
        stub_len = max(0.08, stub_len)  # 1 inch minimum
    else:
        stub_len = min(0.3, connector_distance * 0.2)  
        stub_len = max(0.1, stub_len)  # 1.2 inch minimum
    
    OUT.print_md('Using stub length: {0:.3f} ft for {1}'.format(stub_len, s0))
    
    # Create stubs using INWARD directions (pointing toward elbow center)
    # This should create proper orientation for elbow connection
    
    p0a = p0  # Connector position
    p0b = (p0[0] + d0_inward.X * stub_len, 
           p0[1] + d0_inward.Y * stub_len, 
           p0[2] + d0_inward.Z * stub_len)  # Stub end (inward)
    
    p1a = p1  # Connector position
    p1b = (p1[0] + d1_inward.X * stub_len,
           p1[1] + d1_inward.Y * stub_len, 
           p1[2] + d1_inward.Z * stub_len)  # Stub end (inward)
    
    # Track stub IDs for cleanup
    stub_ids_to_delete = []
    
    try:
        # Create stub ducts - FROM connector position TO inward stub end
        d0_stub = _create_duct_segment(p0a, p0b, level, st, s0, t0)
        d1_stub = _create_duct_segment(p1a, p1b, level, st, s1, t1)
        
        # Store stub IDs for cleanup after elbow creation
        stub_ids_to_delete.extend([d0_stub.Id, d1_stub.Id])
        
        # Get the connectors at the INWARD ends (where elbow will connect)
        c0_inner = _closest_connector(d0_stub, _to_xyz(p0b))  # Inward end
        c1_inner = _closest_connector(d1_stub, _to_xyz(p1b))  # Inward end
        
        if not c0_inner or not c1_inner:
            return None, 'Could not find stub connectors'
        
        # Check the angle between inward-pointing stub connectors
        try:
            # Get the flow directions of the stub connectors
            c0_flow = c0_inner.CoordinateSystem.BasisZ
            c1_flow = c1_inner.CoordinateSystem.BasisZ
            
            # Calculate angle between stub connector flows
            connector_angle = _angle_between_vectors(c0_flow, c1_flow)
            OUT.print_md('Stub connector flow angle: {0:.1f}°'.format(connector_angle))
            
            # For proper elbow creation, stub flows should be angled toward each other
            # If they're parallel (0°) or opposite (180°), we may have orientation issues
            if connector_angle < 15:
                OUT.print_md('WARNING: Stub connectors nearly parallel - may affect elbow orientation')
            elif connector_angle > 165:
                OUT.print_md('WARNING: Stub connectors opposite - may affect elbow orientation')
                
        except Exception as ex:
            OUT.print_md('Could not calculate stub connector angles: {0}'.format(ex))
        
        # Try to create the elbow
        try:
            elbow = doc.Create.NewElbowFitting(c0_inner, c1_inner)
        except Exception as create_ex:
            # If creation fails, try swapping the connectors
            try:
                OUT.print_md('Retrying with swapped connectors')
                elbow = doc.Create.NewElbowFitting(c1_inner, c0_inner)
            except Exception as swap_ex:
                return None, 'Elbow creation failed: {0} (swap: {1})'.format(str(create_ex), str(swap_ex))
        
        if elbow:
            # CRITICAL: Delete the stub ducts after successful elbow creation
            try:
                for stub_id in stub_ids_to_delete:
                    doc.Delete(stub_id)
                OUT.print_md('SUCCESS: Created elbow {0} and cleaned up stubs'.format(rec.get('element_id', 'n/a')))
            except Exception as cleanup_ex:
                OUT.print_md('WARNING: Elbow created but stub cleanup failed: {0}'.format(cleanup_ex))
            
            return elbow, None
        else:
            return None, 'Revit returned None for elbow creation'
            
    except Exception as ex:
        # Clean up any stubs that were created before the error
        try:
            for stub_id in stub_ids_to_delete:
                doc.Delete(stub_id)
        except:
            pass
        return None, 'Elbow creation process failed: {0}'.format(str(ex))

# ---------------- transitions ----------------
def _make_transition(rec, fitting_data):
    cs, s0, s1, t0, t1 = fitting_data
    
    # Global context
    level = _get_level(rec.get('level'))
    if not level:
        lvls = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
        level = lvls[0] if lvls else None
    st = _get_mep_system_type()
    if not (level and st):
        return None, 'Missing Level or System Type'
    
    # Enhanced debugging for transitions
    elem_id = rec.get('element_id', 'n/a')
    OUT.print_md('=== TRANSITION DEBUG {0} ==='.format(elem_id))
    OUT.print_md('From: {0} {1} -> To: {2} {3}'.format(s0, t0, s1, t1))
    
    # Get connector positions and directions
    p0 = _to_tuple3(cs[0].get('origin'))
    p1 = _to_tuple3(cs[1].get('origin'))
    d0_data = _to_tuple3(cs[0].get('direction'))
    d1_data = _to_tuple3(cs[1].get('direction'))
    
    OUT.print_md('Connector positions: P0={0}, P1={1}'.format(p0, p1))
    OUT.print_md('Connector directions: D0={0}, D1={1}'.format(d0_data, d1_data))
    
    if not (p0 and p1):
        return None, 'Connector origins missing'
    
    # Calculate axis and span
    connector_distance = _to_xyz(p0).DistanceTo(_to_xyz(p1))
    axis = _unit(_to_xyz(_vec(p0, p1)))
    axisn = DB.XYZ(-axis.X, -axis.Y, -axis.Z)
    
    # Desired length from export (span between connector origins)
    target_span = _safe_float(rec.get('span_ft')) or connector_distance
    
    OUT.print_md('Connector distance: {0:.3f} ft, Target span: {1:.3f} ft'.format(connector_distance, target_span))
    OUT.print_md('Axis vector: ({0:.3f}, {1:.3f}, {2:.3f})'.format(axis.X, axis.Y, axis.Z))
    
    # Enhanced stub calculation with debugging
    base_stub_len = max(0.25, min(2.0, target_span * 0.2))
    OUT.print_md('Base stub length: {0:.3f} ft (20% of span, clamped 0.25-2.0)'.format(base_stub_len))
    
    # Track stub IDs for cleanup
    stub_ids_to_delete = []
    
    # Enhanced stub orientation logic with validation
    if d0_data and d1_data:
        OUT.print_md('Using connector directions for stub orientation')
        
        # Calculate the angle between connector directions to validate alignment
        d0_vec = _to_xyz(d0_data)
        d1_vec = _to_xyz(d1_data)
        connector_angle = _angle_between_vectors(d0_vec, d1_vec)
        OUT.print_md('Connector direction angle: {0:.1f}°'.format(connector_angle))
        
        # Check if connectors are roughly opposing each other (good for transitions)
        if connector_angle > 150:  # Nearly opposite directions
            OUT.print_md('Connectors are opposing - using direction-based stubs')
            # Use inward-pointing directions
            d0_inward = _to_xyz((-d0_data[0], -d0_data[1], -d0_data[2]))
            d1_inward = _to_xyz((-d1_data[0], -d1_data[1], -d1_data[2]))
            
            p0a = p0  # Start at connector origin
            p0b = (p0[0] + d0_inward.X * base_stub_len, 
                   p0[1] + d0_inward.Y * base_stub_len, 
                   p0[2] + d0_inward.Z * base_stub_len)  # End inward
            p1a = p1  # Start at connector origin
            p1b = (p1[0] + d1_inward.X * base_stub_len,
                   p1[1] + d1_inward.Y * base_stub_len, 
                   p1[2] + d1_inward.Z * base_stub_len)  # End inward
        else:
            OUT.print_md('Connectors not opposing (angle {0:.1f}°) - using axis-based approach'.format(connector_angle))
            # Fall back to axis-based for better alignment
            p0a = p0  # Start at connector origin
            p0b = (p0[0] + axis.X * base_stub_len, p0[1] + axis.Y * base_stub_len, p0[2] + axis.Z * base_stub_len)
            p1a = p1  # Start at connector origin  
            p1b = (p1[0] + axisn.X * base_stub_len, p1[1] + axisn.Y * base_stub_len, p1[2] + axisn.Z * base_stub_len)
    else:
        OUT.print_md('No direction data - using axis-based stub orientation')
        # Fallback to axis-based approach
        p0a = p0  # Start at connector origin
        p0b = (p0[0] + axis.X * base_stub_len, p0[1] + axis.Y * base_stub_len, p0[2] + axis.Z * base_stub_len)
        p1a = p1  # Start at connector origin  
        p1b = (p1[0] + axisn.X * base_stub_len, p1[1] + axisn.Y * base_stub_len, p1[2] + axisn.Z * base_stub_len)
    
    # Validate stub positioning
    stub_distance = _to_xyz(p0b).DistanceTo(_to_xyz(p1b))
    OUT.print_md('Stub separation: {0:.3f} ft (original connector distance: {1:.3f} ft)'.format(
        stub_distance, connector_distance))
    
    OUT.print_md('Stub endpoints: P0B={0}, P1B={1}'.format(
        (round(p0b[0], 3), round(p0b[1], 3), round(p0b[2], 3)),
        (round(p1b[0], 3), round(p1b[1], 3), round(p1b[2], 3))))
    
    try:
        # Create stub ducts
        OUT.print_md('Creating stub ducts...')
        d0 = _create_duct_segment(p0a, p0b, level, st, s0, t0)
        d1 = _create_duct_segment(p1a, p1b, level, st, s1, t1)
        
        # Store stub IDs for cleanup after transition creation
        stub_ids_to_delete.extend([d0.Id, d1.Id])
        
        OUT.print_md('Stub ducts created successfully')
        
        # Connect inner ends with a transition fitting
        c0 = _closest_connector(d0, _to_xyz(p0b))
        c1 = _closest_connector(d1, _to_xyz(p1b))
        
        OUT.print_md('Found connectors: C0={0}, C1={1}'.format(c0 is not None, c1 is not None))
        
        # CRITICAL FIX: If connectors not found at expected positions, try all connectors
        if not c0:
            OUT.print_md('Trying all connectors for d0 stub...')
            for i, conn in enumerate(_nd_iter_connectors(d0)):
                pos = _nd_xyz_to_tuple(conn.Origin)
                OUT.print_md('  d0 connector {0}: {1}'.format(i, pos))
                if not c0:  # Take first available
                    c0 = conn
                    OUT.print_md('  -> Using d0 connector {0}'.format(i))
        
        if not c1:
            OUT.print_md('Trying all connectors for d1 stub...')
            for i, conn in enumerate(_nd_iter_connectors(d1)):
                pos = _nd_xyz_to_tuple(conn.Origin)
                OUT.print_md('  d1 connector {0}: {1}'.format(i, pos))
                if not c1:  # Take first available
                    c1 = conn
                    OUT.print_md('  -> Using d1 connector {0}'.format(i))
        
        if not c0 or not c1:
            return None, 'Could not find transition connectors after exhaustive search'
        
        # Check connector orientations
        try:
            c0_origin = _nd_xyz_to_tuple(c0.Origin) if hasattr(c0, 'Origin') else 'Unknown'
            c1_origin = _nd_xyz_to_tuple(c1.Origin) if hasattr(c1, 'Origin') else 'Unknown'
            c0_dir = _nd_xyz_to_tuple(c0.CoordinateSystem.BasisZ) if hasattr(c0, 'CoordinateSystem') else 'Unknown'
            c1_dir = _nd_xyz_to_tuple(c1.CoordinateSystem.BasisZ) if hasattr(c1, 'CoordinateSystem') else 'Unknown'
            
            OUT.print_md('C0 origin: {0}, direction: {1}'.format(c0_origin, c0_dir))
            OUT.print_md('C1 origin: {0}, direction: {1}'.format(c1_origin, c1_dir))
            
            if hasattr(c0, 'CoordinateSystem') and hasattr(c1, 'CoordinateSystem'):
                connector_angle = _angle_between_vectors(
                    c0.CoordinateSystem.BasisZ,
                    c1.CoordinateSystem.BasisZ
                )
                OUT.print_md('Angle between stub connectors: {0:.1f}°'.format(connector_angle))
        except Exception as debug_ex:
            OUT.print_md('Could not analyze connector orientations: {0}'.format(debug_ex))
        
        # Attempt transition creation
        OUT.print_md('Attempting transition creation...')
        trans = doc.Create.NewTransitionFitting(c0, c1)
        
        if trans is None:
            # Retry with nudge adjustment
            OUT.print_md('First attempt failed - trying with position nudge')
            try:
                nudge = 0.10
                DB.ElementTransformUtils.MoveElement(doc, d0.Id, DB.XYZ(0, 0, -nudge))
                DB.ElementTransformUtils.MoveElement(doc, d1.Id, DB.XYZ(0, 0, nudge))
                OUT.print_md('Applied nudge: d0 down {0} ft, d1 up {0} ft'.format(nudge))
                trans = doc.Create.NewTransitionFitting(c0, c1)
            except Exception as nudge_ex:
                OUT.print_md('Nudge attempt failed: {0}'.format(nudge_ex))
            
            if trans is None:
                # Try swapping connectors
                OUT.print_md('Nudge failed - trying with swapped connectors')
                try:
                    trans = doc.Create.NewTransitionFitting(c1, c0)
                    if trans:
                        OUT.print_md('SUCCESS with swapped connectors')
                except Exception as swap_ex:
                    OUT.print_md('Connector swap failed: {0}'.format(swap_ex))
                    return None, 'All transition creation attempts failed'
        
        if trans is None:
            return None, 'Revit returned None (check Routing Preferences for {0} to {1})'.format(s0, s1)
        
        OUT.print_md('Transition created successfully')
        
        # CRITICAL: Delete the stub ducts after successful transition creation
        try:
            for stub_id in stub_ids_to_delete:
                doc.Delete(stub_id)
            OUT.print_md('SUCCESS: Cleaned up stub ducts after transition creation')
        except Exception as cleanup_ex:
            OUT.print_md('WARNING: Transition created but stub cleanup failed: {0}'.format(cleanup_ex))
        
        # Analyze the created transition
        try:
            tcm = getattr(trans.MEPModel, 'ConnectorManager', None) if hasattr(trans, 'MEPModel') else None
            if tcm:
                trans_connectors = list(tcm.Connectors)
                OUT.print_md('Transition has {0} connectors'.format(len(trans_connectors)))
                if len(trans_connectors) >= 2:
                    actual_length = trans_connectors[0].Origin.DistanceTo(trans_connectors[1].Origin)
                    OUT.print_md('Actual transition length: {0:.3f} ft'.format(actual_length))
                    
                    # Try to adjust length
                    length_diff = target_span - actual_length
                    OUT.print_md('Length difference: {0:.3f} ft (target - actual)'.format(length_diff))
                    
                    # Enhanced length parameter setting
                    length_set = _try_set_length(trans, target_span)
                    OUT.print_md('Length parameter setting result: {0}'.format(length_set))
                    
                    # Check if parameter setting worked
                    if length_set:
                        try:
                            # Re-measure after parameter change
                            updated_length = trans_connectors[0].Origin.DistanceTo(trans_connectors[1].Origin)
                            OUT.print_md('Length after parameter adjustment: {0:.3f} ft'.format(updated_length))
                            length_diff = target_span - updated_length
                        except:
                            pass
                    
                    # Note: Geometric adjustment removed since stubs are deleted
                    # The transition should be correctly positioned with the stub cleanup
                    if abs(length_diff) > 0.05:  # 0.6 inch tolerance
                        OUT.print_md('NOTE: Final length differs by {0:.3f} ft - may need manual adjustment or check Routing Preferences'.format(length_diff))
                    else:
                        OUT.print_md('SUCCESS: Transition length within tolerance')
                else:
                    OUT.print_md('Could not access transition connectors for length analysis')
            else:
                OUT.print_md('Could not access transition connector manager')
                
        except Exception as analysis_ex:
            OUT.print_md('Transition analysis failed: {0}'.format(analysis_ex))
        
        OUT.print_md('=== END TRANSITION DEBUG {0} ==='.format(elem_id))
        return trans, None
    
    except Exception as ex:
        # Clean up any stubs that were created before the error
        try:
            for stub_id in stub_ids_to_delete:
                doc.Delete(stub_id)
            OUT.print_md('Cleaned up stubs after transition creation failure')
        except:
            pass
        OUT.print_md('=== TRANSITION FAILED {0}: {1} ==='.format(elem_id, str(ex)))
        return None, 'Transition creation failed: {0}'.format(str(ex))

def _nd_xyz_to_tuple(xyz):
    """Helper function to convert XYZ to tuple for debugging"""
    try:
        return (round(xyz.X, 3), round(xyz.Y, 3), round(xyz.Z, 3))
    except:
        return 'Invalid XYZ'

# ---------------- connection optimization ----------------
def _optimize_straight_connections():
    """Optimize straight duct connections to fittings by adjusting lengths"""
    optimized_count = 0
    
    # Get all ducts and fittings in the model
    duct_collector = DB.FilteredElementCollector(doc).OfClass(DB.Mechanical.Duct)
    fitting_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_DuctFitting)
    
    all_ducts = list(duct_collector)
    all_fittings = list(fitting_collector)
    
    OUT.print_md('Found {} ducts and {} fittings for connection analysis'.format(len(all_ducts), len(all_fittings)))
    
    for duct in all_ducts:
        try:
            # Get duct endpoints
            loc = getattr(duct, 'Location', None)
            if not isinstance(loc, DB.LocationCurve):
                continue
            
            curve = loc.Curve
            start_pt = curve.GetEndPoint(0)
            end_pt = curve.GetEndPoint(1)
            
            # Check for nearby fittings at each end
            start_fitting = _find_nearby_fitting(start_pt, all_fittings)
            end_fitting = _find_nearby_fitting(end_pt, all_fittings)
            
            # Only modify if connecting to fittings (not duct-to-duct)
            connection_changes = []
            
            if start_fitting:
                # Find the closest connector on the fitting
                start_target = _find_closest_fitting_connector(start_fitting, start_pt)
                if start_target:
                    connection_changes.append(('start', start_target))
            
            if end_fitting:
                # Find the closest connector on the fitting  
                end_target = _find_closest_fitting_connector(end_fitting, end_pt)
                if end_target:
                    connection_changes.append(('end', end_target))
            
            # Apply connection changes if any fittings found
            if connection_changes:
                success = _adjust_duct_to_fittings(duct, connection_changes)
                if success:
                    optimized_count += 1
                    OUT.print_md('Optimized duct {} connections'.format(duct.Id.IntegerValue))
                    
        except Exception as ex:
            OUT.print_md('Error optimizing duct {}: {}'.format(duct.Id.IntegerValue, ex))
            continue
    
    return optimized_count

def _find_nearby_fitting(point, fittings, tolerance=2.0):
    """Find fitting within tolerance distance of point"""
    closest_fitting = None
    closest_distance = tolerance
    
    for fitting in fittings:
        try:
            # Get fitting connectors to find closest point
            fitting_connectors = list(_nd_iter_connectors(fitting))
            if not fitting_connectors:
                continue
            
            for connector in fitting_connectors:
                distance = connector.Origin.DistanceTo(point)
                if distance < closest_distance:
                    closest_distance = distance
                    closest_fitting = fitting
        except:
            continue
    
    return closest_fitting

def _find_closest_fitting_connector(fitting, target_point):
    """Find the closest connector on a fitting to the target point"""
    closest_connector = None
    closest_distance = 1e9
    
    try:
        for connector in _nd_iter_connectors(fitting):
            distance = connector.Origin.DistanceTo(target_point)
            if distance < closest_distance:
                closest_distance = distance
                closest_connector = connector
    except:
        pass
    
    return closest_connector

def _adjust_duct_to_fittings(duct, connection_changes):
    """Adjust duct length to connect properly to fittings"""
    try:
        loc = duct.Location
        if not isinstance(loc, DB.LocationCurve):
            return False
        
        curve = loc.Curve
        original_start = curve.GetEndPoint(0)
        original_end = curve.GetEndPoint(1)
        
        new_start = original_start
        new_end = original_end
        
        # Apply connection changes
        for end_type, target_connector in connection_changes:
            target_point = target_connector.Origin
            
            if end_type == 'start':
                new_start = target_point
                OUT.print_md('  Adjusting start point: {:.3f} ft movement'.format(original_start.DistanceTo(target_point)))
            elif end_type == 'end':
                new_end = target_point
                OUT.print_md('  Adjusting end point: {:.3f} ft movement'.format(original_end.DistanceTo(target_point)))
        
        # Check if adjustment is reasonable (not too extreme)
        original_length = curve.Length
        new_length = new_start.DistanceTo(new_end)
        length_change = abs(new_length - original_length)
        
        # Don't make extreme length changes (more than 50% change or more than 5 feet)
        if length_change > original_length * 0.5 or length_change > 5.0:
            OUT.print_md('  Skipping extreme length change: {:.3f} ft to {:.3f} ft'.format(original_length, new_length))
            return False
        
        # Create new curve with adjusted endpoints
        new_curve = DB.Line.CreateBound(new_start, new_end)
        
        # Update the duct location
        loc.Curve = new_curve
        
        OUT.print_md('  Length adjusted: {:.3f} ft -> {:.3f} ft ({:+.3f} ft)'.format(
            original_length, new_length, new_length - original_length))
        
        return True
        
    except Exception as ex:
        OUT.print_md('  Failed to adjust duct: {}'.format(ex))
        return False

# ---------------- straights ----------------
def _create_duct_from_record(rec):
    try:
        p0 = _to_xyz(rec.get('curve_start'))
        p1 = _to_xyz(rec.get('curve_end'))
        if not p0 or not p1:
            return None, 'Missing curve points'
        
        lvl = _get_level(rec.get('level'))
        if not lvl:
            lvls = list(DB.FilteredElementCollector(doc).OfClass(DB.Level))
            lvl = lvls[0] if lvls else None
        
        st = _get_mep_system_type()
        if not (lvl and st):
            return None, 'Missing level/system'
        
        # Choose shape by provided dimensions
        shp = 'Round' if _safe_float(rec.get('diameter_ft')) else 'Rectangular'
        d = _create_duct_segment(_to_tuple3(p0), _to_tuple3(p1), lvl, st, shp,
                                (_safe_float(rec.get('width_ft')), _safe_float(rec.get('height_ft')),
                                 _safe_float(rec.get('diameter_ft'))))
        return d, None
    except Exception as ex:
        return None, str(ex)

# ---------------- file IO ----------------
def _read_jsonl(path):
    out = []
    with open(path, 'r') as f:
        for ln in f:
            ln = ln.strip()
            if ln:
                out.append(json.loads(ln))
    return out

def _read_csv(path):
    out = []
    with open(path, 'r') as f:
        rd = csv.DictReader(f)
        for row in rd:
            r = {}
            for k, v in row.items():
                if v in ('', 'None'):
                    r[k] = None
                elif k in ('element_id', 'connector_count'):
                    try:
                        r[k] = int(v)
                    except:
                        r[k] = v
                elif k in ('curve_length_ft', 'span_ft', 'diameter_ft', 'width_ft', 'height_ft', 'angle_deg'):
                    try:
                        r[k] = float(v)
                    except:
                        r[k] = v
                elif k in ('curve_start', 'curve_end'):
                    try:
                        if v and v.startswith('(') and v.endswith(')'):
                            a = v[1:-1].split(',')
                            r[k] = (float(a[0]), float(a[1]), float(a[2]))
                        else:
                            r[k] = None
                    except:
                        r[k] = None
                else:
                    r[k] = v
            out.append(r)
    return out

# ---------------- main ----------------
def run_import_native_duct():
    OUT.freeze()
    OUT.print_md('**Enhanced Native Duct Import - Straights + Elbows + Transitions**')
    OUT.print_md('*Respects Routing Preferences for fitting selection*')
    
    f = forms.pick_file()
    if not f:
        OUT.unfreeze()
        return
    OUT.print_md('Reading file: `{}`'.format(f))
    
    if f.lower().endswith('.jsonl'):
        recs = _read_jsonl(f)
    elif f.lower().endswith('.csv'):
        recs = _read_csv(f)
    else:
        forms.alert('Pick a .jsonl or .csv', title='Importer')
        OUT.unfreeze()
        return
    
    if not recs:
        forms.alert('No records found', title='Importer')
        OUT.unfreeze()
        return
    
    OUT.print_md('Found {} records to import'.format(len(recs)))
    
    # Analyze the records
    straight_recs = [r for r in recs if _is_straight_duct(r)]
    fitting_recs = [r for r in recs if r.get('kind') == 'DuctFitting']
    
    # Count fittings by type
    elbow_count = 0
    transition_count = 0
    for rec in fitting_recs:
        try:
            fitting_type, fitting_data = _classify_fitting(rec)
            if fitting_type.startswith('Elbow') or fitting_type == 'ElbowGeneric':
                elbow_count += 1
            elif fitting_type == 'Transition':
                transition_count += 1
        except:
            pass
    
    OUT.print_md('- Straight ducts: {}'.format(len(straight_recs)))
    OUT.print_md('- Elbows: {}'.format(elbow_count))
    OUT.print_md('- Transitions: {}'.format(transition_count))
    
    # Step 1: Element Type Selection using button interface with multi-selection capability
    available_types = []
    if len(straight_recs) > 0:
        available_types.append(('Straights', len(straight_recs)))
    if elbow_count > 0:
        available_types.append(('Elbows', elbow_count))
    if transition_count > 0:
        available_types.append(('Transitions', transition_count))
    
    if not available_types:
        forms.alert('No valid elements found to import', title='Import')
        OUT.unfreeze()
        return
    
    # Multi-selection through sequential button clicks
    selected_type_names = []
    remaining_options = available_types[:]
    
    OUT.print_md('**Element Selection Process:**')
    
    while remaining_options:
        # Create options for current selection
        current_options = []
        for type_name, count in remaining_options:
            current_options.append("{} ({} available)".format(type_name, count))
        
        # Add control options
        if selected_type_names:
            current_options.append("Finished Selecting ({} types chosen)".format(len(selected_type_names)))
        current_options.append("Select All Remaining")
        current_options.append("Cancel Import")
        
        # Show current selection status
        if selected_type_names:
            status_msg = "Currently selected: {}\\nChoose additional types or finish selection".format(", ".join(selected_type_names))
        else:
            status_msg = "Select fittings to import\\nChoose first element type"
        
        selection = forms.CommandSwitchWindow.show(
            current_options,
            message=status_msg
        )
        
        if not selection or selection == "Cancel Import":
            OUT.print_md('Import cancelled by user')
            OUT.unfreeze()
            return
        elif selection == "Select All Remaining":
            # Add all remaining types
            for type_name, count in remaining_options:
                if type_name not in selected_type_names:
                    selected_type_names.append(type_name)
                    OUT.print_md('Added: {} ({} available)'.format(type_name, count))
            break
        elif selection.startswith("Finished Selecting"):
            if not selected_type_names:
                forms.alert('Please select at least one element type', title='Selection Required')
                continue
            break
        else:
            # Extract type name from selection
            selected_type = selection.split(' (')[0]
            if selected_type not in selected_type_names:
                selected_type_names.append(selected_type)
                OUT.print_md('Selected: {}'.format(selection))
                
                # Remove from remaining options
                remaining_options = [(t, c) for t, c in remaining_options if t != selected_type]
            else:
                forms.alert('{} already selected'.format(selected_type), title='Already Selected')
    
    if not selected_type_names:
        OUT.print_md('No element types selected - import cancelled')
        OUT.unfreeze()
        return
    
    OUT.print_md('**Final Selection:** {}'.format(", ".join(selected_type_names)))
    
    # Step 2: Generation Order Selection (only if multiple types selected)
    if len(selected_type_names) > 1:
        # Create all possible order combinations based on selected types
        order_options = []
        
        # Helper function to create readable order strings
        def create_order_option(order_list, description=""):
            filtered_order = [t for t in order_list if t in selected_type_names]
            if len(filtered_order) == len(selected_type_names):
                order_str = " → ".join(filtered_order)
                return "{} {}".format(order_str, description).strip()
            return None
        
        # Add standard order options that match selected types
        standard_option = create_order_option(['Straights', 'Elbows', 'Transitions'], "(Recommended)")
        if standard_option:
            order_options.append(standard_option)
        
        complex_first = create_order_option(['Transitions', 'Elbows', 'Straights'], "(Complex First)")
        if complex_first:
            order_options.append(complex_first)
        
        fittings_first = create_order_option(['Elbows', 'Straights', 'Transitions'], "(Fittings First)")
        if fittings_first:
            order_options.append(fittings_first)
        
        # Add all other possible combinations for selected types
        if len(selected_type_names) == 2:
            if 'Straights' in selected_type_names and 'Elbows' in selected_type_names:
                if "Elbows → Straights" not in [opt.split(' (')[0] for opt in order_options]:
                    order_options.append("Elbows → Straights")
            elif 'Straights' in selected_type_names and 'Transitions' in selected_type_names:
                if "Transitions → Straights" not in [opt.split(' (')[0] for opt in order_options]:
                    order_options.append("Transitions → Straights")
            elif 'Elbows' in selected_type_names and 'Transitions' in selected_type_names:
                if "Transitions → Elbows" not in [opt.split(' (')[0] for opt in order_options]:
                    order_options.append("Transitions → Elbows")
                if "Elbows → Transitions" not in [opt.split(' (')[0] for opt in order_options]:
                    order_options.append("Elbows → Transitions")
        elif len(selected_type_names) == 3:
            # Add remaining combinations for all 3 types
            all_combinations = [
                "Straights → Transitions → Elbows",
                "Elbows → Transitions → Straights", 
                "Transitions → Straights → Elbows"
            ]
            for combo in all_combinations:
                if combo not in [opt.split(' (')[0] for opt in order_options]:
                    order_options.append(combo)
        
        if not order_options:
            # Fallback - just use the selected types in alphabetical order
            generation_order = sorted(selected_type_names)
        else:
            order_choice = forms.CommandSwitchWindow.show(
                order_options,
                message="Select Generation Order"
            )
            
            if not order_choice:
                OUT.print_md('Import cancelled by user')
                OUT.unfreeze()
                return
            
            # Parse the selection to get the order
            order_str = order_choice.split(' (')[0]  # Remove description in parentheses
            generation_order = [t.strip() for t in order_str.split('→')]
    else:
        generation_order = selected_type_names
    
    # Step 3: Final confirmation
    summary = 'Import Configuration:\n\n'
    summary += 'Selected Types: {}\n'.format(', '.join(selected_type_names))
    summary += 'Generation Order: {}\n\n'.format(' → '.join(generation_order))
    summary += 'Total Elements: {}\n'.format(
        sum([len(straight_recs) if 'Straights' in selected_type_names else 0,
             elbow_count if 'Elbows' in selected_type_names else 0,
             transition_count if 'Transitions' in selected_type_names else 0])
    )
    summary += '\nNote: All temporary stub ducts will be automatically cleaned up after fitting creation.'
    
    confirmed = forms.alert(
        summary,
        title='Confirm Import Settings',
        yes=True, 
        no=True
    )
    
    if not confirmed:
        OUT.print_md('Import cancelled by user')
        OUT.unfreeze()
        return
    
    # Display selected options
    OUT.print_md('**User Selection:**')
    OUT.print_md('- Selected types: {}'.format(', '.join(selected_type_names)))
    OUT.print_md('- Generation order: {}'.format(' → '.join(generation_order)))
    
    straight_cnt = 0
    elbow_cnt = 0
    trans_cnt = 0
    skipped = 0
    
    t = DB.Transaction(doc, 'Enhanced Import Duct (User Selected)')
    t.Start()
    try:
        # Process in user-specified order
        for phase_num, element_type in enumerate(generation_order, 1):
            OUT.print_md('**Phase {}: Creating {}...**'.format(phase_num, element_type))
            
            if element_type == 'Straights' and 'Straights' in selected_type_names:
                for rec in straight_recs:
                    try:
                        el, err = _create_duct_from_record(rec)
                        if el:
                            straight_cnt += 1
                        else:
                            OUT.print_md(':x: Skip straight {}: {}'.format(rec.get('element_id', 'n/a'), err))
                            skipped += 1
                    except Exception as ex:
                        OUT.print_md(':x: Error in straight {}: {}'.format(rec.get('element_id', 'n/a'), ex))
                        skipped += 1
            
            elif element_type == 'Elbows' and 'Elbows' in selected_type_names:
                for rec in fitting_recs:
                    try:
                        fitting_type, fitting_data = _classify_fitting(rec)
                        
                        if fitting_type.startswith('Elbow') or fitting_type == 'ElbowGeneric':
                            el, err = _make_elbow(rec, fitting_type, fitting_data)
                            if el:
                                elbow_cnt += 1
                            else:
                                OUT.print_md(':x: Skip elbow {}: {}'.format(rec.get('element_id', 'n/a'), err))
                                skipped += 1
                    except Exception as ex:
                        OUT.print_md(':x: Error in elbow {}: {}'.format(rec.get('element_id', 'n/a'), ex))
                        skipped += 1
            
            elif element_type == 'Transitions' and 'Transitions' in selected_type_names:
                for rec in fitting_recs:
                    try:
                        fitting_type, fitting_data = _classify_fitting(rec)
                        
                        if fitting_type == 'Transition':
                            el, err = _make_transition(rec, fitting_data)
                            if el:
                                trans_cnt += 1
                            else:
                                OUT.print_md(':x: Skip transition {}: {}'.format(rec.get('element_id', 'n/a'), err))
                                skipped += 1
                    except Exception as ex:
                        OUT.print_md(':x: Error in transition {}: {}'.format(rec.get('element_id', 'n/a'), ex))
                        skipped += 1
        
        t.Commit()
        
        # Step 3: Optimize straight duct connections to fittings
        if 'Straights' in selected_type_names and (elbow_cnt > 0 or trans_cnt > 0):
            OUT.print_md('**Step 3: Optimizing straight duct connections to fittings...**')
            
            # Start a new transaction for connection optimization
            t2 = DB.Transaction(doc, 'Optimize Straight Duct Connections')
            t2.Start()
            
            try:
                optimized_count = _optimize_straight_connections()
                t2.Commit()
                OUT.print_md('**Connection Optimization Complete:** {} straights adjusted'.format(optimized_count))
            except Exception as opt_ex:
                OUT.print_md(':x: **Connection Optimization Failed**: {}'.format(opt_ex))
                try:
                    t2.RollBack()
                except:
                    pass
        
        OUT.print_md('**Import Complete**')
        OUT.print_md('- Straights created: {}'.format(straight_cnt))
        OUT.print_md('- Elbows created: {}'.format(elbow_cnt))
        OUT.print_md('- Transitions created: {}'.format(trans_cnt))
        OUT.print_md('- Skipped: {}'.format(skipped))
        OUT.print_md('**All temporary stub ducts have been automatically cleaned up**')
        
    except Exception as ex:
        OUT.print_md(':x: **Transaction Failed**: {}'.format(ex))
        try:
            t.RollBack()
        except:
            pass
    
    OUT.unfreeze()

# ================================================================
#  ENGINE 3: EXPORT FAMILY DATA (inlined)
# ================================================================

def _xyz_to_tuple(xyz):
    if xyz:
        return (round(xyz.X, 6), round(xyz.Y, 6), round(xyz.Z, 6))
    return (None, None, None)

def _xyz_to_dict(xyz, prefix=""):
    if xyz:
        return {
            prefix + "x": round(xyz.X, 6),
            prefix + "y": round(xyz.Y, 6),
            prefix + "z": round(xyz.Z, 6)
        }
    return {prefix + "x": None, prefix + "y": None, prefix + "z": None}

def _bb6(bb):
    if not bb:
        return (None,) * 6
    mn, mx = bb.Min, bb.Max
    return (round(mn.X, 6), round(mn.Y, 6), round(mn.Z, 6),
            round(mx.X, 6), round(mx.Y, 6), round(mx.Z, 6))

def _safe_get(o, name, default=None):
    try:
        return getattr(o, name)
    except:
        return default

def _safe_name(o, fallback=None):
    if o is None:
        return fallback
    try:
        n = o.Name
        if n:
            return n
    except:
        pass
    try:
        p = o.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_NAME)
        if p:
            n = p.AsString()
            if n:
                return n
    except:
        pass
    return fallback

def _param_val(p):
    """Extract parameter value with type info"""
    try:
        st = p.StorageType
        if st == DB.StorageType.Double:
            return p.AsDouble()
        if st == DB.StorageType.Integer:
            return p.AsInteger()
        if st == DB.StorageType.String:
            return p.AsString()
        if st == DB.StorageType.ElementId:
            eid = p.AsElementId()
            return int(eid.IntegerValue) if eid else None
    except:
        pass
    return None

def _get_connector_domain(conn):
    """Get connector domain as string"""
    try:
        domain = conn.Domain
        if domain == DB.Domain.DomainPiping:
            return "Piping"
        elif domain == DB.Domain.DomainHvac:
            return "Hvac"
        elif domain == DB.Domain.DomainElectrical:
            return "Electrical"
        elif domain == DB.Domain.DomainCableTrayConduit:
            return "CableTrayConduit"
        else:
            return str(domain)
    except:
        return None

def _get_connector_shape(conn):
    """Get connector shape as string"""
    try:
        shape = conn.Shape
        if shape == DB.ConnectorProfileType.Round:
            return "Round"
        elif shape == DB.ConnectorProfileType.Rectangular:
            return "Rectangular"
        elif shape == DB.ConnectorProfileType.Oval:
            return "Oval"
        else:
            return str(shape)
    except:
        return None

def _get_connector_system_type(conn):
    """Get MEP system type"""
    try:
        # Try pipe system type
        if hasattr(conn, 'PipeSystemType'):
            return str(conn.PipeSystemType)
    except:
        pass
    try:
        # Try duct system type
        if hasattr(conn, 'DuctSystemType'):
            return str(conn.DuctSystemType)
    except:
        pass
    return None

def _extract_connectors_from_family(famdoc, fid, fname):
    """Extract connector definitions from family document"""
    connectors = []
    try:
        # Get connector elements in family
        for elem in DB.FilteredElementCollector(famdoc).OfClass(DB.ConnectorElement):
            try:
                conn_data = {
                    "family_id": fid,
                    "family_name": fname,
                    "connector_id": int(elem.Id.IntegerValue),
                    "domain": None,
                    "shape": None,
                    "radius": None,
                    "width": None,
                    "height": None,
                    "system_type": None,
                }
                
                # Get origin
                try:
                    origin = elem.Origin
                    conn_data.update(_xyz_to_dict(origin, "origin_"))
                except:
                    conn_data.update({"origin_x": None, "origin_y": None, "origin_z": None})
                
                # Get direction/orientation
                try:
                    coord_sys = elem.CoordinateSystem
                    if coord_sys:
                        conn_data.update(_xyz_to_dict(coord_sys.BasisZ, "direction_"))
                except:
                    conn_data.update({"direction_x": None, "direction_y": None, "direction_z": None})
                
                # Get domain from connector element parameters
                try:
                    domain_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_DOMAIN)
                    if domain_param:
                        conn_data["domain"] = domain_param.AsValueString()
                except:
                    pass
                
                # Get shape
                try:
                    shape_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_PROFILE_TYPE)
                    if shape_param:
                        conn_data["shape"] = shape_param.AsValueString()
                except:
                    pass
                
                # Get dimensions based on shape
                try:
                    # Radius for round connectors
                    radius_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_RADIUS)
                    if radius_param:
                        conn_data["radius"] = radius_param.AsDouble()
                except:
                    pass
                
                try:
                    # Width/Height for rectangular
                    width_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_WIDTH)
                    height_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_HEIGHT)
                    if width_param:
                        conn_data["width"] = width_param.AsDouble()
                    if height_param:
                        conn_data["height"] = height_param.AsDouble()
                except:
                    pass
                
                # Get system classification
                try:
                    sys_param = elem.get_Parameter(DB.BuiltInParameter.CONNECTOR_SYSTEM_CLASSIFICATION)
                    if sys_param:
                        conn_data["system_type"] = sys_param.AsValueString()
                except:
                    pass
                
                connectors.append(conn_data)
            except Exception as e:
                pass
    except:
        pass
    return connectors

def _extract_connectors_from_instance(inst):
    """Extract connector info from a placed family instance"""
    connectors = []
    try:
        mep_model = inst.MEPModel
        if not mep_model:
            return connectors
        conn_mgr = mep_model.ConnectorManager
        if not conn_mgr:
            return connectors
        
        for conn in conn_mgr.Connectors:
            try:
                conn_data = {
                    "element_id": int(inst.Id.IntegerValue),
                    "connector_index": conn.Id,
                    "domain": _get_connector_domain(conn),
                    "shape": _get_connector_shape(conn),
                    "system_type": _get_connector_system_type(conn),
                    "is_connected": conn.IsConnected,
                }
                
                # Origin
                try:
                    conn_data.update(_xyz_to_dict(conn.Origin, "origin_"))
                except:
                    conn_data.update({"origin_x": None, "origin_y": None, "origin_z": None})
                
                # Direction
                try:
                    coord_sys = conn.CoordinateSystem
                    if coord_sys:
                        conn_data.update(_xyz_to_dict(coord_sys.BasisZ, "direction_"))
                except:
                    conn_data.update({"direction_x": None, "direction_y": None, "direction_z": None})
                
                # Dimensions
                try:
                    if conn.Shape == DB.ConnectorProfileType.Round:
                        conn_data["radius"] = conn.Radius
                        conn_data["width"] = None
                        conn_data["height"] = None
                    elif conn.Shape == DB.ConnectorProfileType.Rectangular:
                        conn_data["radius"] = None
                        conn_data["width"] = conn.Width
                        conn_data["height"] = conn.Height
                    elif conn.Shape == DB.ConnectorProfileType.Oval:
                        conn_data["radius"] = None
                        conn_data["width"] = conn.Width
                        conn_data["height"] = conn.Height
                except:
                    conn_data["radius"] = None
                    conn_data["width"] = None
                    conn_data["height"] = None
                
                # Flow direction for MEP
                try:
                    conn_data["flow_direction"] = str(conn.Direction)
                except:
                    conn_data["flow_direction"] = None
                
                connectors.append(conn_data)
            except:
                pass
    except:
        pass
    return connectors

# ------------- deep family scrape -------------
def _collect_family_deep(fam):
    out = {
        "families": [], 
        "family_params": [], 
        "family_param_values": [],  # NEW: parameter values per type
        "types": [], 
        "type_params": [],
        "ref_planes": [], 
        "ref_lines": [], 
        "forms": [], 
        "loops": [],
        "connectors": []  # NEW: connector definitions
    }
    fid = int(fam.Id.IntegerValue)
    fname = _safe_name(fam)
    fcat = _safe_name(_safe_get(fam, "FamilyCategory", None))
    out["families"].append({"family_id": fid, "family_name": fname, "category": fcat})

    famdoc = None
    try:
        famdoc = doc.EditFamily(fam)
        mgr = famdoc.FamilyManager

        # Collect family parameter definitions
        param_defs = []
        try:
            for p in mgr.GetParameters():
                try:
                    defn = p.Definition
                    group = None
                    try:
                        group = str(defn.ParameterGroup)
                    except:
                        pass
                    storage = None
                    try:
                        storage = str(p.StorageType)
                    except:
                        pass
                    formula = None
                    try:
                        formula = mgr.GetFormula(p)
                    except:
                        pass
                    
                    param_info = {
                        "family_id": fid, 
                        "family_name": fname,
                        "name": defn.Name if defn else None,
                        "group": group,
                        "is_instance": getattr(p, "IsInstance", None),
                        "is_report": getattr(p, "IsReporting", None),
                        "storage": storage,
                        "formula": formula or None,
                    }
                    out["family_params"].append(param_info)
                    param_defs.append((p, defn.Name if defn else None))
                except:
                    pass
        except:
            pass

        # Collect parameter VALUES for each family type
        try:
            for fam_type in mgr.Types:
                try:
                    mgr.CurrentType = fam_type
                    type_name = fam_type.Name if fam_type else "Default"
                    
                    for p, pname in param_defs:
                        if not pname:
                            continue
                        try:
                            val = None
                            st = p.StorageType
                            if st == DB.StorageType.Double:
                                val = mgr.CurrentType.AsDouble(p) if mgr.CurrentType else None
                                if val is None:
                                    try:
                                        val = p.AsDouble()
                                    except:
                                        pass
                            elif st == DB.StorageType.Integer:
                                val = mgr.CurrentType.AsInteger(p) if mgr.CurrentType else None
                                if val is None:
                                    try:
                                        val = p.AsInteger()
                                    except:
                                        pass
                            elif st == DB.StorageType.String:
                                val = mgr.CurrentType.AsString(p) if mgr.CurrentType else None
                                if val is None:
                                    try:
                                        val = p.AsString()
                                    except:
                                        pass
                            elif st == DB.StorageType.ElementId:
                                eid = mgr.CurrentType.AsElementId(p) if mgr.CurrentType else None
                                val = int(eid.IntegerValue) if eid else None
                            
                            out["family_param_values"].append({
                                "family_id": fid,
                                "family_name": fname,
                                "type_name": type_name,
                                "param_name": pname,
                                "storage": str(p.StorageType),
                                "value": val
                            })
                        except:
                            pass
                except:
                    pass
        except:
            pass

        # Reference planes
        try:
            for rp in DB.FilteredElementCollector(famdoc).OfClass(DB.ReferencePlane):
                try:
                    b = _bb6(rp.get_BoundingBox(None))
                    bx, by, bz = _xyz_to_tuple(_safe_get(rp, "BubbleEnd", None))
                    fx, fy, fz = _xyz_to_tuple(_safe_get(rp, "FreeEnd", None))
                    cvx, cvy, cvz = _xyz_to_tuple(_safe_get(rp, "CutVector", None))
                    out["ref_planes"].append({
                        "family_id": fid, "family_name": fname,
                        "id": int(rp.Id.IntegerValue), "name": _safe_name(rp),
                        "bubble_end_x": bx, "bubble_end_y": by, "bubble_end_z": bz,
                        "free_end_x": fx, "free_end_y": fy, "free_end_z": fz,
                        "cut_vec_x": cvx, "cut_vec_y": cvy, "cut_vec_z": cvz,
                        "bb_min_x": b[0], "bb_min_y": b[1], "bb_min_z": b[2],
                        "bb_max_x": b[3], "bb_max_y": b[4], "bb_max_z": b[5],
                    })
                except:
                    pass
        except:
            pass

        # Reference lines (model curves) - use CurveElement for Revit 2025+ compatibility
        try:
            for ce in DB.FilteredElementCollector(famdoc).OfClass(DB.CurveElement):
                try:
                    # Filter to only ModelCurve types
                    if not isinstance(ce, DB.ModelCurve):
                        continue
                    mc = ce
                    crv = mc.GeometryCurve
                    sx, sy, sz = _xyz_to_tuple(crv.GetEndPoint(0)) if crv else (None, None, None)
                    ex, ey, ez = _xyz_to_tuple(crv.GetEndPoint(1)) if crv else (None, None, None)
                    b = _bb6(mc.get_BoundingBox(None))
                    out["ref_lines"].append({
                        "family_id": fid, "family_name": fname,
                        "id": int(mc.Id.IntegerValue), "name": _safe_name(mc),
                        "is_reference_line": getattr(mc, "IsReferenceLine", None),
                        "curve_type": crv.GetType().Name if crv else None,
                        "start_x": sx, "start_y": sy, "start_z": sz,
                        "end_x": ex, "end_y": ey, "end_z": ez,
                        "bb_min_x": b[0], "bb_min_y": b[1], "bb_min_z": b[2],
                        "bb_max_x": b[3], "bb_max_y": b[4], "bb_max_z": b[5],
                    })
                except:
                    pass
        except:
            pass

        # Forms + loops (Extrusions - SOLIDS and VOIDS)
        for form_type in ["Extrusion", "Blend", "Revolve", "Sweep", "SweptBlend"]:
            cls = getattr(DB, form_type, None)
            if not cls:
                continue
            try:
                for el in DB.FilteredElementCollector(famdoc).OfClass(cls):
                    try:
                        b = _bb6(el.get_BoundingBox(None))
                        dirx = diry = dirz = None
                        depth = None
                        
                        # Get solid/void status
                        is_solid = True
                        is_void = False
                        try:
                            is_solid = getattr(el, "IsSolid", True)
                            is_void = not is_solid
                        except:
                            pass
                        
                        if form_type == "Extrusion":
                            try:
                                d = getattr(el, "ExtrusionDirection", None)
                                dirx, diry, dirz = _xyz_to_tuple(d)
                            except:
                                pass
                            try:
                                # Get start and end offsets for full extrusion info
                                depth = getattr(el, "EndOffset", 0) - getattr(el, "StartOffset", 0)
                                if depth == 0:
                                    depth = getattr(el, "ExtrusionDepth", None)
                            except:
                                try:
                                    depth = getattr(el, "ExtrusionDepth", None)
                                except:
                                    pass
                        
                        # Get subcategory (for visibility control)
                        subcat = None
                        try:
                            sc = el.Subcategory
                            if sc:
                                subcat = sc.Name
                        except:
                            pass
                        
                        # Get material
                        material_id = None
                        try:
                            mat_param = el.get_Parameter(DB.BuiltInParameter.MATERIAL_ID_PARAM)
                            if mat_param:
                                material_id = mat_param.AsElementId().IntegerValue
                        except:
                            pass
                        
                        out["forms"].append({
                            "family_id": fid, 
                            "family_name": fname,
                            "id": int(el.Id.IntegerValue),
                            "kind": form_type,
                            "is_solid": is_solid,
                            "is_void": is_void,
                            "subcategory": subcat,
                            "material_id": material_id,
                            "dir_x": dirx, "dir_y": diry, "dir_z": dirz,
                            "depth": depth,
                            "bb_min_x": b[0], "bb_min_y": b[1], "bb_min_z": b[2],
                            "bb_max_x": b[3], "bb_max_y": b[4], "bb_max_z": b[5],
                        })

                        # Extract sketch loops for extrusions
                        if form_type == "Extrusion":
                            try:
                                sk = getattr(el, "Sketch", None)
                                prof = getattr(sk, "Profile", None) if sk else None
                                if prof:
                                    loop_i = 0
                                    for ca in prof:
                                        loop_i += 1
                                        for c in ca:
                                            try:
                                                ctype = c.GetType().Name
                                                s = c.GetEndPoint(0)
                                                e = c.GetEndPoint(1)
                                                
                                                loop_data = {
                                                    "family_id": fid, 
                                                    "family_name": fname,
                                                    "form_id": int(el.Id.IntegerValue), 
                                                    "form_kind": form_type,
                                                    "is_void": is_void,
                                                    "loop_index": loop_i, 
                                                    "curve_type": ctype,
                                                    "start_x": s.X, "start_y": s.Y, "start_z": s.Z,
                                                    "end_x": e.X, "end_y": e.Y, "end_z": e.Z
                                                }
                                                
                                                if ctype == "Arc":
                                                    try:
                                                        mid = c.Evaluate(0.5, True)
                                                        loop_data["mid_x"] = mid.X
                                                        loop_data["mid_y"] = mid.Y
                                                        loop_data["mid_z"] = mid.Z
                                                    except:
                                                        loop_data["mid_x"] = None
                                                        loop_data["mid_y"] = None
                                                        loop_data["mid_z"] = None
                                                
                                                out["loops"].append(loop_data)
                                            except:
                                                pass
                            except:
                                pass
                    except:
                        pass
            except:
                pass

        # Extract connectors from family document
        family_connectors = _extract_connectors_from_family(famdoc, fid, fname)
        out["connectors"].extend(family_connectors)

    finally:
        try:
            if famdoc:
                famdoc.Close(False)
        except:
            pass

    # Types + type params (from project types)
    for sid in fam.GetFamilySymbolIds():
        sym = doc.GetElement(sid)
        out["types"].append({
            "family_id": fid, "family_name": fname,
            "symbol_id": int(sym.Id.IntegerValue),
            "type_name": _safe_name(sym),
            "category": _safe_name(_safe_get(sym, "Category", None))
        })
        try:
            for p in sym.Parameters:
                try:
                    defn = p.Definition
                    group = None
                    try:
                        group = str(defn.ParameterGroup)
                    except:
                        pass
                    out["type_params"].append({
                        "family_id": fid, "family_name": fname,
                        "symbol_id": int(sym.Id.IntegerValue), "type_name": _safe_name(sym),
                        "name": defn.Name if defn else None,
                        "group": group,
                        "storage": str(p.StorageType),
                        "value": _param_val(p),
                    })
                except:
                    pass
        except:
            pass

    return out

# ------------- selection + file writing -------------
def _collect_targets(mode):
    if mode == "Pick Elements":
        refs = uidoc.Selection.PickObjects(DB.Selection.ObjectType.Element, "Pick family instances")
        return [doc.GetElement(r.ElementId) for r in refs]
    elif mode == "Use Current Selection":
        sel = [doc.GetElement(eid) for eid in uidoc.Selection.GetElementIds()]
        return [e for e in sel if isinstance(e, DB.FamilyInstance)]
    elif mode == "All Family Instances In View":
        v = doc.ActiveView
        flt = DB.ElementClassFilter(DB.FamilyInstance)
        return list(DB.FilteredElementCollector(doc, v.Id).WherePasses(flt).WhereElementIsNotElementType())
    else:
        return None

def _write_csv(path, headers, rows):
    if PY3:
        with open(path, 'w', newline='', encoding='utf-8') as f:
            w = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
            w.writeheader()
            for r in rows:
                w.writerow(r)
    else:
        with open(path, 'wb') as f:
            w = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
            w.writeheader()
            for r in rows:
                encoded_row = {}
                for k, v in r.items():
                    if isinstance(v, unicode):
                        encoded_row[k] = v.encode('utf-8')
                    else:
                        encoded_row[k] = v
                w.writerow(encoded_row)

def run_export_family():
    modes = ["Pick Elements", "Use Current Selection", "All Family Instances In View"]
    choice = forms.CommandSwitchWindow.show(modes, message="Export Family Data (v10) — choose scope")
    if not choice:
        return
    insts = _collect_targets(choice) or []
    if not insts:
        return forms.alert("No family instances found.")

    OUT.print_md("**Export Family Data v10**")
    OUT.print_md("- Revit Version: {}".format(REVIT_VERSION))
    OUT.print_md("- Selected {} instances".format(len(insts)))

    # Build set of families
    fams = {}
    for i in insts:
        if getattr(i, "Symbol", None):
            fams[i.Symbol.Family.Id.IntegerValue] = i.Symbol.Family

    OUT.print_md("- Found {} unique families".format(len(fams)))

    base = forms.save_file(file_ext='jsonl', default_name='family_data_v10.jsonl')
    if not base:
        return
    root, _ = os.path.splitext(base)

    # Collect all data
    families = []
    family_params = []
    family_param_values = []
    types = []
    type_params = []
    ref_planes = []
    ref_lines = []
    forms_rows = []
    loops_rows = []
    connectors_family = []
    instances = []
    instance_connectors = []

    with open(root + ".jsonl", 'w') as jf:
        for f in fams.values():
            OUT.print_md("  - Processing family: **{}**".format(_safe_name(f)))
            rows = _collect_family_deep(f)
            families.extend(rows["families"])
            family_params.extend(rows["family_params"])
            family_param_values.extend(rows["family_param_values"])
            types.extend(rows["types"])
            type_params.extend(rows["type_params"])
            ref_planes.extend(rows["ref_planes"])
            ref_lines.extend(rows["ref_lines"])
            forms_rows.extend(rows["forms"])
            loops_rows.extend(rows["loops"])
            connectors_family.extend(rows["connectors"])
            
            # Write JSONL record
            jf.write(json.dumps({
                "record_type": "FAMILY",
                "family_id": rows["families"][0]["family_id"],
                "family_name": rows["families"][0]["family_name"],
                "category": rows["families"][0]["category"],
                "family_parameters": rows["family_params"],
                "family_param_values": rows["family_param_values"],
                "connectors": rows["connectors"],
                "geometry": {
                    "ref_planes": rows["ref_planes"],
                    "ref_lines": rows["ref_lines"],
                    "forms": rows["forms"],
                    "loops": rows["loops"]
                }
            }) + "\n")
            
            for t in rows["types"]:
                jf.write(json.dumps(dict({"record_type": "TYPE"}, **t)) + "\n")

        # Process instances
        for inst in insts:
            try:
                loc = inst.Location
                px = py = pz = fx = fy = fz = hx = hy = hz = None
                if isinstance(loc, DB.LocationPoint):
                    px, py, pz = _xyz_to_tuple(loc.Point)
                    fo = getattr(inst, "FacingOrientation", None)
                    ho = getattr(inst, "HandOrientation", None)
                    if fo:
                        fx, fy, fz = _xyz_to_tuple(fo)
                    if ho:
                        hx, hy, hz = _xyz_to_tuple(ho)
                
                fam = inst.Symbol.Family if getattr(inst, "Symbol", None) else None
                sym = inst.Symbol if getattr(inst, "Symbol", None) else None
                
                level_id = None
                level_name = None
                try:
                    lid = getattr(inst, "LevelId", None)
                    if lid and lid.IntegerValue > 0:
                        level_id = int(lid.IntegerValue)
                        lvl = doc.GetElement(lid)
                        if lvl:
                            level_name = lvl.Name
                except:
                    pass
                
                host_id = None
                try:
                    h = getattr(inst, "Host", None)
                    if h:
                        host_id = int(h.Id.IntegerValue)
                except:
                    pass
                
                # Get rotation angle
                rotation = None
                try:
                    if isinstance(loc, DB.LocationPoint):
                        rotation = loc.Rotation
                except:
                    pass
                
                rec = {
                    "element_id": int(inst.Id.IntegerValue),
                    "family_id": int(fam.Id.IntegerValue) if fam else None,
                    "family_name": _safe_name(fam),
                    "symbol_id": int(sym.Id.IntegerValue) if sym else None,
                    "type_name": _safe_name(sym),
                    "category": _safe_name(getattr(inst, "Category", None)),
                    "level_id": level_id,
                    "level_name": level_name,
                    "host_id": host_id,
                    "point_x": px, "point_y": py, "point_z": pz,
                    "facing_x": fx, "facing_y": fy, "facing_z": fz,
                    "hand_x": hx, "hand_y": hy, "hand_z": hz,
                    "rotation": rotation
                }
                instances.append(rec)
                
                # Get instance connectors
                inst_conns = _extract_connectors_from_instance(inst)
                for ic in inst_conns:
                    ic["family_name"] = _safe_name(fam)
                    ic["type_name"] = _safe_name(sym)
                instance_connectors.extend(inst_conns)
                
                jf.write(json.dumps(dict({"record_type": "INSTANCE"}, **rec)) + "\n")
            except:
                pass

    # Write CSV sidecars
    def w(suffix, headers, rows):
        _write_csv(root + suffix, headers, rows)

    w("_families.csv", ["family_id", "family_name", "category"], families)
    w("_family_params.csv", ["family_id", "family_name", "name", "group", "is_instance", "is_report", "storage", "formula"], family_params)
    w("_family_param_values.csv", ["family_id", "family_name", "type_name", "param_name", "storage", "value"], family_param_values)
    w("_types.csv", ["family_id", "family_name", "symbol_id", "type_name", "category"], types)
    w("_type_params.csv", ["family_id", "family_name", "symbol_id", "type_name", "name", "group", "storage", "value"], type_params)
    w("_ref_planes.csv", ["family_id", "family_name", "id", "name", "bubble_end_x", "bubble_end_y", "bubble_end_z", "free_end_x", "free_end_y", "free_end_z", "cut_vec_x", "cut_vec_y", "cut_vec_z", "bb_min_x", "bb_min_y", "bb_min_z", "bb_max_x", "bb_max_y", "bb_max_z"], ref_planes)
    w("_ref_lines.csv", ["family_id", "family_name", "id", "name", "is_reference_line", "curve_type", "start_x", "start_y", "start_z", "end_x", "end_y", "end_z", "bb_min_x", "bb_min_y", "bb_min_z", "bb_max_x", "bb_max_y", "bb_max_z"], ref_lines)
    w("_forms.csv", ["family_id", "family_name", "id", "kind", "is_solid", "is_void", "subcategory", "material_id", "dir_x", "dir_y", "dir_z", "depth", "bb_min_x", "bb_min_y", "bb_min_z", "bb_max_x", "bb_max_y", "bb_max_z"], forms_rows)
    w("_loops.csv", ["family_id", "family_name", "form_id", "form_kind", "is_void", "loop_index", "curve_type", "start_x", "start_y", "start_z", "end_x", "end_y", "end_z", "mid_x", "mid_y", "mid_z"], loops_rows)
    w("_connectors_family.csv", ["family_id", "family_name", "connector_id", "domain", "shape", "radius", "width", "height", "system_type", "origin_x", "origin_y", "origin_z", "direction_x", "direction_y", "direction_z"], connectors_family)
    w("_instances.csv", ["element_id", "family_id", "family_name", "symbol_id", "type_name", "category", "level_id", "level_name", "host_id", "point_x", "point_y", "point_z", "facing_x", "facing_y", "facing_z", "hand_x", "hand_y", "hand_z", "rotation"], instances)
    w("_connectors_instance.csv", ["element_id", "family_name", "type_name", "connector_index", "domain", "shape", "system_type", "is_connected", "radius", "width", "height", "flow_direction", "origin_x", "origin_y", "origin_z", "direction_x", "direction_y", "direction_z"], instance_connectors)

    OUT.print_md("---")
    OUT.print_md("**Export complete (v10)**")
    OUT.print_md("- Families: {}".format(len(families)))
    OUT.print_md("- Family parameters: {} definitions, {} values".format(len(family_params), len(family_param_values)))
    OUT.print_md("- Types: {}".format(len(types)))
    OUT.print_md("- Forms (solids/voids): {}".format(len(forms_rows)))
    OUT.print_md("- Sketch loops: {}".format(len(loops_rows)))
    OUT.print_md("- Family connectors: {}".format(len(connectors_family)))
    OUT.print_md("- Instances: {}".format(len(instances)))
    OUT.print_md("- Instance connectors: {}".format(len(instance_connectors)))
    OUT.print_md("- Files saved to: `{}`".format(root))

# ================================================================
#  ENGINE 4: IMPORT FAMILY DATA (inlined)
# ================================================================

def _f(x):
    if x is None: return None
    if isinstance(x, (int, float)): return float(x)
    s = str(x).strip()
    if s == "" or s.lower() == "none": return None
    try: return float(s)
    except: return None

def _b(x):
    if x is None: return False
    if isinstance(x, bool): return x
    return str(x).strip().lower() in ('true', '1', 'yes')

def _read_records_any(path):
    try:
        with codecs.open(path, 'r', 'utf-8-sig') as f:
            data = json.load(f)
        if isinstance(data, list): return data
        if isinstance(data, dict): return [data]
    except ValueError: pass
    recs = []
    with codecs.open(path, 'r', 'utf-8-sig') as f:
        for ln, line in enumerate(f, 1):
            s = line.strip()
            if not s: continue
            try: recs.append(json.loads(s))
            except: pass
    return recs

def _split_records(recs):
    families, types, instances = [], [], []
    for rec in recs:
        rt = rec.get('record_type')
        if rt == 'FAMILY': families.append(rec)
        elif rt == 'TYPE': types.append(rec)
        elif rt == 'INSTANCE': instances.append(rec)
    return families, types, instances

def _find_template_file(candidates):
    try: root = app.FamilyTemplatePath
    except: root = None
    search_roots = []
    if root and os.path.isdir(root): search_roots.append(root)
    common_paths = [
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English".format(REVIT_VERSION),
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates\English_I".format(REVIT_VERSION),
        r"C:\ProgramData\Autodesk\RVT {}\Family Templates".format(REVIT_VERSION),
    ]
    for p in common_paths:
        if os.path.isdir(p) and p not in search_roots:
            search_roots.append(p)
    for candidate in candidates:
        for search_root in search_roots:
            for dirpath, _, filenames in os.walk(search_root):
                for f in filenames:
                    if f.lower() == candidate.lower():
                        return os.path.join(dirpath, f)
    return None

def _load_family(project_doc, path):
    class FamOpt(DB.IFamilyLoadOptions):
        def OnFamilyFound(self, familyInUse, overwriteParameterValues):
            overwriteParameterValues = True
            return True
        def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
            overwriteParameterValues = True
            return True
    fam_ref = clr.Reference[DB.Family]()
    ok = project_doc.LoadFamily(path, FamOpt(), fam_ref)
    return fam_ref.Value if ok else None

_TOL = 1e-6

def _create_test_box(famdoc, size=1.0):
    """Create a simple test box"""
    OUT.print_md("  Creating TEST BOX...")
    p0, p1 = DB.XYZ(0, 0, 0), DB.XYZ(size, 0, 0)
    p2, p3 = DB.XYZ(size, size, 0), DB.XYZ(0, size, 0)
    carr = DB.CurveArray()
    carr.Append(DB.Line.CreateBound(p0, p1))
    carr.Append(DB.Line.CreateBound(p1, p2))
    carr.Append(DB.Line.CreateBound(p2, p3))
    carr.Append(DB.Line.CreateBound(p3, p0))
    profiles = DB.CurveArrArray()
    profiles.Append(carr)
    t = DB.Transaction(famdoc, 'Test Box')
    t.Start()
    try:
        sp = DB.SketchPlane.Create(famdoc, DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ.Zero))
        ext = famdoc.FamilyCreate.NewExtrusion(True, profiles, sp, size)
        t.Commit()
        OUT.print_md("  TEST BOX: **SUCCESS**")
        return ext
    except Exception as e:
        OUT.print_md("  TEST BOX FAILED: {}".format(e))
        t.RollBack()
        return None

def _pt_tuple(row, prefix):
    x = _f(row.get(prefix + '_x'))
    y = _f(row.get(prefix + '_y'))
    z = _f(row.get(prefix + '_z'))
    if x is None or y is None or z is None: return None
    return (x, y, z)

def _dominant_z(rows):
    counts = {}
    for r in rows:
        for key in ['start', 'end', 'mid']:
            pt = _pt_tuple(r, key)
            if pt:
                z = round(pt[2], 6)
                counts[z] = counts.get(z, 0) + 1
    if not counts: return 0.0
    return max(counts.items(), key=lambda kv: kv[1])[0]

def _dist2(a, b):
    return (a.X-b.X)**2 + (a.Y-b.Y)**2 + (a.Z-b.Z)**2

def _order_loop_curves(rows, plane_z):
    curves = []
    for r in rows:
        ct = r.get('curve_type')
        s = _pt_tuple(r, 'start')
        e = _pt_tuple(r, 'end')
        m = _pt_tuple(r, 'mid')
        if not s or not e: continue
        s_xyz = DB.XYZ(s[0], s[1], plane_z)
        e_xyz = DB.XYZ(e[0], e[1], plane_z)
        if _dist2(s_xyz, e_xyz) < _TOL: continue
        if ct == 'Arc' and m:
            m_xyz = DB.XYZ(m[0], m[1], plane_z)
            try: c = DB.Arc.Create(s_xyz, e_xyz, m_xyz)
            except: c = DB.Line.CreateBound(s_xyz, e_xyz)
        else:
            c = DB.Line.CreateBound(s_xyz, e_xyz)
        curves.append(c)
    if len(curves) < 3: return None
    out = [curves.pop(0)]
    while curves:
        endpt = out[-1].GetEndPoint(1)
        match_idx, rev = None, False
        for i, c in enumerate(curves):
            if _dist2(c.GetEndPoint(0), endpt) < _TOL:
                match_idx, rev = i, False
                break
            if _dist2(c.GetEndPoint(1), endpt) < _TOL:
                match_idx, rev = i, True
                break
        if match_idx is None:
            out.extend(curves)
            break
        nxt = curves.pop(match_idx)
        if rev:
            try: nxt = nxt.CreateReversed()
            except: nxt = DB.Line.CreateBound(nxt.GetEndPoint(1), nxt.GetEndPoint(0))
        out.append(nxt)
    carr = DB.CurveArray()
    for c in out: carr.Append(c)
    return carr

def _build_extrusion_from_loops(famdoc, loops_data, form_data):
    is_void = _b(form_data.get('is_void'))
    is_solid = not is_void
    depth = _f(form_data.get('depth')) or 1.0
    
    OUT.print_md("    Loop extrusion: {} curves, depth={}".format(len(loops_data), depth))
    
    plane_z = _dominant_z(loops_data)
    by_loop = {}
    for r in loops_data:
        idx = int(r.get('loop_index', 1))
        by_loop.setdefault(idx, []).append(r)
    
    profiles = DB.CurveArrArray()
    for idx in sorted(by_loop.keys()):
        carr = _order_loop_curves(by_loop[idx], plane_z)
        if carr and carr.Size >= 3:
            profiles.Append(carr)
            OUT.print_md("      Loop {}: {} curves OK".format(idx, carr.Size))
        else:
            OUT.print_md("      Loop {}: FAILED".format(idx))
    
    if profiles.Size == 0:
        OUT.print_md("    No valid profiles!")
        return None
    
    t = DB.Transaction(famdoc, 'Extrusion')
    t.Start()
    try:
        sp = DB.SketchPlane.Create(famdoc, DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, plane_z)))
        ext = famdoc.FamilyCreate.NewExtrusion(is_solid, profiles, sp, abs(depth))
        t.Commit()
        OUT.print_md("    **LOOP EXTRUSION OK** ({})".format('void' if is_void else 'solid'))
        return ext
    except Exception as e:
        OUT.print_md("    **LOOP EXTRUSION FAILED**: {}".format(e))
        t.RollBack()
        return None

def _build_from_bbox(famdoc, form_data):
    """Create box from bounding box - ALWAYS try this"""
    is_void = _b(form_data.get('is_void'))
    is_solid = not is_void
    
    x0 = _f(form_data.get('bb_min_x'))
    y0 = _f(form_data.get('bb_min_y'))
    z0 = _f(form_data.get('bb_min_z'))
    x1 = _f(form_data.get('bb_max_x'))
    y1 = _f(form_data.get('bb_max_y'))
    z1 = _f(form_data.get('bb_max_z'))
    
    OUT.print_md("    BBox: min({},{},{}) max({},{},{})".format(x0,y0,z0,x1,y1,z1))
    
    if None in (x0,y0,z0,x1,y1,z1):
        OUT.print_md("    BBox has None values!")
        return None
    
    # Ensure minimum size
    dx = abs(x1 - x0)
    dy = abs(y1 - y0)
    dz = abs(z1 - z0)
    
    if dx < 0.01: 
        OUT.print_md("    BBox X too small, expanding")
        x1 = x0 + 1.0
    if dy < 0.01: 
        OUT.print_md("    BBox Y too small, expanding")
        y1 = y0 + 1.0
    if dz < 0.01: 
        OUT.print_md("    BBox Z too small, defaulting to 1")
        dz = 1.0
    else:
        dz = abs(z1 - z0)
    
    p0, p1 = DB.XYZ(x0, y0, z0), DB.XYZ(x1, y0, z0)
    p2, p3 = DB.XYZ(x1, y1, z0), DB.XYZ(x0, y1, z0)
    
    carr = DB.CurveArray()
    carr.Append(DB.Line.CreateBound(p0, p1))
    carr.Append(DB.Line.CreateBound(p1, p2))
    carr.Append(DB.Line.CreateBound(p2, p3))
    carr.Append(DB.Line.CreateBound(p3, p0))
    
    profiles = DB.CurveArrArray()
    profiles.Append(carr)
    
    t = DB.Transaction(famdoc, 'BBox Extrusion')
    t.Start()
    try:
        sp = DB.SketchPlane.Create(famdoc, DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0, 0, z0)))
        ext = famdoc.FamilyCreate.NewExtrusion(is_solid, profiles, sp, dz)
        t.Commit()
        OUT.print_md("    **BBOX EXTRUSION OK**")
        return ext
    except Exception as e:
        OUT.print_md("    **BBOX FAILED**: {}".format(e))
        t.RollBack()
        return None

def _process_family(project_doc, family_rec, global_template):
    fname = family_rec.get('family_name')
    category = family_rec.get('category')
    
    OUT.print_md("---")
    OUT.print_md("## Family: **{}**".format(fname))
    OUT.print_md("Category: {}".format(category))
    
    # Check if exists
    for f in DB.FilteredElementCollector(project_doc).OfClass(DB.Family):
        try:
            if f.Name == fname:
                OUT.print_md("Already exists - skipping proxy creation")
                return f
        except: pass
    
    # Get template
    templates = ["Metric Generic Model.rft", "Generic Model.rft"]
    tpath = global_template or _find_template_file(templates)
    if not tpath:
        OUT.print_md("**ERROR: No template found**")
        return None
    
    OUT.print_md("Template: {}".format(os.path.basename(tpath)))
    
    try:
        famdoc = app.NewFamilyDocument(tpath)
    except Exception as e:
        OUT.print_md("**ERROR creating family doc**: {}".format(e))
        return None
    
    # ========== GET GEOMETRY DATA ==========
    geometry = family_rec.get('geometry', {})
    forms_list = geometry.get('forms', [])
    loops_list = geometry.get('loops', [])
    
    OUT.print_md("### Data Found")
    OUT.print_md("- **Forms**: {}".format(len(forms_list)))
    OUT.print_md("- **Loops**: {}".format(len(loops_list)))
    
    # Dump raw data for debugging
    if forms_list:
        OUT.print_md("### Raw Forms Data")
        for i, f in enumerate(forms_list):
            OUT.print_md("```")
            OUT.print_md("Form[{}]: id={}, kind={}, is_void={}".format(i, f.get('id'), f.get('kind'), f.get('is_void')))
            OUT.print_md("  depth={}, bbox=({},{},{}) to ({},{},{})".format(
                f.get('depth'),
                f.get('bb_min_x'), f.get('bb_min_y'), f.get('bb_min_z'),
                f.get('bb_max_x'), f.get('bb_max_y'), f.get('bb_max_z')
            ))
            OUT.print_md("```")
    
    if loops_list:
        OUT.print_md("### Raw Loops Data (first 5)")
        for i, L in enumerate(loops_list[:5]):
            OUT.print_md("```")
            OUT.print_md("Loop[{}]: form_id={}, loop_index={}, type={}".format(
                i, L.get('form_id'), L.get('loop_index'), L.get('curve_type')))
            OUT.print_md("  start=({},{},{}), end=({},{},{})".format(
                L.get('start_x'), L.get('start_y'), L.get('start_z'),
                L.get('end_x'), L.get('end_y'), L.get('end_z')
            ))
            OUT.print_md("```")
    
    # Build lookup
    loops_by_form = {}
    for L in loops_list:
        fid = L.get('form_id')
        if fid is not None:
            try: fid = int(fid)
            except: pass
            loops_by_form.setdefault(fid, []).append(L)
    
    OUT.print_md("### Loops by Form ID")
    OUT.print_md("```")
    for k, v in loops_by_form.items():
        OUT.print_md("  form_id {}: {} loops".format(k, len(v)))
    OUT.print_md("```")
    
    # ========== BUILD GEOMETRY ==========
    built_count = 0
    
    if forms_list:
        OUT.print_md("### Building Geometry")
        
        # Process solids first, then voids
        solids = [f for f in forms_list if not _b(f.get('is_void'))]
        voids = [f for f in forms_list if _b(f.get('is_void'))]
        
        OUT.print_md("- {} solids, {} voids".format(len(solids), len(voids)))
        
        for form in solids + voids:
            form_id = form.get('id')
            try: form_id_int = int(form_id)
            except: form_id_int = form_id
            
            form_kind = form.get('kind', 'Unknown')
            is_void = _b(form.get('is_void'))
            
            OUT.print_md("#### Form {} ({}) {}".format(form_id, form_kind, '[VOID]' if is_void else '[SOLID]'))
            
            # Try loops first for Extrusions
            loops_for_form = loops_by_form.get(form_id) or loops_by_form.get(form_id_int) or []
            OUT.print_md("  Loops for this form: {}".format(len(loops_for_form)))
            
            built = False
            if form_kind == 'Extrusion' and loops_for_form:
                ext = _build_extrusion_from_loops(famdoc, loops_for_form, form)
                if ext:
                    built_count += 1
                    built = True
            
            # Always try bbox if loops didn't work
            if not built:
                OUT.print_md("  Trying BBox fallback...")
                ext = _build_from_bbox(famdoc, form)
                if ext:
                    built_count += 1
    else:
        OUT.print_md("### NO FORMS FOUND IN DATA!")
    
    # Create test box if nothing was built
    if built_count == 0:
        OUT.print_md("### Nothing built - creating TEST BOX as placeholder")
        _create_test_box(famdoc, 1.0)
    
    OUT.print_md("### Total built: {}".format(built_count))
    
    # ========== SAVE AND LOAD ==========
    temp_dir = os.getenv('TEMP') or os.getcwd()
    safe = (fname or 'Family').replace('\\','_').replace('/','_').replace(':','_')
    fam_path = os.path.join(temp_dir, '{}_{}.rfa'.format(safe, uuid.uuid4().hex[:8]))
    
    try:
        opts = DB.SaveAsOptions()
        opts.OverwriteExistingFile = True
        famdoc.SaveAs(fam_path, opts)
        OUT.print_md("Saved: {}".format(fam_path))
    except Exception as e:
        OUT.print_md("**SAVE ERROR**: {}".format(e))
        famdoc.Close(False)
        return None
    
    famdoc.Close(False)
    fam_loaded = _load_family(project_doc, fam_path)
    try: os.remove(fam_path)
    except: pass
    
    if fam_loaded:
        OUT.print_md("**LOADED INTO PROJECT OK**")
    else:
        OUT.print_md("**LOAD FAILED**")
    
    return fam_loaded

def run_import_family():
    OUT.print_md("# Family Data Import v10d")
    OUT.print_md("**Revit Version:** {}".format(REVIT_VERSION))
    
    jsonl_path = forms.pick_file(file_ext='jsonl', restore_dir=True, title='Pick Family Data JSONL')
    if not jsonl_path: return
    
    global_template = forms.pick_file(file_ext='rft', restore_dir=True, 
                                      title='Pick .rft template (or Cancel to auto-detect)')
    
    OUT.print_md("---")
    OUT.print_md("## Loading JSONL")
    records = _read_records_any(jsonl_path)
    families, types, instances = _split_records(records)
    
    OUT.print_md("- Records: {}".format(len(records)))
    OUT.print_md("- Families: {}".format(len(families)))
    OUT.print_md("- Types: {}".format(len(types)))
    OUT.print_md("- Instances: {}".format(len(instances)))
    
    # Show structure of first family
    if families:
        OUT.print_md("### First Family Record Structure")
        OUT.print_md("```")
        for k, v in families[0].items():
            if isinstance(v, list):
                OUT.print_md("  {}: list[{}]".format(k, len(v)))
            elif isinstance(v, dict):
                OUT.print_md("  {}: dict{{{}}}".format(k, ', '.join(v.keys())))
            else:
                OUT.print_md("  {}: {}".format(k, type(v).__name__))
        OUT.print_md("```")
    
    cache_fam = {}
    
    t = DB.Transaction(doc, 'Import Families (v10d)')
    t.Start()
    for frec in families:
        fam = _process_family(doc, frec, global_template)
        if fam:
            cache_fam[frec.get('family_name')] = fam
    t.Commit()
    
    # Place instances
    OUT.print_md("---")
    OUT.print_md("## Placing Instances")
    
    t = DB.Transaction(doc, 'Place Instances')
    t.Start()
    placed = 0
    for irec in instances:
        fname = irec.get('family_name')
        fam = cache_fam.get(fname)
        if not fam: continue
        sym_ids = list(fam.GetFamilySymbolIds())
        if not sym_ids: continue
        sym = doc.GetElement(sym_ids[0])
        if not sym.IsActive:
            sym.Activate()
            doc.Regenerate()
        x, y, z = _f(irec.get('point_x')), _f(irec.get('point_y')), _f(irec.get('point_z'))
        if None in (x, y, z): continue
        try:
            doc.Create.NewFamilyInstance(DB.XYZ(x,y,z), sym, DB.Structure.StructuralType.NonStructural)
            placed += 1
        except: pass
    t.Commit()
    
    OUT.print_md("Placed: {}".format(placed))
    OUT.print_md("---")
    OUT.print_md("# **COMPLETE**")

# ================================================================
#  ENGINE DISPATCH MAP
# ================================================================
ENGINES = {
    ("Native Duct", "Export"): run_export_native_duct,
    ("Native Duct", "Import"): run_import_native_duct,
    ("Family", "Export"):      run_export_family,
    ("Family", "Import"):      run_import_family,
}

DESCRIPTIONS = {
    ("Native Duct", "Export"):
        "Export native duct geometry and fitting data to JSONL + CSV.\n\n"
        "Captures connector positions, directions, coordinate systems, "
        "shapes, sizes, fitting classifications, and curve geometry.\n\n"
        "Selection: Pick, Box Select, All In View, Visible In View, "
        "or Current Selection.",
    ("Native Duct", "Import"):
        "Import native duct geometry from JSONL or CSV files.\n\n"
        "Creates straights, elbows, and transitions using Revit's "
        "routing preferences. Supports selective import by element "
        "type and configurable generation order.",
    ("Family", "Export"):
        "Export full MEP family data to JSONL + CSV sidecar files.\n\n"
        "Captures family parameters with default values per type, "
        "extrusion geometry (solids and voids), XYZ instance "
        "positioning with orientation, and MEP connectors.",
    ("Family", "Import"):
        "Import family data from JSONL and recreate proxy families.\n\n"
        "Rebuilds extrusion geometry from loop data or bounding box "
        "fallback, loads families into the project, and places "
        "instances at recorded positions.",
}


# ================================================================
#  WINDOW CLASS
# ================================================================
class DataWindow(object):
    """Main window for the unified Data tool."""

    def __init__(self, xaml_path):
        reader = StreamReader(xaml_path)
        self.window = XamlReader.Load(reader.BaseStream)
        reader.Close()

        self._pending_engine = None
        self._get_controls()
        self._wire_events()
        self._update_description()

    def _get_controls(self):
        find = self.window.FindName
        self.rb_duct      = find("RbNativeDuct")
        self.rb_family    = find("RbFamily")
        self.rb_export    = find("RbExport")
        self.rb_import    = find("RbImport")
        self.txt_desc     = find("TxtDescription")
        self.btn_run      = find("BtnRun")
        self.btn_close    = find("BtnClose")
        self.txt_status   = find("TxtStatus")

    def _wire_events(self):
        self.btn_run.Click   += self._on_run
        self.btn_close.Click += self._on_close
        self.window.KeyDown  += self._on_key
        self.rb_duct.Checked     += self._on_radio_changed
        self.rb_family.Checked   += self._on_radio_changed
        self.rb_export.Checked   += self._on_radio_changed
        self.rb_import.Checked   += self._on_radio_changed

    def _data_type(self):
        if self.rb_duct.IsChecked:
            return "Native Duct"
        return "Family"

    def _operation(self):
        if self.rb_export.IsChecked:
            return "Export"
        return "Import"

    def _on_radio_changed(self, sender, args):
        self._update_description()

    def _update_description(self):
        key = (self._data_type(), self._operation())
        self.txt_desc.Text = DESCRIPTIONS.get(key, "")
        self.btn_run.Content = "{} {} Data".format(
            self._operation(), self._data_type())

    def _on_run(self, sender, args):
        key = (self._data_type(), self._operation())
        engine = ENGINES.get(key)
        if not engine:
            self.txt_status.Text = "No engine for {}".format(key)
            return

        # For Export Native Duct with Pick/Box, we need Revit UI context
        # Use hide/show pattern for all engines since they may show dialogs
        self._pending_engine = engine
        self.window.Hide()

    def _on_close(self, sender, args):
        self.window.Close()

    def _on_key(self, sender, args):
        if args.Key == System.Windows.Input.Key.Escape:
            self.window.Close()

    def show(self):
        while True:
            self._pending_engine = None
            self.window.ShowDialog()
            if self._pending_engine:
                try:
                    self._pending_engine()
                except Exception as ex:
                    OUT.print_md(':x: **Engine error:** {}'.format(ex))
                self.txt_status.Text = "Ready for another operation"
                continue
            else:
                break


# ================================================================
#  MAIN
# ================================================================
def main():
    win = DataWindow(XAML_PATH)
    win.show()


if __name__ == "__main__":
    main()
