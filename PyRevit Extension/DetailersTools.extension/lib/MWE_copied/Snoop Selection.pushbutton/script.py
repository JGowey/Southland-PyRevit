# -*- coding: utf-8 -*-
# Snoop Selection — Advanced Inspector with Comparison Mode
#
# New in this build:
#   • COMPARISON MODE for multiple elements
#   • Three views: Normal (single element drill-down), Similarities, Differences
#   • Similarities shows properties/parameters with IDENTICAL values across all selected elements
#   • Differences shows properties/parameters with DIFFERENT values across selected elements
#   • Comparison views show side-by-side values for easy analysis

from pyrevit import revit, DB, forms
from System.Collections.ObjectModel import ObservableCollection
from System import String, Enum
import System, System.IO as SIO
from System.Reflection import BindingFlags

try:
    from Autodesk.Revit.UI.Selection import ObjectType
except:
    ObjectType = None

# ---------------------- runtime flags ----------------------
# Changed to False for better performance - methods and geometry can be very slow
# Users can enable these in the UI if needed for deep inspection
FORCE_INCLUDE_METHODS   = False  # ✅ Don't auto-invoke methods (can be slow)
FORCE_INCLUDE_NONGET    = False  # ✅ Skip non-getter methods
FORCE_INCLUDE_GEOMETRY  = False  # ✅ Skip geometry (VERY slow in Revit)
FORCE_INCLUDE_STATIC    = False
FORCE_INCLUDE_EVENTS    = False
FORCE_INCLUDE_INDEXERS  = True   # Keep indexers (usually fast)
FORCE_NONPUBLIC         = False
FORCE_ALLOW_PARAM_CALLS = False

# ---------------------- data model ----------------------
class RowItem(object):
    def __init__(self, member, value, typ, state="", raw=None):
        self.Member = member or ""
        self.Value  = value  or ""
        self.Type   = typ    or ""
        self.State  = state or ""
        self.Raw    = raw
        self.Details = None

class ComparisonRow(object):
    """Row for comparison views showing values across multiple elements"""
    def __init__(self, member, values_dict, typ, state=""):
        self.Member = member or ""
        self.Type = typ or ""
        self.State = state or ""
        # values_dict: {elem_id: value_string}
        self.Values = values_dict or {}
        # Create display columns for each element
        self.AllValues = ", ".join(["{}:{}".format(k, v) for k, v in sorted(values_dict.items())])

SAFE_METHOD_PREFIXES = ("Get", "get_")

# ---------------------- helpers ----------------------
def _is_primitive(v):
    try:
        import numbers
        return isinstance(v, (str, String, bool)) or isinstance(v, numbers.Number)
    except:
        return isinstance(v, (str, String, bool, int, float))

def _is_sequence(v):
    if v is None: return False
    if isinstance(v, (str, String)): return False
    return hasattr(v, "__iter__")

def _is_geometry_member(name):
    n = (name or "").lower()
    return ("geometry" in n) or ("solid" in n) or ("mesh" in n) or ("bbox" in n) or ("boundingbox" in n)

def _elem_name(el):
    try:
        n = getattr(el, "Name", None)
        if n: return n
        p = el.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
        if p: return p.AsValueString() or p.AsString() or ""
    except: pass
    return el.__class__.__name__

def _obj_name(obj):
    if isinstance(obj, DB.Element): return _elem_name(obj)
    try:
        n = getattr(obj, "Name", None)
        if n: return str(n)
    except: pass
    try:    return obj.__class__.__name__
    except: return str(type(obj))

def _format_xyz(xyz):
    try:
        return "({:.3f}, {:.3f}, {:.3f})".format(xyz.X, xyz.Y, xyz.Z)
    except:
        return str(xyz)

def _str(v):
    try:
        if v is None: return ""
        if isinstance(v, (str, String)): return str(v)
        try:
            import numbers
            if isinstance(v, numbers.Number): return str(v)
        except: pass
        if isinstance(v, DB.XYZ):       return _format_xyz(v)
        if isinstance(v, DB.ElementId): return "ElementId({})".format(v.IntegerValue)
        if isinstance(v, DB.Element):   return "{} [{}]".format(_elem_name(v), v.Id.IntegerValue)
        if isinstance(v, Enum):         return str(v)
        if _is_sequence(v):
            try:
                it = list(v)
                s  = ", ".join([_str(x) for x in it[:5]])
                if len(it) > 5: s += ", ... ({} items)".format(len(it))
                return "[" + s + "]"
            except: return str(v)
        return str(v)
    except:
        return ""

def _safe_eval(getter, allow_geom=False):
    try:
        val = getter()
        if not allow_geom and isinstance(val, DB.GeometryObject):
            return (None, "null")
        return (val, "")
    except Exception as ex:
        return (str(ex), "error")

def _classify_link(v):
    if v is None: return ("", None)
    if isinstance(v, DB.Element) or isinstance(v, DB.ElementId): return ("link", v)
    if _is_primitive(v): return ("", None)
    return ("link", v)

def _flags_mask(include_static, include_nonpublic):
    f = BindingFlags.Public | BindingFlags.Instance
    if include_static:   f |= BindingFlags.Static
    if include_nonpublic:f |= BindingFlags.NonPublic
    return f

# ---------------------- comparison helpers ----------------------
def _get_param_value_str(element, param_name):
    """Get parameter value as comparable string"""
    try:
        param = element.LookupParameter(param_name)
        if not param:
            return None
        
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString() or ""
        elif st == DB.StorageType.Integer:
            return str(param.AsInteger())
        elif st == DB.StorageType.Double:
            return param.AsValueString() or str(param.AsDouble())
        elif st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            return "ElementId({})".format(eid.IntegerValue) if eid else ""
    except:
        pass
    return None

def _get_property_value_str(element, prop_name):
    """Get property value as comparable string"""
    try:
        t = element.GetType()
        prop = t.GetProperty(prop_name)
        if prop and prop.CanRead:
            val = prop.GetValue(element, None)
            return _str(val)
    except:
        pass
    return None

def _compare_elements(elements):
    """
    Compare multiple elements and return:
    (similarities_dict, differences_dict)
    
    similarities_dict: {property_name: common_value}
    differences_dict: {property_name: {elem_id: value}}
    """
    if not elements or len(elements) < 2:
        return {}, {}
    
    # Collect all parameter names
    all_param_names = set()
    for elem in elements:
        try:
            for p in elem.Parameters:
                if p.Definition:
                    all_param_names.add(p.Definition.Name)
        except:
            pass
    
    # Collect all property names (common ones)
    common_props = ["Name", "Category", "Level", "WorksetId"]
    
    similarities = {}
    differences = {}
    
    # Compare parameters
    for param_name in all_param_names:
        values = {}
        for elem in elements:
            val = _get_param_value_str(elem, param_name)
            if val is not None:
                values[elem.Id.IntegerValue] = val
        
        if len(values) > 0:
            unique_vals = set(values.values())
            if len(unique_vals) == 1:
                # All same
                similarities[param_name] = list(unique_vals)[0]
            else:
                # Different
                differences[param_name] = values
    
    # Compare common properties
    for prop_name in common_props:
        values = {}
        for elem in elements:
            val = _get_property_value_str(elem, prop_name)
            if val is not None:
                values[elem.Id.IntegerValue] = val
        
        if len(values) > 0:
            unique_vals = set(values.values())
            if len(unique_vals) == 1:
                similarities[prop_name] = list(unique_vals)[0]
            else:
                differences[prop_name] = values
    
    return similarities, differences

# ---------------------- view picking for BoundingBox(View) ----------------------
def _all_candidate_views(doc):
    results, seen_ids = [], set()
    try:
        av = getattr(revit, "active_view", None)
        if isinstance(av, DB.View) and not av.IsTemplate:
            results.append((av, u"Active: {}".format(av.Name)))
            seen_ids.add(av.Id.IntegerValue)
    except: pass
    try:
        uiviews = revit.uidoc.GetOpenUIViews()
        for uiv in uiviews:
            try:
                v = doc.GetElement(uiv.ViewId)
                if v and not v.IsTemplate:
                    iid = v.Id.IntegerValue
                    if iid not in seen_ids:
                        results.append((v, u"Open: {}".format(v.Name)))
                        seen_ids.add(iid)
            except: pass
    except: pass
    try:
        col = DB.FilteredElementCollector(doc).OfClass(DB.View)
        for v in col:
            if v and not v.IsTemplate:
                iid = v.Id.IntegerValue
                if iid not in seen_ids:
                    results.append((v, u"{} — {}".format(v.ViewType, v.Name)))
                    seen_ids.add(iid)
    except: pass
    return results

def _resolve_bbox_with_active_or_prompt(owner_el, getter_methodinfo):
    try:
        av = getattr(revit, "active_view", None)
        if isinstance(av, DB.View):
            try:
                bb = getter_methodinfo.Invoke(owner_el, System.Array[object]([av]))
                if bb is not None:
                    return bb, u"Active view"
            except: pass
        doc = revit.doc
        choices = ["Model (no view)"]; mapping = {"Model (no view)": None}
        for v, label in _all_candidate_views(doc):
            choices.append(label); mapping[label] = v
        picked = forms.SelectFromList.show(choices, title="Pick a View for BoundingBox",
                                           multiselect=False, button_name='Use')
        if not picked: return None, None
        view_obj = mapping.get(picked, None)
        bb = getter_methodinfo.Invoke(owner_el,
                                      System.Array[object]([view_obj]) if view_obj is not None
                                      else System.Array[object]([None]))
        return bb, picked
    except:
        return None, None

# ---------------------- row builders (existing single-element logic) ----------------------
def _rows_for_sequence(seq, q=""):
    q = (q or "").lower()
    rows = [RowItem("Sequence", "", "", "section")]
    try: rows.append(RowItem("(count)", str(len(seq)), "int", ""))
    except: pass
    try: items = list(seq)
    except:
        items = []
        try:
            for it in seq: items.append(it)
        except: pass
    for i, item in enumerate(items):
        m, sval, typ = "[{}]".format(i), _str(item), _obj_name(item)
        st, drill = _classify_link(item)
        if q and (q not in m.lower() and q not in sval.lower() and q not in (typ or "").lower()): continue
        rows.append(RowItem(m, sval, typ, st, drill))
    return rows

def _rows_fields(t, obj, include_static, include_nonpublic, q=""):
    q = (q or "").lower()
    flags = _flags_mask(include_static, include_nonpublic)
    try: fields = list(t.GetFields(flags))
    except: fields = []
    if not fields: return []
    rows = [RowItem("Fields", "", "", "section")]
    for f in sorted(fields, key=lambda x: x.Name.lower()):
        try:
            val, st = _safe_eval(lambda f=f: f.GetValue(None if f.IsStatic else obj), allow_geom=True)
            sval = _str(val); state, drill = _classify_link(val)
            if st: state = st
            elif sval == "": state = "null"
            if q and (q not in f.Name.lower() and q not in (sval or "").lower() and q not in f.FieldType.Name.lower()): continue
            rows.append(RowItem(f.Name, sval, f.FieldType.Name, state, drill))
        except Exception as ex:
            msg = str(ex)
            if q and (q not in f.Name.lower() and q not in msg.lower()): continue
            rows.append(RowItem(f.Name, msg, f.FieldType.Name, "error"))
    return rows

def _rows_properties(t, obj, include_geom, include_static, include_nonpublic, q=""):
    q = (q or "").lower()
    flags = _flags_mask(include_static, include_nonpublic)
    props = []
    try:
        for p in t.GetProperties(flags):
            try:
                if not p.CanRead: continue
                if p.GetIndexParameters().Length != 0: continue
                if not include_geom and _is_geometry_member(p.Name): continue
                def _mk(prop=p):
                    return lambda: prop.GetValue(None if (include_static and prop.GetGetMethod(True).IsStatic) else obj, None)
                val, st = _safe_eval(_mk(), allow_geom=include_geom)
                typ = p.PropertyType.Name if p.PropertyType is not None else ""
                props.append((p.Name, val, typ, st))
            except: pass
    except: pass
    if not props: return []
    props.sort(key=lambda x: x[0].lower())
    rows = [RowItem("Properties", "", "", "section")]
    for name, val, typ, st in props:
        sval = _str(val)
        state = st if st else ("null" if sval == "" else "")
        link_state, drill = _classify_link(val)
        if link_state == "link" and st == "": state = "link"
        if q and (q not in name.lower() and q not in (sval or "").lower() and q not in (typ or "").lower()): continue
        rows.append(RowItem(name, sval, typ, state, drill))
    return rows

def _collect_string_keys_for(obj, cap=5):
    try:
        keys = getattr(obj, "Keys", None)
        if keys is not None:
            out = []
            for i, k in enumerate(list(keys)):
                if i >= cap: break
                if isinstance(k, (str, String)): out.append(k)
            if out: return out
    except: pass
    try:
        out = []
        for i, kv in enumerate(obj):
            if i >= cap: break
            try:
                k = getattr(kv, "Key", None)
                if isinstance(k, (str, String)): out.append(k)
            except: pass
        if out: return out
    except: pass
    return []

def _rows_indexers(t, obj, include_geom, include_static, include_nonpublic, q=""):
    q = (q or "").lower()
    flags = _flags_mask(include_static, include_nonpublic)
    idxs = []
    try:
        for p in t.GetProperties(flags):
            if not p.CanRead: continue
            if p.GetIndexParameters().Length == 0: continue
            sig = "{}({})".format(p.Name, ", ".join([ip.ParameterType.Name for ip in p.GetIndexParameters()]))
            idxs.append((p, sig))
    except: pass
    if not idxs: return []
    rows = [RowItem("Indexers", "", "", "section")]
    for p, sig in sorted(idxs, key=lambda x: x[1].lower()):
        try:
            preview = ""; state = ""
            ip = p.GetIndexParameters()
            if ip.Length == 1 and ip[0].ParameterType.Name in ("Int32","Int16"):
                count = None
                for cname in ("Count","Length"):
                    try:
                        count = getattr(obj, cname, None)
                        if isinstance(count, int): break
                    except: pass
                if isinstance(count, int):
                    end = min(count, 5); samples = []
                    for i in range(end):
                        def _mk(i=i, prop=p): return lambda: prop.GetValue(obj, System.Array[object]([i]))
                        val, st = _safe_eval(_mk(), allow_geom=include_geom)
                        samples.append("[{}]={}".format(i, _str(val)))
                    preview = "; ".join(samples)
                else:
                    state = "null"
            elif ip.Length == 1 and ip[0].ParameterType.Name in ("String","string"):
                keys = _collect_string_keys_for(obj, cap=5)
                if keys:
                    samples = []
                    for k in keys:
                        def _mk(k=k, prop=p): return lambda: prop.GetValue(obj, System.Array[object]([k]))
                        val, st = _safe_eval(_mk(), allow_geom=include_geom)
                        samples.append("[\"{}\"]{}".format(k, _str(val)))
                    preview = "; ".join(samples)
                else:
                    state = "null"
            if q and (q not in p.Name.lower() and q not in sig.lower() and q not in preview.lower()): continue
            rows.append(RowItem(sig, preview, p.PropertyType.Name, state))
        except Exception as ex:
            msg = str(ex)
            if q and (q not in p.Name.lower() and q not in sig.lower() and q not in msg.lower()): continue
            rows.append(RowItem(sig, msg, p.PropertyType.Name, "error"))
    return rows

def _rows_methods(t, obj, include_geom, include_static, include_nonpublic, include_nonget, allow_invoke_params, q=""):
    q = (q or "").lower()
    flags = _flags_mask(include_static, include_nonpublic)
    meths = []
    try:
        for m in t.GetMethods(flags):
            try:
                name = m.Name
                pars = m.GetParameters()
                is_zero = pars.Length == 0
                if is_zero:
                    if (not include_nonget) and (not name.startswith(SAFE_METHOD_PREFIXES)): continue
                    if not include_geom and _is_geometry_member(name): continue
                    rtype = ""
                    try: rtype = m.ReturnType.Name
                    except: pass
                    if rtype and rtype.lower() != "void":
                        def _mk(mm=m): return lambda: mm.Invoke(None if (include_static and mm.IsStatic) else obj, None)
                        val, st = _safe_eval(_mk(), allow_geom=include_geom)
                    else:
                        val, st = ("(void)", "null")
                    meths.append((name + "()", val, rtype or "", st, None))
                else:
                    if not allow_invoke_params: continue
                    simple = True; ptypes = []
                    for pinfo in pars:
                        ptypes.append(pinfo.ParameterType)
                        pn = pinfo.ParameterType.Name.lower()
                        if pn not in ("string","int32","int16","double","single","boolean"):
                            simple = False
                    state = "link" if simple else "null"
                    meths.append(("{}({})".format(name, ", ".join([p.ParameterType.Name for p in pars])),
                                  "(double-click to invoke)" if simple else "",
                                  m.ReturnType.Name if m.ReturnType is not None else "",
                                  "", ("method", obj, m, ptypes) if simple else None))
            except: pass
    except: pass
    if not meths: return []
    meths.sort(key=lambda x: x[0].lower())
    rows = [RowItem("Methods", "", "", "section")]
    for disp, val, typ, st, raw in meths:
        sval = _str(val)
        state = st if st else ("null" if sval == "" else "")
        if raw is not None and state == "": state = "link"
        if q and (q not in disp.lower() and q not in (sval or "").lower() and q not in (typ or "").lower()): continue
        rows.append(RowItem(disp, sval, typ, state, raw))
    return rows

def _rows_events(t, include_static, include_nonpublic, q=""):
    q = (q or "").lower()
    flags = _flags_mask(include_static, include_nonpublic)
    evs = []
    try:
        for ev in t.GetEvents(flags):
            et = None
            try: et = ev.EventHandlerType
            except: pass
            evs.append((ev.Name, et.Name if et else ""))
    except: pass
    if not evs: return []
    evs.sort(key=lambda x: x[0].lower())
    rows = [RowItem("Events", "", "", "section")]
    for name, typ in evs:
        if q and (q not in name.lower() and q not in (typ or "").lower()): continue
        rows.append(RowItem(name, "", typ, "null"))
    return rows

def _rows_parameters(el, q=""):
    q = (q or "").lower()
    rows = [RowItem("Parameters", "", "", "section")]
    try:
        def pstr(p):
            st = p.StorageType
            if st == DB.StorageType.String:   return p.AsString() or ""
            if st == DB.StorageType.Integer:  return str(p.AsInteger())
            if st == DB.StorageType.Double:   return p.AsValueString() or str(p.AsDouble())
            if st == DB.StorageType.ElementId:
                eid = p.AsElementId()
                return "ElementId({})".format(eid.IntegerValue) if eid else ""
            return ""
        params = list(el.Parameters)
        params.sort(key=lambda pp: (pp.Definition.Name if pp.Definition else "").lower())
        for p in params:
            try:
                name = p.Definition.Name if p.Definition else "(no name)"
                sval = pstr(p)
                typ  = str(p.StorageType)
                state = "null" if sval in ("", None) else ""
                raw = p.AsElementId() if p.StorageType == DB.StorageType.ElementId else None
                if isinstance(raw, DB.ElementId) and raw.IntegerValue != -1: state = "link"
                if q and (q not in name.lower() and q not in (sval or "").lower() and q not in (typ or "").lower()): continue
                rows.append(RowItem(name, sval, typ, state, raw))
            except: pass
    except: pass
    return rows

def _rows_bboxxyz_computed(bb, q=""):
    q = (q or "").lower()
    rows = [RowItem("BoundingBoxXYZ (Computed)", "", "", "section")]
    try:
        mn = bb.Min; mx = bb.Max
        dx, dy, dz = (mx.X - mn.X), (mx.Y - mn.Y), (mx.Z - mn.Z)
        centroid = DB.XYZ((mn.X + mx.X)/2.0, (mn.Y + mx.Y)/2.0, (mn.Z + mx.Z)/2.0)
        sa = 2.0 * (dx*dy + dy*dz + dz*dx)
        vol = max(dx,0.0) * max(dy,0.0) * max(dz,0.0)
        verts = [
            DB.XYZ(mn.X, mn.Y, mn.Z), DB.XYZ(mx.X, mn.Y, mn.Z),
            DB.XYZ(mx.X, mx.Y, mn.Z), DB.XYZ(mn.X, mx.Y, mn.Z),
            DB.XYZ(mn.X, mn.Y, mx.Z), DB.XYZ(mx.X, mn.Y, mx.Z),
            DB.XYZ(mx.X, mx.Y, mx.Z), DB.XYZ(mn.X, mx.Y, mx.Z),
        ]
        items = [
            ("Centroid", _str(centroid), "XYZ", centroid),
            ("SurfaceArea", str(sa), "double", sa),
            ("Volume", str(vol), "double", vol),
            ("Vertices", "[{} points]".format(len(verts)), "List<XYZ>", verts),
        ]
        for name, sval, typ, raw in items:
            if q and (q not in name.lower() and q not in sval.lower() and q not in typ.lower()): continue
            rows.append(RowItem(name, sval, typ, "link", raw))
    except: pass
    return rows

def _rows_adapters(obj, include_geom, include_nonpublic, q=""):
    rows = []
    if not include_geom:
        return rows
    q = (q or "").lower()

    try:
        if isinstance(obj, DB.Element):
            flags = _flags_mask(False, include_nonpublic)
            getter = None
            for prop in obj.GetType().GetProperties(flags):
                try:
                    if prop.Name != "BoundingBox": continue
                    ip = prop.GetIndexParameters()
                    if ip.Length == 1 and ip[0].ParameterType == DB.View:
                        getter = prop.GetGetMethod(True)
                        break
                except: pass
            if getter is not None:
                bb_active = None
                try:
                    av = getattr(revit, "active_view", None)
                    if isinstance(av, DB.View):
                        bb_active = getter.Invoke(obj, System.Array[object]([av]))
                except: bb_active = None

                if bb_active is not None:
                    rows.append(RowItem("BoundingBox (Active view)", "BoundingBoxXYZ", "BoundingBoxXYZ", "link", bb_active))
                else:
                    rows.append(RowItem("BoundingBox (Choose view…)", "Pick view (or Model)", "BoundingBoxXYZ",
                                        "link", ("bbox_prompt", obj, getter)))
    except:
        pass

    try:
        if isinstance(obj, DB.BoundingBoxXYZ):
            rows += _rows_bboxxyz_computed(obj, q)
    except:
        pass

    return rows

def _rows_for_object(obj, flags, q=""):
    (include_methods, include_nonget, include_geom, include_static,
     include_events, include_indexers, include_nonpublic, allow_invoke_params) = flags
    q = (q or "").lower()

    if _is_sequence(obj):
        return _rows_for_sequence(obj, q)

    rows = [RowItem("Object", "", "", "section")]
    if isinstance(obj, DB.Element):
        basics = [
            ("ElementId", obj.Id.IntegerValue, obj.Id),
            ("UniqueId", getattr(obj, "UniqueId", ""), None),
            ("Class", obj.__class__.__name__, None),
            ("Category", obj.Category.Name if obj.Category else "", None),
            ("Name", _elem_name(obj), None),
        ]
    else:
        basics = [("Class", obj.__class__.__name__, None),
                  ("Type",  str(type(obj)), None),
                  ("ToString", str(obj), None)]
    for k, v, raw in basics:
        sv = _str(v); state = "null" if sv == "" else ""
        if isinstance(raw, DB.ElementId): state = "link"
        rows.append(RowItem(k, sv, type(v).__name__, state, raw))

    t = obj.GetType() if hasattr(obj, "GetType") else type(obj)
    rows += _rows_adapters(obj, include_geom, include_nonpublic, q)
    rows += _rows_fields(t, obj, include_static, include_nonpublic, q)
    rows += _rows_properties(t, obj, include_geom, include_static, include_nonpublic, q)
    if include_indexers:
        rows += _rows_indexers(t, obj, include_geom, include_static, include_nonpublic, q)
    if include_methods or include_nonget or allow_invoke_params:
        rows += _rows_methods(t, obj, include_geom, include_static, include_nonpublic, include_nonget, allow_invoke_params, q)
    if isinstance(obj, DB.Element): rows += _rows_parameters(obj, q)
    if include_events: rows += _rows_events(t, include_static, include_nonpublic, q)
    return rows

def _bind_rows_for(obj, flags, q):
    oc = ObservableCollection[RowItem]()
    for r in _rows_for_object(obj, flags, q):
        oc.Add(r)
    return oc

def _parse_arg(text, target_type_name):
    if text is None: return None
    s = text.strip()
    tn = (target_type_name or "").lower()
    try:
        if tn in ("string",): return s
        if tn in ("int32","int16"): return int(s)
        if tn in ("double","single"): return float(s)
        if tn in ("boolean","bool"): return s.lower() in ("true","1","yes","y","on")
    except: pass
    return s

def _invoke_with_params(owner, method_info, param_types):
    inputs = []
    for i, pt in enumerate(param_types):
        tname = pt.Name if getattr(pt, "Name", None) else str(pt)
        resp = forms.ask_for_string(prompt="Argument {} ({}):".format(i+1, tname),
                                    title="Invoke Method",
                                    default="")
        if resp is None: return None, "Cancelled"
        inputs.append(_parse_arg(resp, tname))
    arr = System.Array[object](inputs)
    try:
        is_static = getattr(method_info, "IsStatic", False)
        target = None if is_static else owner
        return method_info.Invoke(target, arr), ""
    except Exception as ex:
        return str(ex), "error"

# ---------------------- WPF Window ----------------------
class ProWindow(forms.WPFWindow):
    def __init__(self, xaml_path, elements):
        forms.WPFWindow.__init__(self, xaml_path)

        self._selection = list(elements or [])
        if not self._selection:
            forms.alert("No elements were selected.", title="Snoop Selection")
            self.Close(); return

        self._nav = []
        self._comparison_mode = "Normal"  # "Normal", "Similarities", "Differences"

        # Build left tree
        from System.Windows.Controls import TreeViewItem
        self.ElemTree.Items.Clear()
        for el in self._selection:
            n = TreeViewItem()
            n.Header = "{} (Id: {})".format(_elem_name(el), el.Id.IntegerValue)
            n.Tag    = el
            self.ElemTree.Items.Add(n)
        self.ElemTree.SelectedItemChanged += self._on_elem_change

        if hasattr(self, "SearchBox"): self.SearchBox.TextChanged += self._on_flags_changed
        if hasattr(self, "MemberGrid"):
            self.MemberGrid.MouseDoubleClick += self._on_grid_double
        if hasattr(self, "CopyBtn"):   self.CopyBtn.Click += self._on_copy
        if hasattr(self, "ExportBtn"): self.ExportBtn.Click += self._on_export
        if hasattr(self, "BackBtn"):   self.BackBtn.Click  += self._on_back
        if hasattr(self, "HomeBtn"):   self.HomeBtn.Click  += self._on_home

        # Comparison mode buttons
        if hasattr(self, "NormalModeBtn"): self.NormalModeBtn.Click += self._on_normal_mode
        if hasattr(self, "SimilaritiesBtn"): self.SimilaritiesBtn.Click += self._on_similarities_mode
        if hasattr(self, "DifferencesBtn"): self.DifferencesBtn.Click += self._on_differences_mode

        # Enable/disable comparison buttons
        self._update_comparison_buttons()

        try: self.ElemTree.Items[0].IsSelected = True
        except: pass
        self._force_bind_current()

    def _flags(self):
        q = self.SearchBox.Text if hasattr(self, "SearchBox") else ""
        return (FORCE_INCLUDE_METHODS, FORCE_INCLUDE_NONGET, FORCE_INCLUDE_GEOMETRY,
                FORCE_INCLUDE_STATIC, FORCE_INCLUDE_EVENTS, FORCE_INCLUDE_INDEXERS,
                FORCE_NONPUBLIC, FORCE_ALLOW_PARAM_CALLS), q

    def _update_nav_buttons(self):
        n = len(self._nav)
        if hasattr(self, "BackBtn"): self.BackBtn.IsEnabled = n > 0
        if hasattr(self, "HomeBtn"): self.HomeBtn.IsEnabled = n > 0

    def _update_comparison_buttons(self):
        """Enable comparison buttons only when multiple elements selected"""
        enable = len(self._selection) > 1
        if hasattr(self, "SimilaritiesBtn"): self.SimilaritiesBtn.IsEnabled = enable
        if hasattr(self, "DifferencesBtn"): self.DifferencesBtn.IsEnabled = enable

    def _bind(self, obj, title=None):
        flags, q = self._flags()
        if title is None:
            title = "Details - {}".format(_obj_name(obj))
            if isinstance(obj, DB.Element): title += " (Id: {})".format(obj.Id.IntegerValue)
        if hasattr(self, "RightTitle"): self.RightTitle.Text = title
        if hasattr(self, "MemberGrid"):
            self.MemberGrid.ItemsSource = _bind_rows_for(obj, flags, q)
            try:
                cnt = self.MemberGrid.ItemsSource.Count
                if hasattr(self, "StatusText"): self.StatusText.Text = "{} rows".format(cnt)
            except: pass

    def _bind_comparison(self, mode):
        """Bind comparison view (Similarities or Differences)"""
        if len(self._selection) < 2:
            return

        similarities, differences = _compare_elements(self._selection)
        
        oc = ObservableCollection[RowItem]()
        
        if mode == "Similarities":
            oc.Add(RowItem("Similarities ({} elements)".format(len(self._selection)), "", "", "section"))
            if not similarities:
                oc.Add(RowItem("(no common values)", "All properties differ", "", "null"))
            else:
                for key in sorted(similarities.keys()):
                    val = similarities[key]
                    oc.Add(RowItem(key, val, "Common", ""))
        
        elif mode == "Differences":
            oc.Add(RowItem("Differences ({} elements)".format(len(self._selection)), "", "", "section"))
            if not differences:
                oc.Add(RowItem("(no differences)", "All properties are identical", "", "null"))
            else:
                for key in sorted(differences.keys()):
                    values_dict = differences[key]
                    # Format: "Id:Value, Id:Value, ..."
                    value_str = " | ".join(["Id{}: {}".format(eid, val) for eid, val in sorted(values_dict.items())])
                    oc.Add(RowItem(key, value_str, "Varying", ""))
        
        if hasattr(self, "MemberGrid"):
            self.MemberGrid.ItemsSource = oc
        if hasattr(self, "RightTitle"):
            self.RightTitle.Text = "{} - {} elements".format(mode, len(self._selection))
        if hasattr(self, "StatusText"):
            self.StatusText.Text = "{} rows".format(len(oc))

    def _current_tree_obj(self):
        try:
            sel = self.ElemTree.SelectedItem
            return getattr(sel, "Tag", None)
        except: return None

    def _force_bind_current(self):
        if self._comparison_mode in ("Similarities", "Differences"):
            self._bind_comparison(self._comparison_mode)
        else:
            target = self._nav[-1][0] if self._nav else self._current_tree_obj()
            if target is not None:
                self._bind(target); self._update_nav_buttons()

    def _ensure_details(self, rowitem):
        if rowitem is None or rowitem.State != "link" or rowitem.Details is not None:
            return
        flags, _ = self._flags()
        target = rowitem.Raw

        if isinstance(target, tuple) and len(target) == 3 and target[0] == "bbox_prompt":
            _, owner, getter = target
            bb, used = _resolve_bbox_with_active_or_prompt(owner, getter)
            rowitem.Details = _bind_rows_for(bb, flags, "") if bb is not None else ObservableCollection[RowItem]()
            return

        rowitem.Details = _bind_rows_for(target, flags, "") if target is not None else ObservableCollection[RowItem]()

    def ExpandBtn_Click(self, sender, args):
        try:
            data_item = sender.Tag
            if data_item is None: return
            self._ensure_details(data_item)
            row = self.MemberGrid.ItemContainerGenerator.ContainerFromItem(data_item)
            if row is None:
                self.MemberGrid.ScrollIntoView(data_item)
                row = self.MemberGrid.ItemContainerGenerator.ContainerFromItem(data_item)
            if row is not None:
                from System.Windows import Visibility
                row.DetailsVisibility = Visibility.Visible if row.DetailsVisibility != Visibility.Visible else Visibility.Collapsed
        except:
            pass

    def _on_elem_change(self, sender, args):
        el = self._current_tree_obj()
        if el is not None:
            self._nav = []
            self._comparison_mode = "Normal"
            self._bind(el); self._update_nav_buttons()

    def _on_flags_changed(self, sender, args):
        self._force_bind_current()

    def _clear_search(self):
        if hasattr(self, "SearchBox") and self.SearchBox.Text:
            self.SearchBox.Text = ""

    def _on_normal_mode(self, sender, args):
        self._comparison_mode = "Normal"
        self._nav = []
        el = self._current_tree_obj()
        if el: self._bind(el)
        self._update_nav_buttons()

    def _on_similarities_mode(self, sender, args):
        if len(self._selection) < 2:
            return
        self._comparison_mode = "Similarities"
        self._nav = []
        self._bind_comparison("Similarities")
        self._update_nav_buttons()

    def _on_differences_mode(self, sender, args):
        if len(self._selection) < 2:
            return
        self._comparison_mode = "Differences"
        self._nav = []
        self._bind_comparison("Differences")
        self._update_nav_buttons()

    def _on_grid_double(self, sender, args):
        # Disable drill-down in comparison modes
        if self._comparison_mode in ("Similarities", "Differences"):
            return

        row = self.MemberGrid.SelectedItem if hasattr(self, "MemberGrid") else None
        if row is None: return

        if isinstance(row.Raw, tuple) and len(row.Raw) == 3 and row.Raw[0] == "bbox_prompt":
            _, owner, getter = row.Raw
            bb, used = _resolve_bbox_with_active_or_prompt(owner, getter)
            if bb is None: return
            self._clear_search()
            curr = self._current_tree_obj() if not self._nav else self._nav[-1][0]
            curr_title = (self.RightTitle.Text if hasattr(self, "RightTitle") else "Details")
            label = "BoundingBox ({})".format(used or "chosen")
            self._nav.append((curr, curr_title))
            self._bind(bb, "{} -> {}".format(curr_title.replace("Details - ", ""), label))
            self._update_nav_buttons()
            return

        if isinstance(row.Raw, tuple) and len(row.Raw) == 4 and row.Raw[0] == "method":
            _, owner, minfo, ptypes = row.Raw
            result, st = _invoke_with_params(owner, minfo, ptypes)
            if st == "error": forms.alert("Invocation failed:\n\n{}".format(result), title="Invoke Method"); return
            if st == "Cancelled" or result is None: return
            self._clear_search()
            curr = self._current_tree_obj() if not self._nav else self._nav[-1][0]
            curr_title = (self.RightTitle.Text if hasattr(self, "RightTitle") else "Details")
            self._nav.append((curr, curr_title))
            self._bind(result, "{} -> {}".format(curr_title.replace("Details - ", ""), row.Member))
            self._update_nav_buttons(); return

        if isinstance(row.Raw, DB.ElementId):
            try:
                el = revit.doc.GetElement(row.Raw)
                if el:
                    revit.get_selection().set_to(el.Id)
                    try: revit.uidoc.ShowElements(el.Id)
                    except: pass
                return
            except: pass
        if isinstance(row.Raw, DB.Element):
            try:
                revit.get_selection().set_to(row.Raw.Id)
                try: revit.uidoc.ShowElements(row.Raw.Id)
                except: pass
            except: pass

        if row.State == "link" and row.Raw is not None and not isinstance(row.Raw, DB.ElementId):
            self._clear_search()
            curr = self._current_tree_obj() if not self._nav else self._nav[-1][0]
            curr_title = (self.RightTitle.Text if hasattr(self, "RightTitle") else "Details")
            self._nav.append((curr, curr_title))
            self._bind(row.Raw, "{} -> {}".format(curr_title.replace("Details - ", ""), row.Member))
            self._update_nav_buttons()

    def _on_back(self, sender, args):
        if not self._nav: return
        self._nav.pop()
        tgt, title = (self._nav[-1] if self._nav else (self._current_tree_obj(), None))
        if tgt is not None: self._bind(tgt, title)
        self._update_nav_buttons()

    def _on_home(self, sender, args):
        if not self._nav: return
        first_obj, first_title = self._nav[0]
        self._nav = []
        self._bind(first_obj, first_title)
        self._update_nav_buttons()

    def _on_copy(self, sender, args):
        if not hasattr(self, "MemberGrid"): return
        sel = list(self.MemberGrid.SelectedItems) if self.MemberGrid.SelectedItems else list(self.MemberGrid.Items)
        lines = ["Member\tValue\tType"]
        for r in sel:
            if r.State == "section": continue
            lines.append(u"{}\t{}\t{}".format(r.Member, r.Value, r.Type))
        s = "\n".join(lines)
        try:
            System.Windows.Clipboard.SetText(s)
            if hasattr(self, "StatusText"): self.StatusText.Text = "Copied {} rows".format(len(sel))
        except:
            if hasattr(self, "StatusText"): self.StatusText.Text = "Clipboard copy failed"

    def _on_export(self, sender, args):
        if not hasattr(self, "MemberGrid"): return
        try:
            tmp = SIO.Path.Combine(SIO.Path.GetTempPath(), "snoop_selection_export.csv")
            with open(tmp, "wb") as f:
                f.write(b"Member,Value,Type\r\n")
                for r in list(self.MemberGrid.Items):
                    if r.State == "section": continue
                    line = u"{},{},{}\r\n".format(
                        (r.Member or "").replace(",", ";"),
                        (r.Value  or "").replace(",", ";"),
                        (r.Type   or "").replace(",", ";"))
                    f.write(line.encode("utf-8"))
            if hasattr(self, "StatusText"): self.StatusText.Text = "Exported CSV: " + tmp
        except Exception as ex:
            if hasattr(self, "StatusText"): self.StatusText.Text = "Export failed: " + str(ex)

def _pick_elements_if_needed():
    """Return a list of elements. If nothing is preselected, let user pick."""
    sel = list(revit.get_selection() or [])
    if sel:
        return sel
    try:
        if ObjectType is None:
            raise Exception("Selection API not available")
        refs = revit.uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Select elements to snoop, then click Finish (or press ESC to cancel)."
        )
        return [revit.doc.GetElement(r.ElementId) for r in refs]
    except:
        return []

# ---------------------- run ----------------------
try:
    elements = _pick_elements_if_needed()
    if not elements:
        forms.alert("No elements selected.", title="Snoop Selection")
    else:
        win = ProWindow("window.xaml", elements)
        win.show_dialog()
except Exception as ex:
    forms.alert("Failed to open Snoop Selection: {}".format(ex), title="Snoop Selection")