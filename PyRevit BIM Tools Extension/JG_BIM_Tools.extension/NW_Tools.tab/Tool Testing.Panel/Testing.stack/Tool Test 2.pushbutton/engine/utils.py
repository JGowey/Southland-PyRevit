# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import os, io, json
from pyrevit import revit, DB

# ---------- basics ----------
def safe_str(x):
    try:
        if x is None: return u""
        if isinstance(x, unicode): return x
        return unicode(x)
    except:
        try: return unicode(str(x))
        except: return u"<unprintable>"

def _bundle_root():
    return os.path.dirname(os.path.dirname(__file__))

# external log callback (window injects this if desired)
LOG_FN = [None]   # mutable cell
def _log(msg):
    fn = LOG_FN[0]
    try:
        if fn: fn(safe_str(msg))
    except: pass

# ---------- rules discovery ----------
def _rules_folder_candidates():
    root = _bundle_root()
    return [os.path.join(root,'rules'), os.path.join(root,'config','rules')]

def _exec_rule_file(path):
    try:
        with io.open(path,'r',encoding='utf-8-sig',errors='ignore') as f:
            src=f.read()
        g,l={'__builtins__':{}},{}
        exec src in g,l
        rule=l.get('RULE') or g.get('RULE')
        if isinstance(rule,dict):
            r=dict(rule); r['path']=path; return r
    except Exception:
        pass
    return None

def _infer_group_from_path(path):
    try:
        root=None
        for c in _rules_folder_candidates():
            if path.startswith(c): root=c; break
        if not root: return u""
        rel=path[len(root):].lstrip("\\/")
        parts=rel.split(os.sep)
        if len(parts)>=2: return safe_str(parts[0])
    except Exception:
        pass
    return u""

def discover_rules_from_folder():
    out=[]
    for folder in _rules_folder_candidates():
        if not os.path.isdir(folder): continue
        for fn in os.listdir(folder):
            full=os.path.join(folder,fn)
            if os.path.isdir(full):
                for sub in os.listdir(full):
                    if sub.lower().endswith('.py'):
                        r=_exec_rule_file(os.path.join(full,sub))
                        if r:
                            r.setdefault('group',_infer_group_from_path(r['path']))
                            out.append(r)
            elif fn.lower().endswith('.py'):
                r=_exec_rule_file(full)
                if r:
                    r.setdefault('group',_infer_group_from_path(r['path']))
                    out.append(r)
    return out

def parse_rules_from_disk_fast():
    rules=[]
    root=_bundle_root()
    # optional JSONs
    for p in [os.path.join(root,'DataSync.index.json'),
              os.path.join(root,'rules.json'),
              os.path.join(root,'config','rules.json')]:
        if not os.path.isfile(p): continue
        try:
            with io.open(p,'r',encoding='utf-8-sig',errors='ignore') as f:
                data=json.loads(f.read())
            if isinstance(data,list): rules.extend(data)
            elif isinstance(data,dict):
                if isinstance(data.get('rules'),list):
                    rules.extend(data['rules'])
                else:
                    for v in data.values():
                        if isinstance(v,list): rules.extend(v)
        except Exception: pass

    rules.extend(discover_rules_from_folder())

    # normalize
    for r in rules:
        r.setdefault('name', os.path.splitext(os.path.basename(r.get('path','rule')))[0])
        r.setdefault('description', u'—')
        cats=[]
        try:
            filt=r.get('filter') or {}
            cats=filt.get('categories') or []
        except: pass
        if not isinstance(cats,list): cats=[]
        r['categories']=list(cats)
        r['group']=safe_str(r.get('group') or u"")
    # dedup
    seen=set(); uniq=[]
    for r in rules:
        key=safe_str(r.get('path') or r.get('name'))
        if key in seen: continue
        seen.add(key); uniq.append(r)
    return uniq

# groups config (DataSync.groups.json)
def _groups_json_candidates():
    root=_bundle_root()
    return [os.path.join(root,'config','DataSync.groups.json'),
            os.path.join(root,'DataSync.groups.json')]

def load_groups_json():
    for p in _groups_json_candidates():
        if not os.path.isfile(p): continue
        try:
            with io.open(p,'r',encoding='utf-8-sig',errors='ignore') as f:
                data=json.loads(f.read())
            items = data.get('groups') if isinstance(data,dict) else None
            if not isinstance(items,list): continue
            names=[]; gmap={}
            for g in items:
                name=safe_str(g.get('name') or u"").strip()
                if not name: continue
                names.append(name)
                files=g.get('rules') or []
                s=set()
                for fn in files:
                    s.add(safe_str(os.path.basename(fn)).lower())
                gmap[name]=s
            if names: return names,gmap
        except Exception:
            pass
    return [],{}

def load_groups(rules=None):
    if rules is None:
        rules=parse_rules_from_disk_fast()
    names=[]
    for r in rules or []:
        g=safe_str(r.get('group') or u"")
        if g and g not in names:
            names.append(g)
    if not names:
        names=[u"All Rules"]
        for r in rules or []:
            if not r.get('group'):
                r['group']=u"All Rules"
    return names

# ---------- parameter helpers with ambiguity policy ----------
_param_cache_by_type = {}

def _doc_key(doc):
    try:
        app=doc.Application
        return (app.Username or u""),(doc.Title or u"")
    except:
        return id(doc)

def _type_key(doc, element):
    try:
        tid=element.GetTypeId()
        return (_doc_key(doc), getattr(tid,'IntegerValue',tid))
    except:
        return (_doc_key(doc), -1)

def ensure_guid(text):
    if not text: return None
    try: return DB.Guid(text)
    except: return None

def _iter_params_by_name(element, name):
    # element.LookupParameter(name) may return one; iterate all to detect ambiguity
    for p in element.Parameters:
        try:
            d=p.Definition
            if d and d.Name == name:
                yield p
        except: pass

def _choose_param_from_candidates(cands, prefer_shared):
    """Deterministic selection when multiple parameters share the same name.
       Priority:
         1) prefer_shared True  -> pick shared True first
            prefer_shared False -> pick shared False first
            prefer_shared None  -> shared True before False
         2) Instance parameter before Type parameter
         3) First in list as last resort
    """
    if not cands: return None
    # annotate candidates
    scored=[]
    for p in cands:
        try:
            is_shared = bool(p.IsShared)
        except:
            is_shared = False
        try:
            # Heuristic: instance params usually available on element; type via element.Symbol/Type
            is_instance = True
            if hasattr(p, 'OwnerViewId'):  # not a great signal; we’ll keep simple
                is_instance = True
        except:
            is_instance = True

        # shared score
        if prefer_shared is True:
            s_score = 0 if is_shared else 1
        elif prefer_shared is False:
            s_score = 0 if (not is_shared) else 1
        else:
            s_score = 0 if is_shared else 1
        # instance score (prefer instance)
        i_score = 0 if is_instance else 1
        scored.append((s_score, i_score, p))

    scored.sort(key=lambda t: (t[0], t[1]))
    chosen = scored[0][2]
    # log ambiguity
    if len(cands) > 1:
        try:
            name = cands[0].Definition.Name
        except:
            name = u"(unknown)"
        _log(u"Ambiguous parameter name '{}' -> chose shared={} instance={} by policy."
             .format(safe_str(name),
                     getattr(chosen,'IsShared',False),
                     True))
    return chosen

def _param_from_name_cached(doc, element, name, prefer_shared):
    base=_type_key(doc, element)
    key=(base, name, False, prefer_shared if prefer_shared is not None else u"*")
    cached=_param_cache_by_type.get(key)
    if cached is not None:
        return cached

    # gather candidates
    cands=list(_iter_params_by_name(element, name))
    if not cands:
        _param_cache_by_type[key]=None
        return None

    chosen = _choose_param_from_candidates(cands, prefer_shared)
    _param_cache_by_type[key]=chosen
    return chosen

def _param_from_guid_cached(doc, element, guid):
    base=_type_key(doc, element)
    key=(base, unicode(guid), True, None)
    cached=_param_cache_by_type.get(key)
    if cached is not None:
        return cached
    try:
        p=element.get_Parameter(guid)
        _param_cache_by_type[key]=p
        return p
    except:
        _param_cache_by_type[key]=None
        return None

def get_target_param(element, name=None, guid=None, prefer_shared=None):
    """Resolution order:
         1) GUID (if given)
         2) Name with ambiguity policy above
       prefer_shared: True/False/None.
    """
    doc=revit.doc
    if guid:
        p=_param_from_guid_cached(doc, element, guid)
        if p: return p
    if name:
        return _param_from_name_cached(doc, element, name, prefer_shared)
    return None

def warm_type_param_cache_for_element(element, names_or_guids=None):
    doc=revit.doc
    if not names_or_guids:
        names_or_guids=[u"Comments", u"Type Name"]
    for key in names_or_guids:
        g=ensure_guid(key)
        if g: _param_from_guid_cached(doc, element, g)
        else: _param_from_name_cached(doc, element, key, None)

def _param_equals(param, new_val, tol=1e-6):
    try:
        st=param.StorageType
        if st == DB.StorageType.Double:
            cur=param.AsDouble(); nv=float(new_val)
            return abs(cur-nv) <= tol
        elif st == DB.StorageType.Integer:
            return param.AsInteger() == int(new_val)
        elif st == DB.StorageType.String:
            return (param.AsString() or u"") == (u"" if new_val is None else unicode(new_val))
        elif st == DB.StorageType.ElementId:
            cur=param.AsElementId()
            if hasattr(new_val,'IntegerValue'): return cur == new_val
            return cur.IntegerValue == int(new_val)
        return False
    except:
        return False

def set_param_value(param, value):
    if _param_equals(param, value):
        return True
    try:
        st=param.StorageType
        if st == DB.StorageType.String:
            param.Set(u"" if value is None else unicode(value)); return True
        elif st == DB.StorageType.Integer:
            param.Set(int(value)); return True
        elif st == DB.StorageType.Double:
            param.Set(float(value)); return True
        elif st == DB.StorageType.ElementId:
            if hasattr(value,'IntegerValue'):
                param.Set(value); return True
            param.Set(DB.ElementId(int(value))); return True
        else:
            param.Set(unicode(value)); return True
    except:
        return False
