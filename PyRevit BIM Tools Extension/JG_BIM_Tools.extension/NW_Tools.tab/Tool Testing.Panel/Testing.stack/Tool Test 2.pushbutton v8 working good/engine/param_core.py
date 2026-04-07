# -*- coding: utf-8 -*-
from __future__ import unicode_literals

from pyrevit import revit, DB
import System

# ---------------- basic helpers ----------------
def safe_str(v):
    try:
        if v is None:
            return u""
        if isinstance(v, System.String):
            return unicode(v)
        return unicode(v)
    except Exception:
        try:
            return str(v)
        except Exception:
            return u""

def ensure_guid(val):
    if not val:
        return None
    try:
        if isinstance(val, DB.Guid):
            return val
    except Exception:
        pass
    try:
        return System.Guid(val)
    except Exception:
        return None

# cache: we only warm a type once
_type_warmed = set()
def warm_type_param_cache_for_element(el):
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue not in _type_warmed:
            _type_warmed.add(tid.IntegerValue)
            _ = el.Document.GetElement(tid)
    except Exception:
        pass

# ---------------- resolution ----------------
def _resolve_by_guid(el, guid):
    if not guid:
        return None
    try:
        p = el.get_Parameter(guid)
        if p: return p
    except Exception:
        pass
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue > 0:
            et = el.Document.GetElement(tid)
            if et:
                p = et.get_Parameter(guid)
                if p: return p
    except Exception:
        pass
    return None

def _resolve_instance_comments(el):
    try:
        p = el.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p: return p
    except Exception:
        pass
    try:
        p = el.LookupParameter("Comments")
        if p: return p
    except Exception:
        pass
    return None

def _resolve_by_name(el, name, prefer_shared):
    if not name:
        return None

    # Special-case: Comments → prefer instance
    if name.strip().lower() == "comments":
        pc = _resolve_instance_comments(el)
        if pc: return pc

    # Instance first
    candidate_inst = None
    try:
        candidate_inst = el.LookupParameter(name)
    except Exception:
        candidate_inst = None

    # Type next (we may skip during apply unless you opt in)
    candidate_type = None
    try:
        tid = el.GetTypeId()
        if tid and tid.IntegerValue > 0:
            et = el.Document.GetElement(tid)
            if et:
                candidate_type = et.LookupParameter(name)
    except Exception:
        candidate_type = None

    # Prefer shared if requested
    if prefer_shared:
        try:
            if candidate_inst and candidate_inst.Definition and candidate_inst.Definition.IsShared:
                return candidate_inst
            if candidate_type and candidate_type.Definition and candidate_type.Definition.IsShared:
                return candidate_type
        except Exception:
            pass

    if candidate_inst:
        return candidate_inst
    if candidate_type:
        return candidate_type
    return None

def get_target_param(el, name=None, guid=None, prefer_shared=None):
    if guid:
        p = _resolve_by_guid(el, guid)
        if p: return p
    if name:
        return _resolve_by_name(el, name, bool(prefer_shared))
    return None

# ---------------- guards ----------------
def is_type_parameter_for_element(el, p):
    try:
        return (p.Owner and p.Owner.Id == el.GetTypeId())
    except Exception:
        return False

def is_readonly(p):
    try:
        return bool(p.IsReadOnly)
    except Exception:
        return True

# ---------------- setter ----------------
def set_param_value(p, value):
    """
    Returns True if Set succeeded or value already equal.
    For doubles: value must be in INTERNAL units.
    """
    try:
        st = p.StorageType

        if st == DB.StorageType.String:
            newv = safe_str(value)
            oldv = p.AsString() or u""
            if safe_str(oldv) == newv:
                return True
            return p.Set(newv)

        if st == DB.StorageType.Integer:
            try:
                iv = int(value)
            except Exception:
                try: iv = int(float(value))
                except Exception: return False
            if p.AsInteger() == iv:
                return True
            return p.Set(iv)

        if st == DB.StorageType.Double:
            try:
                dv = float(value)
            except Exception:
                return False
            old = p.AsDouble()
            if abs((old or 0.0) - dv) < 1e-12:
                return True
            return p.Set(dv)

        if st == DB.StorageType.ElementId:
            try:
                iv = int(value)
                eid = DB.ElementId(iv) if iv != 0 else DB.ElementId.InvalidElementId
            except Exception:
                return False
            old = p.AsElementId()
            if (old.IntegerValue if old else 0) == (eid.IntegerValue if eid else 0):
                return True
            return p.Set(eid)

        return False
    except Exception:
        return False

# ---------------- small read-back for verification ----------------
def read_param_as_text(p):
    try:
        st = p.StorageType
        if st == DB.StorageType.String:
            return safe_str(p.AsString() or u"")
        if st == DB.StorageType.Integer:
            return safe_str(p.AsInteger())
        if st == DB.StorageType.Double:
            return safe_str(p.AsDouble())
        if st == DB.StorageType.ElementId:
            eid = p.AsElementId()
            return safe_str(eid.IntegerValue if eid else 0)
    except Exception:
        pass
    return u""
