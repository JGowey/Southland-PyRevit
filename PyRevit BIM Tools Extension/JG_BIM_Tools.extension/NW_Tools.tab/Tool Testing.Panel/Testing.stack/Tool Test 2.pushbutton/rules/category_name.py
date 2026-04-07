# -*- coding: utf-8 -*-
# RULE must be first; ASCII-only so the loader can detect it.

RULE = {
    "name": "Clean Category -> CP_BOM_Category",
    "description": "CPL_Family Category (instance->type) else Category.Name. Fabrication uses Category.Name.",
    "priority": 50,
    "filter": {
        "categories": [
            # --- Fabrication (this is what your test files use) ---
            "MEP Fabrication Pipework",
            "MEP Fabrication Ductwork",
            "MEP Fabrication Hangers",
            "MEP Fabrication Containment",

            # --- Naviate DataFlow system categories (so the rule also works there) ---
            "Air Terminals", "Casework", "Ceilings", "Columns", "Communication Devices",
            "Conduits", "Data Devices", "Doors", "Duct Accessories", "Duct Fittings",
            "Duct Insulations", "Duct Linings", "Duct Systems", "Ducts",
            "Electrical Equipment", "Electrical Fixtures", "Fire Alarm Devices",
            "Flex Ducts", "Flex Pipes", "Furniture", "Lighting Devices",
            "Lighting Fixtures", "Mechanical Equipment", "Pipe Accessories",
            "Pipe Fittings", "Pipe Insulations", "Pipes", "Plumbing Fixtures",
            "Security Devices", "Specialty Equipment", "Sprinklers",
            "Structural Columns", "Structural Framing", "Windows"
        ]
    },
    "target": { "name": "CP_BOM_Category" },
    "combine": "last_wins"
}

SOURCE_PARAM_NAME = "CPL_Family Category"

__type_cpl_cache = None
__cat_name_cache = None
__last_doc_key = None

def _s(x):
    try:
        if x is None:
            return u""
        return x if isinstance(x, unicode) else unicode(x)
    except:
        try:
            return unicode(str(x))
        except:
            return u""

def _doc_key(doc):
    try:
        return (doc.Title or u"", doc.PathName or u"")
    except:
        return id(doc)

def _reset_caches_if_new_doc(doc):
    global __type_cpl_cache, __cat_name_cache, __last_doc_key
    k = _doc_key(doc)
    if k != __last_doc_key:
        __type_cpl_cache = {}
        __cat_name_cache = {}
        __last_doc_key = k

def _get_string_param(el, name):
    try:
        p = el.LookupParameter(name)
        if p and p.HasValue:
            v = p.AsString()
            if v and _s(v).strip() != u"":
                return _s(v)
    except:
        pass
    return None

def _get_category_name_cached(el):
    try:
        cat = el.Category
        if not cat: return None
        key = cat.Id.IntegerValue
        cache = __cat_name_cache
        if key in cache: return cache[key]
        nm = _s(cat.Name)
        cache[key] = nm
        return nm
    except:
        return None

def compute(element, context):
    """
    Returns the string to write into CP_BOM_Category.
    Discovery-safe: no imports at module top; all API access is via element/doc.
    """
    doc = element.Document
    _reset_caches_if_new_doc(doc)

    # Heuristic: treat anything with Category name starting with "MEP Fabrication"
    # as a fabrication element and use Category.Name directly.
    try:
        cat = element.Category
        if cat:
            catname = _get_category_name_cached(element)
            if catname and catname.startswith(u"MEP Fabrication"):
                return catname
    except:
        pass

    # 1) Instance CPL_Family Category
    inst_val = _get_string_param(element, SOURCE_PARAM_NAME)
    if inst_val:
        return inst_val

    # 2) Type CPL_Family Category (cached per type)
    try:
        tid = element.GetTypeId()
        if hasattr(tid, "IntegerValue") and tid.IntegerValue != -1:
            cache = __type_cpl_cache
            if tid.IntegerValue in cache:
                cached = cache[tid.IntegerValue]
                if cached is not None:
                    return cached
            typ = doc.GetElement(tid)
            val = _get_string_param(typ, SOURCE_PARAM_NAME) if typ else None
            cache[tid.IntegerValue] = val  # cache None too
            if val:
                return val
    except:
        pass

    # 3) Category.Name fallback
    nm = _get_category_name_cached(element)
    if nm:
        return nm

    # 4) Fallback
    return u"Unknown Category"
