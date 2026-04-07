# -*- coding: utf-8 -*-
# Discovery-safe: RULE first, ASCII-only, no imports at top.

RULE = {
    "name": "DEBUG Clean Category -> Comments",
    "description": "Same clean-category logic, but writes to Comments (for quick verification).",
    "priority": 50,
    "filter": {
        "categories": [
            "MEP Fabrication Pipework","MEP Fabrication Ductwork","MEP Fabrication Hangers","MEP Fabrication Containment",
            "Air Terminals","Casework","Ceilings","Columns","Communication Devices",
            "Conduits","Data Devices","Doors","Duct Accessories","Duct Fittings",
            "Duct Insulations","Duct Linings","Duct Systems","Ducts",
            "Electrical Equipment","Electrical Fixtures","Fire Alarm Devices",
            "Flex Ducts","Flex Pipes","Furniture","Lighting Devices",
            "Lighting Fixtures","Mechanical Equipment","Pipe Accessories",
            "Pipe Fittings","Pipe Insulations","Pipes","Plumbing Fixtures",
            "Security Devices","Specialty Equipment","Sprinklers",
            "Structural Columns","Structural Framing","Windows"
        ]
    },
    "target": { "name": "Comments" },
    "combine": "last_wins"
}

SOURCE_PARAM_NAME = "CPL_Family Category"
__type_cpl_cache = None
__cat_name_cache  = None
__last_doc_key    = None

def _s(x):
    try:
        if x is None: return u""
        return x if isinstance(x, unicode) else unicode(x)
    except:
        try: return unicode(str(x))
        except: return u""

def _doc_key(doc):
    try: return (doc.Title or u"", doc.PathName or u"")
    except: return id(doc)

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
            if v and _s(v).strip() != u"": return _s(v)
    except: pass
    return None

def _get_category_name_cached(el):
    try:
        cat = el.Category
        if not cat: return None
        key = cat.Id.IntegerValue
        cache = __cat_name_cache
        if key in cache: return cache[key]
        nm = _s(cat.Name); cache[key] = nm
        return nm
    except: return None

def compute(element, context):
    doc = element.Document
    _reset_caches_if_new_doc(doc)

    # Treat “MEP Fabrication …” as fabrication → use Category.Name directly
    try:
        cat = element.Category
        if cat:
            cn = _get_category_name_cached(element)
            if cn and cn.startswith(u"MEP Fabrication"):
                return cn
    except: pass

    # 1) Instance CPL_Family Category
    inst = _get_string_param(element, SOURCE_PARAM_NAME)
    if inst: return inst

    # 2) Type CPL_Family Category (cached)
    try:
        tid = element.GetTypeId()
        if hasattr(tid, "IntegerValue") and tid.IntegerValue != -1:
            cache = __type_cpl_cache
            if tid.IntegerValue in cache:
                c = cache[tid.IntegerValue]
                if c is not None: return c
            typ = doc.GetElement(tid)
            val = _get_string_param(typ, SOURCE_PARAM_NAME) if typ else None
            cache[tid.IntegerValue] = val
            if val: return val
    except: pass

    # 3) Fallback to Category.Name
    cn = _get_category_name_cached(element)
    if cn: return cn

    return u"Unknown Category"
