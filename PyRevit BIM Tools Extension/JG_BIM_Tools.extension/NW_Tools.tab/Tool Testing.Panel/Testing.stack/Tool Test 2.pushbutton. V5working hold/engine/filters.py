# -*- coding: utf-8 -*-
# engine/filters.py — compile fast native Revit filters only
from __future__ import unicode_literals
from pyrevit import DB

# Display-name → BuiltInCategory aliases (expand as needed)
CAT_NAME_ALIASES = {
    u"MEP Fabrication Pipework":  "OST_FabricationPipework",
    u"MEP Fabrication Ductwork":  "OST_FabricationDuctwork",
    u"MEP Fabrication Hangers":   "OST_FabricationHangers",
    u"MEP Fabrication Parts":     "OST_FabricationParts",
    u"Pipes":                     "OST_PipeCurves",
    u"Pipe Fittings":             "OST_PipeFitting",
    u"Pipe Accessories":          "OST_PipeAccessory",
    u"Ducts":                     "OST_DuctCurves",
    u"Duct Fittings":             "OST_DuctFitting",
    u"Duct Accessories":          "OST_DuctAccessory",
}

def _bic_from_display(catname):
    if not catname:
        return None
    key = CAT_NAME_ALIASES.get(catname)
    if not key:
        key = "OST_" + catname.replace(' ', '').replace('-', '')
    try:
        return getattr(DB.BuiltInCategory, key)
    except:
        return None

def compile_union_filter(rules):
    """
    Build a fast native Revit filter from a list of rule dicts that may contain category names.
    Fallback is 'non-types' if no categories provided.
    """
    cats = set()
    for r in (rules or []):
        for n in (r.get('categories') or []):
            cats.add(n)

    bics = []
    for n in cats:
        bic = _bic_from_display(n)
        if bic is not None:
            bics.append(bic)

    if bics:
        cat_filters = [DB.ElementCategoryFilter(b) for b in bics]
        return DB.LogicalOrFilter(cat_filters) if len(cat_filters) > 1 else cat_filters[0]

    # Safe default: operate on instances (not types)
    return DB.ElementIsElementTypeFilter(False)
