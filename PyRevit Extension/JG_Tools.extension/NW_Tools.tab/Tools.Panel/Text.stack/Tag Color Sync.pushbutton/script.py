# -*- coding: utf-8 -*-
"""
Tag Sync — Match Tag Line Color to Host Element Filter Override
================================================================
Scans all IndependentTags in the active view and colors each tag's
projection lines to match the view filter override applied to its
host element. This keeps tag colors visually consistent with the
filtered element colors on sheets and views.

Color resolution per host element:
    1. SurfaceForegroundPatternColor from the matching filter override
    2. Fallback to ProjectionLineColor if surface color is not set
    3. Skip if neither color is valid

Also copies the Halftone setting and forces a Solid line pattern on
the tag so dashed filter patterns don't make tags unreadable.

Revit 2022+ compatible (avoids 2023+ only APIs).
"""

import sys
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")

from pyrevit import revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    IndependentTag,
    ParameterFilterElement,
    OverrideGraphicSettings,
    ElementId,
    Color,
    LinePatternElement,
    LinkElementId
)

doc = revit.doc
view = doc.ActiveView
output = script.get_output()
output.set_title("Tag Sync (Line Color)")

if not view:
    output.print_md("**No active view.**")
    sys.exit()


# =============================================================================
# PRE-BUILD FILTER MEMBERSHIP MAPS
# =============================================================================
# Instead of running a FilteredElementCollector for every (tag, filter)
# combination, we build the maps once upfront:
#   filter_element_sets : filter_id -> set of element ids that pass the filter
#   filter_overrides    : filter_id -> OverrideGraphicSettings for that filter
#   filter_categories   : filter_id -> set of category ids the filter applies to
#
# This turns O(tags * filters) collector calls into O(filters) collector calls.

SOLID_LINE_ID = LinePatternElement.GetSolidPatternId()

view_filter_ids = view.GetFilters()
filter_element_sets = {}
filter_overrides = {}
filter_categories = {}

for f_id in view_filter_ids:
    f_elem = doc.GetElement(f_id)
    if not isinstance(f_elem, ParameterFilterElement):
        continue

    # Cache which categories this filter applies to
    filter_categories[f_id] = set(f_elem.GetCategories())

    # Cache the override settings
    ogs = view.GetFilterOverrides(f_id)
    if not ogs:
        continue
    filter_overrides[f_id] = ogs

    # Cache which element ids pass this filter in this view
    try:
        collector = FilteredElementCollector(doc, view.Id).WherePasses(f_elem.GetElementFilter())
        filter_element_sets[f_id] = set(e.Id for e in collector)
    except:
        filter_element_sets[f_id] = set()


# =============================================================================
# RESOLVE HOST ELEMENT -> FILTER COLOR
# =============================================================================

def _resolve_filter_color(host_element):
    """
    Finds the first matching view filter for a host element and returns
    its (color, halftone) tuple, or None if no filter match has a valid color.

    Checks filters in view priority order (first match wins, matching Revit behavior).
    """
    cat = host_element.Category
    if cat is None:
        return None

    cat_id = cat.Id
    elem_id = host_element.Id

    for f_id in view_filter_ids:
        # Skip filters that don't apply to this element's category
        cats = filter_categories.get(f_id)
        if cats is None or cat_id not in cats:
            continue

        # Skip filters where this element doesn't pass the filter criteria
        members = filter_element_sets.get(f_id)
        if members is None or elem_id not in members:
            continue

        # Get the cached override
        ogs = filter_overrides.get(f_id)
        if ogs is None:
            continue

        # Color fallback: surface foreground pattern -> projection line
        color = ogs.SurfaceForegroundPatternColor
        if not color or not color.IsValid:
            color = ogs.ProjectionLineColor
        if not color or not color.IsValid:
            continue

        return (color, ogs.Halftone)

    return None


# =============================================================================
# MAIN: SYNC TAG COLORS
# =============================================================================

tags = FilteredElementCollector(doc, view.Id).OfClass(IndependentTag)
count_updated = 0
count_skipped = 0

with revit.Transaction("Sync Tag Line Colors"):
    for tag in tags:
        try:
            tagged_ids = tag.GetTaggedElementIds()
            if not tagged_ids:
                continue

            # Resolve color from the first tagged element that has a filter match.
            # For multi-host tags, we use the first match (consistent with how
            # Revit displays the tag — it follows the primary host).
            result = None
            for tid in tagged_ids:
                host_id = tid.HostElementId if isinstance(tid, LinkElementId) else tid
                host_element = doc.GetElement(host_id)
                if not host_element:
                    continue
                result = _resolve_filter_color(host_element)
                if result:
                    break

            if result:
                color, halftone = result
                tag_ogs = OverrideGraphicSettings()
                tag_ogs.SetProjectionLineColor(color)
                tag_ogs.SetProjectionLinePatternId(SOLID_LINE_ID)
                tag_ogs.SetHalftone(halftone)
                view.SetElementOverrides(tag.Id, tag_ogs)

                output.print_md("Tag `{}` -> R:{} G:{} B:{}".format(
                    tag.Id.IntegerValue, color.Red, color.Green, color.Blue
                ))
                count_updated += 1
            else:
                count_skipped += 1

        except Exception as e:
            output.print_md("Error on tag {}: {}".format(tag.Id.IntegerValue, e))
            count_skipped += 1

output.print_md("---")
output.print_md("**Done.** {} tag(s) colored, {} skipped.".format(count_updated, count_skipped))
