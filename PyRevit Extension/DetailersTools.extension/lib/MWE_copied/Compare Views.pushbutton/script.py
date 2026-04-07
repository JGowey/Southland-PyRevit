# -*- coding: utf-8 -*-
# Compare two Revit Views (or View Templates) and optionally export differences to CSV

from pyrevit import revit, DB, forms, script
import csv

doc   = revit.doc
uidoc = revit.uidoc
out   = script.get_output()
try:
    out.set_title("Compare Views")
    out.resize(1200, 900)
    out.center()
except:
    pass

# Collect CSV rows as [Section, Item, View A, View B]
csv_diffs = []

# -------------------- Helpers --------------------
def _is_schedule(v):
    """Return True if v is a schedule view (always excluded here)."""
    return isinstance(v, DB.ViewSchedule)

def _nice_label(v):
    """Readable label for selection and headers."""
    vt = getattr(v, 'ViewType', None)
    vtype = str(vt) if vt is not None else "View"
    tag = "TEMPLATE • " if getattr(v, 'IsTemplate', False) else ""
    return u"{0}{1} • {2}".format(tag, vtype, v.Name)

def _get_all_views():
    """Get all non-template views from the project."""
    all_views = DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements()
    return sorted(
        [v for v in all_views if not v.IsTemplate and not _is_schedule(v)],
        key=lambda x: x.Name.lower()
    )

def _get_view_templates():
    """All view templates (no schedules)."""
    return sorted(
        [v for v in DB.FilteredElementCollector(doc)
                     .OfClass(DB.View).ToElements()
         if getattr(v, 'IsTemplate', False) and not _is_schedule(v)],
        key=lambda x: x.Name.lower()
    )

def _pick_two_views_or_templates():
    """Dark, two-button source chooser -> list picker for exactly two items."""
    source = forms.CommandSwitchWindow.show(
        ["Project Views", "View Templates"],
        message="What would you like to compare?",
        title="Compare",
        dark_theme=True,
        width=420,
        height=200
    )
    if not source:
        forms.alert("Cancelled.", exitscript=True)

    cands = _get_all_views() if source == "Project Views" else _get_view_templates()
    if len(cands) < 2:
        forms.alert(
            "Need at least two {} in this model."
            .format("views" if source == "Project Views" else "view templates"),
            exitscript=True
        )

    label_map = {_nice_label(v): v for v in cands}
    labels_sorted = sorted(label_map.keys(), key=lambda s: s.lower())

    picks = forms.SelectFromList.show(
        labels_sorted,
        title="Pick exactly two {}".format("views" if source == "Project Views" else "view templates"),
        multiselect=True,
        button_name="Compare",
        dark_theme=True,
        height=600
    )
    if not picks or len(picks) != 2:
        forms.alert("Please pick exactly two.", exitscript=True)

    return [label_map[p] for p in picks]

def _pick_aspects_to_compare():
    """Shows a checklist of comparison categories for the user to select."""
    aspects = [
        "Model Categories",
        "Annotation Categories",
        "Analytical Model Categories",
        "Imported Categories",
        "Filters",
        "Worksets",
        "Revit Links"
    ]
    selected_aspects = forms.SelectFromList.show(
        aspects,
        title="Select Aspects to Compare",
        multiselect=True,
        button_name="Run Comparison",
        dark_theme=True,
        default_selected=True
    )
    if not selected_aspects:
        forms.alert("No aspects selected. Cancelled.", exitscript=True)
    return selected_aspects

# NEW: This function lets the user choose how to filter the results.
def _pick_comparison_mode():
    """Lets the user choose whether to see all, different, or same settings."""
    mode = forms.CommandSwitchWindow.show(
        ["Show Differences", "Show Same", "Show Everything"],
        message="How would you like to view the results?",
        title="Comparison Mode",
        dark_theme=True
    )
    if not mode:
        forms.alert("Cancelled.", exitscript=True)
    return mode

def _cat_groups():
    """Top-level categories grouped by type."""
    groups = {"Model": [], "Annotation": [], "Analytical": [], "Import/Other": []}
    for cat in doc.Settings.Categories:
        try:
            if not isinstance(cat, DB.Category) or cat.Parent is not None:
                continue
            if cat.CategoryType == DB.CategoryType.Model:
                groups["Model"].append(cat)
            elif cat.CategoryType == DB.CategoryType.Annotation:
                groups["Annotation"].append(cat)
            elif cat.CategoryType == DB.CategoryType.AnalyticalModel:
                groups["Analytical"].append(cat)
            else:
                groups["Import/Other"].append(cat)
        except:
            pass
    for k in groups:
        groups[k].sort(key=lambda c: c.Name.lower())
    return groups

def _cat_hidden(vw, cat):
    try:
        return not vw.IsCategoryVisible(cat.Id) if hasattr(vw, 'IsCategoryVisible') else vw.GetCategoryHidden(cat.Id)
    except:
        return False

def _ogs_as_tuple(ogs):
    """Normalize OverrideGraphicSettings to a hashable tuple."""
    try:
        def _get(getter, default=None):
            try: return getter()
            except: return default
        def _col(c):
            return (c.Red, c.Green, c.Blue) if isinstance(c, DB.Color) else None
        def _eid(eid):
            return eid.IntegerValue if isinstance(eid, DB.ElementId) else None
        return (
            _col(_get(ogs.GetCutLineColor)), _col(_get(ogs.GetProjectionLineColor)),
            _get(ogs.GetCutLineWeight), _get(ogs.GetProjectionLineWeight),
            _get(ogs.IsHalftone), _get(ogs.GetSurfaceTransparency),
            _eid(_get(ogs.GetSurfaceForegroundPatternId)), _eid(_get(ogs.GetSurfaceBackgroundPatternId)),
            _eid(_get(ogs.GetCutForegroundPatternId))
        )
    except:
        return None

def _worksets():
    try:
        return list(DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)) if DB.WorksharingUtils.IsWorkshared(doc) else []
    except:
        return []

def _ws_vis(vw, ws):
    try:
        return str(vw.GetWorksetVisibility(ws.Id))
    except:
        return "UseGlobalSetting"

def _revit_links():
    try:
        return list(DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance))
    except:
        return []

def _link_hidden(vw, inst):
    try:
        return inst.IsHidden(vw)
    except:
        return False

def _add_diff(section, name, val_a, val_b):
    csv_diffs.append([section, name, val_a, val_b])

def _print_heading(view_a, view_b):
    out.print_md("# Compare Views")
    out.print_md("* View A: **{}**".format(_nice_label(view_a)))
    out.print_md("* View B: **{}**".format(_nice_label(view_b)))
    out.print_md("\n## Results\n")

# -------------------- Comparisons --------------------

def _compare_category_settings(view_a, view_b, selected_aspects, mode):
    aspect_map = {
        "Model": "Model Categories", "Annotation": "Annotation Categories",
        "Analytical": "Analytical Model Categories", "Import/Other": "Imported Categories"
    }
    all_category_groups = _cat_groups()

    for group_key, aspect_name in aspect_map.items():
        if aspect_name not in selected_aspects:
            continue

        out.print_md("### {}".format(aspect_name))
        cats_in_group = all_category_groups.get(group_key, [])
        if not cats_in_group:
            out.print_md("- No categories in this group.")
            continue

        # --- Visibility ---
        out.print_md("#### Visibility")
        vis_items_printed = 0
        for c in cats_in_group:
            ha, hb = _cat_hidden(view_a, c), _cat_hidden(view_b, c)
            is_diff = ha != hb
            if (mode == "Show Differences" and is_diff) or \
               (mode == "Show Same" and not is_diff) or \
               mode == "Show Everything":
                vis_items_printed += 1
                state_a, state_b = ("Hidden" if ha else "Shown"), ("Hidden" if hb else "Shown")
                tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
                out.print_md(u"- {}{}: A={} / B={}".format(c.Name, tag, state_a, state_b))
                if is_diff: _add_diff(aspect_name + " - Visibility", c.Name, state_a, state_b)
        if not vis_items_printed: out.print_md("- No items to show for this mode.")

        # --- Overrides ---
        out.print_md("#### Overrides")
        ovr_items_printed = 0
        for c in cats_in_group:
            ta, tb = _ogs_as_tuple(view_a.GetCategoryOverrides(c.Id)), _ogs_as_tuple(view_b.GetCategoryOverrides(c.Id))
            is_diff = ta != tb
            if (mode == "Show Differences" and is_diff) or \
               (mode == "Show Same" and not is_diff) or \
               mode == "Show Everything":
                ovr_items_printed += 1
                tag = " (DIFFERENT)" if is_diff else ""
                out.print_md(u"- {}{}".format(c.Name, tag))
                if is_diff: _add_diff(aspect_name + " - Overrides", c.Name, "Different", "Different")
        if not ovr_items_printed: out.print_md("- No items to show for this mode.")

def _compare_filters(view_a, view_b, mode):
    out.print_md("\n### Filters")
    fa_ids = {eid.IntegerValue for eid in (view_a.GetFilters() or [])}
    fb_ids = {eid.IntegerValue for eid in (view_b.GetFilters() or [])}
    
    def _name(eid_int):
        e = doc.GetElement(DB.ElementId(eid_int)); return e.Name if e else "<missing>"

    all_ids = sorted(list(fa_ids | fb_ids))
    items_printed = 0
    for i in all_ids:
        in_a, in_b = i in fa_ids, i in fb_ids
        status = ""
        if in_a and not in_b: status = "Only in A"
        elif not in_a and in_b: status = "Only in B"
        else:
            oa, ob = view_a.GetFilterOverrides(DB.ElementId(i)), view_b.GetFilterOverrides(DB.ElementId(i))
            status = "Different Overrides" if _ogs_as_tuple(oa) != _ogs_as_tuple(ob) else "Same Overrides"
        
        is_diff = status != "Same Overrides"
        if (mode == "Show Differences" and is_diff) or \
           (mode == "Show Same" and not is_diff) or \
           mode == "Show Everything":
            items_printed += 1
            tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
            out.print_md(u"- {}: {}{}".format(_name(i), status, tag))
            if is_diff: _add_diff("Filters", _name(i), status, "N/A")
    if not items_printed: out.print_md("- No items to show for this mode.")

def _compare_worksets(view_a, view_b, mode):
    out.print_md("\n### Worksets")
    wss = _worksets()
    if not wss:
        out.print_md("- Not workshared."); return
    items_printed = 0
    for ws in sorted(wss, key=lambda w: w.Name.lower()):
        a, b = _ws_vis(view_a, ws), _ws_vis(view_b, ws)
        is_diff = a != b
        if (mode == "Show Differences" and is_diff) or \
           (mode == "Show Same" and not is_diff) or \
           mode == "Show Everything":
            items_printed += 1
            tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
            out.print_md(u"- {}{}: A={} / B={}".format(ws.Name, tag, a, b))
            if is_diff: _add_diff("Worksets", ws.Name, a, b)
    if not items_printed: out.print_md("- No items to show for this mode.")

def _compare_links(view_a, view_b, mode):
    out.print_md("\n### Revit Links")
    links = _revit_links()
    if not links:
        out.print_md("- No links."); return
    items_printed = 0
    for inst in sorted(links, key=lambda x: x.Name.lower()):
        a, b = _link_hidden(view_a, inst), _link_hidden(view_b, inst)
        is_diff = a != b
        if (mode == "Show Differences" and is_diff) or \
           (mode == "Show Same" and not is_diff) or \
           mode == "Show Everything":
            items_printed += 1
            state_a, state_b = ("Hidden" if a else "Shown"), ("Hidden" if b else "Shown")
            tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
            out.print_md(u"- {}{}: A={} / B={}".format(inst.Name, tag, state_a, state_b))
            if is_diff: _add_diff("Links", inst.Name, state_a, state_b)
    if not items_printed: out.print_md("- No items to show for this mode.")

# -------------------- Run --------------------
view_a, view_b = _pick_two_views_or_templates()
selected_aspects = _pick_aspects_to_compare()
comparison_mode = _pick_comparison_mode()

_print_heading(view_a, view_b)

category_aspects = ["Model Categories", "Annotation Categories", "Analytical Model Categories", "Imported Categories"]
if any(aspect in selected_aspects for aspect in category_aspects):
    _compare_category_settings(view_a, view_b, selected_aspects, comparison_mode)

if "Filters" in selected_aspects:
    _compare_filters(view_a, view_b, comparison_mode)

if "Worksets" in selected_aspects:
    _compare_worksets(view_a, view_b, comparison_mode)

if "Revit Links" in selected_aspects:
    _compare_links(view_a, view_b, comparison_mode)

# -------------------- Save CSV? --------------------
if csv_diffs:
    save_choice = forms.CommandSwitchWindow.show(
        ["Save CSV…", "No"],
        message="Do you want to save the differences to CSV?",
        title="Export",
        dark_theme=True,
        width=380,
        height=170
    )
    if save_choice == "Save CSV…":
        path = forms.save_file(file_ext='csv', title="Save Differences as CSV")
        if path:
            try:
                with open(path, mode='w', newline='', encoding='utf-8') as f:
                    w = csv.writer(f)
                    w.writerow(["Section", "Item", "View A Value", "View B Value"])
                    w.writerows(csv_diffs)
                out.print_md("\n**CSV saved to:** `{}`".format(path))
            except Exception as ex:
                forms.alert("Failed to write CSV:\n{}".format(ex))
else:
    out.print_md("\n_No differences found to export._")
