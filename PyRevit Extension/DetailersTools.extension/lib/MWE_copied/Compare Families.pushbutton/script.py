# -*- coding: utf-8 -*-
# Compare two Revit Families and their types/parameters.

from pyrevit import revit, DB, forms, script
import csv

doc   = revit.doc
uidoc = revit.uidoc
out   = script.get_output()
try:
    out.set_title("Compare Families")
    out.resize(1200, 900)
    out.center()
except:
    pass

# Collect CSV rows for export
csv_diffs = []

# -------------------- Helpers --------------------
def _safe_name(obj, fallback=""):
    """Return a readable name for any Revit object, or fallback if not available."""
    try:
        return getattr(obj, "Name", None) or fallback
    except Exception:
        return fallback

def _safe_param_dict(paramset):
    """Map parameter name -> Parameter, skipping params with no Definition/Name."""
    result = {}
    for p in paramset:
        try:
            d = p.Definition
            if d and d.Name:
                result[d.Name] = p
        except Exception:
            continue
    return result

# NEW: Helper to group families by their category name.
def _get_families_by_category():
    """Gets all user-created families and groups them into a dictionary by category name."""
    families_by_cat = {}
    all_families = DB.FilteredElementCollector(doc).OfClass(DB.Family).ToElements()
    
    for f in all_families:
        if f.IsUserCreated and f.FamilyCategory:
            cat_name = f.FamilyCategory.Name
            if cat_name not in families_by_cat:
                families_by_cat[cat_name] = []
            families_by_cat[cat_name].append(f)
            
    return families_by_cat

# MODIFIED: This function now implements the two-step selection process.
def _pick_two_families():
    """Lets the user first select categories, then select two families from within those categories."""
    families_by_cat = _get_families_by_category()
    if not families_by_cat:
        forms.alert("No categorized, loadable families found in this model.", exitscript=True)

    # Step 1: Pick one or more categories
    sorted_categories = sorted(families_by_cat.keys())
    selected_cats = forms.SelectFromList.show(
        sorted_categories,
        title="Step 1: Select Model Categories",
        multiselect=True,
        button_name="Next",
        dark_theme=True,
        height=600
    )
    if not selected_cats:
        forms.alert("No categories selected. Cancelled.", exitscript=True)

    # Step 2: Pick two families from the selected categories
    candidate_families = []
    for cat in selected_cats:
        candidate_families.extend(families_by_cat.get(cat, []))
    
    if len(candidate_families) < 2:
        forms.alert("The selected categories contain fewer than two families.", exitscript=True)

    # Sort families by name for the final selection list
    sorted_candidates = sorted(candidate_families, key=lambda f: f.Name)
    
    picks = forms.SelectFromList.show(
        [_safe_name(f) for f in sorted_candidates],
        title="Step 2: Pick Exactly Two Families to Compare",
        multiselect=True,
        button_name="Compare",
        dark_theme=True,
        height=600
    )
    if not picks or len(picks) != 2:
        forms.alert("Please pick exactly two families.", exitscript=True)

    # Return the full Family objects
    return [f for f in sorted_candidates if _safe_name(f) in picks]


def _pick_aspects_to_compare():
    """Shows a checklist of what to compare."""
    aspects = [
        "Family Properties (Category, Shared, etc.)",
        "Family Parameters",
        "Family Types (and their parameters)"
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

def _get_param_value(param):
    """Safely gets a parameter's value as a displayable string."""
    if not param:
        return "N/A"
    try:
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString() or ""
        elif st == DB.StorageType.Double:
            return str(round(param.AsDouble(), 4))
        elif st == DB.StorageType.Integer:
            return str(param.AsInteger())
        elif st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != DB.ElementId.InvalidElementId:
                elem = doc.GetElement(eid)
                return _safe_name(elem, fallback=str(eid.IntegerValue))
            return "None"
        else:
            return "Unhandled Type"
    except Exception:
        return "Error Reading Value"

def _add_diff(section, name, val_a, val_b):
    csv_diffs.append([section, name, val_a, val_b])

def _print_heading(fam_a, fam_b):
    out.print_md("# Compare Families")
    out.print_md("* Family A: **{}**".format(fam_a.Name))
    out.print_md("* Family B: **{}**".format(fam_b.Name))
    out.print_md("\n## Results\n")

# -------------------- Comparisons --------------------

def _compare_family_properties(fam_a, fam_b, mode):
    """Compares top-level properties of the two families."""
    out.print_md("### Family Properties (Category, Shared, etc.)")
    
    props_to_check = {
        "Family Category": lambda f: f.FamilyCategory.Name if f.FamilyCategory else "None",
        "Is Shared": lambda f: "Yes" if f.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED) and f.get_Parameter(DB.BuiltInParameter.FAMILY_SHARED).AsInteger() == 1 else "No",
        "Is In-Place": lambda f: str(f.IsInPlace),
        "Part Type": lambda f: f.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE).AsValueString() if f.get_Parameter(DB.BuiltInParameter.FAMILY_CONTENT_PART_TYPE) else "N/A"
    }

    items_printed = 0
    for prop_name, getter in props_to_check.items():
        val_a, val_b = getter(fam_a), getter(fam_b)
        is_diff = val_a != val_b

        if (mode == "Show Differences" and is_diff) or \
           (mode == "Show Same" and not is_diff) or \
           mode == "Show Everything":
            items_printed += 1
            tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
            out.print_md(u"- {}{}: A='{}' / B='{}'".format(prop_name, tag, val_a, val_b))
            if is_diff: _add_diff("Family Properties", prop_name, val_a, val_b)
            
    if not items_printed: out.print_md("- No items to show for this mode.")

def _compare_family_parameters(fam_a, fam_b, mode):
    """Compares the parameters defined at the family level."""
    out.print_md("\n### Family Parameters")
    
    params_a = _safe_param_dict(fam_a.Parameters)
    params_b = _safe_param_dict(fam_b.Parameters)
    all_param_names = sorted(list(set(params_a.keys()) | set(params_b.keys())))

    items_printed = 0
    for name in all_param_names:
        p_a, p_b = params_a.get(name), params_b.get(name)
        val_a, val_b = _get_param_value(p_a), _get_param_value(p_b)
        is_diff = val_a != val_b

        if (mode == "Show Differences" and is_diff) or \
           (mode == "Show Same" and not is_diff) or \
           mode == "Show Everything":
            items_printed += 1
            tag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
            out.print_md(u"- {}{}: A='{}' / B='{}'".format(name, tag, val_a, val_b))
            if is_diff: _add_diff("Family Parameters", name, val_a, val_b)

    if not items_printed: out.print_md("- No items to show for this mode.")

def _compare_family_types(fam_a, fam_b, mode):
    """Compares the types within each family and their parameters."""
    out.print_md("\n### Family Types (and their parameters)")

    types_a, types_b = {}, {}
    try:
        for tid in fam_a.GetFamilySymbolIds():
            sym = doc.GetElement(tid)
            nm = _safe_name(sym)
            if sym and nm:
                types_a[nm] = sym
    except Exception: pass
    try:
        for tid in fam_b.GetFamilySymbolIds():
            sym = doc.GetElement(tid)
            nm = _safe_name(sym)
            if sym and nm:
                types_b[nm] = sym
    except Exception: pass

    all_type_names = sorted(list(set(types_a.keys()) | set(types_b.keys())))
    items_printed = 0

    for type_name in all_type_names:
        type_a, type_b = types_a.get(type_name), types_b.get(type_name)

        if type_a and type_b:
            params_a = _safe_param_dict(type_a.Parameters)
            params_b = _safe_param_dict(type_b.Parameters)
            all_param_names = sorted(list(set(params_a.keys()) | set(params_b.keys())))

            type_has_diffs = False
            param_results = []
            for name in all_param_names:
                p_a, p_b = params_a.get(name), params_b.get(name)
                val_a, val_b = _get_param_value(p_a), _get_param_value(p_b)
                is_diff = (val_a != val_b)
                type_has_diffs |= is_diff
                param_results.append((name, val_a, val_b, is_diff))

            if (mode == "Show Differences" and type_has_diffs) or \
               (mode == "Show Same" and not type_has_diffs) or \
               mode == "Show Everything":
                items_printed += 1
                tag = " (DIFFERENCES FOUND)" if type_has_diffs and mode == "Show Everything" else ""
                out.print_md("#### Type: {}{}".format(type_name, tag))
                for name, val_a, val_b, is_diff in param_results:
                    if (mode == "Show Differences" and is_diff) or \
                       (mode == "Show Same" and not is_diff) or \
                       mode == "Show Everything":
                        dtag = " (DIFFERENT)" if is_diff and mode == "Show Everything" else ""
                        out.print_md(u"  - {}{}: A='{}' / B='{}'".format(name, dtag, val_a, val_b))
                        if is_diff:
                            _add_diff("Type: " + type_name, name, val_a, val_b)
        else:
            if mode in ("Show Differences", "Show Everything"):
                items_printed += 1
                status = "Only in Family A" if type_a else "Only in Family B"
                tag = " (DIFFERENT)" if mode == "Show Everything" else ""
                out.print_md("#### Type: {}{}".format(type_name, tag))
                out.print_md("  - Status: {}".format(status))
                _add_diff("Family Types", type_name, status, "")

    if not items_printed:
        out.print_md("- No items to show for this mode.")

# -------------------- Run --------------------
fam_a, fam_b = _pick_two_families()
selected_aspects = _pick_aspects_to_compare()
comparison_mode = _pick_comparison_mode()

_print_heading(fam_a, fam_b)

if "Family Properties (Category, Shared, etc.)" in selected_aspects:
    _compare_family_properties(fam_a, fam_b, comparison_mode)

if "Family Parameters" in selected_aspects:
    _compare_family_parameters(fam_a, fam_b, comparison_mode)

if "Family Types (and their parameters)" in selected_aspects:
    _compare_family_types(fam_a, fam_b, comparison_mode)

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
                    w.writerow(["Section", "Item", "Family A Value", "Family B Value"])
                    w.writerows(csv_diffs)
                out.print_md("\n**CSV saved to:** `{}`".format(path))
            except Exception as ex:
                forms.alert("Failed to write CSV:\n{}".format(ex))
else:
    out.print_md("\n_No differences found to export._")
