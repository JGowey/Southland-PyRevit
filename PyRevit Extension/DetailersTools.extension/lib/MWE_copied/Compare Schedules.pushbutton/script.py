# -*- coding: utf-8 -*-
# Compare 2+ Revit Schedules:
# - Fields, Filters, Sorting/Grouping, Field Formatting, Appearance
# - Dark-theme pickers, differences/similarities/both filter
# - Optional CSV export (prompted)

from pyrevit import revit, DB, forms, script
import csv

doc = revit.doc
out = script.get_output()
try:
    out.set_title("Compare Schedules")
    out.resize(1280, 900)
    out.center()
except:
    pass

# ----------------- helpers -----------------
def _eid(i):
    try:
        return i.IntegerValue if isinstance(i, DB.ElementId) else None
    except:
        return None

def _col(c):
    try:
        return (c.Red, c.Green, c.Blue) if isinstance(c, DB.Color) else None
    except:
        return None

def _status(values):
    if not values:
        return "same"
    f = values[0]
    for v in values[1:]:
        if v != f:
            return "diff"
    return "same"

def _print_table(section, headers, rows, show_mode, csv_rows):
    if not rows:
        return
    out.print_md("\n### " + section)
    header = "| Item | " + " | ".join(headers) + " | Status |"
    sep = "|" + "---|" * (len(headers) + 2)
    out.print_md(header)
    out.print_md(sep)
    for name, vals in rows:
        st = _status(vals)
        if show_mode == "Differences only" and st == "same":
            continue
        if show_mode == "Similarities only" and st == "diff":
            continue
        out.print_md("| {} | {} | {} |".format(name, " | ".join([unicode(v) for v in vals]),
                                               ("**diff**" if st == "diff" else "same")))
        csv_rows.append([section, name] + [unicode(v) for v in vals] + [st])

# ----------------- picking -----------------
def _all_schedules():
    try:
        col = DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule)
        return sorted([s for s in col], key=lambda s: s.Name.lower())
    except:
        return []

def pick_schedules():
    cands = _all_schedules()
    if len(cands) < 2:
        forms.alert("Need at least two schedules in this model.", exitscript=True)
    labels = [s.Name for s in cands]
    mapby = dict(zip(labels, cands))
    picks = forms.SelectFromList.show(labels,
                                      "Pick Schedules (2 or more)",
                                      multiselect=True, button_name="Use",
                                      dark_theme=True, height=600)
    if not picks or len(picks) < 2:
        forms.alert("Please pick at least two schedules.", exitscript=True)
    return [mapby[p] for p in picks]

SCHEDULE_ASPECTS = [
    ("Fields", True),
    ("Filters", True),
    ("Sorting/Grouping", True),
    ("Field Formatting", False),
    ("Appearance", False),
]

def pick_aspects():
    items = [a for a, _ in SCHEDULE_ASPECTS]
    picks = forms.SelectFromList.show(items,
                                      "Choose aspects to compare",
                                      multiselect=True, button_name="Next",
                                      dark_theme=True, height=450)
    if not picks:
        forms.alert("Nothing selected.", exitscript=True)
    return picks

def pick_mode():
    mode = forms.CommandSwitchWindow.show(
        ["Differences only", "Similarities only", "Both"],
        "Compare Options",
        dark_theme=True, width=420, height=220
    )
    if not mode:
        forms.alert("Cancelled.", exitscript=True)
    return mode

# ----------------- harvest (schedules) -----------------
def _defn(s):
    try:
        return s.Definition
    except:
        return None

def harvest_fields(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            f = d.GetField(i)
            nm = f.GetName() if hasattr(f, "GetName") else "<field>"
            hdr = f.ColumnHeading if hasattr(f, "ColumnHeading") else ""
            hidden = bool(getattr(f, "IsHidden", False))
            calc = bool(getattr(f, "IsCalculatedField", False))
            outl.append((nm, hdr, hidden, calc))
        except:
            pass
    return outl

def harvest_filters(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFilterCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            flt = d.GetFilter(i)
            fname = "<field>"
            try:
                fid = flt.FieldId
                fld = d.GetField(fid)
                fname = fld.GetName() if hasattr(fld, "GetName") else fname
            except:
                pass
            ftype = str(getattr(flt, "FilterType", ""))
            vstr = ""
            try:
                vstr = flt.GetStringValue()
            except:
                try:
                    vstr = str(flt.GetDoubleValue())
                except:
                    try:
                        vstr = str(flt.GetIntegerValue())
                    except:
                        vstr = ""
            outl.append((fname, ftype, vstr))
        except:
            pass
    return outl

def harvest_sorting(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetSortGroupFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            sg = d.GetSortGroupField(i)
            fname = "<field>"
            try:
                fld = d.GetField(sg.FieldId)
                fname = fld.GetName() if hasattr(fld, "GetName") else fname
            except:
                pass
            order = "Ascending" if getattr(sg, "SortOrder", 0) == DB.SortOrder.Ascending else "Descending"
            header = bool(getattr(sg, "ShowHeader", False))
            footer = bool(getattr(sg, "ShowFooter", False))
            outl.append((fname, order, header, footer))
        except:
            pass
    return outl

def harvest_formatting(s):
    outl = []
    d = _defn(s)
    if not d:
        return outl
    try:
        cnt = d.GetFieldCount()
    except:
        cnt = 0
    for i in range(cnt):
        try:
            f = d.GetField(i)
            nm = f.GetName() if hasattr(f, "GetName") else "<field>"
            width = getattr(f, "GridColumnWidth", None)
            align = str(getattr(f, "HorizontalAlignment", None))
            custom = bool(getattr(f, "HasCustomFormatting", False))
            outl.append((nm, width, align, custom))
        except:
            pass
    return outl

def harvest_appearance(s):
    m = {}
    try:
        m["ShowTitle"] = bool(getattr(s, "ShowTitle", False))
    except:
        pass
    try:
        m["Title"] = getattr(s, "Title", "")
    except:
        pass
    try:
        m["ShowHeaders"] = bool(getattr(s, "ShowHeaders", True))
    except:
        pass
    try:
        m["ShowGrandTotals"] = bool(getattr(s, "ShowGrandTotals", False))
    except:
        pass
    return m

# ----------------- compare -----------------
def compare_schedules(items, labels, aspects, mode, csv_rows):
    if "Fields" in aspects:
        names, per = set(), []
        for s in items:
            lst = harvest_fields(s); per.append(lst); names |= set([x[0] for x in lst])
        rows = []
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                if hit:
                    _, hdr, hidden, calc = hit[0]
                    vals.append("hdr='{}'; hidden={}; calc={}".format(hdr, hidden, calc))
                else:
                    vals.append("Missing")
            rows.append((nm, vals))
        _print_table("Schedule Fields", labels, rows, mode, csv_rows)

    if "Filters" in aspects:
        names, per = set(), []
        for s in items:
            lst = harvest_filters(s); per.append(lst); names |= set([x[0] for x in lst])
        rows = []
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                same_field = [x for x in lst if x[0] == nm]
                vals.append("; ".join(["{} {}".format(x[1], x[2]) for x in same_field]) if same_field else "None")
            rows.append((nm, vals))
        _print_table("Schedule Filters", labels, rows, mode, csv_rows)

    if "Sorting/Grouping" in aspects:
        names, per = set(), []
        for s in items:
            lst = harvest_sorting(s); per.append(lst); names |= set([x[0] for x in lst])
        rows = []
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                vals.append("{}; hdr={}; ftr={}".format(hit[0][1], hit[0][2], hit[0][3]) if hit else "None")
            rows.append((nm, vals))
        _print_table("Sorting / Grouping", labels, rows, mode, csv_rows)

    if "Field Formatting" in aspects:
        names, per = set(), []
        for s in items:
            lst = harvest_formatting(s); per.append(lst); names |= set([x[0] for x in lst])
        rows = []
        for nm in sorted(names, key=lambda s: s.lower()):
            vals = []
            for lst in per:
                hit = [x for x in lst if x[0] == nm]
                vals.append("w={}; align={}; custom={}".format(hit[0][1], hit[0][2], hit[0][3]) if hit else "None")
            rows.append((nm, vals))
        _print_table("Field Formatting", labels, rows, mode, csv_rows)

    if "Appearance" in aspects:
        keys, per = set(), []
        for s in items:
            m = harvest_appearance(s); per.append(m); keys |= set(m.keys())
        rows = []
        for k in sorted(keys, key=lambda s: s.lower()):
            rows.append((k, [m.get(k, "") for m in per]))
        _print_table("Appearance", labels, rows, mode, csv_rows)

# ----------------- run -----------------
def main():
    schedules = pick_schedules()
    labels = [s.Name for s in schedules]
    aspects = pick_aspects()
    mode = pick_mode()

    csv_rows = []
    out.print_md("# Compare Schedules")
    for i, lbl in enumerate(labels, 1):
        out.print_md("- **S{}**: {}".format(i, lbl))

    compare_schedules(schedules, labels, aspects, mode, csv_rows)

    if csv_rows:
        ans = forms.alert("Save results to CSV?", options=["Yes", "No"], warn_icon=False)
        if ans == "Yes":
            path = forms.save_file(file_ext="csv", title="Save Comparison CSV")
            if path:
                with open(path, "w") as f:
                    w = csv.writer(f, lineterminator="\n")
                    w.writerow(["Section", "Item"] + labels + ["Status"])
                    for r in csv_rows:
                        w.writerow(r)
                out.print_md("\n**CSV saved:** `{}`".format(path))
    else:
        out.print_md("\n_No rows to save._")

main()
