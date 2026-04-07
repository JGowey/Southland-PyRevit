# pyRevit: Dump PROJECT parameter group identifiers for this Revit version
import clr
from Autodesk.Revit import DB

rev_year = int(__revit__.Application.VersionNumber)

print("\n=== PROJECT parameter groups ===")
print("RevitYear:", rev_year)
print("")

if rev_year >= 2024:
    from System.Reflection import BindingFlags

    def label_for_group(ftid):
        try:
            return DB.LabelUtils.GetLabelForGroup(ftid)
        except:
            return None

    gt_type = clr.GetClrType(DB.GroupTypeId)
    props = gt_type.GetProperties(BindingFlags.Public | BindingFlags.Static)

    rows = []
    for p in props:
        try:
            gid = p.GetValue(None, None)
        except:
            continue
        lab = label_for_group(gid)
        if lab:
            rows.append((lab, "DB.GroupTypeId." + p.Name, gid.TypeId))

    rows.sort(key=lambda r: r[0].lower())
    for lab, api_name, typeid in rows:
        print("{:<30} | {:<30} | {}".format(lab, api_name, typeid))

else:
    Bipg = DB.BuiltInParameterGroup
    names = [n for n in dir(Bipg) if n.startswith("PG_")]

    rows = []
    for nm in names:
        try:
            val = getattr(Bipg, nm)
        except:
            continue
        try:
            lab = DB.LabelUtils.GetLabelFor(val)
        except:
            lab = None
        rows.append((lab or "(no label)", "DB.BuiltInParameterGroup." + nm, str(int(val))))

    rows.sort(key=lambda r: ((r[0] or "").lower(), r[1]))
    for lab, api_name, int_val in rows:
        print("{:<30} | {:<40} | {}".format(lab, api_name, int_val))
