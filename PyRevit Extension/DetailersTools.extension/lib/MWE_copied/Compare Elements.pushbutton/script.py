# -*- coding: utf-8 -*-
# Compare 2-100 Fabrication Elements with WPF Window
# Features: Side-by-side comparison, filtering modes, CSV export

from pyrevit import forms, revit, DB, script
from System import Enum
from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import DataGridTextColumn, DataGridLength, DataGridLengthUnitType
from System.Windows.Data import Binding
from System.Windows import Thickness
import System
import csv, datetime, codecs, os

doc = revit.doc
uidoc = revit.uidoc

MIN_PARTS, MAX_PARTS = 2, 100

# ==================== Data Model ====================
class ComparisonRow(object):
    """Represents one property row in the comparison"""
    def __init__(self, property_name, values, is_different):
        self.Property = property_name
        self.Values = values  # list of values for each element
        self.IsDifferent = is_different
        # Create dynamic properties for each element
        for i, val in enumerate(values):
            setattr(self, "Element{}".format(i), val)

# ==================== Selection ====================
def _is_fab(e):
    return isinstance(e, DB.FabricationPart)

def pick_fabrication_parts():
    """Pick fabrication parts from current selection or prompt user"""
    cur = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    cur = [e for e in cur if _is_fab(e)]
    
    if MIN_PARTS <= len(cur) <= MAX_PARTS:
        return cur
    
    try:
        from Autodesk.Revit.UI.Selection import ObjectType
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            "Pick {}-{} Fabrication Parts, then Finish".format(MIN_PARTS, MAX_PARTS)
        )
        picked = [doc.GetElement(r.ElementId) for r in refs]
    except:
        picked = []
    
    picked = [e for e in picked if _is_fab(e)]
    
    if not (MIN_PARTS <= len(picked) <= MAX_PARTS):
        forms.alert(
            'Please pick between {} and {} Fabrication Parts.'.format(MIN_PARTS, MAX_PARTS),
            exitscript=True
        )
    
    return picked

# ==================== Formatting ====================
def _trim_zeros(numstr):
    if "." not in numstr:
        return numstr
    s = numstr.rstrip("0").rstrip(".")
    return s if s else "0"

def _double_to_proj(val):
    try:
        fo = doc.GetUnits().GetFormatOptions(DB.SpecTypeId.Length)
        unit_id = fo.GetUnitTypeId()
        x = float(val)
        if abs(x) < 1e-9:
            return "0"
        s = "{0:.3f}".format(DB.UnitUtils.ConvertFromInternalUnits(x, unit_id))
        return _trim_zeros(s)
    except:
        try:
            s = "{0:.3f}".format(float(val))
            return _trim_zeros(s)
        except:
            return str(val)

def _string_or_value(elem, param_name):
    try:
        p = elem.LookupParameter(param_name)
        if p and p.HasValue:
            try:
                v = p.AsValueString()
                if v: return str(v)
            except:
                pass
            try:
                st = p.StorageType
                if st == DB.StorageType.String:
                    return p.AsString() or ""
                if st == DB.StorageType.Double:
                    return _double_to_proj(p.AsDouble())
                if st == DB.StorageType.Integer:
                    return str(p.AsInteger())
            except:
                pass
    except:
        pass
    return ""

# ==================== Fabrication Helpers ====================
def _cfg():
    try:
        return DB.FabricationConfiguration.GetFabricationConfiguration(doc)
    except:
        return None

def _resolve_id_to_name(prop_name, id_value):
    cfg = _cfg()
    if cfg is None:
        return str(id_value)
    
    try:
        if prop_name == "ServiceId":
            return cfg.GetServiceName(id_value) or str(id_value)
        if prop_name == "GaugeId":
            return cfg.GetGaugeName(id_value) or str(id_value)
        if prop_name in ("MaterialId", "ItemMaterialId"):
            return cfg.GetMaterialName(id_value) or str(id_value)
        if prop_name == "SpecificationId":
            return cfg.GetSpecificationName(id_value) or str(id_value)
        if prop_name == "InsulationSpecificationId":
            return cfg.GetInsulationSpecificationName(id_value) or str(id_value)
        if prop_name == "LiningSpecificationId":
            return cfg.GetLiningSpecificationName(id_value) or str(id_value)
    except:
        pass
    
    return str(id_value)

def _get_connector_name_by_id(elem, end_id_int):
    try:
        cfg = _cfg()
        cm = getattr(elem, "ConnectorManager", None)
        if not (cfg and cm):
            return "-"
        
        selected = None
        for c in cm.Connectors:
            try:
                if int(getattr(c, "Id", -1)) == end_id_int:
                    selected = c
                    break
            except:
                pass
        
        if selected is None:
            conns = list(cm.Connectors)
            if len(conns) > end_id_int:
                selected = conns[end_id_int]
        
        if selected is None:
            return "-"
        
        fci = selected.GetFabricationConnectorInfo()
        if fci:
            try:
                return cfg.GetFabricationConnectorName(fci.BodyConnectorId) or "-"
            except:
                return "-"
    except:
        pass
    
    return "-"

def _end_allowance(elem, endnum):
    names = [
        "End {} Connector Allowance", "Connector {} Allowance", "End {} Allowance",
        "Joint {} Allowance", "End {} Joint Allowance", "Allowance End {}"
    ]
    
    for pat in names:
        p = elem.LookupParameter(pat.format(endnum))
        if p and p.HasValue:
            try:
                st = p.StorageType
                if st == DB.StorageType.Double:
                    return _double_to_proj(p.AsDouble())
                if st == DB.StorageType.Integer:
                    return str(p.AsInteger())
                if st == DB.StorageType.String:
                    return p.AsString()
            except:
                pass
    
    return "-"

def _get_assembly_name(elem):
    try:
        aid = getattr(elem, "AssemblyInstanceId", None)
        if isinstance(aid, DB.ElementId) and aid.IntegerValue > 0:
            inst = doc.GetElement(aid)
            if inst and hasattr(inst, "Name") and inst.Name:
                return str(inst.Name)
    except:
        pass
    return "-"

# ==================== Property Skip/Hide ====================
SKIP_PROPS = set([
    "ConnectorManager", "GeometryChecksum", "HangerRodKit", "LiningArea", "Origin",
    "Oversize", "PartGuid", "PartStatus", "SpoolName", "ValidationStatus",
    "DoubleWallMaterialArea", "InsulationArea", "ItemCustomId",
    "ServiceName", "ServiceType", "Specification", "ServiceAbbreviation"
])

SKIP_PREFIXES = ("Product",)
HIDE_PROPS = set(["DisplayName", "DisplayNameSource", "PartType", "ItemName"])

def _skip_prop(pname):
    if pname in SKIP_PROPS:
        return True
    for pref in SKIP_PREFIXES:
        if pname.startswith(pref):
            return True
    return False

# ==================== DisplayName & Family ====================
DESC_PARAM_NAMES = [
    "Description", "Part Description", "Item Description",
    "MAP Description", "Fabrication Description", "Button Description"
]

def _get_param_string(elem, pname):
    try:
        p = elem.LookupParameter(pname)
        if p and p.HasValue:
            if p.StorageType == DB.StorageType.String:
                return p.AsString() or ""
            elif p.StorageType == DB.StorageType.Integer:
                return str(p.AsInteger())
            elif p.StorageType == DB.StorageType.Double:
                return _double_to_proj(p.AsDouble())
    except:
        pass
    return ""

def _get_fab_display_name(elem):
    # Try description params first
    for nm in DESC_PARAM_NAMES:
        val = _get_param_string(elem, nm)
        if val and val.strip():
            return val.strip()
    
    # Try ItemName
    try:
        if hasattr(elem, "ItemName") and elem.ItemName:
            return str(elem.ItemName)
    except:
        pass
    
    # Try ITM filename
    try:
        if hasattr(elem, "ItemSourceFileName") and elem.ItemSourceFileName:
            base = os.path.splitext(os.path.basename(elem.ItemSourceFileName))[0]
            if base:
                return str(base)
    except:
        pass
    
    return "-"

# ==================== Harvest Properties ====================
def harvest_fp_api(elem):
    d = {}
    
    # Get FabricationPart properties
    try:
        t = elem.GetType()
        props = [p for p in t.GetProperties() if p.DeclaringType == t and p.CanRead]
        
        for p in props:
            pname = str(p.Name)
            if _skip_prop(pname):
                continue
            
            try:
                raw = p.GetValue(elem, None)
            except:
                raw = None
            
            if raw is None:
                val = "-"
            else:
                try:
                    if isinstance(raw, (int, long)):
                        if pname.endswith("Id"):
                            val = _resolve_id_to_name(pname, raw)
                        else:
                            val = str(raw)
                    elif isinstance(raw, float):
                        lower = pname.lower()
                        if lower in ("width", "height", "diameter", "insulationthickness",
                                   "liningthickness", "length", "centerlinelength", "materialthickness"):
                            val = _double_to_proj(raw)
                        else:
                            val = str(raw)
                    elif isinstance(raw, bool):
                        val = "True" if raw else "False"
                    elif isinstance(raw, DB.ElementId):
                        val = str(raw.IntegerValue)
                    else:
                        val = str(raw)
                except:
                    val = str(raw) if raw is not None else "-"
            
            d[pname] = val
    except:
        pass
    
    # Format key parameters
    for nm in ("Size", "FreeSize", "OverallSize", "CenterlineLength",
               "MaterialThickness", "InsulationThickness", "LiningThickness"):
        val = _string_or_value(elem, nm)
        if val:
            d[nm] = val
    
    # Connectors and allowances
    d["EndConnector1"] = _get_connector_name_by_id(elem, 0)
    d["EndConnector2"] = _get_connector_name_by_id(elem, 1)
    d["End1Allowance"] = _end_allowance(elem, 1)
    d["End2Allowance"] = _end_allowance(elem, 2)
    
    # Assembly
    d["Assembly Name"] = _get_assembly_name(elem)
    
    # Names
    try:
        disp_val = _get_fab_display_name(elem)
        d["ItemName"] = str(getattr(elem, "ItemName", "") or "-")
        
        try:
            d["PartType"] = str(Enum.GetName(DB.FabricationPartType, elem.PartType) or
                              elem.PartType.ToString())
        except:
            d["PartType"] = "-"
        
        d["DisplayName"] = str(disp_val if disp_val else "-")
        
        # Family = DisplayName (trim before ':')
        fam_only = d["DisplayName"]
        if fam_only and ":" in fam_only:
            fam_only = fam_only.split(":", 1)[0].strip()
        
        d["Family"] = fam_only or "-"
    except:
        d["ItemName"] = d["PartType"] = d["DisplayName"] = d["Family"] = "-"
    
    return d

# ==================== Property Order ====================
PARAM_ORDER = [
    "DomainType", "Family", "Alias", "ItemNumber", "Assembly Name",
    "EndConnector1", "EndConnector2", "End1Allowance", "End2Allowance",
    "CutType", "CenterlineLength", "Size", "FreeSize", "OverallSize",
    "ServiceId", "Material", "MaterialGauge", "MaterialThickness",
    "Weight", "SheetMetalArea", "HasDoubleWall", "DoubleWallMaterial",
    "DoubleWallMaterialThickness", "HasInsulation", "InsulationSpecification",
    "InsulationThickness", "InsulationType", "HasLining", "LiningThickness",
    "LiningType", "LevelOffset", "BottomOfPartElevation", "TopOfPartElevation",
    "Slope", "IsBoughtOut", "Notes"
]

# ==================== WPF Window ====================
class CompareWindow(forms.WPFWindow):
    def __init__(self, xaml_path, elements, data):
        forms.WPFWindow.__init__(self, xaml_path)
        
        self.elements = elements
        self.data = data
        self.all_rows = []
        self.current_mode = "Everything"
        
        # Prepare data
        self._prepare_data()
        
        # Update UI
        self._update_title()
        self._show_mode("Everything")
    
    def _prepare_data(self):
        """Prepare comparison data"""
        # Get all properties
        prop_set = set()
        for d in self.data:
            prop_set.update(d.keys())
        
        prop_set = prop_set.difference(HIDE_PROPS)
        
        # Order properties
        all_props = [p for p in PARAM_ORDER if p in prop_set]
        remaining = [p for p in prop_set if p not in PARAM_ORDER]
        all_props.extend(sorted(remaining, key=lambda s: s.lower()))
        
        # Create rows
        for prop in all_props:
            vals = [self.data[i].get(prop, "-") for i in range(len(self.elements))]
            is_diff = len(set(vals)) > 1
            
            row = ComparisonRow(prop, vals, is_diff)
            self.all_rows.append(row)
    
    def _update_title(self):
        """Update window title"""
        self.SubtitleText.Text = "Comparing {} Fabrication Parts".format(len(self.elements))
    
    def _show_mode(self, mode):
        """Filter and show rows based on mode"""
        self.current_mode = mode
        
        if mode == "Everything":
            filtered = self.all_rows
        elif mode == "Differences":
            filtered = [r for r in self.all_rows if r.IsDifferent]
        elif mode == "Same":
            filtered = [r for r in self.all_rows if not r.IsDifferent]
        else:
            filtered = self.all_rows
        
        # Update filter status
        diff_count = sum(1 for r in self.all_rows if r.IsDifferent)
        same_count = len(self.all_rows) - diff_count
        
        if mode == "Everything":
            self.FilterStatusText.Text = "{} properties ({} diff, {} same)".format(
                len(filtered), diff_count, same_count
            )
        elif mode == "Differences":
            self.FilterStatusText.Text = "{} different properties".format(len(filtered))
        elif mode == "Same":
            self.FilterStatusText.Text = "{} identical properties".format(len(filtered))
        
        # Bind to grid
        self._bind_grid(filtered)
        
        # Update status
        self.StatusText.Text = "Showing {} of {} properties".format(
            len(filtered), len(self.all_rows)
        )
    
    def _bind_grid(self, rows):
        """Bind rows to DataGrid"""
        # Clear existing columns
        self.ResultsGrid.Columns.Clear()
        
        # Add Property column
        prop_col = DataGridTextColumn()
        prop_col.Header = "Property"
        prop_col.Binding = Binding("Property")
        prop_col.Width = DataGridLength(200)
        self.ResultsGrid.Columns.Add(prop_col)
        
        # Add element columns
        for i, elem in enumerate(self.elements):
            col = DataGridTextColumn()
            col.Header = "E{} (Id: {})".format(i + 1, elem.Id.IntegerValue)
            col.Binding = Binding("Element{}".format(i))
            col.Width = DataGridLength(1, DataGridLengthUnitType.Star)
            self.ResultsGrid.Columns.Add(col)
        
        # Bind data
        oc = ObservableCollection[ComparisonRow]()
        for row in rows:
            oc.Add(row)
        
        self.ResultsGrid.ItemsSource = oc
    
    # Event handlers
    def OnShowEverything(self, sender, args):
        self._show_mode("Everything")
    
    def OnShowDifferences(self, sender, args):
        self._show_mode("Differences")
    
    def OnShowSame(self, sender, args):
        self._show_mode("Same")
    
    def OnExport(self, sender, args):
        """Export to CSV"""
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_path = forms.save_file(
            file_ext="csv",
            default_name="FabParts_Compare_{}elem_{}.csv".format(len(self.elements), ts),
            title="Save Comparison Results"
        )
        
        if csv_path:
            try:
                with codecs.open(csv_path, "w", "utf-8") as f:
                    w = csv.writer(f, lineterminator="\n")
                    
                    # Header
                    header = ["Property"] + ["E{} (Id: {})".format(i + 1, elem.Id.IntegerValue)
                                            for i, elem in enumerate(self.elements)]
                    header.append("Status")
                    w.writerow(header)
                    
                    # Data
                    for row in self.all_rows:
                        csv_row = [row.Property] + row.Values
                        csv_row.append("Different" if row.IsDifferent else "Same")
                        w.writerow(csv_row)
                
                self.StatusText.Text = "Exported to: {}".format(csv_path)
                forms.toast("CSV exported successfully", title="Export Complete")
            except Exception as ex:
                forms.alert("Failed to export CSV:\n{}".format(ex), title="Export Error")
    
    def OnClose(self, sender, args):
        """Close window"""
        self.Close()

# ==================== Main ====================
def run():
    # Pick elements
    elements = pick_fabrication_parts()
    
    if not elements:
        forms.alert("No elements selected.", exitscript=True)
    
    # Harvest data
    data = [harvest_fp_api(el) for el in elements]
    
    # Show window
    import os
    script_dir = os.path.dirname(__file__)
    xaml_path = os.path.join(script_dir, "window.xaml")
    
    if not os.path.exists(xaml_path):
        forms.alert("window.xaml not found at: {}".format(xaml_path), exitscript=True)
    
    window = CompareWindow(xaml_path, elements, data)
    window.ShowDialog()

if __name__ == "__main__":
    run()