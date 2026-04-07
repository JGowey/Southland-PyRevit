# -*- coding: utf-8 -*-
"""Coil Duct - Calculate overall cut lengths for coil machine.

Analyzes rectangular fabrication duct, reads end connections via
FabricationConnectorInfo, and calculates overall cut length including
connector allowances. Checks whether each piece fits on the coil.

Overall Cut Length = Body Length + End 1 Allowance + End 2 Allowance
Must be <= Coil Width (typically 60" or 48").

After analysis:
  - "Should be Coil": fits coil, valid connectors, but NOT CutType 2.
    Gets tagged and selected.
  - "Should not be Coil": IS CutType 2 but has connectors NOT in the
    enabled list. Gets tagged.

Tags are written to the user-selected text parameter.

Author: Jeremiah Griffith
Version: 1.3.0
"""

from __future__ import division
from pyrevit import revit, DB, forms, script
import clr
import re
import json
import os

clr.AddReference("System")
from System.Collections.Generic import List
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter

doc = revit.doc
uidoc = revit.uidoc
out = script.get_output()

CONFIG_FILE = os.path.join(os.path.dirname(__file__), "coil_config.json")
TAG_SHOULD_BE = "Should be Coil"
TAG_SHOULD_NOT = "Should not be Coil"
TAG_TOO_LONG = "Fitting too Long"

# ==============================================================================
#   DEFAULT CONNECTOR ALLOWANCES
# ==============================================================================

DEFAULT_CONNECTORS = {
    "None":        {"enabled": True,  "allowance": 0.0},
    "RAW":         {"enabled": True,  "allowance": 0.0},
    "RAW 1 NOTCH": {"enabled": True,  "allowance": 0.0},
    "DM 35":       {"enabled": True,  "allowance": 0.0},
    "S&D":         {"enabled": True,  "allowance": 0.5},
    "S&S":         {"enabled": True,  "allowance": 0.5},
    "TDC":         {"enabled": True,  "allowance": 1.875},
}

DEFAULT_COIL_WIDTH = 60.0


# ==============================================================================
#   CONNECTOR ENTRY CLASS
# ==============================================================================

class ConnectorEntry(object):
    """Bindable row for the connector DataGrid."""
    def __init__(self, short_name, full_name="", allowance=0.0, is_enabled=False, connector_id=-1):
        self.ShortName = short_name
        self.FullName = full_name
        self.Allowance = allowance
        self.IsEnabled = is_enabled
        self.ConnectorId = connector_id


# ==============================================================================
#   FABRICATION HELPERS
# ==============================================================================

def get_fab_config():
    try:
        return DB.FabricationConfiguration.GetFabricationConfiguration(doc)
    except:
        return None


def get_connector_name_from_connector(cfg, connector):
    """Get full connector name string from a Connector object."""
    try:
        fci = connector.GetFabricationConnectorInfo()
        if fci:
            name = cfg.GetFabricationConnectorName(fci.BodyConnectorId)
            cid = fci.BodyConnectorId
            return (name or "", cid)
    except:
        pass
    return ("", -1)


def extract_short_name(full_name):
    """
    Extract the short connector name from the full string.
    'HVAC - NW - TDC: TDC' -> 'TDC'
    None or empty -> 'None'
    """
    if not full_name or not full_name.strip():
        return "None"
    parts = full_name.split(":")
    if len(parts) >= 2:
        return parts[-1].strip()
    return full_name.strip()


def discover_connectors_from_model():
    """
    Scan all FabricationParts in the model to discover unique connectors
    on RECTANGULAR connector profiles only.
    Returns dict: {short_name: {"full_name": str, "connector_id": int}}
    """
    cfg = get_fab_config()
    if not cfg:
        return {}

    found = {}
    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork) \
        .WhereElementIsNotElementType()

    for e in collector:
        if not isinstance(e, DB.FabricationPart):
            continue
        try:
            cm = e.ConnectorManager
            if not cm:
                continue
            for c in cm.Connectors:
                # Only include rectangular connectors
                try:
                    if c.Shape != DB.ConnectorProfileType.Rectangular:
                        continue
                except:
                    continue

                full_name, cid = get_connector_name_from_connector(cfg, c)
                short = extract_short_name(full_name)
                if short and short not in found:
                    found[short] = {"full_name": full_name, "connector_id": cid}
        except:
            continue

    return found


def discover_connectors_from_database():
    """
    Brute-force scan the fabrication database for all connector definitions.
    Tries connector IDs 0-5000 and returns any that resolve to a name.
    Filters to names containing 'NW' (rectangular duct connectors) based on
    the naming convention seen in the database: 'HVAC - NW - ...'

    Returns dict: {short_name: {"full_name": str, "connector_id": int}}
    """
    cfg = get_fab_config()
    if not cfg:
        return {}

    found = {}
    for cid in range(5000):
        try:
            name = cfg.GetFabricationConnectorName(cid)
            if not name or not name.strip():
                continue
            short = extract_short_name(name)
            if short and short not in found:
                found[short] = {"full_name": name, "connector_id": cid}
        except:
            continue

    return found


def discover_text_parameters():
    """
    Find all writable text (string) parameters available on fabrication ductwork.
    Samples up to 5 elements to build a union of parameter names.
    Returns sorted list of parameter name strings.
    """
    collector = DB.FilteredElementCollector(doc) \
        .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork) \
        .WhereElementIsNotElementType()

    param_names = set()
    count = 0
    for e in collector:
        if not isinstance(e, DB.FabricationPart):
            continue
        for p in e.Parameters:
            try:
                if p.StorageType == DB.StorageType.String and not p.IsReadOnly:
                    name = p.Definition.Name
                    if name:
                        param_names.add(name)
            except:
                continue
        count += 1
        if count >= 5:
            break

    # Also add the built-in Comments parameter by name
    param_names.add("Comments")

    return sorted(param_names)


def is_rect(e):
    """Check if element is rectangular."""
    try:
        size = e.OverallSize or ""
        if "x" in size.lower():
            return True
    except:
        pass
    try:
        for c in e.ConnectorManager.Connectors:
            if c.Shape == DB.ConnectorProfileType.Rectangular:
                return True
    except:
        pass
    return False


def is_straight(e):
    """Check if a FabricationPart is a straight (not transition, elbow, tap, etc.)."""
    try:
        return e.IsAStraight()
    except:
        return False


def get_cut_type(e):
    """Get CutType as int. Returns -1 if unavailable."""
    try:
        return int(e.CutType)
    except:
        return -1


def centerline_length_inches(e):
    """Get centerline length in inches."""
    try:
        cl = e.CenterlineLength
        if cl and cl > 0:
            return round(cl * 12.0, 4)
    except:
        pass
    return None


def get_both_ends(cfg, elem):
    """
    Get connector info for a fabrication part.
    Returns list of (full_name, short_name, connector_id) tuples.
    """
    results = []
    try:
        cm = elem.ConnectorManager
        if not cm:
            return results
        for c in cm.Connectors:
            full_name, cid = get_connector_name_from_connector(cfg, c)
            short = extract_short_name(full_name)
            results.append((full_name, short, cid))
    except:
        pass
    return results


def write_param(elem, param_name, text):
    """Write text to a named parameter. Tries LookupParameter first, then BuiltInParameter for Comments."""
    p = elem.LookupParameter(param_name)
    if p and p.StorageType == DB.StorageType.String and not p.IsReadOnly:
        p.Set(text)
        return True

    # Fallback for Comments
    if param_name == "Comments":
        p = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if p and not p.IsReadOnly:
            p.Set(text)
            return True

    return False


class FabPartFilter(ISelectionFilter):
    def AllowElement(self, elem):
        try:
            return elem.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_FabricationDuctwork)
        except:
            return False
    def AllowReference(self, reference, position):
        return False


# ==============================================================================
#   WPF WINDOW
# ==============================================================================

class CoilDuctWindow(forms.WPFWindow):
    def __init__(self, xaml_file, preselection=None):
        self.elements = preselection or []
        self._next_action = None
        self._connector_entries = []
        self.selected_param_name = "Comments"

        forms.WPFWindow.__init__(self, xaml_file)

        # Load saved defaults
        saved = self._load_config(CONFIG_FILE)

        if saved and "coil_width" in saved:
            self.CoilWidth.Text = str(saved["coil_width"])
        else:
            self.CoilWidth.Text = str(DEFAULT_COIL_WIDTH)

        # Populate parameter dropdown
        self._populate_param_combo(saved)

        # Build connector list
        self._build_connector_list(saved)
        self._update_selection_count()
        self._refresh_grid()

    # --- Parameter dropdown ---

    def _populate_param_combo(self, saved_config=None):
        params = discover_text_parameters()
        self.TagParamCombo.Items.Clear()
        for name in params:
            self.TagParamCombo.Items.Add(name)

        # Restore saved selection or default to Comments
        saved_param = None
        if saved_config and "tag_parameter" in saved_config:
            saved_param = saved_config["tag_parameter"]

        target = saved_param or "Comments"
        if target in params:
            self.TagParamCombo.SelectedItem = target
        elif params:
            self.TagParamCombo.SelectedIndex = 0

    # --- Config I/O ---

    def _load_config(self, path):
        if not os.path.exists(path):
            return None
        try:
            with open(path, 'r') as f:
                return json.load(f)
        except:
            return None

    def _save_config(self, path):
        data = self._build_save_data()
        try:
            with open(path, 'w') as f:
                json.dump(data, f, indent=4)
            return True
        except:
            return False

    def _build_save_data(self):
        self._read_grid_state()
        connectors = {}
        for entry in self._connector_entries:
            connectors[entry.ShortName] = {
                "enabled": entry.IsEnabled,
                "allowance": entry.Allowance,
                "full_name": entry.FullName,
                "connector_id": entry.ConnectorId,
            }
        try:
            cw = float(self.CoilWidth.Text)
        except:
            cw = DEFAULT_COIL_WIDTH

        tag_param = self.TagParamCombo.SelectedItem or "Comments"

        return {
            "coil_width": cw,
            "tag_parameter": tag_param,
            "connectors": connectors,
        }

    # --- Connector List Building ---

    def _build_connector_list(self, saved_config=None):
        self._connector_entries = []
        merged = {}

        for short_name, info in DEFAULT_CONNECTORS.items():
            merged[short_name] = {
                "enabled": info["enabled"],
                "allowance": info["allowance"],
                "full_name": "",
                "connector_id": -1,
            }

        if saved_config and "connectors" in saved_config:
            for short_name, info in saved_config["connectors"].items():
                if short_name in merged:
                    merged[short_name]["enabled"] = info.get("enabled", merged[short_name]["enabled"])
                    merged[short_name]["allowance"] = info.get("allowance", merged[short_name]["allowance"])
                    merged[short_name]["full_name"] = info.get("full_name", "")
                    merged[short_name]["connector_id"] = info.get("connector_id", -1)
                else:
                    merged[short_name] = {
                        "enabled": info.get("enabled", False),
                        "allowance": info.get("allowance", 0.0),
                        "full_name": info.get("full_name", ""),
                        "connector_id": info.get("connector_id", -1),
                    }

        for short_name in sorted(merged.keys()):
            info = merged[short_name]
            self._connector_entries.append(ConnectorEntry(
                short_name=short_name,
                full_name=info["full_name"],
                allowance=info["allowance"],
                is_enabled=info["enabled"],
                connector_id=info["connector_id"],
            ))

    def _merge_discovered_connectors(self, discovered):
        existing_names = {e.ShortName for e in self._connector_entries}
        added = 0
        for short_name, info in discovered.items():
            if short_name not in existing_names:
                self._connector_entries.append(ConnectorEntry(
                    short_name=short_name,
                    full_name=info.get("full_name", ""),
                    allowance=0.0,
                    is_enabled=False,
                    connector_id=info.get("connector_id", -1),
                ))
                added += 1
            else:
                for entry in self._connector_entries:
                    if entry.ShortName == short_name:
                        if not entry.FullName and info.get("full_name"):
                            entry.FullName = info["full_name"]
                        if entry.ConnectorId < 0 and info.get("connector_id", -1) >= 0:
                            entry.ConnectorId = info["connector_id"]
                        break

        self._connector_entries.sort(key=lambda x: x.ShortName)
        return added

    def _refresh_grid(self):
        self.ConnectorGrid.ItemsSource = None
        self.ConnectorGrid.ItemsSource = self._connector_entries

    def _update_selection_count(self):
        n = len(self.elements)
        self.SelectionCount.Text = "{} elements selected".format(n)
        self.RunBtn.IsEnabled = n > 0

    def _read_grid_state(self):
        try:
            self.ConnectorGrid.CommitEdit()
        except:
            pass

        for entry in self._connector_entries:
            try:
                entry.Allowance = float(entry.Allowance)
            except (ValueError, TypeError):
                entry.Allowance = 0.0
            try:
                entry.IsEnabled = bool(entry.IsEnabled)
            except:
                entry.IsEnabled = False

        self.selected_param_name = self.TagParamCombo.SelectedItem or "Comments"

    # --- Button Handlers ---

    def ScanModel_Click(self, sender, e):
        discovered = discover_connectors_from_model()
        added = self._merge_discovered_connectors(discovered)
        self._refresh_grid()
        total = len(discovered)
        self.ProfileStatus.Text = "Model: {} rect connectors ({} new)".format(total, added)

    def ScanDb_Click(self, sender, e):
        discovered = discover_connectors_from_database()
        added = self._merge_discovered_connectors(discovered)
        self._refresh_grid()
        total = len(discovered)
        self.ProfileStatus.Text = "Database: {} connectors ({} new)".format(total, added)

    def EnableAll_Click(self, sender, e):
        for entry in self._connector_entries:
            entry.IsEnabled = True
        self._refresh_grid()

    def DisableAll_Click(self, sender, e):
        for entry in self._connector_entries:
            entry.IsEnabled = False
        self._refresh_grid()

    def Import_Click(self, sender, e):
        path = forms.pick_file(file_ext='json')
        if path:
            data = self._load_config(path)
            if data:
                if "coil_width" in data:
                    self.CoilWidth.Text = str(data["coil_width"])
                if "tag_parameter" in data:
                    target = data["tag_parameter"]
                    if target in [self.TagParamCombo.Items[i] for i in range(self.TagParamCombo.Items.Count)]:
                        self.TagParamCombo.SelectedItem = target
                self._build_connector_list(data)
                self._refresh_grid()
                self.ProfileStatus.Text = "Imported: {}".format(os.path.basename(path))
            else:
                forms.alert("Could not read profile from:\n{}".format(path))

    def Export_Click(self, sender, e):
        path = forms.save_file(file_ext='json', default_name='coil_profile')
        if path:
            if self._save_config(path):
                self.ProfileStatus.Text = "Exported: {}".format(os.path.basename(path))
            else:
                forms.alert("Error saving profile.")

    def SaveDefault_Click(self, sender, e):
        if self._save_config(CONFIG_FILE):
            self.ProfileStatus.Text = "Defaults saved."
        else:
            forms.alert("Error saving defaults.")

    def Pick_Click(self, sender, e):
        self._read_grid_state()
        self._next_action = "pick"
        self.Close()

    def Run_Click(self, sender, e):
        self._read_grid_state()
        self._next_action = "run"
        self.Close()

    def Cancel_Click(self, sender, e):
        self._next_action = None
        self.Close()


# ==============================================================================
#   FAILURE SUPPRESSION
# ==============================================================================

class _SuppressWarnings(DB.IFailuresPreprocessor):
    """Delete all warnings (including Duplicate Mark) during tagging."""
    def PreprocessFailures(self, failuresAccessor):
        has_errors = False
        try:
            for msg in failuresAccessor.GetFailureMessages():
                try:
                    sev = msg.GetSeverity()
                    if sev == DB.FailureSeverity.Warning:
                        failuresAccessor.DeleteWarning(msg)
                    else:
                        has_errors = True
                except:
                    has_errors = True
        except:
            pass
        if has_errors:
            return DB.FailureProcessingResult.ProceedWithRollBack
        return DB.FailureProcessingResult.Continue


# ==============================================================================
#   ANALYSIS
# ==============================================================================

def analyze(parts, coil_width, connector_map):
    """
    Analyze fabrication parts for coil eligibility.

    Only straight duct can be coil candidates. Transitions, elbows, taps, etc.
    are filtered out. CutType 2 non-straights are explicitly flagged as
    "Should not be Coil" with reason "Not a straight".

    Five buckets:
        already_coil    - CutType 2, straight, valid connectors, fits coil
        should_be_coil  - NOT CutType 2, straight, valid connectors, fits coil
        should_not_coil - CutType 2 but not straight or has bad connectors
        too_long        - straight, valid connectors but overall exceeds coil width
        skipped_not_rect - count of non-rectangular parts

    Each result tuple (except skipped count):
        (elem, body_len, short1, short2, allow1, allow2, overall, size, cut_type)
    bad_connector rows use:
        (elem, body_len, short1, short2, reason, size, cut_type)
    """
    cfg = get_fab_config()
    already_coil = []
    should_be_coil = []
    should_not_coil = []
    too_long = []
    skipped_not_rect = 0

    for e in parts:
        if not isinstance(e, DB.FabricationPart):
            continue

        if not is_rect(e):
            skipped_not_rect += 1
            continue

        body_len = centerline_length_inches(e)
        size = e.OverallSize or e.FreeSize or "?"
        ct = get_cut_type(e)
        straight = is_straight(e)

        ends = get_both_ends(cfg, e) if cfg else []

        if len(ends) < 2:
            if ct == 2:
                should_not_coil.append(
                    (e, body_len, "?", "?", "Could not read connectors", size, ct)
                )
            continue

        _, short1, _ = ends[0]
        _, short2, _ = ends[1]

        valid1 = short1 in connector_map
        valid2 = short2 in connector_map

        # CutType 2 but not a straight = should NOT be coil
        if ct == 2 and not straight:
            should_not_coil.append(
                (e, body_len, short1, short2, "Not a straight", size, ct)
            )
            continue

        # CutType 2 straight but invalid connectors = should NOT be coil
        if ct == 2 and (not valid1 or not valid2):
            bad_parts = []
            if not valid1:
                bad_parts.append("End 1: {}".format(short1))
            if not valid2:
                bad_parts.append("End 2: {}".format(short2))
            should_not_coil.append(
                (e, body_len, short1, short2, ", ".join(bad_parts), size, ct)
            )
            continue

        # Non-CutType-2 with bad connectors: skip (not actionable)
        if not valid1 or not valid2:
            continue

        # Not a straight: skip (only straights can be coil candidates)
        if not straight:
            continue

        # From here: straight, valid connectors
        if body_len is None:
            continue

        allow1 = connector_map[short1]
        allow2 = connector_map[short2]
        overall = body_len + allow1 + allow2

        row = (e, body_len, short1, short2, allow1, allow2, overall, size, ct)

        if overall > coil_width + 0.25:
            too_long.append(row)
        elif ct == 2:
            already_coil.append(row)
        else:
            should_be_coil.append(row)

    return (already_coil, should_be_coil, should_not_coil, too_long, skipped_not_rect)


# ==============================================================================
#   REPORT
# ==============================================================================

def ct_label(ct):
    """Human-readable CutType label."""
    if ct == 0:
        return "None"
    elif ct == 1:
        return "Value"
    elif ct == 2:
        return "Auto"
    return str(ct)


def print_report(coil_width, total_count, already_coil, should_be_coil,
                 should_not_coil, too_long, skipped,
                 tagged_should, tagged_should_not, tagged_too_long,
                 param_name):
    out.print_md("# Coil Duct Analyzer")
    out.print_md("**Coil Width:** {:.0f}\"  |  **Tag Parameter:** {}".format(coil_width, param_name))
    out.print_md("**Total Parts Analyzed:** {}".format(total_count))
    if skipped:
        out.print_md("*Skipped {} non-rectangular parts*".format(skipped))

    action_count = len(should_be_coil) + len(should_not_coil) + len(too_long)
    if action_count:
        out.print_md("**{} actionable pieces colored (green = should be coil, red = should not, purple = too long)**".format(action_count))

    out.print_md("---")

    # --- SHOULD BE COIL (the primary actionable output) ---
    if should_be_coil:
        out.print_md("## [!] Should be Coil ({} pieces)".format(
            len(should_be_coil)))
        if tagged_should:
            out.print_md('*Tagged {} pieces with "{}" in {}*'.format(
                tagged_should, TAG_SHOULD_BE, param_name))
        out.print_md("| ID | Size | CutType | Body (in) | End 1 | End 2 | +Allow 1 | +Allow 2 | Overall (in) |")
        out.print_md("|---|---|---|---|---|---|---|---|---|")

        total_linear = 0.0
        for (e, body, s1, s2, a1, a2, overall, size, ct) in sorted(should_be_coil, key=lambda x: -x[6]):
            total_linear += overall
            out.print_md("| {} | {} | {} | {:.2f} | {} | {} | {:.3f} | {:.3f} | **{:.3f}** |".format(
                out.linkify(e.Id), size, ct_label(ct), body, s1, s2, a1, a2, overall))

        out.print_md("")
        out.print_md("**Total linear cut:** {:.2f}\" ({:.1f}')".format(
            total_linear, total_linear / 12.0))
    else:
        out.print_md("## No pieces need to be changed to coil.")

    # --- SHOULD NOT BE COIL ---
    if should_not_coil:
        out.print_md("---")
        out.print_md("## [X] Should NOT be Coil ({} pieces)".format(
            len(should_not_coil)))
        if tagged_should_not:
            out.print_md('*Tagged {} pieces with "{}" in {}*'.format(
                tagged_should_not, TAG_SHOULD_NOT, param_name))
        out.print_md("| ID | Size | End 1 | End 2 | Issue |")
        out.print_md("|---|---|---|---|---|")
        for (e, body, s1, s2, reason, size, ct) in should_not_coil:
            out.print_md("| {} | {} | {} | {} | {} |".format(
                out.linkify(e.Id), size, s1, s2, reason))

    # --- ALREADY ON COIL ---
    if already_coil:
        out.print_md("---")
        out.print_md("## Already Coil / CutType 2 ({} pieces)".format(len(already_coil)))
        out.print_md("| ID | Size | Body (in) | End 1 | End 2 | +Allow 1 | +Allow 2 | Overall (in) |")
        out.print_md("|---|---|---|---|---|---|---|---|")

        total_linear = 0.0
        for (e, body, s1, s2, a1, a2, overall, size, ct) in sorted(already_coil, key=lambda x: -x[6]):
            total_linear += overall
            out.print_md("| {} | {} | {:.2f} | {} | {} | {:.3f} | {:.3f} | **{:.3f}** |".format(
                out.linkify(e.Id), size, body, s1, s2, a1, a2, overall))

        out.print_md("")
        out.print_md("**Total linear cut:** {:.2f}\" ({:.1f}')".format(
            total_linear, total_linear / 12.0))

    # --- TOO LONG FOR COIL ---
    if too_long:
        out.print_md("---")
        out.print_md("## Exceeds Coil Width ({} pieces)".format(len(too_long)))
        if tagged_too_long:
            out.print_md('*Tagged {} pieces with "{}" in {}*'.format(
                tagged_too_long, TAG_TOO_LONG, param_name))
        out.print_md("| ID | Size | CutType | Body (in) | End 1 | End 2 | +Allow 1 | +Allow 2 | Overall (in) | Over by |")
        out.print_md("|---|---|---|---|---|---|---|---|---|---|")
        for (e, body, s1, s2, a1, a2, overall, size, ct) in sorted(too_long, key=lambda x: -x[6]):
            over = overall - coil_width
            out.print_md("| {} | {} | {} | {:.2f} | {} | {} | {:.3f} | {:.3f} | **{:.3f}** | {:.3f} |".format(
                out.linkify(e.Id), size, ct_label(ct), body, s1, s2, a1, a2, overall, over))


# ==============================================================================
#   MAIN
# ==============================================================================

def main():
    xaml_path = os.path.join(os.path.dirname(__file__), "window.xaml")
    if not os.path.exists(xaml_path):
        forms.alert("Could not find window.xaml", title="Error")
        return

    sel = uidoc.Selection.GetElementIds()
    preselection = [doc.GetElement(eid) for eid in sel]
    preselection = [p for p in preselection if isinstance(p, DB.FabricationPart)]

    while True:
        win = CoilDuctWindow(xaml_path, preselection)
        win.ShowDialog()

        action = getattr(win, "_next_action", None)

        if not action:
            break

        if action == "pick":
            try:
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element, FabPartFilter(),
                    "Select Fabrication Parts"
                )
                preselection = [doc.GetElement(r.ElementId) for r in refs]
            except:
                pass
            continue

        if action == "run":
            parts = preselection

            if not parts:
                choice = forms.CommandSwitchWindow.show(
                    ["All Fab Duct in Model", "Cancel"],
                    message="No elements selected."
                )
                if not choice or "Cancel" in choice:
                    continue

                collector = DB.FilteredElementCollector(doc) \
                    .OfCategory(DB.BuiltInCategory.OST_FabricationDuctwork) \
                    .WhereElementIsNotElementType()
                parts = [e for e in collector if isinstance(e, DB.FabricationPart)]

            if not parts:
                forms.alert("No fabrication parts found.", title="Coil Duct")
                continue

            try:
                coil_width = float(win.CoilWidth.Text)
            except:
                coil_width = DEFAULT_COIL_WIDTH

            param_name = win.selected_param_name

            connector_map = {}
            for entry in win._connector_entries:
                if entry.IsEnabled:
                    connector_map[entry.ShortName] = entry.Allowance

            # Run analysis
            already_coil, should_be_coil, should_not_coil, too_long, skipped = analyze(
                parts, coil_width, connector_map
            )

            # Tag elements in a single transaction
            tagged_should = 0
            tagged_should_not = 0
            tagged_too_long = 0

            if should_be_coil or should_not_coil or too_long:
                t = DB.Transaction(doc, "Coil Duct - Tag Results")
                opts = t.GetFailureHandlingOptions()
                opts.SetFailuresPreprocessor(_SuppressWarnings())
                t.SetFailureHandlingOptions(opts)
                t.Start()
                try:
                    for (e, _, _, _, _, _, _, _, _) in should_be_coil:
                        if write_param(e, param_name, TAG_SHOULD_BE):
                            tagged_should += 1

                    for (e, _, _, _, _, _, _) in should_not_coil:
                        if write_param(e, param_name, TAG_SHOULD_NOT):
                            tagged_should_not += 1

                    for (e, _, _, _, _, _, _, _, _) in too_long:
                        if write_param(e, param_name, TAG_TOO_LONG):
                            tagged_too_long += 1

                    t.Commit()
                except:
                    t.RollBack()

            # Color actionable pieces in the active view
            # Green = should be coil, Red = should not be coil, Purple = too long
            active_view = doc.ActiveView
            green = DB.Color(0, 180, 0)
            red = DB.Color(220, 40, 40)
            purple = DB.Color(150, 0, 200)

            if should_be_coil or should_not_coil or too_long:
                t2 = DB.Transaction(doc, "Coil Duct - Color Overrides")
                opts2 = t2.GetFailureHandlingOptions()
                opts2.SetFailuresPreprocessor(_SuppressWarnings())
                t2.SetFailureHandlingOptions(opts2)
                t2.Start()
                try:
                    ogs_green = DB.OverrideGraphicSettings()
                    ogs_green.SetProjectionLineColor(green)
                    ogs_green.SetSurfaceForegroundPatternColor(green)
                    ogs_green.SetCutLineColor(green)
                    ogs_green.SetCutForegroundPatternColor(green)

                    ogs_red = DB.OverrideGraphicSettings()
                    ogs_red.SetProjectionLineColor(red)
                    ogs_red.SetSurfaceForegroundPatternColor(red)
                    ogs_red.SetCutLineColor(red)
                    ogs_red.SetCutForegroundPatternColor(red)

                    ogs_purple = DB.OverrideGraphicSettings()
                    ogs_purple.SetProjectionLineColor(purple)
                    ogs_purple.SetSurfaceForegroundPatternColor(purple)
                    ogs_purple.SetCutLineColor(purple)
                    ogs_purple.SetCutForegroundPatternColor(purple)

                    for (e, _, _, _, _, _, _, _, _) in should_be_coil:
                        active_view.SetElementOverrides(e.Id, ogs_green)

                    for (e, _, _, _, _, _, _) in should_not_coil:
                        active_view.SetElementOverrides(e.Id, ogs_red)

                    for (e, _, _, _, _, _, _, _, _) in too_long:
                        active_view.SetElementOverrides(e.Id, ogs_purple)

                    t2.Commit()
                except:
                    t2.RollBack()

            # Print report
            print_report(coil_width, len(parts), already_coil, should_be_coil,
                         should_not_coil, too_long, skipped,
                         tagged_should, tagged_should_not, tagged_too_long,
                         param_name)

            break


if __name__ == "__main__":
    main()
