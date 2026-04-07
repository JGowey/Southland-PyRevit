# -*- coding: utf-8 -*-
"""Nested Link Scan - Finds Attachment-type sub-links bleeding through Overlay links.

When a Revit link is set to Overlay reference type, its own nested sub-links
(those marked Attachment inside that linked file's Manage Links dialog) bleed
into the host model to some capacity -- appearing as Null Workset entries in
schedules and loading without the host model's knowledge or control.

This tool scans all Overlay-type Revit links in the active model, opens each
loaded linked document via the API, and reports any Attachment sub-links found
inside. An optional "Fix Selected" path attempts to write the fix directly to
the linked document; in practice this only succeeds if the linked file is
locally modifiable (not checked into a workshared cloud model).

Author: Jeremiah Griffith
Version: 1.0.0
"""

from __future__ import division
from pyrevit import revit, DB, forms, script
import clr
import os

clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights, FontStyles

doc = revit.doc


# ---------------------------------------------------------------------------
# Shared brush constants -- frozen so they are safe to share across WPF elements
# ---------------------------------------------------------------------------
def _brush(r, g, b):
    b_ = SolidColorBrush(Color.FromArgb(255, r, g, b))
    b_.Freeze()
    return b_

C_ORANGE = _brush(230,  81,   0)   # Attachment problem  (#E65100)
C_NORMAL = _brush( 33,  33,  33)   # Standard text       (#212121)
C_GRAY   = _brush(158, 158, 158)   # Not loaded / unknown (#9E9E9E)
C_GREEN  = _brush( 46, 125,  50)   # Clean / OK          (#2E7D32)
C_BLUE   = _brush( 21, 101, 192)   # Overlay nested (fine) (#1565C0)
C_DIM    = _brush(144, 164, 174)   # Placeholder / informational text (#90A4AE)


# ---------------------------------------------------------------------------
# Row data model
# ---------------------------------------------------------------------------
class LinkRow(object):
    """One row in the results DataGrid.

    row_kind values:
        'problem'      -- nested link is Attachment type (the bug Clyde describes)
        'clean_nested' -- nested link is Overlay type (expected/fine)
        'no_sub_links' -- parent Overlay link has zero nested links
        'not_loaded'   -- parent Overlay link is not loaded; cannot inspect
        'scan_error'   -- exception while scanning this link
    """

    def __init__(self, parent_name, parent_loaded, nested_name, ref_type,
                 row_kind, linked_doc=None, nested_type_id=None):
        self.ParentLink    = parent_name
        self.NestedLink    = nested_name
        self.RefType       = ref_type
        self.ParentLoaded  = "Yes" if parent_loaded else "No"
        self.IsProblem     = (row_kind == "problem")
        self.LinkedDoc     = linked_doc
        self.NestedTypeId  = nested_type_id

        # Italic for placeholder / informational rows
        if row_kind in ("no_sub_links", "not_loaded", "scan_error"):
            self.NestedStyle = FontStyles.Italic
        else:
            self.NestedStyle = FontStyles.Normal

        if row_kind == "problem":
            self.ParentColor   = C_NORMAL
            self.NestedColor   = C_ORANGE
            self.RefTypeColor  = C_ORANGE
            self.RefTypeWeight = FontWeights.Bold
            self.LoadedColor   = C_GREEN
            self.Action        = "Notify BIM team: change to Overlay"
            self.ActionColor   = C_ORANGE

        elif row_kind == "clean_nested":
            self.ParentColor   = C_NORMAL
            self.NestedColor   = C_DIM
            self.RefTypeColor  = C_BLUE
            self.RefTypeWeight = FontWeights.Normal
            self.LoadedColor   = C_GREEN
            self.Action        = "No action needed"
            self.ActionColor   = C_GREEN

        elif row_kind == "no_sub_links":
            self.ParentColor   = C_DIM
            self.NestedColor   = C_GRAY
            self.RefTypeColor  = C_GRAY
            self.RefTypeWeight = FontWeights.Normal
            self.LoadedColor   = C_GREEN
            self.Action        = "No action needed"
            self.ActionColor   = C_GREEN

        elif row_kind == "not_loaded":
            self.ParentColor   = C_GRAY
            self.NestedColor   = C_GRAY
            self.RefTypeColor  = C_GRAY
            self.RefTypeWeight = FontWeights.Normal
            self.LoadedColor   = C_GRAY
            self.Action        = "Load link and re-scan"
            self.ActionColor   = C_GRAY

        else:  # scan_error
            self.ParentColor   = C_GRAY
            self.NestedColor   = C_GRAY
            self.RefTypeColor  = C_GRAY
            self.RefTypeWeight = FontWeights.Normal
            self.LoadedColor   = C_GRAY
            self.Action        = "Error during scan"
            self.ActionColor   = C_GRAY


# ---------------------------------------------------------------------------
# Core scanner (pure logic, no UI dependency)
# ---------------------------------------------------------------------------
def scan_nested_links(host_doc):
    """Scan all Overlay RevitLinks for nested Attachment sub-links.

    Returns (rows, stats) where:
        rows  -- list of LinkRow objects
        stats -- dict: overlay_count, scanned_count, problem_count
    """
    rows = []
    overlay_count = 0
    scanned_count = 0
    problem_count = 0

    # All RevitLinkType elements in the host model
    all_link_types = list(
        DB.FilteredElementCollector(host_doc)
          .OfClass(DB.RevitLinkType)
    )

    # Build map: TypeId integer -> first loaded linked Document
    # Using integer values as dict keys (ElementId equality is unreliable in IronPython)
    instance_map = {}
    for inst in DB.FilteredElementCollector(host_doc).OfClass(DB.RevitLinkInstance):
        key = inst.GetTypeId().IntegerValue
        if key not in instance_map:
            ld = inst.GetLinkDocument()
            if ld is not None:
                instance_map[key] = ld

    for lt in all_link_types:
        try:
            att = lt.AttachmentType
        except Exception:
            # Not a Revit file link type (IFC, DWG, etc.) -- skip silently
            continue

        if att != DB.AttachmentType.Overlay:
            continue  # Only care about links the user has set to Overlay

        overlay_count += 1
        parent_name = lt.Name
        linked_doc  = instance_map.get(lt.Id.IntegerValue, None)

        if linked_doc is None:
            rows.append(LinkRow(
                parent_name   = parent_name,
                parent_loaded = False,
                nested_name   = "(link not loaded -- cannot inspect nested sub-links)",
                ref_type      = "--",
                row_kind      = "not_loaded"
            ))
            continue

        scanned_count += 1

        # Scan nested RevitLinkType elements inside the linked document
        try:
            nested_types = list(
                DB.FilteredElementCollector(linked_doc)
                  .OfClass(DB.RevitLinkType)
            )
        except Exception as ex:
            rows.append(LinkRow(
                parent_name   = parent_name,
                parent_loaded = True,
                nested_name   = "(scan error: {})".format(str(ex)[:80]),
                ref_type      = "--",
                row_kind      = "scan_error"
            ))
            continue

        if not nested_types:
            rows.append(LinkRow(
                parent_name   = parent_name,
                parent_loaded = True,
                nested_name   = "(no nested Revit links found in this file)",
                ref_type      = "--",
                row_kind      = "no_sub_links"
            ))
            continue

        added_any = False
        for nlt in nested_types:
            try:
                nested_att  = nlt.AttachmentType
                nested_name = nlt.Name
                is_problem  = (nested_att == DB.AttachmentType.Attachment)
                ref_str     = "Attachment" if is_problem else "Overlay"

                if is_problem:
                    problem_count += 1

                rows.append(LinkRow(
                    parent_name    = parent_name,
                    parent_loaded  = True,
                    nested_name    = nested_name,
                    ref_type       = ref_str,
                    row_kind       = "problem" if is_problem else "clean_nested",
                    linked_doc     = linked_doc if is_problem else None,
                    nested_type_id = nlt.Id    if is_problem else None
                ))
                added_any = True
            except Exception:
                continue  # Skip individual nested links that fail to read

        if not added_any:
            rows.append(LinkRow(
                parent_name   = parent_name,
                parent_loaded = True,
                nested_name   = "(could not read sub-link data)",
                ref_type      = "--",
                row_kind      = "scan_error"
            ))

    stats = {
        "overlay_count": overlay_count,
        "scanned_count": scanned_count,
        "problem_count": problem_count
    }
    return rows, stats


# ---------------------------------------------------------------------------
# Fix attempt: write AttachmentType.Overlay to nested link in linked document
# ---------------------------------------------------------------------------
def try_fix_row(row):
    """Attempt to change nested link AttachmentType to Overlay in the linked doc.

    Returns (success, message_string).
    Typical failure reason: linked document is read-only (workshared/ACC-hosted).
    """
    ld  = row.LinkedDoc
    nid = row.NestedTypeId

    if ld is None or nid is None:
        return False, "No linked document reference stored for this row."

    try:
        if not ld.IsModifiable:
            return False, (
                "Linked document is read-only (workshared and not checked out, "
                "or opened via a read-only path). Apply the fix inside that file."
            )

        nlt = ld.GetElement(nid)
        if nlt is None:
            return False, "Could not retrieve the nested link type element by ID."

        t = DB.Transaction(ld, "Fix Nested Link Attachment Type")
        t.Start()
        try:
            nlt.AttachmentType = DB.AttachmentType.Overlay
            t.Commit()
            return True, "Changed to Overlay."
        except Exception as ex:
            t.RollBack()
            return False, "Transaction failed: {}".format(str(ex))

    except Exception as ex:
        return False, "Unexpected error: {}".format(str(ex))


# ---------------------------------------------------------------------------
# WPF Window
# ---------------------------------------------------------------------------
class NestedLinkScanWindow(forms.WPFWindow):

    def __init__(self, xaml_file):
        forms.WPFWindow.__init__(self, xaml_file)
        self._rows   = ObservableCollection[object]()
        self._scanned = False
        self.ResultsGrid.ItemsSource = self._rows

    # ------------------------------------------------------------------ Scan
    def Scan_Click(self, sender, e):
        self.StatusText.Text     = "Scanning Overlay links..."
        self.ScanBtn.IsEnabled   = False
        self.ExportBtn.IsEnabled = False
        self.FixBtn.IsEnabled    = False
        self._rows.Clear()
        self._scanned = False

        try:
            rows, stats = scan_nested_links(doc)
        except Exception as ex:
            forms.alert(
                "An error occurred during the scan:\n\n{}".format(str(ex)),
                title="Nested Link Scanner"
            )
            self.StatusText.Text   = "Scan failed. See error above."
            self.ScanBtn.IsEnabled = True
            return

        for r in rows:
            self._rows.Add(r)

        oc = stats["overlay_count"]
        sc = stats["scanned_count"]
        pc = stats["problem_count"]

        self.OverlayCount.Text = str(oc)
        self.ScannedCount.Text = str(sc)
        self.ProblemCount.Text = str(pc)
        pc_brush = _brush(230, 81, 0) if pc > 0 else _brush(46, 125, 50)
        self.ProblemCount.Foreground = pc_brush

        self._scanned            = True
        self.ScanBtn.IsEnabled   = True
        self.ExportBtn.IsEnabled = True

        if oc == 0:
            self.StatusText.Text = "No Overlay-type Revit links found in this model."
        elif pc == 0:
            self.StatusText.Text = (
                "Scan complete. {} Overlay link(s) inspected -- no nested Attachment issues found."
                .format(sc)
            )
        else:
            self.StatusText.Text = (
                "Scan complete. {} nested Attachment sub-link(s) detected across {} Overlay link(s). "
                "Orange rows require BIM team action."
                .format(pc, sc)
            )

    # ------------------------------------------------ Selection changed
    def ResultsGrid_SelectionChanged(self, sender, e):
        if not self._scanned:
            return
        has_problem = any(r.IsProblem for r in self.ResultsGrid.SelectedItems)
        self.FixBtn.IsEnabled = has_problem

    # ------------------------------------------------ Attempt fix
    def Fix_Click(self, sender, e):
        problem_rows = [r for r in self.ResultsGrid.SelectedItems if r.IsProblem]
        if not problem_rows:
            return

        fixed_names = []
        fail_msgs   = []

        for row in problem_rows:
            ok, msg = try_fix_row(row)
            if ok:
                fixed_names.append(row.NestedLink)
            else:
                fail_msgs.append("  - {}: {}".format(row.NestedLink, msg))

        parts = []
        if fixed_names:
            parts.append("Successfully changed to Overlay ({} item(s)):".format(len(fixed_names)))
            for n in fixed_names:
                parts.append("  + {}".format(n))
            parts.append("")

        if fail_msgs:
            parts.append("Could not change ({} item(s)):".format(len(fail_msgs)))
            parts.extend(fail_msgs)
            parts.append("")
            parts.append(
                "For read-only linked files, the fix must be applied by the team\n"
                "who owns that file. Ask them to open their model, go to Manage\n"
                "Links, and change the listed sub-link(s) from Attachment to Overlay."
            )

        forms.alert("\n".join(parts), title="Fix Results")

        if fixed_names:
            # Refresh the grid after successful changes
            self.Scan_Click(sender, e)

    # ------------------------------------------------ Export report
    def Export_Click(self, sender, e):
        out = script.get_output()
        out.print_md("# Nested Link Scan Report")
        out.print_md("**Host Model:** `{}`".format(doc.Title))
        out.print_md("")
        out.print_md("| Metric | Value |")
        out.print_md("|---|---|")
        out.print_md("| Overlay links found | {} |".format(self.OverlayCount.Text))
        out.print_md("| Loaded and scanned  | {} |".format(self.ScannedCount.Text))
        out.print_md("| Nested Attachments  | {} |".format(self.ProblemCount.Text))
        out.print_md("")

        problem_rows = [r for r in self._rows if r.IsProblem]
        if problem_rows:
            out.print_md("## Attachment Sub-Links Detected")
            out.print_md(
                "| Parent Link (Overlay in Host) | Nested Sub-Link | Ref Type | Action |"
            )
            out.print_md("|---|---|:---:|---|")
            for r in problem_rows:
                out.print_md("| {} | {} | {} | {} |".format(
                    r.ParentLink, r.NestedLink, r.RefType, r.Action
                ))
            out.print_md("")
            out.print_md(
                "> **How to fix:** The BIM team who owns each flagged linked file must open "
                "their model, go to **Manage > Manage Links**, and change each listed "
                "sub-link from **Attachment** to **Overlay**. The Revit API cannot write "
                "this change to a workshared linked document from the host."
            )
        else:
            out.print_md("No nested Attachment issues detected. Model is clean.")

        out.print_md("")
        out.print_md("---")
        out.print_md(
            "*Generated by DetailersTools Nested Link Scanner "
            "(Jeremiah Griffith / Southland Industries NW Division)*"
        )

        self.StatusText.Text = "Report exported to pyRevit output window."

    # ------------------------------------------------ Close
    def Close_Click(self, sender, e):
        self.Close()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
xaml_path = os.path.join(os.path.dirname(__file__), "window.xaml")
win = NestedLinkScanWindow(xaml_path)
win.ShowDialog()
