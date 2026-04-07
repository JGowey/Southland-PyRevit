# -*- coding: utf-8 -*-
"""
Toggle Case - WPF UI with profile support and element selection.
Transforms selected TextNote elements between UPPERCASE, Title Case,
lowercase, or Auto-Cycle (Shift+F3 behavior).
"""
import os
import json

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType

doc   = revit.doc
uidoc = revit.uidoc

SCRIPT_DIR   = os.path.dirname(__file__)
XAML_PATH    = os.path.join(SCRIPT_DIR, "window.xaml")
PROFILES_DIR = os.path.join(SCRIPT_DIR, "CaseProfiles")

if not os.path.exists(PROFILES_DIR):
    os.makedirs(PROFILES_DIR)


# ── selection filter ─────────────────────────────────────────────────

class TextNoteFilter(ISelectionFilter):
    def AllowElement(self, elem):
        return isinstance(elem, DB.TextNote)
    def AllowReference(self, ref, point):
        return False


# ── case helpers ─────────────────────────────────────────────────────

def detect_case(text):
    if not any(c.isalpha() for c in text):
        return "upper"
    if text == text.upper():
        return "upper"
    if text == text.lower():
        return "lower"
    if text == text.title():
        return "title"
    return "mixed"


def next_case(current):
    cycle = ["upper", "title", "lower"]
    try:
        return cycle[(cycle.index(current) + 1) % len(cycle)]
    except ValueError:
        return "upper"


def apply_case(text, mode):
    if mode == "upper":  return text.upper()
    if mode == "lower":  return text.lower()
    if mode == "title":  return text.title()
    return text


def preview_case(text, mode):
    if mode == "cycle":
        dominant = detect_case(text) if any(c.isalpha() for c in text) else "mixed"
        target = next_case(dominant)
        label = {"upper": "UPPERCASE", "title": "Title Case", "lower": "lowercase"}
        return "{} (next: {})".format(apply_case(text, target), label.get(target, target))
    return apply_case(text, mode)


# ── profile helpers ──────────────────────────────────────────────────

def list_profiles():
    return sorted([f[:-5] for f in os.listdir(PROFILES_DIR) if f.endswith(".json")])

def load_profile(name):
    path = os.path.join(PROFILES_DIR, name + ".json")
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return None

def save_profile(name, data):
    path = os.path.join(PROFILES_DIR, name + ".json")
    with open(path, "w") as f:
        json.dump(data, f, indent=2)

def delete_profile(name):
    path = os.path.join(PROFILES_DIR, name + ".json")
    if os.path.exists(path):
        os.remove(path)


# ── window ───────────────────────────────────────────────────────────

class ToggleCaseWindow(forms.WPFWindow):

    def __init__(self, xaml_path, preselected=None):
        forms.WPFWindow.__init__(self, xaml_path)
        self._text_notes = list(preselected) if preselected else []
        self._pending_pick = False
        self._refresh_profiles()
        self._update_count()
        self._update_preview()

    # ── internal helpers ─────────────────────────────────────────────

    def _refresh_profiles(self):
        current = self.ProfileCombo.SelectedItem
        self.ProfileCombo.Items.Clear()
        for name in list_profiles():
            self.ProfileCombo.Items.Add(name)
        if current and current in list_profiles():
            self.ProfileCombo.SelectedItem = current

    def _update_count(self):
        count = len(self._text_notes)
        if count == 0:
            self.CountPill.Visibility = forms.framework.System.Windows.Visibility.Collapsed
        else:
            self.CountPill.Visibility = forms.framework.System.Windows.Visibility.Visible
            self.CountLabel.Text = "{} element{}".format(count, "s" if count != 1 else "")

    def _get_mode(self):
        if self.ModeUpper.IsChecked: return "upper"
        if self.ModeTitle.IsChecked: return "title"
        if self.ModeLower.IsChecked: return "lower"
        return "cycle"

    def _set_mode(self, mode):
        self.ModeUpper.IsChecked = (mode == "upper")
        self.ModeTitle.IsChecked = (mode == "title")
        self.ModeLower.IsChecked = (mode == "lower")
        self.ModeCycle.IsChecked = (mode == "cycle")

    def _update_preview(self):
        text = self.PreviewInput.Text or ""
        self.PreviewOutput.Text = preview_case(text, self._get_mode()) if text else ""

    # ── event handlers ───────────────────────────────────────────────

    def OnModeChanged(self, sender, e):
        self._update_preview()

    def OnPreviewChanged(self, sender, e):
        self._update_preview()

    def OnSelectElements(self, sender, e):
        if self._text_notes:
            if not forms.alert(
                "Replace the current selection?",
                title="Select Elements",
                yes=True, no=True, ok=False
            ):
                return
        self._pending_pick = True
        self.Close()

    def OnProfileSelected(self, sender, e):
        name = self.ProfileCombo.SelectedItem
        if not name:
            return
        data = load_profile(name)
        if not data:
            return
        self._set_mode(data.get("mode", "cycle"))
        preview = data.get("preview_text", "")
        if preview:
            self.PreviewInput.Text = preview
        self._update_preview()

    def OnSaveProfile(self, sender, e):
        name = forms.ask_for_string(
            default=self.ProfileCombo.SelectedItem or "",
            prompt="Profile name:",
            title="Save Profile"
        )
        if not name:
            return
        name = name.strip()
        save_profile(name, {
            "mode": self._get_mode(),
            "preview_text": self.PreviewInput.Text or ""
        })
        self._refresh_profiles()
        self.ProfileCombo.SelectedItem = name
        forms.toast("Saved: " + name, title="Toggle Case")

    def OnDeleteProfile(self, sender, e):
        name = self.ProfileCombo.SelectedItem
        if not name:
            return
        if not forms.alert(
            'Delete profile "{}"?'.format(name),
            title="Toggle Case",
            yes=True, no=True, ok=False
        ):
            return
        delete_profile(name)
        self._refresh_profiles()
        forms.toast("Deleted: " + name, title="Toggle Case")

    def OnApply(self, sender, e):
        mode = self._get_mode()

        # If no explicitly selected notes, fall back to current Revit selection
        text_notes = list(self._text_notes)
        if not text_notes:
            for eid in uidoc.Selection.GetElementIds():
                elem = doc.GetElement(eid)
                if isinstance(elem, DB.TextNote):
                    text_notes.append(elem)

        if not text_notes:
            forms.alert(
                "No TextNote elements selected.\n"
                "Use 'Select Elements in Revit' or pre-select text notes before opening.",
                title="Toggle Case"
            )
            return

        # For cycle mode, detect dominant state across all notes
        if mode == "cycle":
            counts = {"upper": 0, "title": 0, "lower": 0, "mixed": 0}
            for tn in text_notes:
                counts[detect_case(tn.Text)] += 1
            target = next_case(max(counts, key=lambda k: counts[k]))
        else:
            target = mode

        with DB.Transaction(doc, "Toggle Case") as t:
            t.Start()
            for tn in text_notes:
                new_text = apply_case(tn.Text, target)
                if new_text != tn.Text:
                    tn.Text = new_text
            t.Commit()

        label = {"upper": "UPPERCASE", "title": "Title Case", "lower": "lowercase"}
        forms.toast(
            "{} note{} -> {}".format(
                len(text_notes),
                "s" if len(text_notes) != 1 else "",
                label.get(target, target)
            ),
            title="Toggle Case"
        )
        self.Close()

    def OnCancel(self, sender, e):
        self.Close()


# ── entry point with pick loop ───────────────────────────────────────

if not os.path.exists(XAML_PATH):
    forms.alert("window.xaml not found:\n" + XAML_PATH, title="Toggle Case")
else:
    # Collect any TextNotes already selected in Revit as preselection
    preselected = []
    for eid in uidoc.Selection.GetElementIds():
        elem = doc.GetElement(eid)
        if isinstance(elem, DB.TextNote):
            preselected.append(elem)

    while True:
        win = ToggleCaseWindow(XAML_PATH, preselected=preselected)
        win.ShowDialog()

        if win._pending_pick:
            # Window closed to allow picking -- go to Revit
            try:
                refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    TextNoteFilter(),
                    "Select TextNote elements, then press Finish"
                )
                preselected = [doc.GetElement(r.ElementId) for r in refs]
            except Exception:
                # User cancelled the pick -- reopen with previous selection
                pass
        else:
            # Normal close (Apply or Cancel)
            break
