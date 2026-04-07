# -*- coding: ascii -*-
# Family Parameter Browser & Editor (RFA, Revit 2025)
# IronPython 2.7 for pyRevit | Windows Forms UI
# Features:
# - Browse all FamilyParameters for the current Family Type
# - Read values and formulas; edit values and formulas
# - Grid lines; auto-resize columns
# - Click any column header to sort (ascending/descending toggle)
#   Columns: Name | Kind | UserMod | Value | Formula

import sys
import traceback
import clr

# Revit API
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

# .NET
clr.AddReference("System")
clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")

from System import Double, Array, String
from System.Collections import IComparer
from System.Drawing import Size, Point
from System.Windows.Forms import (
    Application, Form, Label, ComboBox, ListView, ColumnHeader,
    Button, TextBox, ListViewItem, View, AnchorStyles,
    MessageBox, MessageBoxButtons, MessageBoxIcon,
    FormStartPosition, ComboBoxStyle, DialogResult,
    ColumnHeaderAutoResizeStyle
)

# ---------- Revit handles (pyRevit provides __revit__) ----------
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

def is_family_doc(d):
    try:
        return d.IsFamilyDocument
    except:
        return False

if not is_family_doc(doc):
    MessageBox.Show("This tool must be run inside an open Family (RFA) document.", "Family Editor Only", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    sys.exit(0)

fm = doc.FamilyManager

# ------------------------ helpers ------------------------

def get_family_types():
    try:
        return [t for t in fm.Types]
    except:
        return []

def get_param_kind(fp):
    try:
        return "Instance" if fp.IsInstance else "Type"
    except:
        return "Unknown"

def get_param_usermod(fp):
    try:
        return "False" if fp.UserModifiable is False else "True"
    except:
        return "Unknown"

def try_get_formula(fp):
    # Primary: FamilyManager.GetFormula
    try:
        s = fm.GetFormula(fp)
        if s:
            return s
    except:
        pass
    # Fallback: FamilyParameter.Formula (when available)
    try:
        s2 = getattr(fp, "Formula", None)
        if s2:
            return s2
    except:
        pass
    # If parameter is formula-driven but text not exposed, show marker
    try:
        if getattr(fp, "IsDeterminedByFormula", False):
            return "<<formula>>"
    except:
        pass
    return ""

def storage_type_of(fp):
    try:
        return fp.StorageType
    except:
        return None

def format_double_value(spec_id, val):
    try:
        return UnitFormatUtils.Format(doc, spec_id, val, False, False)
    except:
        try:
            return str(val)
        except:
            return ""

def get_value_string_any(ft, fp):
    # 1) Revit formatted string
    try:
        vs = ft.AsValueString(fp)
        if vs:
            return vs
    except:
        pass

    # 2) Fallbacks by storage type
    st = storage_type_of(fp)
    spec_id = None
    try:
        spec_id = fp.Definition.GetDataType()
    except:
        spec_id = None

    try:
        if st == StorageType.Double:
            try:
                dval = ft.AsDouble(fp)
                return format_double_value(spec_id, dval)
            except:
                return ""
        elif st == StorageType.Integer:
            try:
                ival = ft.AsInteger(fp)
                # Show Yes/No nicely when it is a boolean spec id
                try:
                    tid = fp.Definition.GetDataType()
                    tid_str = tid.TypeId if tid else ""
                except:
                    tid_str = ""
                if "spec:bool" in tid_str or "spec:yesno" in tid_str:
                    return "True" if ival == 1 else "False"
                return str(ival)
            except:
                return ""
        elif st == StorageType.String:
            try:
                sval = ft.AsString(fp)
                return sval if sval is not None else ""
            except:
                return ""
        elif st == StorageType.ElementId:
            try:
                eid = ft.AsElementId(fp)
                if eid and eid.IntegerValue > 0:
                    try:
                        el = doc.GetElement(eid)
                        if el is not None and hasattr(el, "Name") and el.Name:
                            return el.Name
                    except:
                        pass
                return ""
            except:
                return ""
        else:
            return ""
    except:
        return ""

def parse_and_set_value(tx, fp, input_text):
    ft = fm.CurrentType
    if ft is None:
        raise Exception("No active Family Type is selected in Family Manager.")

    st = storage_type_of(fp)
    if st is None:
        raise Exception("Unknown storage type.")

    spec_id = None
    try:
        spec_id = fp.Definition.GetDataType()
    except:
        spec_id = None

    if not tx.HasStarted():
        tx.Start("Set Family Parameter Value")

    try:
        if st == StorageType.Double:
            ok = False
            parsed_val = None
            try:
                byref_val = clr.Reference[Double](0.0)
                ok = UnitFormatUtils.TryParse(doc, spec_id, input_text, byref_val)
                if ok:
                    parsed_val = byref_val.Value
            except Exception:
                ok = False
                parsed_val = None

            if not ok:
                try:
                    parsed_val = float(input_text)
                    ok = True
                except:
                    ok = False

            if not ok:
                raise Exception("Could not parse numeric value. Try including units (e.g., 5 ft, 1200 mm).")

            fm.Set(fp, parsed_val)

        elif st == StorageType.Integer:
            try:
                ival = int(input_text)
            except:
                s = (input_text or "").strip().lower()
                if s in ["true", "yes", "1"]:
                    ival = 1
                elif s in ["false", "no", "0"]:
                    ival = 0
                else:
                    raise Exception("Could not parse integer or boolean. Use 0/1 or True/False.")
            fm.Set(fp, ival)

        elif st == StorageType.String:
            fm.Set(fp, input_text if input_text is not None else "")

        elif st == StorageType.ElementId:
            raise Exception("Setting ElementId parameters from free text is not supported in this tool.")

        else:
            raise Exception("Unsupported storage type.")
    except:
        tx.RollBack()
        raise
    else:
        tx.Commit()

def set_formula(tx, fp, formula_text):
    if not tx.HasStarted():
        tx.Start("Set Family Parameter Formula")
    try:
        if formula_text is None:
            formula_text = ""
        if len(formula_text.strip()) == 0:
            fm.SetFormula(fp, None)
        else:
            fm.SetFormula(fp, formula_text)
    except:
        tx.RollBack()
        raise
    else:
        tx.Commit()

# ------------------------ sorting support ------------------------

def _try_parse_float_simple(s):
    if not s:
        return None
    t = s.strip().replace(",", "")
    # strip enclosing quotes often seen in value strings for plain text
    if len(t) >= 2 and ((t[0] == '"' and t[-1] == '"') or (t[0] == "'" and t[-1] == "'")):
        t = t[1:-1]
    # quick reject if it contains letters (units), we keep as string
    for ch in t:
        if ch.isalpha():
            return None
    try:
        return float(t)
    except:
        return None

class ListViewColumnSorter(IComparer):
    def __init__(self):
        self.Column = 0
        self.Ascending = True

    def Compare(self, x, y):
        try:
            a = x.SubItems[self.Column].Text or ""
            b = y.SubItems[self.Column].Text or ""
        except:
            a = ""
            b = ""
        # For Value column, try numeric compare when possible
        if self.Column == 3:
            fa = _try_parse_float_simple(a)
            fb = _try_parse_float_simple(b)
            if fa is not None and fb is not None:
                cmpnum = (fa > fb) - (fa < fb)
                return cmpnum if self.Ascending else -cmpnum
        # Case-insensitive string compare
        al = a.lower()
        bl = b.lower()
        if al == bl:
            return 0
        res = -1 if al < bl else 1
        return res if self.Ascending else -res

# ------------------------ UI ------------------------

class ParamEditorForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Family Parameter Browser and Editor (RFA)"
        self.MinimumSize = Size(1000, 640)
        self.StartPosition = FormStartPosition.CenterScreen

        self.lblType = Label()
        self.lblType.Text = "Family Type:"
        self.lblType.Location = Point(12, 12)
        self.lblType.AutoSize = True

        self.cboType = ComboBox()
        self.cboType.DropDownStyle = ComboBoxStyle.DropDownList
        self.cboType.Location = Point(100, 8)
        self.cboType.Width = 360
        self.cboType.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right

        self.btnRefresh = Button()
        self.btnRefresh.Text = "Refresh"
        self.btnRefresh.Location = Point(470, 7)
        self.btnRefresh.Width = 90
        self.btnRefresh.Anchor = AnchorStyles.Top | AnchorStyles.Right

        self.lv = ListView()
        self.lv.View = View.Details
        self.lv.FullRowSelect = True
        self.lv.HideSelection = False
        self.lv.MultiSelect = False
        self.lv.GridLines = True
        self.lv.Location = Point(12, 40)
        self.lv.Size = Size(960, 420)
        self.lv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        # Columns: Name | Kind | UserMod | Value | Formula
        ch1 = ColumnHeader(); ch1.Text = "Name"; ch1.Width = 260
        ch2 = ColumnHeader(); ch2.Text = "Kind"; ch2.Width = 80
        ch3 = ColumnHeader(); ch3.Text = "UserMod"; ch3.Width = 70
        ch4 = ColumnHeader(); ch4.Text = "Value"; ch4.Width = 260
        ch5 = ColumnHeader(); ch5.Text = "Formula"; ch5.Width = 270
        self.lv.Columns.AddRange((ch1, ch2, ch3, ch4, ch5))

        # Sorting
        self.sorter = ListViewColumnSorter()
        self.lv.ListViewItemSorter = self.sorter
        self.lv.ColumnClick += self.on_column_click

        self.lblSel = Label(); self.lblSel.Text = "Selected Parameter:"; self.lblSel.Location = Point(12, 468); self.lblSel.AutoSize = True
        self.txtSel = TextBox(); self.txtSel.Location = Point(140, 465); self.txtSel.Width = 300; self.txtSel.ReadOnly = True
        self.txtSel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        self.lblVal = Label(); self.lblVal.Text = "Value:"; self.lblVal.Location = Point(12, 498); self.lblVal.AutoSize = True
        self.txtVal = TextBox(); self.txtVal.Location = Point(140, 495); self.txtVal.Width = 460
        self.txtVal.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        self.btnApplyVal = Button(); self.btnApplyVal.Text = "Apply Value"; self.btnApplyVal.Location = Point(610, 493); self.btnApplyVal.Width = 120
        self.btnApplyVal.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.lblForm = Label(); self.lblForm.Text = "Formula:"; self.lblForm.Location = Point(12, 528); self.lblForm.AutoSize = True
        self.txtForm = TextBox(); self.txtForm.Location = Point(140, 525); self.txtForm.Width = 460
        self.txtForm.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right

        self.btnApplyForm = Button(); self.btnApplyForm.Text = "Apply Formula"; self.btnApplyForm.Location = Point(610, 523); self.btnApplyForm.Width = 120
        self.btnApplyForm.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.btnClose = Button(); self.btnClose.Text = "Close"; self.btnClose.Location = Point(852, 560); self.btnClose.Width = 120
        self.btnClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right

        self.Controls.Add(self.lblType)
        self.Controls.Add(self.cboType)
        self.Controls.Add(self.btnRefresh)
        self.Controls.Add(self.lv)
        self.Controls.Add(self.lblSel)
        self.Controls.Add(self.txtSel)
        self.Controls.Add(self.lblVal)
        self.Controls.Add(self.txtVal)
        self.Controls.Add(self.btnApplyVal)
        self.Controls.Add(self.lblForm)
        self.Controls.Add(self.txtForm)
        self.Controls.Add(self.btnApplyForm)
        self.Controls.Add(self.btnClose)

        self.family_types = get_family_types()

        self.btnRefresh.Click += self.on_refresh
        self.cboType.SelectedIndexChanged += self.on_type_changed
        self.lv.SelectedIndexChanged += self.on_row_selected
        self.btnApplyVal.Click += self.on_apply_value
        self.btnApplyForm.Click += self.on_apply_formula
        self.btnClose.Click += self.on_close

        self.populate_types()
        self.refresh_list()

    # ---------- sorting handler ----------
    def on_column_click(self, sender, args):
        col = int(args.Column)
        if self.sorter.Column == col:
            self.sorter.Ascending = not self.sorter.Ascending
        else:
            self.sorter.Column = col
            self.sorter.Ascending = True
        try:
            self.lv.Sort()
        except:
            pass

    def populate_types(self):
        self.cboType.Items.Clear()
        current_index = -1
        cur = fm.CurrentType
        for i, t in enumerate(self.family_types):
            name = t.Name if t.Name is not None else "(Unnamed Type)"
            self.cboType.Items.Add(name)
            try:
                if cur is not None and (t == cur or (t.Name == cur.Name)):
                    current_index = i
            except:
                pass

        if current_index >= 0:
            self.cboType.SelectedIndex = current_index
        elif self.cboType.Items.Count > 0:
            self.cboType.SelectedIndex = 0

    def ensure_current_type(self):
        idx = self.cboType.SelectedIndex
        if idx < 0 or idx >= len(self.family_types):
            return
        sel_type = self.family_types[idx]
        try:
            cur = fm.CurrentType
            if cur is None or not (sel_type == cur or sel_type.Name == cur.Name):
                t = Transaction(doc, "Switch Family Type")
                t.Start()
                fm.CurrentType = sel_type
                t.Commit()
        except Exception as ex:
            MessageBox.Show("Could not switch Family Type:\n" + str(ex), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    def refresh_list(self, keep_selection=False):
        self.ensure_current_type()

        # Remember selection by parameter pointer
        prev_fp = None
        if keep_selection and self.lv.SelectedItems.Count > 0:
            try:
                prev_fp = self.lv.SelectedItems[0].Tag
            except:
                prev_fp = None

        self.lv.BeginUpdate()
        try:
            self.lv.Items.Clear()

            ft = fm.CurrentType
            fps = [fp for fp in fm.Parameters]

            for fp in fps:
                name = fp.Definition.Name if fp.Definition and fp.Definition.Name else "(Unnamed)"
                kind = get_param_kind(fp)
                umod = get_param_usermod(fp)
                valstr = get_value_string_any(ft, fp)
                formula = try_get_formula(fp)

                row = [name, kind, umod, valstr, formula]
                arr = Array[String](row)
                item = ListViewItem(arr)
                item.Tag = fp  # keep the parameter with the row even when sorted
                self.lv.Items.Add(item)
        finally:
            self.lv.EndUpdate()

        # Apply current sorter (maintain sort when refreshing)
        try:
            self.lv.Sort()
        except:
            pass

        # Re-select previous item if possible
        if prev_fp is not None:
            try:
                for it in self.lv.Items:
                    if it.Tag == prev_fp:
                        it.Selected = True
                        break
            except:
                pass

        # Auto-size columns to content
        try:
            self.lv.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
        except:
            pass

        if self.lv.SelectedIndices.Count == 0:
            self.txtSel.Text = ""
            self.txtVal.Text = ""
            self.txtForm.Text = ""

    def on_refresh(self, sender, args):
        self.refresh_list(keep_selection=True)

    def on_type_changed(self, sender, args):
        self.refresh_list(keep_selection=False)

    def on_row_selected(self, sender, args):
        if self.lv.SelectedItems.Count == 0:
            self.txtSel.Text = ""
            self.txtVal.Text = ""
            self.txtForm.Text = ""
            return
        it = self.lv.SelectedItems[0]
        fp = it.Tag
        if fp is None:
            return
        name = fp.Definition.Name if fp.Definition and fp.Definition.Name else "(Unnamed)"
        self.txtSel.Text = name

        try:
            ft = fm.CurrentType
            self.txtVal.Text = get_value_string_any(ft, fp)
        except:
            self.txtVal.Text = ""

        try:
            self.txtForm.Text = try_get_formula(fp)
        except:
            self.txtForm.Text = ""

    def on_apply_value(self, sender, args):
        if self.lv.SelectedItems.Count == 0:
            MessageBox.Show("Select a parameter first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        fp = self.lv.SelectedItems[0].Tag
        if fp is None:
            return
        text = self.txtVal.Text if self.txtVal.Text is not None else ""

        try:
            has_formula_flag = getattr(fp, "IsDeterminedByFormula", False)
        except:
            has_formula_flag = False

        current_formula = try_get_formula(fp)
        if has_formula_flag or (current_formula and len(current_formula.strip()) > 0):
            res = MessageBox.Show("This parameter is formula-driven.\nYou must clear the formula before setting a direct value.\n\nClear formula now?", "Formula Present", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            if res == DialogResult.Yes:
                try:
                    tx = Transaction(doc, "Clear Formula")
                    set_formula(tx, fp, "")
                except Exception as ex:
                    MessageBox.Show("Failed to clear formula:\n" + str(ex), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    return
            else:
                return

        try:
            tx = Transaction(doc, "Set Family Parameter Value")
            parse_and_set_value(tx, fp, text)
        except Exception as ex:
            MessageBox.Show("Failed to set value:\n" + str(ex), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        self.refresh_list(keep_selection=True)

    def on_apply_formula(self, sender, args):
        if self.lv.SelectedItems.Count == 0:
            MessageBox.Show("Select a parameter first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        fp = self.lv.SelectedItems[0].Tag
        if fp is None:
            return
        text = self.txtForm.Text if self.txtForm.Text is not None else ""

        try:
            tx = Transaction(doc, "Set Family Parameter Formula")
            set_formula(tx, fp, text)
        except Exception as ex:
            MessageBox.Show("Failed to set formula:\n" + str(ex), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            return

        self.refresh_list(keep_selection=True)

    def on_close(self, sender, args):
        self.Close()

# ------------------------ run (pyRevit: use ShowDialog) ------------------------

try:
    form = ParamEditorForm()
    form.ShowDialog()
except Exception as e:
    MessageBox.Show("Unexpected error:\n" + traceback.format_exc(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
