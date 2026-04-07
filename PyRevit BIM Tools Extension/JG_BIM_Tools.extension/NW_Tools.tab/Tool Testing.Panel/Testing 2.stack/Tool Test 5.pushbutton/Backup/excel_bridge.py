# -*- coding: utf-8 -*-
"""
RFA Type Manager - Excel Bridge (excel_bridge.py)
Pure IronPython 2.7. Zero external dependencies.
Writes/reads .xlsx via built-in zipfile + Open XML strings.
"""

import zipfile
import re
import json

try:
    unicode
except NameError:
    unicode = str
try:
    unichr
except NameError:
    unichr = chr

# =============================================================================
# UNIT CONVERSION
# =============================================================================

UNIT_SCALES = {"inches": 12.0, "feet": 1.0}

def _unit_scale(unit_mode):
    return UNIT_SCALES.get(unit_mode, 12.0)

# =============================================================================
# XML / ENCODING HELPERS
# =============================================================================

def _xe(s):
    if s is None:
        return u""
    s = unicode(s)
    s = s.replace(u"&", u"&amp;")
    s = s.replace(u"<", u"&lt;")
    s = s.replace(u">", u"&gt;")
    s = s.replace(u'"', u"&quot;")
    return s

def _enc(s):
    if isinstance(s, bytes):
        return s
    return s.encode("utf-8")

def _col_letter(n):
    r = u""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        r = unichr(65 + rem) + r
    return r

def _ref(row, col):
    return u"{}{}".format(_col_letter(col), row)

# =============================================================================
# SHARED STRINGS
# =============================================================================

class SharedStrings(object):
    def __init__(self):
        self._list  = []
        self._index = {}

    def idx(self, s):
        s = u"" if s is None else unicode(s)
        if s not in self._index:
            self._index[s] = len(self._list)
            self._list.append(s)
        return self._index[s]

    def xml(self):
        n = len(self._list)
        parts = [
            u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            u'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            u' count="{0}" uniqueCount="{0}">'.format(n),
        ]
        for s in self._list:
            parts.append(u'<si><t xml:space="preserve">{}</t></si>'.format(_xe(s)))
        parts.append(u'</sst>')
        return u"\n".join(parts)

# =============================================================================
# STYLES
# Index: 0=default 1=header 2=hdr-name 3=hdr-formula 4=number 5=integer
#        6=formula-cell 7=alt-row 8=meta-key
# =============================================================================

S_DEFAULT = 0
S_HEADER  = 1
S_HDR_NM  = 2
S_HDR_FMT = 3
S_NUMBER  = 4
S_INTEGER = 5
S_FORMULA = 6
S_ALT     = 7
S_METAKEY = 8

def _styles_xml():
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.000"/>
  </numFmts>
  <fonts count="6">
    <font><sz val="9"/><name val="Arial"/></font>
    <font><sz val="9"/><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font>
    <font><sz val="9"/><b/><color rgb="FF333333"/><name val="Arial"/></font>
    <font><sz val="9"/><i/><color rgb="FF505090"/><name val="Arial"/></font>
    <font><sz val="8"/><color rgb="FF888888"/><name val="Arial"/></font>
    <font><sz val="8"/><b/><color rgb="FF333333"/><name val="Arial"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2F3D4E"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF3A5068"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEBEBEB"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF0F4FA"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFCCCCCC"/></left>
      <right style="thin"><color rgb="FFCCCCCC"/></right>
      <top style="thin"><color rgb="FFCCCCCC"/></top>
      <bottom style="thin"><color rgb="FFCCCCCC"/></bottom>
    </border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="9">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="1" xfId="0"><alignment vertical="center"/></xf>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="1" xfId="0" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="1" fillId="3" borderId="1" xfId="0" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="1"   fontId="0" fillId="0" borderId="1" xfId="0"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0"   fontId="5" fillId="4" borderId="1" xfId="0" applyFill="1" applyFont="1"><alignment vertical="center"/></xf>
  </cellXfs>
</styleSheet>"""

# =============================================================================
# SHEET BUILDER
# Correct Open XML element order (ECMA-376):
#   sheetPr? > dimension? > sheetViews? > sheetFormatPr? > cols? >
#   sheetData > sheetCalcPr? > ... > autoFilter? > ...
# sheetViews MUST come before sheetData.
# =============================================================================

class SheetWriter(object):
    def __init__(self, ss):
        self._ss   = ss
        self._rows = []

    def _sc(self, ref, val, style):
        return u'<c r="{}" s="{}" t="s"><v>{}</v></c>'.format(ref, style, self._ss.idx(val))

    def _nc(self, ref, val, style):
        return u'<c r="{}" s="{}"><v>{}</v></c>'.format(ref, style, val)

    def add_row(self, rn, cells, height=None):
        ht = u' ht="{}" customHeight="1"'.format(height) if height else u""
        parts = [u'<row r="{}"{}>'.format(rn, ht)]
        for cn, val, style in cells:
            r = _ref(rn, cn)
            if val is None:
                parts.append(u'<c r="{}" s="{}"/>'.format(r, style))
            elif style in (S_NUMBER, S_INTEGER) and isinstance(val, (int, float)):
                parts.append(self._nc(r, val, style))
            else:
                parts.append(self._sc(r, unicode(val) if val is not None else u"", style))
        parts.append(u"</row>")
        self._rows.append(u"".join(parts))

    def xml(self, col_widths=None, freeze_col=0, freeze_row=0,
            num_data_cols=1, num_data_rows=0):
        """
        Builds worksheet XML with correct element order per ECMA-376 spec:
        sheetViews -> cols -> sheetData -> autoFilter
        """
        parts = [
            u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
            u' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        ]

        # sheetViews FIRST (before cols and sheetData)
        if freeze_col > 0 or freeze_row > 0:
            tl = _ref(freeze_row + 1, freeze_col + 1)
            parts.append(u'<sheetViews>')
            parts.append(u'<sheetView workbookViewId="0">')
            parts.append(
                u'<pane xSplit="{}" ySplit="{}" topLeftCell="{}"'
                u' activePane="bottomRight" state="frozen"/>'.format(
                    freeze_col, freeze_row, tl)
            )
            parts.append(u'</sheetView>')
            parts.append(u'</sheetViews>')

        # cols
        if col_widths:
            parts.append(u"<cols>")
            for cn in sorted(col_widths):
                parts.append(
                    u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(
                        cn, col_widths[cn])
                )
            parts.append(u"</cols>")

        # sheetData
        parts.append(u"<sheetData>")
        parts.extend(self._rows)
        parts.append(u"</sheetData>")

        # autoFilter (after sheetData, covers header row + data)
        if num_data_rows > 0 and num_data_cols > 0:
            end_ref = _ref(num_data_rows + 1, num_data_cols)
            parts.append(u'<autoFilter ref="A1:{}"/>'.format(end_ref))

        parts.append(u"</worksheet>")
        return u"\n".join(parts)

# =============================================================================
# STATIC XML PARTS
# =============================================================================

_CONTENT_TYPES = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    u'<Default Extension="xml" ContentType="application/xml"/>'
    u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    u'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    u'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    u'</Types>'
)

_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    u'</Relationships>'
)

_WORKBOOK = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    u' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    u'<sheets>'
    u'<sheet name="Types" sheetId="1" r:id="rId1"/>'
    u'<sheet name="Meta" sheetId="2" r:id="rId2"/>'
    u'</sheets>'
    u'</workbook>'
)

_WB_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    u'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
    u'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    u'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    u'</Relationships>'
)

# =============================================================================
# EXPORT
# =============================================================================

def export_to_xlsx(snapshot, xlsx_path, unit_mode="inches", selected_params=None):
    """
    Writes a TypeSnapshot to a formatted .xlsx file.
    selected_params=None exports ALL writable params.
    Formula params are always included as visible gray columns (expression shown).
    """
    import datetime as _dt

    params_all  = snapshot.get("parameters", [])
    types_all   = snapshot.get("types",      [])
    family_name = snapshot.get("family_name", u"Unknown")
    unit_scale  = _unit_scale(unit_mode)
    unit_label  = u"in" if unit_mode == "inches" else u"ft"

    formula_params  = [p for p in params_all if p.get("has_formula") or p.get("read_only")]
    writable_params = [p for p in params_all if not p.get("has_formula") and not p.get("read_only")]

    if selected_params is not None:
        sel = set(selected_params)
        writable_params = [p for p in writable_params if p["name"] in sel]

    # All columns: writable first, then formula (gray, read-only in intent)
    display_params = writable_params + formula_params
    num_cols = 1 + len(display_params)  # col A = Type Name

    ss = SharedStrings()
    tw = SheetWriter(ss)
    mw = SheetWriter(ss)

    # Column widths
    cw = {1: 28}
    for ci, p in enumerate(display_params, 2):
        cw[ci] = max(14, min(36, len(p["name"]) + 6))

    # Header row (row 1)
    hdr = [(1, u"Type Name", S_HDR_NM)]
    for ci, p in enumerate(display_params, 2):
        st   = p.get("storage_type", "String")
        is_f = p.get("has_formula") or p.get("read_only")
        if st == "Double" and not is_f:
            lbl = p["name"] + u" (" + unit_label + u")"
        else:
            lbl = p["name"]
        hdr.append((ci, lbl, S_HDR_FMT if is_f else S_HEADER))
    tw.add_row(1, hdr, height=22)

    # Data rows
    for ri, tr in enumerate(types_all):
        er   = ri + 2
        alt  = S_ALT if ri % 2 == 1 else S_DEFAULT
        vals = tr.get("values", {})
        row  = [(1, tr.get("name", u""), alt)]

        for ci, p in enumerate(display_params, 2):
            name = p["name"]
            st   = p.get("storage_type", "String")
            is_f = p.get("has_formula") or p.get("read_only")

            if is_f:
                expr = p.get("formula_expr", u"")
                row.append((ci, (u"=" + expr) if expr else u"[formula]", S_FORMULA))
                continue

            raw = vals.get(name)
            if st == "Double":
                v = round(float(raw) * unit_scale, 6) if raw is not None else 0.0
                row.append((ci, v, S_NUMBER))
            elif st in ("Integer", "ElementId"):
                row.append((ci, int(raw) if raw is not None else 0, S_INTEGER))
            else:
                row.append((ci, unicode(raw) if raw is not None else u"", alt))

        tw.add_row(er, row, height=16)

    # Blank new-row hint
    tw.add_row(len(types_all) + 2, [(1, u"", S_DEFAULT)], height=16)

    # Meta sheet
    def mr(r, k, v):
        mw.add_row(r, [(1, unicode(k), S_METAKEY), (2, unicode(v), S_DEFAULT)])

    mr(1,  u"family_name",         family_name)
    mr(2,  u"export_timestamp",    _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    mr(3,  u"unit_mode",           unit_mode)
    mr(4,  u"unit_scale",          unicode(unit_scale))
    mr(5,  u"revit_year",          unicode(snapshot.get("revit_year", 0)))
    mr(6,  u"type_count",          unicode(len(types_all)))
    mr(7,  u"param_count",         unicode(len(params_all)))
    mr(9,  u"parameters_json",     json.dumps(params_all))
    mr(10, u"export_columns_json", json.dumps([p["name"] for p in writable_params]))

    # Generate sheet XML — pass dimensions so autoFilter covers actual data
    types_xml = tw.xml(
        col_widths=cw,
        freeze_col=1,
        freeze_row=1,
        num_data_cols=num_cols,
        num_data_rows=len(types_all),
    )
    meta_xml = mw.xml(col_widths={1: 22, 2: 120})

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",        _enc(_CONTENT_TYPES))
        zf.writestr("_rels/.rels",                _enc(_RELS))
        zf.writestr("xl/workbook.xml",            _enc(_WORKBOOK))
        zf.writestr("xl/_rels/workbook.xml.rels", _enc(_WB_RELS))
        zf.writestr("xl/worksheets/sheet1.xml",   _enc(types_xml))
        zf.writestr("xl/worksheets/sheet2.xml",   _enc(meta_xml))
        zf.writestr("xl/sharedStrings.xml",       _enc(ss.xml()))
        zf.writestr("xl/styles.xml",              _enc(_styles_xml()))

# =============================================================================
# IMPORT
# =============================================================================

def _unescape(s):
    s = s.replace(u"&amp;",  u"&")
    s = s.replace(u"&lt;",   u"<")
    s = s.replace(u"&gt;",   u">")
    s = s.replace(u"&quot;", u'"')
    s = s.replace(u"&apos;", u"'")
    return s

def import_from_xlsx(xlsx_path):
    """Reads a .xlsx exported by this module and returns a TypeSnapshot dict."""
    with zipfile.ZipFile(xlsx_path, "r") as zf:
        names = zf.namelist()

        # Shared strings
        ss_list = []
        if "xl/sharedStrings.xml" in names:
            raw = zf.read("xl/sharedStrings.xml").decode("utf-8")
            for m in re.finditer(r"<si>.*?</si>", raw, re.DOTALL):
                tm = re.search(r"<t[^>]*>(.*?)</t>", m.group(0), re.DOTALL)
                ss_list.append(_unescape(tm.group(1)) if tm else u"")

        def _cell_val(cell_str):
            ta = re.search(r'\bt="(\w+)"', cell_str)
            vm = re.search(r"<v>(.*?)</v>", cell_str)
            if not vm:
                return u""
            raw = vm.group(1).strip()
            if ta and ta.group(1) == "s":
                try:
                    return ss_list[int(raw)]
                except (IndexError, ValueError):
                    return raw
            return raw

        def _parse(sheet_xml):
            rows = {}
            for rm in re.finditer(r"<row\b[^>]*>.*?</row>", sheet_xml, re.DOTALL):
                rstr = rm.group(0)
                rn   = re.search(r'\br="(\d+)"', rstr)
                if not rn:
                    continue
                rnum = int(rn.group(1))
                rd   = {}
                for cm in re.finditer(r"<c\b[^>]*(?:/>|>.*?</c>)", rstr, re.DOTALL):
                    cs  = cm.group(0)
                    ref = re.search(r'\br="([A-Z]+\d+)"', cs)
                    if not ref:
                        continue
                    col = re.match(r"([A-Z]+)", ref.group(1)).group(1)
                    rd[col] = _cell_val(cs)
                rows[rnum] = rd
            return rows

        # Meta
        meta_rows = _parse(zf.read("xl/worksheets/sheet2.xml").decode("utf-8"))
        meta = {}
        for rn in sorted(meta_rows):
            k = meta_rows[rn].get("A", "").strip()
            v = meta_rows[rn].get("B", "").strip()
            if k:
                meta[k] = v

        if "unit_scale" not in meta:
            raise ValueError("Missing Meta sheet — was this file exported by RFA Type Manager?")

        unit_scale  = float(meta.get("unit_scale", 12.0) or 12.0)
        params_all  = json.loads(meta.get("parameters_json", "[]"))
        family_name = meta.get("family_name", "Unknown")
        revit_year  = int(meta.get("revit_year", 2024) or 2024)

        param_by_name  = {p["name"]: p for p in params_all}
        writable_names = set(p["name"] for p in params_all
                             if not p.get("read_only") and not p.get("has_formula"))

        # Types sheet
        sheet_rows = _parse(zf.read("xl/worksheets/sheet1.xml").decode("utf-8"))
        if not sheet_rows:
            raise ValueError("Types sheet is empty.")

        col_to_param = {}
        for col, hval in sheet_rows.get(1, {}).items():
            if col == "A":
                continue
            name = (hval or "").strip()
            if " (" in name and name.endswith(")"):
                name = name[:name.rfind(" (")].strip()
            if name in writable_names:
                col_to_param[col] = name

        type_rows = []
        for rn in sorted(sheet_rows):
            if rn == 1:
                continue
            row       = sheet_rows[rn]
            type_name = (row.get("A", "") or "").strip()
            if not type_name:
                continue
            values = {}
            for col, pname in col_to_param.items():
                raw = (row.get(col, "") or "").strip()
                p   = param_by_name.get(pname)
                if not p:
                    continue
                st = p.get("storage_type", "String")
                try:
                    if st == "Double":
                        values[pname] = float(raw) / unit_scale if raw else 0.0
                    elif st in ("Integer", "ElementId"):
                        values[pname] = int(float(raw)) if raw else 0
                    else:
                        values[pname] = raw
                except (ValueError, ZeroDivisionError):
                    values[pname] = None
            type_rows.append({"name": type_name, "values": values})

    return {
        "family_name": family_name,
        "revit_year":  revit_year,
        "parameters":  params_all,
        "types":       type_rows,
    }
