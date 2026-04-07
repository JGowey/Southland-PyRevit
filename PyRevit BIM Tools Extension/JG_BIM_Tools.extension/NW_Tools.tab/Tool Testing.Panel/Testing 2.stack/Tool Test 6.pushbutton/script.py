# -*- coding: utf-8 -*-
"""
Southland Hanger Renumbering Tool
Revit 2022–2026 | PyRevit | IronPython

PIPELINE DESIGN:

1. INPUT STAGE
   - Get selected elements
   - Filter valid hangers
   - Retrieve job number

2. GROUPING STAGE
   - Group by service abbreviation

3. NUMBERING STAGE
   - Determine next available sequence
   - Generate description string
   - Generate point number

4. WRITE STAGE
   - Apply values in single transaction

OUTPUT FORMAT:

GTP_PointDescription_0:
    3/4"-16"-Beam Clamp

GTP_PointNumber_0:
    <ProjectNumber>-<Service>-###   (001, 002, 003...)

Author: Clean structured version
"""

# ==========================================================
# IMPORTS
# ==========================================================

import re
from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    BuiltInParameter,
    Transaction,
    UnitUtils
)

logger = script.get_logger()


# ==========================================================
# CONFIGURATION
# ==========================================================

# ---- Input Parameters on Hanger ----
P_HANGER_SIZE = "CP_Hanger Size"
P_ROD_DIAM    = "CP_Rod Diameter"
P_MID_ATTACH  = "CP_Rod Middle Attachment"
P_SERVICE     = "CP_Service Abbv"

# ---- Output Parameters on Hanger ----
P_OUT_DESC = "GTP_PointDescription_0"
P_OUT_NUM  = "GTP_PointNumber_0"

# ---- Number Formatting ----
NUMBER_SEPARATOR = "-"
SEQUENCE_WIDTH = 3           # 001, 002, 003
ROUND_DENOM = 16             # nearest 1/16"


# ==========================================================
# UTILITY FUNCTIONS
# ==========================================================

def get_selected_elements(doc):
    """
    Robust selection getter.
    Handles ElementIds or Elements depending on pyRevit version.
    """
    sel = revit.get_selection()

    try:
        ids = list(sel.element_ids)
        return [doc.GetElement(eid) for eid in ids]
    except Exception:
        pass

    try:
        items = list(sel)
        if items and hasattr(items[0], "IntegerValue"):
            return [doc.GetElement(x) for x in items]
        return items
    except Exception:
        return []


def lookup_param(elem, name):
    try:
        return elem.LookupParameter(name)
    except Exception:
        return None


def param_text(param):
    """
    Safely returns readable parameter text.
    """
    if not param:
        return ""
    try:
        val = param.AsString()
        if val:
            return val.strip()
    except Exception:
        pass
    try:
        val = param.AsValueString()
        if val:
            return val.strip()
    except Exception:
        pass
    try:
        return str(param.AsInteger())
    except Exception:
        return ""
    

# ==========================================================
# LENGTH CONVERSION (Pure Inches Display)
# ==========================================================

def gcd(a, b):
    while b:
        a, b = b, a % b
    return a


def round_fraction(value, denom):
    return int(round(value * denom)) / float(denom)


def inches_to_string(inches):
    """
    Convert float inches to mixed fraction string.
    Example:
        0.75  -> 3/4"
        16.0  -> 16"
        16.5  -> 16 1/2"
    """
    inches = round_fraction(inches, ROUND_DENOM)
    whole = int(inches)
    frac = inches - whole
    num = int(round(frac * ROUND_DENOM))

    if num == 0:
        return '{}"'.format(whole)

    g = gcd(num, ROUND_DENOM)
    num //= g
    den = ROUND_DENOM // g

    if whole == 0:
        return '{}/{}"'.format(num, den)

    return '{} {}/{}"'.format(whole, num, den)


def length_param_to_inches_string(param):
    """
    Converts:
        - Double (internal feet)
        - Revit formatted string (1' - 4")
    Into:
        Pure inches display (16")
    """
    if not param:
        return ""

    # Case 1: Double storage (internal feet)
    try:
        if param.StorageType.ToString() == "Double":
            feet = param.AsDouble()
            inches = UnitUtils.ConvertFromInternalUnits(
                feet, revit.units.units.inches
            )
            return inches_to_string(inches)
    except Exception:
        pass

    # Case 2: String like 1' - 4"
    raw = param_text(param)
    if not raw:
        return ""

    cleaned = raw.replace("-", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    match = re.search(r"(-?\d+)\s*'\s*(.*)", cleaned)
    if match:
        feet = int(match.group(1))
        remainder = match.group(2).replace('"', '').strip()
        inch_val = parse_inches_component(remainder)
        if inch_val is None:
            return raw
        total_inches = feet * 12.0 + inch_val
        return inches_to_string(total_inches)

    # No feet present
    inch_val = parse_inches_component(cleaned.replace('"', ''))
    if inch_val is None:
        return raw
    return inches_to_string(inch_val)


def parse_inches_component(value):
    """
    Parses:
        3/4
        16
        16 1/2
    """
    value = value.strip()

    mixed = re.match(r"(\d+)\s+(\d+)/(\d+)", value)
    if mixed:
        return float(mixed.group(1)) + float(mixed.group(2)) / float(mixed.group(3))

    frac = re.match(r"(\d+)/(\d+)", value)
    if frac:
        return float(frac.group(1)) / float(frac.group(2))

    integer = re.match(r"^\d+$", value)
    if integer:
        return float(integer.group(0))

    return None


# ==========================================================
# CORE LOGIC
# ==========================================================

def is_valid_hanger(elem):
    """
    Determines whether element is a Southland hanger
    by required parameter presence.
    """
    required = [
        P_ROD_DIAM,
        P_HANGER_SIZE,
        P_SERVICE,
        P_OUT_DESC,
        P_OUT_NUM
    ]
    for pname in required:
        if not lookup_param(elem, pname):
            return False
    return True


def get_job_number(doc):
    """
    Returns ProjectInformation.Number.
    Prompts user if blank.
    """
    pi = doc.ProjectInformation

    try:
        if pi.Number and pi.Number.strip():
            return pi.Number.strip()
    except Exception:
        pass

    try:
        param = pi.get_Parameter(BuiltInParameter.PROJECT_NUMBER)
        value = param_text(param)
        if value:
            return value
    except Exception:
        pass

    # Prompt if still blank
    value = forms.ask_for_string(
        default="",
        prompt="ProjectInformation.Number is blank.\nEnter Job Number:",
        title="Job Number Required"
    )

    if not value:
        script.exit()

    return value.strip()


def build_description(elem):
    """
    Builds:
        3/4"-16"-Beam Clamp
    """
    rod = length_param_to_inches_string(lookup_param(elem, P_ROD_DIAM))
    hanger = length_param_to_inches_string(lookup_param(elem, P_HANGER_SIZE))
    attach = param_text(lookup_param(elem, P_MID_ATTACH))

    parts = []
    if rod: parts.append(rod)
    if hanger: parts.append(hanger)
    if attach: parts.append(attach)

    return "-".join(parts)


def get_next_sequence(elems, prefix):
    """
    Finds next available sequence number within selection.
    """
    max_seq = 0
    pattern = r"^" + re.escape(prefix) + re.escape(NUMBER_SEPARATOR) + r"(\d+)$"

    for e in elems:
        current = param_text(lookup_param(e, P_OUT_NUM))
        match = re.match(pattern, current)
        if match:
            seq = int(match.group(1))
            if seq > max_seq:
                max_seq = seq

    return max_seq + 1


# ==========================================================
# EXECUTION PIPELINE
# ==========================================================

doc = revit.doc

# ---- Stage 1: Selection ----
elements = get_selected_elements(doc)
if not elements:
    forms.alert("Select hanger instances before running the tool.", exitscript=True)

hangers = [e for e in elements if is_valid_hanger(e)]
if not hangers:
    forms.alert("Selection contains no valid Southland hangers.", exitscript=True)

# ---- Stage 2: Job Number ----
job_number = get_job_number(doc)

# ---- Stage 3: Group by Service ----
groups = {}
for h in hangers:
    svc = param_text(lookup_param(h, P_SERVICE))
    if not svc:
        svc = "UNK"
    groups.setdefault(svc, []).append(h)

# ---- Confirmation ----
if not forms.alert(
    "Renumber {} selected hangers?\n\nFormat:\nDescription: 3/4\"-16\"-Beam Clamp\nNumber: Job-Service-###"
    .format(len(hangers)),
    yes=True, no=True
):
    script.exit()

# ---- Stage 4: Write Transaction ----
t = Transaction(doc, "Renumber Southland Hangers")
t.Start()

try:
    for service, elems in groups.items():

        prefix = "{}{}{}".format(job_number, NUMBER_SEPARATOR, service)
        next_seq = get_next_sequence(elems, prefix)

        elems = sorted(elems, key=lambda x: x.Id.IntegerValue)

        for elem in elems:
            desc = build_description(elem)
            number = "{}{}{}".format(
                prefix,
                NUMBER_SEPARATOR,
                str(next_seq).zfill(SEQUENCE_WIDTH)
            )

            lookup_param(elem, P_OUT_DESC).Set(desc)
            lookup_param(elem, P_OUT_NUM).Set(number)

            next_seq += 1

    t.Commit()

except Exception as ex:
    logger.exception(ex)
    t.RollBack()
    forms.alert("Failed:\n{}".format(ex), exitscript=True)

forms.alert("Success.\nRenumbered {} hangers.".format(len(hangers)))