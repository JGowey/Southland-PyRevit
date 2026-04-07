# RFA Hanger Sync — Shared Knowledge Document
## For use in DetailersTools (pyRevit) and See 'n' Say (C#)

**Version:** 4.2.0 (pyRevit) | See 'n' Say port pending  
**Author:** Jeremiah Griffith  
**Location (pyRevit):** `GMYSTYB Tools.pulldown/12HangerSync.pushbutton/`

---

## Overview

Syncs RFA Generic Model hanger families to their associated duct or pipe content. Reads geometry and size from selected content elements and writes the correct elevation, clamp size, trapeze width, insulation, and XY position to each matched hanger.

Two workflows:

| Workflow | Trigger | Description |
|----------|---------|-------------|
| **Batch** | User pre-selects duct/pipe runs | Finds all nearby hangers, shows grid, user reviews then syncs all |
| **Pick Pair** | No pre-selection | User picks one part then one hanger; insulation dialog guides snap choice |

---

## Architecture — Content-First (Critical)

**Always select content first, then find hangers.** Never the reverse.

Scanning fabrication collectors (`OST_FabricationDuctwork`, `OST_FabricationPipework`) from a hanger-first approach triggers Revit's ContentDoc queue (`Resolve Missing Servers`, `Fix Content Doc Tree`, `Upgrade for external parameters`, etc.). These operations queue at model open and detonate at the first transaction boundary, crashing Revit with `ExceptionCode=0xe0434352` (non-continuable CLR exception, flags `0x00000081`). Python/C# exception handlers cannot intercept it.

**Content-first avoids this entirely** — the scan uses `OST_GenericModel` only, which has no connection to the fabrication content system.

---

## Crash History: 0xe0434352

Journals 0696–0722. Root causes in discovery order:

| # | Trigger | Fix |
|---|---------|-----|
| 1 | Writing `CP_Hung Object Diameter` / `CP_Host Nominal Diameter` | Nested sub-family regen — write in separate try/except, never rollback from it |
| 2 | `TransactionGroup.Start()` | Any transaction boundary flushes Revit's queued internal ops — removed TG entirely |
| 3 | `Save Modified Doc` worksharing in progress | Check `doc.IsModifiable` before sync |
| 4 | ContentDoc queue built at model open | Content-first architecture — never query fab collectors during scan |
| 5 | Undo + Override Host transaction | Undoing previous syncs queues `ContentDoc_RegenerateFamilyElemsInUndoAndGather` — removed Override Host entirely |

**Rule:** Never start a transaction after the user has executed Undo in the same Revit session without checking AppCDATA for queued ContentDoc operations.

---

## Parameters Written

| Parameter | Content Source | Notes |
|-----------|---------------|-------|
| `CP_Hung Object Bottom Elev` | Slope-interpolated bottom Z | Primary elevation target |
| `BOHElevation` | Same | Fallback |
| `CP_Bottom Elevation` | Same | Fallback |
| `CP_Hung Object Diameter` | `nominal_dia + 2 × insulation_ft` | Clevis clamp size. Can trigger MODIFICATION IS FORBIDDEN regen — write in isolated try/except |
| `CP_Host Nominal Diameter` | Same | Same warning |
| `CP_Trapeze Width Desired` | `Main Primary Width` (ft) | Trapeze hangers on rectangular duct |
| `CP_Host Insulation Thickness` | From selected insulation param | Reference; rod length driven by BOHElevation |
| XY Position | Perpendicular to content curve only | Uses `ElementTransformUtils.MoveElement` with Z=0 |

**Never written:**
- `CP_Hanger Insulation Adjust` — leave as `"By Spec"`, it does not drive geometry
- Any GTP param — writing to GTP triggers element borrow → ContentDoc queue → crash

---

## Size Parameters by Content Type

### Fabrication Pipe

| Param | Value | Notes |
|-------|-------|-------|
| `Product Entry` | Nominal size string e.g. `"8"` | **Use this for nominal dia.** Never includes insulation |
| `Main Primary Diameter` | Can reflect insulated OD | Do NOT use for nominal |
| `Overall Size` | `"8"ø"` — contains `ø` character | Kills `float()` parsing |
| `Outside Diameter` | True OD including wall | Not nominal |

**Nominal in feet:** `float(product_entry.strip('"')) / 12.0`

### Fabrication Duct

| Param | Value | Notes |
|-------|-------|-------|
| `Main Primary Width` | Duct width in **feet** | Feeds `CP_Trapeze Width Desired` |
| `Main Primary Depth` | Duct height in feet | Used for depth, not trapeze |
| `Overall Size` | `"24x12"` or `"14"ø"` | Strip non-numeric chars before parsing |

### Native Duct/Pipe

Use `Diameter`, `Outside Diameter`, `Width`, `Height` — all stored in feet.

---

## Elevation Interpolation (Sloped Content)

**For fabrication parts — do not compute `centerline_Z - half_depth`.** That requires profile geometry knowledge and fails for round duct.

Instead, interpolate directly between `Lower End Bottom Elevation` and `Upper End Bottom Elevation`:

```
bottom_Z(t) = lower_end_bottom + t × (upper_end_bottom - lower_end_bottom)
```

Revit computes these params correctly for all profile shapes (round, rectangular, insulated). `t` is the normalized [0, 1] parameter along the content curve at the hanger's XY position.

**For native duct/pipe** (no fabrication end params): `centerline_Z(t) - half_depth` with `half_depth = OD/2` for pipe, `Height/2` for duct.

---

## XY Projection — Always 2D

All curve projections must be XY-only. `curve.Project(pt)` is 3D — on a sloped pipe it snaps to the nearest endpoint in 3D space, causing hangers to fail the on-segment test and appear to be mismatched.

```
p0, p1 = curve endpoints
ux, uy = XY unit vector along pipe
dot    = (pt.X - p0.X)*ux + (pt.Y - p0.Y)*uy
t      = dot / length_XY          # normalized [0,1]
cx     = p0.X + dot*ux            # XY centerline point across from hanger
cy     = p0.Y + dot*uy
cz     = p0.Z + t*(p1.Z - p0.Z)  # interpolated Z for slope
```

On-segment test: reject hangers where `t < -0.01 or t > 1.01`.

---

## XY Recenter — Perpendicular Only

Move the hanger only in the direction perpendicular to the pipe. Do not snap to the nearest point on the bounded curve (that moves the hanger along the run to the nearest segment endpoint).

```
move_vector = XYZ(cx - pt.X, cy - pt.Y, 0.0)   # Z = 0 preserves rod-top
ElementTransformUtils.MoveElement(doc, hanger.Id, move_vector)
```

---

## Proximity Search

- **Hanger search:** `OST_GenericModel` FamilyInstances only. Never `OST_FabricationDuctwork` or `OST_FabricationPipework` in the scan path.
- **Default radius:** 1.0 ft XY. User-configurable. Saved to settings.
- **Two-stage filter:**
  1. `BoundingBoxIntersectsFilter` on content bbox expanded by radius — fast collector pass
  2. XY perpendicular distance ≤ radius AND on-segment `t` test — rejects hangers past segment endpoints
- **On-segment test:** compare bounded vs unbounded XY projection. If they differ by > 0.01 ft the hanger is past an endpoint — belongs to a different segment.

---

## Insulation Handling

### Auto-detection
Geometry derivation from fabrication part end elevation params:

```
insulation_ft = lower_end_bottom_elevation - lower_end_bottom_of_insulation_elevation
```

Returns 0.0 when uninsulated (params are equal). Reliable for both pipe and duct.

### Insulated clamp sizing
```
insulated_dia = nominal_dia_ft + 2 × insulation_ft
```
Written to `CP_Hung Object Diameter` and `CP_Host Nominal Diameter`.

### Insulated elevation (BOHElevation)
```
hung_bottom_z = pipe_bottom_z - insulation_ft
```
The rod must reach the bottom of the insulation layer, which sits below the bare pipe surface.

### Pick Pair workflow insulation
If auto-detection returns 0 and insulation may exist, `InsulationChoiceWindow` shows:
- Detected thickness (or "could not be determined")
- If undetermined: param dropdown sorted by usage count desc, then alphabetically
- Radio buttons: "Outside of insulation" vs "Pipe/duct surface"

---

## Family and Type Name Lookup

`sym.FamilyName`, `sym.Family.Name`, and `ELEM_FAMILY_PARAM` all fail silently in IronPython for loaded Generic Model families. Reliable pattern (from Renumber tool):

```python
# IronPython
type_id   = hanger.GetTypeId()
elem_type = doc.GetElement(type_id)
family    = elem_type.FamilyName if hasattr(elem_type, 'FamilyName') else ""
type_name = elem_type.Name       if hasattr(elem_type, 'Name')       else ""
```

```csharp
// C#
var elemType  = doc.GetElement(hanger.GetTypeId()) as FamilySymbol;
string family = elemType?.FamilyName ?? "";
string type   = elemType?.Name       ?? "";
```

---

## Settings Persistence

Saved to `hanger_sync_settings.json` in the pushbutton directory (pyRevit) or `%AppData%\JGriffithTech\settings\HangerSync.json` (See 'n' Say).

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `proximity_ft` | float | 1.0 | Max XY hanger search radius |
| `insulation_on` | bool | false | Insulation checkbox state |
| `insulation_param` | string | `""` | Last selected insulation param name |
| `col_family` | bool | true | Family column visible |
| `col_type` | bool | true | Type column visible |
| `col_status` | bool | false | Status detail column visible |

Initial scan uses the saved `proximity_ft` — no Re-Scan needed on first open.

---

## IronPython-Specific Gotchas

- `RowHeight="Auto"` is invalid for `DataGrid` — use a fixed pixel value (`RowHeight="22"`)
- `out.print_code()` does not accept a `language=` keyword argument
- `curve.Project(pt)` is 3D — always use manual XY projection for sloped elements
- For self-contained inline XAML (no separate .xaml file): write to temp file, pass to `WPFWindow.__init__`, unlink after `ShowDialog()`
- `_SyncSuppressor` returning `Continue` is the only safe return value for MODIFICATION IS FORBIDDEN failures — `ProceedWithRollBack` and `ProceedWithCommit` both cause crash cascades

---

## Transaction Rules

- **One flat `Transaction` per hanger** — no `TransactionGroup` (its `Start()` is a transaction boundary)
- **Attach `IFailuresPreprocessor` before `t.Start()`** — options set after `Start()` are ignored in Revit 2024.3
- **Size writes in isolated try/except** — MODIFICATION IS FORBIDDEN from nested sub-family regen must not roll back the elevation write
- **Never write GTP params in a transaction** — element borrow builds ContentDoc queue

---

## What Was Removed (and Why)

| Feature | Reason Removed |
|---------|---------------|
| SI_Version == 0 guard | Not needed — 1 ft radius prevents false positives |
| RK service guard | Not needed — same reason |
| Override Host | Picking content in a transaction after Undo → ContentDoc queue → crash (journal 0722) |
| TransactionGroup | Its Start() detonates ContentDoc queue (journal 0707) |
| GTP_PointElement_0 writes | Element borrow triggers ContentDoc queue |
| `cp.Project(pt)` for proximity | 3D projection fails for sloped pipes |
| `Main Primary Diameter` for nominal size | Reflects insulated OD when insulation present |

