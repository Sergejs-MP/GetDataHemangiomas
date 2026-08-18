# Decisions

## 2026-08-18 — Input-driven missing-ROI extraction

- The CLI requires exactly three paths: `PatientID,PlanID` CSV, one-column ROI
  ID list, and output CSV. Output/input path collisions are rejected.
- Course ID is the first single digit found in plan ID (`111_...` maps to course
  `1`); no default course or plan is substituted.
- Plan and `Structure.Id` matching are exact and case-insensitive. Aliases,
  prefixes, candidate selection, and side-effect candidate files are excluded.
- Output stays wide and preserves one row per input plan. Every requested ROI
  contributes status, volume, D2, D50, D60, DVH coverage, and DVH sampling
  coverage columns so workbook rows remain directly reconcilable and coverage
  warnings remain auditable.
- Lookup and data failures remain explicit output rows rather than being skipped.
- Output units are explicit (`cm3`, `Gy`). Absolute cGy values are divided by
  100; unsupported or non-finite dose values remain blank with a status.
- Dose metrics require valid finite Gy/cGy curve points and the requested
  relative volume to be bracketed; no boundary value is fabricated. Valid
  coverage values below the 0.999 project QA threshold retain bracketed metrics
  with an explicit warning. Non-finite or out-of-range coverage blocks dose
  export.
- Plan status is `WARNING` only when every ROI metric is complete and at least
  one ROI has a coverage warning; missing or partial metrics remain `PARTIAL`.
- One cumulative DVH is reused for D2, D50, and D60. A failed patient close stops
  further patient opens while preserving one status row per remaining request.
- The requested DVH bin is 0.01 for Gy grids and 1 for cGy grids. Unit discovery
  temporarily sets `DoseValuePresentation` to absolute, checks `TotalDose`,
  `DosePerFraction`, then `DoseMax3D`, and restores the prior presentation. The
  DVH call independently requests absolute dose. All returned cGy points are
  converted to Gy; percent/unknown units remain rejected. Bin-width scaling still
  requires cGy acceptance testing because ESAPI documentation does not state the
  bin-width unit.
- Output is written to a unique same-directory temporary file and atomically
  moved/replaced only after the full batch is serialized.
- Patient-bearing inputs and exports live in anchored, Git-ignored
  `/ValeriaData/`; only documentation and the ignore rule are tracked.
