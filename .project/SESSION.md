# Current State

Updated: 2026-08-18

## Work completed

- Reconciled the first Windows extraction with the source workbook: all 56
  patient-plan keys and all 30 requested headers match with no duplicates.
- Confirmed that volumes were extracted, but every requested D2/D50/D60 value is
  blank because the run returned `DOSE_UNIT_UNSUPPORTED` for every usable ROI.
- Changed DVH bin-width unit selection to use a scoped absolute presentation,
  check `TotalDose`, `DosePerFraction`, then `DoseMax3D`, and restore the prior
  presentation. DVH output remains normalized to Gy.

## Current status

The original workbook remains unchanged. No combined or statistical workbook is
ready to share because the dose extraction must be rerun and validated first.
The unit-source fix has passed static review and is published for the Windows
build and rerun.

## Known issues

- This macOS host cannot compile or run the .NET Framework 4.8 / ESAPI 17
  executable.
- The standalone spreadsheet-authoring runtime is unavailable in this Codex
  task, so Excel output must use a later enabled task or an explicit live Excel
  session.
- The cGy bin-width branch and representative DVH values remain unverified.

## Recommended next step

Rebuild on Windows, rerun the 56 plans, and require nonblank verified dose
metrics before creating the combined and descriptive-statistics workbooks.
