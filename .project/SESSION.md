# Current State

Updated: 2026-08-18

## Work completed

- Reworked the extractor around three explicit paths: patient-plan CSV, exact ROI
  ID list, and output CSV.
- Replaced hard-coded course, plan, and ROI selection with input-driven exact
  lookups; course ID comes from the first digit in plan ID.
- Added reconciliable plan/ROI statuses and volume, D2, D50, and D60 output.
- Normalized absolute Gy/cGy DVH points to Gy, rejected incomplete coverage and
  unbracketed/non-finite metrics, and reused one DVH per ROI.
- Prepared the local ignored `ValeriaData` request and ROI-list files from the
  supplied workbook.
- Documented usage, input contracts, statuses, and verification limits.

## Current status

Source and data-schema checks pass. The prepared 56 request rows and 11 ROI IDs
reconcile to all 30 blank workbook columns. The local data directory is ignored
and absent from tracked changes.

## Known issues

- This macOS host has no .NET Framework/ESAPI toolchain, so the executable has
  not been compiled or run here.
- Representative DVH values and the 99.9% coverage threshold still require
  confirmation in the intended Eclipse environment.
- ESAPI documents the DVH bin width as a requested width but does not explicitly
  state its unit; the cGy scaling branch requires the acceptance check in
  `TODO.md`.

## Recommended next step

Complete the Windows/Eclipse acceptance checks in `TODO.md` before using the
export for analysis.
