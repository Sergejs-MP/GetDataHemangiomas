# Current State

Updated: 2026-08-18

## Work completed

- Reconciled the replacement Windows extraction: 56 unique patient-plan rows
  match the source workbook, with 451 complete ROI dose triplets in Gy.
- Confirmed that 148 of 165 non-`OK` ROI results received the former combined
  incomplete/non-finite coverage status; the legacy output does not record the
  raw coverage values or test whether those DVH curves bracket the metrics.
- Changed valid sub-threshold coverage from a hard blocker to an explicit
  warning while retaining only finite, curve-bracketed D2/D50/D60 values.
- Added per-ROI DVH coverage and sampling-coverage audit fields and distinct
  plan-level `WARNING` semantics.

## Current status

The original workbook remains unchanged. The updated extractor source has
passed static review, but the new 85-column output has not yet been built or run
on Windows. The current replacement CSV remains the prior 63-column extraction;
the number of warning-status metrics that will pass strict bracketing is not yet
known.

## Known issues

- This macOS host cannot compile or run the .NET Framework 4.8 / ESAPI 17
  executable.
- The standalone spreadsheet-authoring runtime is unavailable in this Codex
  task, so Excel output must use a later enabled task or an explicit live Excel
  session.
- The cGy bin-width branch, the 0.999 sampling-coverage threshold, and
  representative warning-status DVH values remain unverified in Eclipse.

## Recommended next step

Rebuild on Windows, rerun the 56 plans, review the recorded coverage fractions,
and spot-check warning-status D2/D50/D60 values in Eclipse before creating the
combined and descriptive-statistics workbooks.
