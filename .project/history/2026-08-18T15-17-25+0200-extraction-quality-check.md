# Extraction quality check

- Reconciled the adjacent extraction CSV with the source workbook without
  changing patient data.
- Verified 56/56 patient-plan joins, no duplicate keys, and complete mapping of
  the 30 requested workbook headers.
- Found that all new dose metrics were blank and all usable ROIs reported
  `DOSE_UNIT_UNSUPPORTED`; only volume extraction succeeded.
- Updated bin-width unit selection to use absolute `PlanSetup.TotalDose` before
  the presentation-dependent `DoseMax3D.Unit` fallback.
- Deferred combined/statistical workbooks until a corrected extraction is rerun
  and the spreadsheet-authoring runtime is available.
