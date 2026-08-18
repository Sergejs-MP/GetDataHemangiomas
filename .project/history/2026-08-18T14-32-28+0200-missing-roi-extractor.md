# Missing ROI extractor session — 2026-08-18

Replaced the fixed two-argument extractor with an input-driven three-argument
workflow. Added robust two-column CSV loading, course derivation, exact plan/ROI
matching, dynamic wide output, explicit failure statuses, CSV escaping, and
input/output collision protection. Removed hard-coded aliases and candidate-file
side effects.

Normalized absolute Gy/cGy DVH points to Gy, made output units explicit, rejected
non-finite, incompletely covered, and unbracketed metrics, and reduced extraction
to one cumulative DVH request per ROI. Patient-close uncertainty now stops
further patient opens while retaining reconciliable rows.
Output replacement is staged through a unique same-directory temporary file so
an existing result is not truncated by a failed write.

Prepared ignored local inputs from the supplied workbook: 56 unique patient-plan
requests in source order and 11 ROI IDs covering all 30 fully blank measurement
columns. Verified exact source/input reconciliation, ROI coverage, Git ignore
behavior, expected 63-column output schema, and a clean diff check without
printing patient values.

Native compilation and ESAPI/DVH execution remain pending because this host is
macOS without .NET Framework 4.8 or ESAPI 17.0.
