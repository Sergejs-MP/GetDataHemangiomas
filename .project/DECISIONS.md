# Decisions

## 2026-08-18 — Native-first physical dose extraction

- The CLI requires exactly three paths: a `PatientID,PlanID` CSV, a one-column
  ROI ID list, and an output CSV. Course ID is the first digit found in Plan ID.
- Plan and `Structure.Id` matching are exact and case-insensitive. Missing and
  empty structures remain explicit and are never sent to a dose fallback.
- A nonempty ROI first uses one native cumulative DVH requested in absolute
  dose and relative volume. D2/D50/D60 require adjacent finite Gy/cGy bracket
  points; cGy is divided by 100 and no boundary value is fabricated.
- Native coverage is audit metadata, not a dose gate. Values below 0.999,
  unavailable values, and values above 1 warn without erasing complete native
  metrics. `DVHSamplingCoverage > 1` is retained and warned, not rejected,
  because ESAPI does not document that field as normalized to 0–1.
- A physical line sampler runs only when the native triplet is incomplete. It
  temporarily selects absolute presentation, uses segment and dose profiles on
  a cell-centred structure-bounds lattice, accepts only finite non-negative
  Gy/cGy samples, restores the original presentation, and never uses EQD2, BED,
  fractionation, PlanSum, registration, or biological weighting.
- Line samples use a 1,024-bin differential histogram. D2/D50/D60 use the lower
  edge of the bin containing the corresponding hottest-volume rank, except that
  rank one uses the exact sampled maximum. A complete line triplet replaces the
  incomplete native triplet as a single consistent source; if line sampling
  fails, defensible partial native values are retained.
- Known in-structure points without valid line dose reduce
  `LineSamplingCoverage`; metrics from remaining valid samples are retained only
  with an incomplete-sampling warning. Incomplete structure profiles,
  inconsistent passes, unsupported units, or no valid samples block line
  metrics.
- `LineInsideVolumeEstimate` and its ratio to `Structure.Volume` audit lattice
  representation only; no automatic volume-ratio or minimum line-coverage
  threshold is imposed. Those remain downstream research validation criteria.
- Line-derived metrics are always explicitly sourced and warned. They are an
  uncommissioned research fallback and must not be pooled with native values
  without source-stratified agreement and sensitivity checks.
- Output stays wide and input-ordered. It has 11 common fields plus 18 audit and
  metric fields per ROI: 209 columns for the supplied 11-ROI list.
- ROI/plan status precedence is missing data before warnings: incomplete values
  are `PARTIAL`; complete warned values are `WARNING`; only complete warning-free
  native values are `OK`.
- Output is written atomically through a same-directory temporary file.
  Patient-bearing inputs and exports remain in Git-ignored `/ValeriaData/`.
