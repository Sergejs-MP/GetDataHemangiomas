# Native DVH and physical line fallback

Date: 2026-08-18

- Confirmed the extraction order: exact nonempty structure, native absolute
  cumulative DVH, then physical absolute line sampling only for an incomplete
  native D2/D50/D60 triplet.
- Changed DVH coverage from a hard gate to retained warning metadata, including
  finite `DVHSamplingCoverage > 1`.
- Ported the single-`PlanSetup` physical portion of the supplied line sampler.
  Excluded EQD2, BED, PlanSum, registration, weighting, and fractionation logic.
- Added Gy/cGy validation, absolute-presentation restoration, two-pass sample
  identity checks, 1,024-bin rank metrics, source/warning fields, and line audit
  values. The 11-ROI schema is 209 columns.
- Updated README and project decisions. Portable diff, XML, source-contract,
  API-surface, privacy, and Git-ignore checks passed.
- Windows Release x64 compilation, ESAPI execution, Eclipse/native agreement,
  cGy, convergence, and runtime validation remain required.
