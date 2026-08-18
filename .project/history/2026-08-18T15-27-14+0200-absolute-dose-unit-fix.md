# Absolute dose unit fix

- Scoped `DoseValuePresentation.Absolute` while resolving the plan dose unit and
  restored the original presentation in `finally`.
- Resolved the unit once per plan from `TotalDose`, `DosePerFraction`, then
  `DoseMax3D`.
- Kept cumulative DVH requests explicitly absolute and retained Gy/cGy-only
  validation, cGy-to-Gy conversion, coverage checks, and interpolation guards.
- Static ESAPI 17 review passed; Windows rebuild and rerun remain required.
