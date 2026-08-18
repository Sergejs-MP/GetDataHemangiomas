# Coverage-warning dose export

- Profiled the replacement 56-row extraction without exposing patient or plan
  identifiers: 451 ROI dose triplets are present in Gy, while 148 ROI results
  carry the legacy combined coverage status.
- Replaced the hard sub-threshold coverage return with strict per-metric curve
  bracketing and explicit warning statuses; invalid coverage still blocks dose.
- Added round-trip per-ROI DVH coverage and sampling-coverage audit fields,
  producing 85 columns for the supplied 11-ROI list.
- Added plan-level `WARNING` only when all metrics are present and coverage is
  the sole qualification; missing or partial results remain `PARTIAL`.
- Static source, schema, documentation, XML, ignore, and private-identifier
  checks passed. Windows ESAPI compilation, rerun, and Eclipse validation remain
  required.
