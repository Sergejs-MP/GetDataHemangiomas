# Current State

Updated: 2026-08-18

## Work completed

- Reconciled the latest 56-row, 85-column export. All 353 former
  `DVH_COVERAGE_INVALID` results were caused only by finite
  `DVHSamplingCoverage > 1`; native `DVHCoverage` was 1 for every nonempty ROI.
- Changed native coverage checks to audit warnings that never suppress a
  complete finite, bracketed absolute-Gy DVH triplet.
- Implemented the confirmed ROI sequence: exact structure lookup, native
  absolute DVH, then a physical absolute line-histogram fallback only when the
  native D2/D50/D60 triplet is incomplete.
- Added explicit metric source, warning, native-DVH, and line-sampling audit
  fields. The 11-ROI output contract is now 209 columns.
- Passed portable source/API, header/value cardinality, documentation, XML,
  diff, ignored-data, and identifier-leak checks.

## Current status

The original workbook and patient-bearing inputs remain unchanged and outside
version control. Extractor v1.2 source and documentation are updated locally.
The portable static review is complete. The new 209-column executable has not
been compiled or run on Windows/ESAPI.

## Known issues

- This macOS host cannot compile or run the .NET Framework 4.8 / ESAPI 17
  executable.
- The physical line histogram is a non-native approximation. Agreement with
  Eclipse/native DVH, small-ROI sampling convergence, and suitable research
  acceptance limits are not established.
- Absolute cGy profile conversion, presentation restoration, runtime, and all
  fallback failure paths remain unverified on ESAPI 17.

## Recommended next step

Build Release x64 on the ESAPI workstation, rerun all 56 requests, validate the
209-column source/status invariants, and compare line-derived D2/D50/D60 with
native DVH/Eclipse values before merging or analysing fallback metrics.
