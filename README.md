# GetDataHemangiomas

Command-line ESAPI extractor for collecting the same ROI measurements across a
list of patient plans.

## Requirements

- Windows x64
- .NET Framework 4.8 / Visual Studio MSBuild
- Varian ESAPI 17.0 at the paths configured in `GetDataHemangiomas.csproj`
- Eclipse database access appropriate for the patients being read

## Usage

```text
GetDataHemangiomas.exe <patients-plans.csv> <roi-ids.txt> <output.csv>
```

Example:

```text
GetDataHemangiomas.exe "ValeriaData\patient_plans.csv" "ValeriaData\roi_ids.txt" "ValeriaData\missing_roi_data.csv"
```

`patients-plans.csv` contains `PatientID,PlanID`. The header is optional, and
quoted CSV fields are supported. Course ID is derived from the first digit in
Plan ID; for example, plan `111_Example` is searched in course `1`. Plan ID
matching is exact and case-insensitive.

`roi-ids.txt` contains one exact Eclipse `Structure.Id` per line. Blank lines,
lines beginning with `#`, a `ROIID` header, and case-insensitive duplicates are
ignored. ROI matching is exact and case-insensitive; aliases and partial matches
are not used.

## Extraction sequence

Each requested ROI follows the same ordered decision path:

1. Find the exact structure. A missing structure is `ROI_NOT_FOUND`; an empty
   structure is `ROI_EMPTY`. No dose fallback is attempted for either case.
2. For a nonempty structure with valid plan dose, request one native Eclipse
   cumulative DVH using absolute dose and relative volume. D2, D50, and D60 are
   accepted only when each requested volume is bracketed by adjacent finite
   Gy/cGy curve points. Absolute cGy values are divided by 100.
3. If the native DVH supplies the complete triplet, use it. Coverage anomalies
   remain warnings and do not erase finite, bracketed native values.
4. Only if the native triplet is incomplete or unavailable, attempt the
   physical line-histogram fallback. If that produces a complete triplet, use
   all three fallback metrics and mark `DoseSource=LINE` plus `LINE_FALLBACK`.
   If it fails, retain any defensible partial native values as
   `DoseSource=DVH_PARTIAL`; otherwise use `DoseSource=NONE`.

All dose metrics are physical absolute dose in Gy. The extractor contains no
EQD2, BED, alpha/beta, fractionation, PlanSum, or registration calculation.

## Native DVH coverage

Finite `DVHCoverage` and `DVHSamplingCoverage` values are written exactly as
returned by ESAPI using invariant round-trip precision. Values below the 0.999
project QA threshold, unavailable values, and values above 1 produce explicit
warning codes but do not suppress a complete native Gy triplet.

ESAPI documents `DVHCoverage` as normalized from 0 to 1, but does not document
the same range for `DVHSamplingCoverage`. Consequently, sampling-coverage values
above 1 are retained and warned for audit; they are not treated as invalid or
as a reason to invoke the fallback.

## Physical line fallback

The fallback saves the plan's current dose presentation, sets it to absolute,
and restores it in `finally`. It uses `Structure.GetSegmentProfile` to identify
in-structure points and `Dose.GetDoseProfile` to read dose on cell-centred
z-lines across the structure bounds at approximately the plan dose-grid
spacing. Only finite, non-negative Gy/cGy samples are retained; cGy is divided
by 100.

After a complete structure-membership pass, available dose samples are reduced
to a 1,024-bin differential histogram. D2, D50, and D60 are selected from the
hottest 2%, 50%, and 60% sample ranks using the selected bin's lower edge; when
the requested rank is one, the exact sampled maximum is used. The maximum sample
and bin width are exported so the quantization is auditable. If
some known in-structure points have no usable dose, the available-sample
triplet is still reported with `LINE_SAMPLING_INCOMPLETE`; the valid/inside
ratio is exported as `LineSamplingCoverage`. There is deliberately no minimum
coverage or sample-count gate: downstream research QA must use the exported
counts and ratio as exclusion or sensitivity criteria. An incomplete structure
profile, inconsistent two-pass sampling, any unsupported profile unit, or no
valid samples leaves the fallback unavailable.

`LineInsideVolumeEstimate` is the inside-sample count multiplied by the lattice
cell volume. `LineVolumeRatio` compares that estimate with ESAPI
`Structure.Volume`; it audits how the lattice represents the ROI, not how much
of the ROI has valid dose. Both are diagnostic only, with no automatic pass/fail
threshold.

This histogram is a non-native research fallback. It uses equal-weight binary
sample points on a structure-bounds lattice and does not reproduce Eclipse's
native voxel/partial-volume DVH algorithm. Small, thin, irregular, or
high-resolution structures—and D2 in particular—may be sensitive to grid
alignment, sample count, and bin width. `LINE` values always produce an ROI
warning and prevent a plan from being `OK`; the plan is `WARNING` only when all
ROIs are complete and otherwise remains `PARTIAL`. Line values must be validated
against native DVH/Eclipse values with predefined agreement and
sampling-convergence criteria before research use. Clinical use requires formal
local commissioning.

## Output

The output preserves one row per input plan and the input order. Common fields
are:

```text
PatientID, CourseID, PlanID, MatchedPlanID, PlanStatus, Message,
ExtractorVersion, DoseBasis, FallbackMethod, VolumeUnit, DoseUnit
```

`ExtractorVersion=1.2.0`, `DoseBasis=PHYSICAL_ABSOLUTE`,
`FallbackMethod=PHYSICAL_LINE_HISTOGRAM`, `VolumeUnit=cm3`, and `DoseUnit=Gy`
identify the data contract.

Each ROI adds 18 fields:

```text
Status, WarningCodes, DoseSource, Vol, D2, D50, D60,
DVHStatus, DVHCoverage, DVHSamplingCoverage,
LineStatus, LineInsideSamples, LineValidDoseSamples, LineSamplingCoverage,
LineInsideVolumeEstimate, LineVolumeRatio, LineMaxDose, LineBinWidth
```

The supplied 11-ROI list therefore produces 209 columns: 11 common fields plus
18 fields per ROI. Downstream work must map columns by header name, not by
position. `WarningCodes` is pipe-delimited. `DoseSource` is `DVH`, `LINE`,
`DVH_PARTIAL`, `NONE`, or blank when dose extraction was not attempted.

## Status values

- ROI `OK`: volume and D2/D50/D60 are complete with no warnings.
- ROI `WARNING`: all requested values are present, but coverage, fallback,
  sampling, volume, or another audit warning requires review. `LINE` can never
  produce plain `OK` at ROI or plan level.
- ROI `PARTIAL`: at least one dose value exists, but volume or part of the dose
  triplet is missing.
- ROI `DOSE_UNAVAILABLE`: the structure exists but no dose metric is available.
- `ROI_NOT_FOUND`, `ROI_EMPTY`: the requested structure is absent or unusable.
- `DVHStatus` independently records native `OK`, `WARNING`, `PARTIAL`,
  `UNAVAILABLE`, `DOSE_UNIT_UNSUPPORTED`, `PRESENTATION_RESTORE_ERROR`, or
  `ERROR`.
- `LineStatus` is `NOT_NEEDED`, `OK`, or `WARNING` when a line result is
  complete; otherwise it records the unavailable/error reason.
- Plan `OK`: every requested ROI is complete and warning-free.
- Plan `WARNING`: every requested ROI is complete and at least one has a
  warning.
- Plan `PARTIAL`: at least one requested ROI remains incomplete after fallback.
- `PATIENT_NOT_FOUND`, `COURSE_NOT_FOUND`, `PLAN_NOT_FOUND`,
  `STRUCTURE_SET_MISSING`, `DOSE_UNAVAILABLE`, `CLOSE_PATIENT_ERROR`,
  `SESSION_ABORTED`, and `ERROR` retain one reconcilable row for operational or
  lookup failures.

When merging into the workbook, copy only finite metric cells and never convert
`ROI_NOT_FOUND`, `ROI_EMPTY`, or unavailable dose to zero. Preserve
`DoseSource`, `WarningCodes`, and the line audit fields. Do not pool native-DVH
and line-derived values without source-stratified QA and sensitivity analysis.
Match headers case-insensitively because the workbook mixes `_vol` and `_Vol`.

The unit, DVH, and profile APIs are documented by ESAPI 17 under
[`DoseValue`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.Types.DoseValue.html),
[`DVHData`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.DVHData.html),
[`Dose`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.Dose.html),
and [`Structure`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.Structure.html).

The executable rejects an output path that would overwrite either input file.
It writes through a same-directory temporary file so an existing output is only
replaced after the complete CSV has been written. Input/validation failures,
fatal errors, and patient-close aborts return a nonzero exit code; ordinary
per-plan lookup/data failures remain in the complete CSV and must be reviewed by
status.

The local `ValeriaData` directory is ignored by Git because it may contain
patient identifiers and generated exports.
