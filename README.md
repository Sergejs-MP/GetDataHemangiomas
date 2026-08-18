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

The output preserves one row per input plan and the input order. It includes a
plan-level status and, for each requested ROI, an ROI status plus volume, D2,
D50, and D60. `VolumeUnit` is `cm3`; `DoseUnit` is `Gy`. Absolute cGy DVH values
are converted to Gy, while percent, unknown, undefined, NaN, and infinite dose
values are rejected. Each ROI uses one cumulative DVH, and dose metrics are
exported only when dose and sampling coverage are both at least 99.9% and the
requested relative volume is bracketed by the curve.

Before selecting the DVH bin width, the helper temporarily sets the plan dose
presentation to absolute and restores the original presentation afterward. It
uses the first defined Gy/cGy unit from `TotalDose`, `DosePerFraction`, or
`DoseMax3D`. The DVH request itself also explicitly requests absolute dose, and
all exported dose values are normalized to Gy.

The unit and coverage handling follows the official ESAPI 17 documentation for
[`DoseValue`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.Types.DoseValue.html)
and [`DVHData`](https://docs.developer.varian.com/api/17.0/VMS.TPS.Common.Model.API.DVHData.html).
Spot-check representative values in Eclipse before using the export for
analysis.

The local `ValeriaData` directory is ignored by Git because it may contain
patient identifiers and generated exports.

When merging the export into the supplied workbook, match headers
case-insensitively: the workbook mixes `_vol` and `_Vol` capitalization.

## Status values

- `OK`: all requested values were extracted.
- `PARTIAL`: dose is available, but at least one requested ROI was not fully
  extracted.
- `PATIENT_NOT_FOUND`, `COURSE_NOT_FOUND`, `PLAN_NOT_FOUND`: lookup failed.
- `STRUCTURE_SET_MISSING`: the plan has no structure set.
- `DOSE_UNAVAILABLE`: volume may be present, but dose is absent or invalid.
- `ROI_NOT_FOUND`, `ROI_EMPTY`, `ROI_INVALID_VOLUME`: per-ROI structure status.
- `DVH_PARTIAL`, `DVH_UNAVAILABLE`, `DVH_INCOMPLETE_COVERAGE`: per-ROI DVH
  status.
- `DOSE_UNIT_UNSUPPORTED`: the absolute dose unit could not be normalized to Gy.
- `CLOSE_PATIENT_ERROR`: patient cleanup failed; later requests are retained as
  `SESSION_ABORTED` without opening more patients.
- `ERROR`: an unexpected per-plan error occurred; the row is retained.

The executable rejects an output path that would overwrite either input file.
It writes through a same-directory temporary file so an existing output is only
replaced after the complete CSV has been written. Input/validation failures,
fatal errors, and patient-close aborts return a nonzero exit code; ordinary
per-plan lookup/data failures remain in the complete CSV and must be reviewed by
status.
