# TODO

- [ ] Build Release x64 with .NET Framework 4.8 and ESAPI 17 references.
- [ ] Run the documented command against `ValeriaData/patient_plans.csv` and
  `ValeriaData/roi_ids.txt` on the authorized Eclipse workstation.
- [ ] Confirm 56 output data rows, 63 columns, stable input order, and review all
  non-`OK` statuses.
- [ ] Spot-check volume, D2, D50, and eye D60 for at least one representative
  plan against Eclipse, including the absolute dose unit.
- [ ] Exercise a cGy-configured plan and confirm that a 1 cGy DVH bin produces
  the same Gy-normalized result as the 0.01 Gy branch.
- [ ] Merge values into a copy of the source workbook by patient-plan key and
  case-insensitive header (`_vol`/`_Vol` varies); preserve the original workbook.
