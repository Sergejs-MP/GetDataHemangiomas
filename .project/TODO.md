# TODO

- [x] Review, commit, and push the scoped absolute-dose unit fix.
- [x] Complete static review of the coverage-warning dose export change.
- [ ] Build Release x64 with .NET Framework 4.8 and ESAPI 17 references.
- [ ] Rerun the 56 requests and confirm that expected D2/D50/D60 fields are
  populated for coverage-warning ROIs; investigate any remaining dose-unit or
  invalid-coverage status before analysis.
- [ ] Confirm 56 output rows, 85 columns, stable input order, and review all
  `WARNING` and `PARTIAL` statuses plus recorded coverage fractions.
- [ ] Spot-check volume, D2, D50, and eye D60 against Eclipse, including the
  absolute dose unit and at least one sub-threshold coverage case.
- [ ] Exercise a cGy-configured plan and validate the 1 cGy bin-width branch.
- [ ] Create a preserved-source combined workbook and a separate descriptive
  statistics/chart workbook after the extraction passes validation.
