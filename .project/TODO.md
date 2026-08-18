# TODO

- [x] Review, commit, and push the scoped absolute-dose unit fix.
- [ ] Build Release x64 with .NET Framework 4.8 and ESAPI 17 references.
- [ ] Rerun the 56 requests and confirm that expected D2/D50/D60 fields are
  populated; investigate any remaining dose-unit status before analysis.
- [ ] Confirm 56 output rows, 63 columns, stable input order, and review all
  non-`OK` statuses.
- [ ] Spot-check volume, D2, D50, and eye D60 against Eclipse, including the
  absolute dose unit.
- [ ] Exercise a cGy-configured plan and validate the 1 cGy bin-width branch.
- [ ] Create a preserved-source combined workbook and a separate descriptive
  statistics/chart workbook after the extraction passes validation.
