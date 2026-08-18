# TODO

- [x] Implement exact structure lookup followed by native absolute DVH and a
  physical absolute line-histogram fallback for incomplete native triplets.
- [x] Remove coverage as a hard dose gate and retain raw coverage warnings.
- [x] Complete static/API/schema/privacy review of extractor v1.2.
- [ ] Build Release x64 with .NET Framework 4.8 and ESAPI 17 references.
- [ ] Rerun the 56 requests and confirm 56 rows, 209 columns, stable key order,
  physical-Gy units, and consistent status/source/warning fields.
- [ ] Spot-check native volume, D2, D50, and eye D60 values in Eclipse.
- [ ] Compare `LINE` against native DVH on representative large, small, thin,
  irregular, and high-resolution ROIs; predefine agreement limits and test a
  half-spacing validation variant for convergence, especially for D2.
- [ ] Exercise Gy and cGy plans, presentation restoration, partial line dose
  coverage, unsupported units, missing profiles, and two-pass mismatch paths.
- [ ] Review performance for the complete 56-plan batch.
- [ ] Create preserved-source combined and descriptive-statistics workbooks
  only after native and fallback values pass the appropriate validation gates.
