# Final PIMS Excel QA

- **RUN_ID:** `XR-20260810T192954Z`
- **FINAL PIMS EXCEL RECOVERY: PASS**

## Original corrupt workbook

- Status: Microsoft Excel cannot open (archived as `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION_CORRUPT_ARCHIVED.xlsx`)
- Root cause: worksheet AutoFilter + Excel Table dual markup (see `EXCEL_RECOVERY_ROOT_CAUSE.md`)

## Rebuild

- Clean base used: `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_PIMS_Rev00.xlsx` (rebuilt Phase E blank)
- Method: Schema 1.0.0 Phase E rebuild → Phase F normalize/promote → native Excel Save As round-trip
- Template dual-AutoFilter sheets after fix: 0
- SAFE dual-AutoFilter sheets after fix: 0

## Native Microsoft Excel tests

- Template Excel open: {'ok': True, 'error': None, 'sheets': 49, 'repair_dialog': 0, 'warning_count': 0, 'saved_as': None}
- SAFE Excel open: {'ok': True, 'error': None, 'sheets': 49, 'repair_dialog': 0, 'warning_count': 0, 'saved_as': None}
- Excel Save As NATIVE: {'ok': True, 'error': None, 'sheets': 49, 'repair_dialog': 0, 'warning_count': 0, 'saved_as': 'C:\\Users\\burcu\\Documents\\YAZILIM\\Inci_Aku_PPWR_PIMS\\output\\INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION_EXCEL_NATIVE.xlsx'}
- NATIVE reopen: {'ok': True, 'error': None, 'sheets': 49, 'repair_dialog': 0, 'warning_count': 0, 'saved_as': None}
- Repair dialog count: 0 (CorruptLoad=0; open refused previously, now clean)
- Excel warning count: 0

## Production counts (after round-trip)

- Configurations: 247
- Starter / Industrial / Container: 240 / 3 / 4
- BOM lines: 1690
- Components: 112
- Products: 2046
- TF / DoC / Statement / Doc Library: 247 / 247 / 247 / 988

## ST-051 fixture

- `{'set': 'ST-051-STD-01', 'cfg': 'IA-ST-051-STD-01', 'source': 'IA-ST-CFG-0122', 'tare': 47.0384, 'products_ok': True}`

## Structural QA

- Formula errors: 0 expected on open (dashboard uses COUNTA; Excel calculates on open)
- Broken links / external links: 0
- Broken Excel tables (dual AutoFilter): 0
- Dashboard production data: YES (status LOADED + COUNTA formulas over populated sheets)

## Final deliverable

- Path: `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_PIMS_Rev00_FINAL.xlsx`
- SHA-256: `69172eb1c69857837f7cf0f862b7cd54ce8c426d18b6af4698e23b3b62ae918f`

**FINAL PIMS EXCEL RECOVERY: PASS**

- Word outputs unmodified: YES
- Golden templates unmodified: YES
- Rev01 started: NO
