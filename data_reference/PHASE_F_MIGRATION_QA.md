# Phase F Migration QA

- **RUN_ID:** `PF-20260810T174412Z`
- **Timestamp (UTC):** 2026-08-10T17:44:20Z
- **Schema version:** 1.0.0

## PRIMARY SOURCE QUALIFICATION

PRIMARY SOURCE QUALIFICATION: PASS
File: C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\input\production\INCI_AKU_PPWR_Final_Configuration_Register_Rev00_GOLDEN_VARIANTS_FINAL.xlsx
Total configurations: 247
Starter: 240
Industrial: 3
Container / Loading: 4
Exact BOM rows: 1690
Product map rows: 2046
Invalid suffixes: 0
Duplicate set codes: 0
Duplicate final IDs: 0
Duplicate Variant Basis within family: 0
Document-ID errors: 0
Configs missing BOM: 0
ST-051-STD-01: {'present': True, 'Final Configuration ID': 'IA-ST-051-STD-01', 'Source Configuration ID': 'IA-ST-CFG-0122', 'Packaging Mass kg': 47.0384, 'bom_components': ['4000033', '4000037', '4000130', '4000300', '4000301', '4000450', '4000782', '4001311'], 'source_from_bom': ['IA-ST-CFG-0122']}
ST-051-STD-02: {'present': True, 'Final Configuration ID': 'IA-ST-051-STD-02', 'Source Configuration ID': 'IA-ST-CFG-0123', 'bom_components': ['4000033', '4000037', '4000130', '4000300', '4000301', '4000450', '4001108', '4001311'], 'source_from_bom': ['IA-ST-CFG-0123']}

## SOURCE INVENTORY

Files inventoried: 7

## COUNTS

- Final configurations: 247
- Starter: 240
- Industrial: 3
- Container / Loading: 4
- Components: 112
- Products: 2046
- Exact BOM lines: 1690
- Component-material rows: 51
- Technical files: 247
- DoCs: 247
- Document library: 988
- Statements: 247
- Commercial scenarios: 1917

## IDENTITY / VARIANT BASIS

- Duplicate Variant Basis errors: 0
- Document-ID errors: 0
- Codec sample roundtrip: True

## WEIGHT / BOM

- Blocking errors: 0
- BOM rows staged: 1690

## DPP TRACEABILITY

{'ok': True, 'product': '1011935', 'set_code': 'ST-051-STD-01', 'final_id': 'IA-ST-051-STD-01', 'source_cfg': 'IA-ST-CFG-0122', 'bom_lines': 8, 'tare_kg': 47.0384, 'expected_tare_kg': 47.0384}

## WORKBOOK TECHNICAL QA

- Production workbook: `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION.xlsx`
- PACKAGING_CONFIGURATION rows: 247
- UI sheets preserved: ['00_README', '01_DASHBOARD', '02_RELEASE_CONTROL', '03_DATA_DICTIONARY', '04_IMPORT_GUIDE']
- Formula error literals: 0
- Broken external links: 0 (no external links added)

## TESTS

- unittest exit code: 0

## CONFIRMATIONS

- Production data loaded: YES
- Word generation run: NO
- PDF generation run: NO
- Phase G started: NO
