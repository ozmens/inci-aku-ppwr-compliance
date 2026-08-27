# PHASE O3 — V4 QA

**PHASE O3 TECHNICAL GATE: PASS**
**TECHNICAL FILE NOMINAL LOAD CLEANUP: PASS**
**UI INTERACTION GATE: PASS**
**MANUAL VISUAL ACCEPTANCE: PENDING**

## Absolute data locks
- Counts unchanged: True
- Baseline: `{'packaging_configurations': 247, 'bom_lines': 1690, 'components': 112, 'products': 2046, 'documents': 988}`
- After: `{'packaging_configurations': 247, 'bom_lines': 1690, 'components': 112, 'products': 2046, 'documents': 988}`
- DoC/Label/Statement regenerated: False
- Promoted: False
- Revision: Rev.00 / R00

## Part A — Technical Files
- TF count: 247
- Nominal Load occurrences: **0** (expected 0)
- Render OK / Fail: 247 / 0
- Golden TF SHA-256: `7c95ffc4f0c4d00de442c67a4d0445dee304eee09585dbf224d48ebfdde4156d`
- Hash sync DOCUMENT_LIBRARY: `{'tf_hashes_updated': 0, 'tf_hashes_same': 247, 'non_tf_rows_untouched': 741, 'tf_rows_skipped_missing_map': 0}`
- Manifest TF hashes updated: 247
- Note: `INCI_AKU_PPWR_PIMS_Rev00_FINAL.xlsx` DOCUMENT_LIBRARY uses `pending://metadata-only/…` URIs (no path-based TF hashes). Path-based hash registry updated in V3/V4 candidates + Phase I manifest (247 TF / 741 non-TF unchanged).

### Sample TF regression
- **STARTER**: exists=True nominal_hits=0 labels=['Ambalaj Seti Kodu / Packaging Set Code', 'Konfigürasyon Kimliği / Configuration ID', 'Konfigürasyon / Packaging Configuration', 'Ayırt edici özellik / Variant basis', 'Kaynak BOM / Veri Soy Ağacı Kimliği / Source BOM / Data Line', 'Toplam ambalaj darası / Total packaging tare', 'Üretici / Dosya sahibi / Manufacturer / Dossier owner', 'Doküman revizyonu / Document revision']
- **INDUSTRIAL**: exists=True nominal_hits=0 labels=['Ambalaj Seti Kodu / Packaging Set Code', 'Konfigürasyon Kimliği / Configuration ID', 'Konfigürasyon / Packaging Configuration', 'Ayırt edici özellik / Variant basis', 'Kaynak BOM / Veri Soy Ağacı Kimliği / Source BOM / Data Line', 'Toplam ambalaj darası / Total packaging tare', 'Üretici / Dosya sahibi / Manufacturer / Dossier owner', 'Doküman revizyonu / Document revision']
- **CONTAINER**: exists=True nominal_hits=0 labels=['Ambalaj Seti Kodu / Packaging Set Code', 'Konfigürasyon Kimliği / Configuration ID', 'Konfigürasyon / Packaging Configuration', 'Ayırt edici özellik / Variant basis', 'Kaynak BOM / Veri Soy Ağacı Kimliği / Source BOM / Data Line', 'Toplam ambalaj darası / Total packaging tare', 'Üretici / Dosya sahibi / Manufacturer / Dossier owner', 'Doküman revizyonu / Document revision']

## Part B — UI V4
- Workbook: `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_PIMS_Rev00_FINAL_EXECUTIVE_V4_CANDIDATE.xlsx`
- Delivery: `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_FINAL_DELIVERY_REV00_EXECUTIVE_V4`
- Home → module links: 13 / 13
- Word links: 988 / 988 (broken=0, absolute=0)
- Visible non-Tahoma: 0
- Shape/table intersections: 0
- Duplicate titles: 0
- Native Excel open: {'ok': True, 'error': None, 'sheets': 62}
- Hyperlink levels: `{'A_ui_shape_hyperlinks': 32, 'A_ui_shape_broken': 0, 'B_ui_cell_hyperlinks': 42, 'ui_shape_valid': True}`
- Interaction flow: **PASS**

### Interaction steps
- {'step': 'HOME→Document Center', 'ok': True, 'sheet': 'DOCUMENT_CENTER'}
- {'step': 'OPEN Technical File link', 'ok': True, 'target': '03_CONTAINER/CNT-20-EUR-01/01_Technical_File.docx'}
- {'step': 'Return HOME', 'ok': True, 'sheet': '00_HOME'}
- {'step': 'HOME→Search', 'ok': True, 'sheet': 'SEARCH'}
- {'step': 'Search ST-051-STD-01', 'ok': True, 'result': 'IA-ST-051-STD-01'}
- {'step': 'Search→HOME', 'ok': True, 'sheet': '00_HOME'}
- {'step': 'Shape.Locked on action card', 'ok': True}

## Previews
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_HOME.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_NAVIGATION.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_SEARCH.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_DOCUMENT_CENTER.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_DOC_ENGINE_MAP.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_SHIPMENTS.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_TECHNICAL_FILES.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_DECLARATIONS_OF_CONFORMITY.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_V4_LABELS.png`
- `C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\PHASE_O3_V4_PREVIEW\PHASE_O3_TF_STARTER_page1.png`

## Build log
- === PART A: Technical File Nominal Load cleanup ===
- Golden TF updated: 7c95ffc4f0c4… → 7c95ffc4f0c4…
- Runtime rebuilt; TF nominal XML hits=0
- PHASE_G_GOLDEN_HASHES TECHNICAL_FILE patched
- Regenerating 247 Technical Files…
-   TF merge 25/247
-   TF merge 50/247
-   TF merge 75/247
-   TF merge 100/247
-   TF merge 125/247
-   TF merge 150/247
-   TF merge 175/247
-   TF merge 200/247
-   TF merge 225/247
-   TF merge 247/247
- Smoke-rendering 247 TFs via Word COM…
-   TF smoke render 50/247 ok=50 fail=0
-   TF smoke render 100/247 ok=100 fail=0
-   TF smoke render 150/247 ok=150 fail=0
-   TF smoke render 200/247 ok=200 fail=0
-   TF smoke render 247/247 ok=247 fail=0
- TF regen: count=247 nominal=0 render_fail=0 errors=0
- DOCUMENT_LIBRARY hash sync (INCI_AKU_PPWR_PIMS_Rev00_FINAL.xlsx): {'tf_hashes_updated': 0, 'tf_hashes_same': 0, 'non_tf_rows_untouched': 988, 'tf_rows_skipped_missing_map': 0}
- DOCUMENT_LIBRARY hash sync (INCI_AKU_PPWR_PIMS_Rev00_FINAL_EXECUTIVE_V3_CANDIDATE.xlsx): {'tf_hashes_updated': 247, 'tf_hashes_same': 0, 'non_tf_rows_untouched': 741, 'tf_rows_skipped_missing_map': 0}
- Manifest TF hashes updated: 247
- TECHNICAL FILE NOMINAL LOAD CLEANUP: PASS
- === PART B: Excel UI V4 ===
- V4 copied from V3 → INCI_AKU_PPWR_PIMS_Rev00_FINAL_EXECUTIVE_V4_CANDIDATE.xlsx
- V4 DOCUMENT_LIBRARY TF hashes: {'tf_hashes_updated': 0, 'tf_hashes_same': 247, 'non_tf_rows_untouched': 741, 'tf_rows_skipped_missing_map': 0}
- Baseline counts: {'packaging_configurations': 247, 'bom_lines': 1690, 'components': 112, 'products': 2046, 'documents': 988}
- Class C polish: [{'sheet': 'SHIPMENTS', 'class': 'C', 'table_start_row': None, 'rows': 0, 'duplicate_titles_cleared': True}, {'sheet': 'DOC_ENGINE_MAP', 'class': 'C', 'table_start_row': 6, 'rows': 9, 'cols': 6, 'duplicate_titles_cleared': True}]
- Class A V4 + protect: {'class_a_shapes_created': 214, 'locked_shapes': 428, 'hyperlinks_added': 32, 'protect': [{'sheet': '00_HOME', 'ok': True, 'mode': 'DrawingObjects+Contents', 'locked_shapes': 135}, {'sheet': 'NAVIGATION', 'ok': True, 'mode': 'DrawingObjects+Contents', 'locked_shapes': 57}, {'sheet': 'SEARCH', 'ok': True, 'mode': 'DrawingObjects+Contents', 'locked_shapes': 22}], 'post_protect_nav_ok': {'HOME→DOCUMENT_CENTER': True, 'HOME→SEARCH': True, 'HOME→PACKAGING': True, 'NAV→TECHNICAL_FILES': True, 'NAV→DECLARATIONS': True, 'SEARCH→DOCUMENT_CENTER': True, 'modules_to_HOME': True, 'modules_to_HOME_count': 13}}
- Delivery: C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output\INCI_AKU_PPWR_FINAL_DELIVERY_REV00_EXECUTIVE_V4
- Interaction flow: {'pass': True, 'status': 'PASS', 'steps': [{'step': 'HOME→Document Center', 'ok': True, 'sheet': 'DOCUMENT_CENTER'}, {'step': 'OPEN Technical File link', 'ok': True, 'target': '03_CONTAINER/CNT-20-EUR-01/01_Technical_File.docx'}, {'step': 'Return HOME', 'ok': True, 'sheet': '00_HOME'}, {'step': 'HOME→Search', 'ok': True, 'sheet': 'SEARCH'}, {'step': 'Search ST-051-STD-01', 'ok': True, 'result': 'IA-ST-051-STD-01'}, {'step': 'Search→HOME', 'ok': True, 'sheet': '00_HOME'}, {'step': 'Shape.Locked on action card', 'ok': True}]}
- Previews: 10

---
**PHASE O3 TECHNICAL GATE: PASS**
**TECHNICAL FILE NOMINAL LOAD CLEANUP: PASS**
**UI INTERACTION GATE: PASS**
**MANUAL VISUAL ACCEPTANCE: PENDING**

STOP — do not promote. Do not start Rev.01. Do not regenerate DoC/Label/Statement.