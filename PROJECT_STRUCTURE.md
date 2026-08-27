# PIMS Project Structure

| Document | Version | Status | Date |
|----------|---------|--------|------|
| PROJECT_STRUCTURE.md | 2.1.0 | Phase E Workbook Platform | 2026-08-10 |

---

## Production release layout (Phase E)

```text
Inci_Aku_PPWR_PIMS/
├── README.md
├── CHANGELOG.md
├── FINAL_DATABASE.md          # Schema 1.0.0 FROZEN — do not edit in Phase E/F without approval
├── PLAN.md / ARCHITECTURE_REVIEW.md / DATABASE.md / NAMING_CONVENTION.md
├── PROJECT_STRUCTURE.md
├── requirements.txt
├── build.py                   # --check | --test | --excel-template | --list-tables
├── config.py
├── config/
│   ├── settings.yaml
│   └── lookup_codes.yaml
├── templates/
├── docs/
├── tests/                     # Phase D + Phase E unit tests
├── output/
│   ├── INCI_AKU_PPWR_PIMS_Rev00.xlsx      # Phase E platform workbook
│   └── PHASE_E_WORKBOOK_QA.md
└── src/
    ├── models/                # Frozen table registry (43)
    ├── builders/
    │   ├── phase_e/           # Phase E workbook platform builder
    │   │   ├── workbook_builder.py
    │   │   ├── db_sheets.py
    │   │   ├── lookup_seed.py
    │   │   ├── validations.py
    │   │   ├── ui_sheets.py
    │   │   ├── styles.py
    │   │   └── qa.py
    │   ├── workbook_builder.py  # Legacy structural builder (flags off)
    │   └── *document* stubs     # Word render still OFF
    ├── services/              # Id / BOM / Weight / Variant / Mapping
    ├── repositories/          # ExcelRepository (Phase E read/write)
    ├── validation/
    └── utils/
```

## Commands

```bash
python build.py --check
python build.py --test
python build.py --excel-template
python build.py --list-tables
```

## Phase E delta (vs Phase D)

- Added `src/builders/phase_e/` workbook generation package
- Added `output/INCI_AKU_PPWR_PIMS_Rev00.xlsx` + `PHASE_E_WORKBOOK_QA.md`
- Added `tests/test_phase_e_workbook.py`
- `ExcelRepository` now opens/saves and reads table rows
- Word / batch / production import still not started (Phase F+)
