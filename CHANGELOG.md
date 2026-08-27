# Changelog

All notable changes to the İnci Akü PPWR Packaging Management System are documented in this file.

Format: [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)  
Versioning: [Semantic Versioning](https://semver.org/)

---

## [Unreleased]

### Planned

- PDF customer export packs (optional downstream)
- ERP shipment import adapters

---

## [3.1.0] — 2026-08-10

### Added

- **Phase J — PIMS application layer + document registry sync**
  - Synced all 988 `DOCUMENT_LIBRARY` rows to Phase I relative DOCX paths + SHA-256
  - Repaired `DOCUMENT_LINK` XOR targets (TF / DoC / Label / Statement)
  - Cleared stale Phase F “Word not generated / metadata-only” notes
  - Restored UI modules: HOME, NAVIGATION, SEARCH, DOCUMENT_CENTER, TECHNICAL_FILES,
    DECLARATIONS_OF_CONFORMITY, LABELS, SHIPMENT_STATEMENTS, PACKAGING_CONFIGURATIONS,
    PRODUCT_MASTER, COMPONENT_MASTER, SHIPMENTS, DOC_ENGINE_MAP
  - Candidate: `output/INCI_AKU_PPWR_PIMS_Rev00_FINAL_UI_CANDIDATE.xlsx`
  - Delivery hyperlink test tree: `output/PHASE_J_DELIVERY_TEST/`
  - CLI: `python build.py --phase-j`
  - QA: `output/PHASE_J_PIMS_INTEGRATION_QA.md`

### Notes

- Does **not** overwrite `INCI_AKU_PPWR_PIMS_Rev00_FINAL.xlsx` until manual UI acceptance.
- Word / Golden unmodified; Rev.01 not started.

---

## [3.0.1] — 2026-08-10

### Fixed

- **Excel PIMS recovery (native Excel-safe workbook)**
  - Root cause: worksheet `autoFilter` written together with Excel Tables → Microsoft Excel refused to open Phase E blank and PRODUCTION workbooks
  - Removed duplicate `ws.auto_filter.ref` from Phase E `db_sheets` and Phase F `promoter`
  - Promoter now deletes ListObject before row structural edits
  - Rebuilt clean template + re-promoted production data
  - Native Excel COM open + Save As round-trip validated
  - Final deliverable: `output/INCI_AKU_PPWR_PIMS_Rev00_FINAL.xlsx`
  - Reports: `output/EXCEL_RECOVERY_ROOT_CAUSE.md`, `output/FINAL_PIMS_EXCEL_QA.md`

### Notes

- Word / Golden outputs unmodified; Rev.01 not started.
- Obsolete prototype `Inci_Aku_PPWR_Packaging_Management_System_Rev00.xlsx` marked NOT FOR DELIVERY.

---

## [3.0.0] — 2026-08-10

### Added

- **Phase I — Full production Word batch / final company delivery**
  - 247 configuration packs × 4 DOCX = **988** customer documents
  - Output tree: `output/PHASE_I_FINAL/` (01_STARTER / 02_INDUSTRIAL / 03_CONTAINER)
  - Manifest: `90_MANIFEST/INCI_AKU_PPWR_DOCUMENT_MANIFEST.xlsx` (+ CSV)
  - QA: `99_QA_REPORT/PHASE_I_BATCH_QA.md` / `.json` / `PHASE_I_DOCUMENT_QA.xlsx`
  - 988 Word COM PDF renders under `99_QA_REPORT/renders/`
  - Final ZIP: `output/INCI_AKU_PPWR_FINAL_COMPANY_DELIVERY_REV00.zip`
  - CLI: `python build.py --phase-i`
  - Flag: `ENABLE_WORD_BATCH_GENERATION=True` (controlled Phase I)

### Notes

- **PHASE I FINAL RELEASE: PASS** — 247/247 packs PASS; all hard counters 0.
- Golden masters unmodified; production PIMS unmodified (read-only generation).
- Rev.01 not started.

---

## [2.9.0] — 2026-08-10

### Added

- **Phase H — Pilot visual acceptance / Golden release gate**
  - Word COM render + OOXML gates for all 12 Phase G pilot DOCX
  - Reports: `output/PHASE_H_ACCEPTANCE/PHASE_H_VISUAL_ACCEPTANCE.md` / `.json`
  - PDF renders under `output/PHASE_H_ACCEPTANCE/renders/`
  - CLI: `python build.py --phase-h`
  - Hard gates: white-on-light=0, visible non-Tahoma=0, no blank pages, no token/sample leaks

### Notes

- **PHASE H RELEASE GATE: PASS** — Golden Word Engine ready for Phase I (not started).
- `ENABLE_WORD_BATCH_GENERATION` remains False.
- Full 247 batch / 988 production DOCX not generated.

---

## [2.8.0] — 2026-08-10

### Added

- **Phase G — Golden Word template integration (pilot only)**
  - Populated ST-012-EUR-01 Golden masters treated as style authority
  - Runtime templates in `templates/word_runtime/` (sample identity tokenized)
  - Merge engine (paragraph/table/header/footer token replace + dynamic BOM/product rows)
  - Pilot packs for `ST-051-STD-01`, `IND-24V-01`, `CNT-20-STD-01` (12 DOCX)
  - QA: `output/PHASE_G_PILOT/QA/PHASE_G_PILOT_QA.md`
  - CLI: `python build.py --word-pilot`
  - Flags: `ENABLE_WORD_PILOT_GENERATION=True`, `ENABLE_WORD_BATCH_GENERATION=False`

### Notes

- Original Golden DOCX under `templates/word_golden/*ST-012*` untouched; GOLDEN aliases are copies.
- Production PIMS not modified.
- Full 247 batch not run at Phase G close.

---

## [2.7.0] — 2026-08-10

### Added

- **Phase F — Production data migration**
  - Content-based Level-1 Golden Register qualification gate (247 / 240 / 3 / 4)
  - Canonical Level-1 file: `INCI_AKU_PPWR_Final_Configuration_Register_Rev00_GOLDEN_VARIANTS_FINAL.xlsx`
  - Migration pipeline under `src/importers/production/`
  - `VariantDescriptionCodec` for DESCRIPTION TR/EN round-trip
  - Production workbook: `output/INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION.xlsx`
  - Inventory / discrepancies / QA reports (`PHASE_F_*`)
  - CLI: `python build.py --migrate-production`

### Notes

- Exact filename `…_GOLDEN_VARIANTS_FINAL.xlsx` was content-identical to the documentation-set
  `…_Register_Rev00.xlsx` (SHA-256 match); Level-1 accepted only after content gate PASS.
- Phase E blank template `INCI_AKU_PPWR_PIMS_Rev00.xlsx` preserved.
- Word / PDF / Phase G not started.

---

## [2.6.0] — 2026-08-10

### Added

- **Phase E — Production PIMS Excel Workbook Platform**
  - `output/INCI_AKU_PPWR_PIMS_Rev00.xlsx` via `python build.py --excel-template`
  - 43 frozen Schema 1.0.0 database sheets + Excel Tables (exact names/columns)
  - UI sheets: `00_README`, `01_DASHBOARD`, `02_RELEASE_CONTROL`, `03_DATA_DICTIONARY`, `04_IMPORT_GUIDE`
  - QA visibility sheet `ZZ_QA_WEIGHT_FIXTURE` (Python WeightService parity = 800 g)
  - Lookup-only seeds (no production masters / BOMs / shipments)
  - `src/builders/phase_e/` builder modules (styles, db sheets, validations, UI, QA)
  - `ExcelRepository` open/save/table-oriented read helpers
  - `output/PHASE_E_WORKBOOK_QA.md`
  - Phase E unit tests (`tests/test_phase_e_workbook.py`)

### CIP / CIF decision

- Schema 1.0.0 defines `LKP_INCOTERM` structure only (no hardcoded code list).
- **CIF** and **CIP** are distinct Incoterms 2020 terms (not aliases).
- Both are seeded as active lookup rows. No schema change. No semantic duplicate.

### Explicitly not done

- Production İnci Akü data import
- 247 final production configurations
- Word / PDF / batch customer outputs
- Phase F

---

## [2.5.0] — 2026-08-10

### Added

- **Phase D — Python Project Architecture / Golden Variant Ready**
  - Domain models, DocumentContext, IdService, VariantBasisService, BomService, WeightService
  - Product mapping + revision + document-link services
  - Validation framework (ValidationResult / ERROR|WARN|INFO)
  - Document builders (TF / DoC / Label / Statement) as architecture stubs — no Word render
  - Unit tests for IDs, Variant Basis, BOM separation, weights, mapping
  - `config/settings.yaml`, `lookup_codes.yaml`; Excel/Word/batch flags OFF
  - README rewritten for Golden Variant + Battery DPP traceability

### Explicitly not done

- Excel workbook generation
- Production data import
- Word/PDF batch outputs

---

## [2.4.0] — 2026-08-02

### Added

- **Final Development Phase — production readiness** (schema frozen)
  - Import templates: `templates/import/Import_Templates.xlsx`
  - Document engine variable catalog (JSON + xlsx + hidden `DOC_ENGINE_VARS`)
  - Export engine shipment package stub + manifest
  - Guides: `SYSTEM_ADMIN_GUIDE.md`, `IMPORT_GUIDE.md`, `DOCUMENT_ENGINE_GUIDE.md`
  - Protect all hidden/system sheets (`PIMS_TECH`)
  - Global Dashboard search (config / scenario / shipment-lot / component)
  - Validation: packaging↔TF coverage, scenario Incoterms

---

## [2.3.0] — 2026-08-02

### Changed

- **Phase 11 — PPWR Operational Model** (business logic only; schema frozen)
  - Docs: `PPWR_OPERATIONAL_MODEL.md`, `USER_WORKFLOW.md`, `IMPLEMENTATION_CHANGES.md`
  - Dashboard wizard: 7 steps starting with Packaging Configuration
  - Lot Number → `SHIPMENT.EXTERNAL_REF`; product resolved from scenario
  - Document hierarchy enforced in UX + validations (`V-SHP-INCI-02` lot warn)

---

## [2.2.0] — 2026-08-02

### Added

- **Phase 10 — Daily Workflow Wizard** on Dashboard (command center)
  - Steps 1–8: Product → Packaging Config → Scenario → Shipment → Statement → Tech File → DoC → Export
  - Live **NEXT ACTION** banner (never wonder what to do next)
  - Yellow dropdowns + auto-resolved green results + GO buttons
  - Shipment draft values auto-assembled from wizard selections
  - Document package readiness checklist
  - No database / architecture changes

---

## [2.1.1] — 2026-08-02

### Changed

- Operating model clarified (no schema changes):
  - **Packaging Configuration** is the primary business object (Product = ERP reference only)
  - Technical File ↔ exactly one Packaging Configuration
  - Commercial Scenario selects DoC variant + shipment document package only (never edits Tech Files)
  - Shipment creation = Product + Packaging Configuration + Commercial Scenario
  - Dashboard/nav order and validation rules (`V-TF-INCI-01`, `V-SHP-INCI-01`) aligned to this model

---

## [2.1.0] — 2026-08-02

### Added

- **Phase 9 — İnci Akü Implementation UX** (no schema changes)
  - Exactly 8 visible worksheets for daily operations
  - Dashboard navigation cards + KPIs + search
  - HOME / Dashboard / BACK on every ops screen
  - Yellow / Blue / Green / Gray cell semantics
  - Packaging recipe lines colocated on Packaging Configuration
  - Auto status panels for Component, Product, Packaging, Shipment, Statement, Tech File, DoC
  - İnci business-mode guidance (Starter / Industrial / Container / Scenario)
  - Expanded continuous validation rules (weight, supplier, evidence, duplicates, lookups)

### Fixed

- Excel case-insensitive sheet rename (`SHIPMENT` → `Shipment`) via two-step rename

---

## [2.0.0] — 2026-08-02

### Added

- **Production release** `Inci_Aku_PPWR_Packaging_Management_System_Rev00`
- Complete Python package (`models`, `builders`, `validation`, `reports`)
- Production workbook with:
  - 43 entity tables + relationship catalog
  - 12 calculation engines (structured references)
  - Native UI (Dashboard / NAV / Search)
  - Lookup & organization seed data
  - Sheet protection + conditional formatting
- `release.py` production orchestrator (build → validate → auto-repair → export)
- `USER_MANUAL.md`
- `TEST_REPORT.md` (generated by release)
- Export package under `export/Inci_Aku_PPWR_Packaging_Management_System_Rev00/`

### Changed

- Project / workbook identity renamed to Packaging Management System Rev00
- `ENABLE_SEED_DATA` and `ENABLE_UI` enabled for production builds

### Fixed

- Lookup status codes made unique across domains
- Release validation auto-rebuild path for failed workbook checks

---

## [1.4.0] — 2026-08-02

### Added

- Native Excel UI (corporate dark-blue theme, no VBA)

---

## [1.3.0] — 2026-08-02

### Added

- Engine calculation sheets with structured-reference formulas

---

## [1.2.0] — 2026-08-02

### Added

- Structural workbook generation (tables, named ranges, FK validations)

---

## [1.1.0] — 2026-08-02

### Added

- Python architecture package

---

## [1.0.0] — 2026-08-02

### Added

- Frozen schema `FINAL_DATABASE.md` (43 tables)

---

## [0.2.0] — 2026-08-02

### Added

- Architecture review / Target Model

---

## [0.1.0] — 2026-08-02

### Added

- Initial architecture documentation set
