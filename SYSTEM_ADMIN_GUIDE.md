# System Administrator Guide — İnci Akü PPWR Packaging Management System

| Document | Version | Date |
|----------|---------|------|
| SYSTEM_ADMIN_GUIDE.md | 1.0.0 | 2026-08-02 |

**Audience:** IT / Application Administrators  
**Schema:** frozen `FINAL_DATABASE.md` 1.0.0 — do not redesign entities.

---

## 1. Production package

| Item | Location |
|------|----------|
| Workbook | `output/Inci_Aku_PPWR_Packaging_Management_System_Rev00.xlsx` |
| Release folder | `export/Inci_Aku_PPWR_Packaging_Management_System_Rev00/` |
| Import templates | `templates/import/Import_Templates.xlsx` |
| Document engine catalog | `templates/document_engine/` |
| Export package stub | `templates/export/shipments/SAMPLE/` |

Rebuild / validate:

```bash
python release.py
```

---

## 2. Sheet protection

| Sheet class | State | Password |
|-------------|-------|----------|
| Dashboard + 7 ops screens | Visible | `PIMS_UI` |
| All other sheets (masters, LKP, ENG, SYS, DOC_ENGINE_VARS) | Hidden | `PIMS_TECH` |

**Rules**

- Operators work only on the 8 visible sheets.  
- Never unhide `ENG_*` / `SYS_*` / `DOC_ENGINE_VARS` for daily users.  
- Yellow cells are unlocked for input; green/gray stay locked.  
- Change passwords only via `config.py` (`UI_SHEET_PASSWORD`, `TECH_SHEET_PASSWORD`) then rebuild.

---

## 3. Visible screens (fixed)

1. Dashboard (command center)  
2. Packaging Configuration (primary object)  
3. Component Master  
4. Product Master (ERP reference)  
5. Shipment  
6. Statements  
7. Technical File  
8. Declaration of Conformity  

---

## 4. Operational model (do not violate)

- Primary object = **Packaging Configuration**  
- Technical File belongs **only** to Packaging Configuration  
- Commercial Scenario selects DoC variant + shipment package only  
- Lot Number = `SHIPMENT.EXTERNAL_REF`  
- See `PPWR_OPERATIONAL_MODEL.md` and `USER_WORKFLOW.md`

---

## 5. Engines (hidden)

| Engine | Purpose |
|--------|---------|
| ENG_MATERIAL_WEIGHT / PACKAGING / TRANSPORT / SHIPMENT | Weights |
| ENG_*_SUMMARY | Material family totals |
| ENG_STATEMENT / TECHNICAL_FILE / DECLARATION | Completeness |
| ENG_VALIDATION | Continuous rules |
| DOC_ENGINE_VARS | Merge tokens for future Word templates |

Do not put formulas inside master Excel Tables.

---

## 6. Import / export / documents

| Guide | Topic |
|-------|-------|
| `IMPORT_GUIDE.md` | Bulk load via templates |
| `DOCUMENT_ENGINE_GUIDE.md` | Variable catalog for TF / DoC / Statement |
| `templates/export/` | Shipment package folder layout + manifest |

---

## 7. Validation

`release.py` runs architecture check + workbook validator.  
Key İnci rules: missing weight, TF ownership, shipment scenario+config, lot number warn, config↔TF coverage, scenario Incoterms.

---

## 8. Support contacts (fill in)

| Role | Name | Contact |
|------|------|---------|
| Packaging Engineering | | |
| Compliance | | |
| IT Admin | | |
| Export / Logistics | | |
