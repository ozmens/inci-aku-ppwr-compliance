# Document Engine Guide — İnci Akü PPWR

| Document | Version | Date |
|----------|---------|------|
| DOCUMENT_ENGINE_GUIDE.md | 1.0.0 | 2026-08-02 |

**Purpose:** Expose merge variables for future Word templates.  
**This engine does NOT generate Word (.docx) files.**

---

## 1. What is delivered

| Artifact | Location |
|----------|----------|
| Variable catalog (JSON) | `templates/document_engine/document_variables.json` |
| Variable workbook | `templates/document_engine/Document_Engine_Variables.xlsx` |
| Hidden sheet in production workbook | `DOC_ENGINE_VARS` (password `PIMS_TECH`) |
| Python API | `src/engines/document_engine.py` |

---

## 2. Document types

| Type | When used | Owned by |
|------|-----------|----------|
| TECHNICAL_FILE | Compliance dossier | Packaging Configuration only |
| DECLARATION_OF_CONFORMITY | Legal variant | Scenario selects variant; TF from config |
| SHIPMENT_STATEMENT | Operational composition | Shipment (+ linked statement) |

---

## 3. Token convention

In a future Word template, use:

```text
{{TF.CODE}}
{{PC.NAME}}
{{SCN.INCOTERM_ID}}
{{SHP.LOT}}
```

Tokens are listed in `DOC_ENGINE_VARS` columns:

- `DOCUMENT_TYPE`  
- `TOKEN`  
- `SOURCE` (workbook table/field or derived engine)  
- `WORD_PLACEHOLDER`  

---

## 4. Technical File variables (examples)

| Token | Source idea |
|-------|-------------|
| `TF.CODE` / `TF.REVISION` | TECHNICAL_FILE |
| `PC.NAME` / `PC.DESCRIPTION` | PACKAGING_CONFIGURATION |
| `PC.COMPONENT_LIST` | Configuration lines |
| `PC.MATERIAL_SUMMARY` | Family summary engines |
| `PC.WEIGHT_TOTAL_G` | Packaging weight engine |
| `PC.SUPPLIER_MATRIX` / `PC.EVIDENCE_MATRIX` | DOCUMENT_LINK |
| `PC.DRAWINGS` / `PC.PHOTOGRAPHS` | Document types via library |

---

## 5. Declaration variables (examples)

| Token | Role |
|-------|------|
| `DOC.NUMBER` | DoC identity |
| `SCN.INCOTERM_ID` | Variant driver (EXW/FCA/FOB/CIP/DAP/DDP…) |
| `TF.CODE` | Supporting TF of Packaging Configuration |
| `PC.NAME` | Arrangement declared |

Scenario **never** rewrites Technical File content — it only selects which DoC variant travels.

---

## 6. Shipment Statement variables (examples)

| Token | Role |
|-------|------|
| `SHP.NUMBER` / `SHP.LOT` | Shipment + Lot (`EXTERNAL_REF`) |
| `STM.COMPOSITION` | Freeze lines |
| `STM.MATERIAL_BREAKDOWN` | ENG_STATEMENT |
| `SCN.CUSTOMER_ID` / `SCN.COUNTRY_ID` | From scenario |

---

## 7. Export engine (companion)

`templates/export/shipments/SAMPLE/manifest.json` describes the shipment package folder layout:

- `statement/`  
- `declaration/`  
- `technical_file/`  
- `evidence/`  

Future Word renderers should fill those folders using this variable catalog.

---

## 8. Python usage (IT)

```python
from engines.document_engine import DocumentEngine

eng = DocumentEngine()
ctx = eng.build_context_stub(
    "SHIPMENT_STATEMENT",
    packaging_configuration_id=1,
    commercial_scenario_id=1,
    shipment_number="SHP-001",
)
# Fill ctx[token] from workbook/DB, then merge into .docx later
```

Regenerate catalogs: `python release.py`
