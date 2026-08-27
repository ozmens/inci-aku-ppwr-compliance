# PIMS Naming Convention (FROZEN)

| Document | Version | Status | Date |
|----------|---------|--------|------|
| NAMING_CONVENTION.md | 1.0.0 | **FROZEN** | 2026-08-02 |

**Authority:** Naming is frozen in `FINAL_DATABASE.md` §8. This file mirrors that standard for quick reference.  
On conflict, `FINAL_DATABASE.md` wins.

---

## 1. General Rules

| Rule | Convention |
|------|------------|
| Case | `UPPER_SNAKE_CASE` for sheet/table and column names |
| Language | Technical identifiers in English |
| Descriptions | May contain Turkish/English text |
| Spaces | Never in sheet/column names |
| Separator | Underscore `_` only |
| Prefixes | `LKP_`, `SYS_`, `VW_`, `RPT_` |
| Stability | Do not rename PK/FK after data load without schema version bump |

---

## 2. Core Entity Sheet Names (Frozen)

| Entity | Sheet Name |
|--------|------------|
| Component | `COMPONENT` |
| Component Material | `COMPONENT_MATERIAL` |
| Product | `PRODUCT` |
| Packaging Configuration | `PACKAGING_CONFIGURATION` |
| Packaging Configuration Line | `PACKAGING_CONFIGURATION_LINE` |
| Transport Configuration | `TRANSPORT_CONFIGURATION` |
| Transport Configuration Line | `TRANSPORT_CONFIGURATION_LINE` |
| Commercial Scenario | `COMMERCIAL_SCENARIO` |
| Shipment | `SHIPMENT` |
| Shipment Line | `SHIPMENT_LINE` |
| Statement | `STATEMENT` |
| Statement Shipment | `STATEMENT_SHIPMENT` |
| Statement Line | `STATEMENT_LINE` |
| Technical File | `TECHNICAL_FILE` |
| Declaration of Conformity | `DECLARATION_OF_CONFORMITY` |
| Document Library | `DOCUMENT_LIBRARY` |
| Document Link | `DOCUMENT_LINK` |

---

## 3. Primary Keys (Frozen)

| Table | PK |
|-------|----|
| COMPONENT | `COMPONENT_ID` |
| COMPONENT_MATERIAL | `COMPONENT_MATERIAL_ID` |
| PRODUCT | `PRODUCT_ID` |
| PACKAGING_CONFIGURATION | `PACKAGING_CONFIGURATION_ID` |
| PACKAGING_CONFIGURATION_LINE | `PACKAGING_CONFIGURATION_LINE_ID` |
| TRANSPORT_CONFIGURATION | `TRANSPORT_CONFIGURATION_ID` |
| TRANSPORT_CONFIGURATION_LINE | `TRANSPORT_CONFIGURATION_LINE_ID` |
| COMMERCIAL_SCENARIO | `COMMERCIAL_SCENARIO_ID` |
| SHIPMENT | `SHIPMENT_ID` |
| SHIPMENT_LINE | `SHIPMENT_LINE_ID` |
| STATEMENT | `STATEMENT_ID` |
| STATEMENT_SHIPMENT | `STATEMENT_SHIPMENT_ID` |
| STATEMENT_LINE | `STATEMENT_LINE_ID` |
| TECHNICAL_FILE | `TECHNICAL_FILE_ID` |
| DECLARATION_OF_CONFORMITY | `DECLARATION_OF_CONFORMITY_ID` |
| DOCUMENT_LIBRARY | `DOCUMENT_ID` |
| DOCUMENT_LINK | `DOCUMENT_LINK_ID` |

FK columns use the **identical** parent PK name (role prefixes only when two FKs point to the same parent).

---

## 4. Deprecated Names (Forbidden)

| Do not use | Use |
|------------|-----|
| `PACKAGING_CONFIG` | `PACKAGING_CONFIGURATION` |
| `LOADING_CONFIG` | `TRANSPORT_CONFIGURATION` |
| `PKG_CONFIG_LINE_ID` | `PACKAGING_CONFIGURATION_LINE_ID` |
| `SCENARIO_ID` | `COMMERCIAL_SCENARIO_ID` |
| `DOC_ID` | `DECLARATION_OF_CONFORMITY_ID` |
| `SHIPMENT_PACKAGING_LINE` | `SHIPMENT_LINE` |
| `COMPONENT_MATERIAL_SHARE` | `COMPONENT_MATERIAL` |
| `TECHNICAL_FILE_LINK` | `DOCUMENT_LIBRARY` / `DOCUMENT_LINK` |
| `PALLET_COMPONENT_ID` | transport line role `PALLET` |
| `DEST_COUNTRY_ID` | `DESTINATION_COUNTRY_ID` |

---

## 5. Measures, Flags, Dates

| Pattern | Example |
|---------|---------|
| `*_G` / `*_KG` / `*_MM` / `*_PCT` | `WEIGHT_G`, `TOTAL_WEIGHT_KG` |
| `IS_*` | `IS_OPTIONAL`, `IS_EDITABLE` |
| `*_AT` | `CREATED_AT`, `CONFIRMED_AT` |
| `EFFECTIVE_FROM` / `EFFECTIVE_TO` | Master/config validity |
| `VALID_FROM` / `VALID_TO` | Commercial scenario validity only |

---

## 6. Excel ListObjects (future build)

Pattern: `T_<TABLE_NAME>` → `T_COMPONENT`, `T_SHIPMENT_LINE`.

---

## 7. Quick Reference

```text
Sheet:     TRANSPORT_CONFIGURATION
PK:        TRANSPORT_CONFIGURATION_ID
Biz key:   CONFIG_GROUP_CODE + REVISION_NO
FK ex:     PACKAGING_CONFIGURATION_ID
Measure:   (derived; no stored total weight)
Lookup:    LKP_TRANSPORT_UNIT_TYPE
Freeze:    SHIPMENT_LINE / STATEMENT_LINE
```

Full standards: `FINAL_DATABASE.md` §8.
