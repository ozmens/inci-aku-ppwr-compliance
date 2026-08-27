# PPWR Operational Model — İnci Akü

| Document | Version | Status | Date |
|----------|---------|--------|------|
| PPWR_OPERATIONAL_MODEL.md | 1.0.0 | Binding (operations) | 2026-08-02 |

**Schema authority remains** `FINAL_DATABASE.md` **(frozen).**  
This document defines **how the business uses** that schema. It does not rename entities or add tables.

---

## 1. Core principle

The primary business object is **Packaging Configuration**.

A Packaging Configuration is one **physical packaging arrangement**, for example:

| Example | Meaning |
|---------|---------|
| Starter — 18 batteries on Euro pallet | Pallet recipe A |
| Starter — 24 batteries on Standard pallet | Pallet recipe B |
| Industrial — single battery on pallet | One battery / pallet |
| Container loading — 40 HC export | Loading materials arrangement |

**Products are not managed as the compliance object.**  
Products are only **linked** to Packaging Configurations (via commercial / transport setup in the frozen schema).

Every Packaging Configuration has **exactly one** Technical File (current effective revision; prior revisions retained for history).

---

## 2. Document hierarchy

```
Component Master
        ↓
Packaging Configuration          ← PRIMARY OBJECT
        ↓
Technical File                   ← owned by Packaging Configuration only
        ↓
Commercial Scenario (Incoterms)  ← never edits Technical File
        ↓
Declaration of Conformity Variant
        ↓
Shipment Statement
```

| Layer | Role |
|-------|------|
| Component Master | Physical packaging parts, weights, suppliers, evidence |
| Packaging Configuration | Recipe (BOM) of the packed unit / loading arrangement |
| Technical File | Compliance dossier for that arrangement |
| Commercial Scenario | Market / customer / **Incoterms** → DoC variant + shipment package |
| Declaration of Conformity | Legal statement variant for the scenario |
| Shipment Statement | Operational packaging composition for the shipment |

---

## 3. Technical File rules

1. A Technical File is **never** created for a Product.  
2. A Technical File belongs **only** to a Packaging Configuration (`PACKAGING_CONFIGURATION_ID` required; leave component/transport subject blank in daily ops).  
3. Technical Files are **version controlled** (`REVISION_NO`). Component or recipe changes → new revision.  
4. Commercial Scenarios **never** change Technical Files.

### Technical File content (assembled / linked)

| Content | How represented (frozen schema) |
|---------|----------------------------------|
| Packaging configuration description | TF title + config name/description |
| Packaging drawings | `DOCUMENT_LIBRARY` + `DOCUMENT_LINK` |
| Packaging photographs | `DOCUMENT_LIBRARY` + `DOCUMENT_LINK` |
| Component list | From `PACKAGING_CONFIGURATION_LINE` |
| Material composition | `COMPONENT_MATERIAL` + engines |
| Material summary | Family summary engines |
| Weight calculations | Component `WEIGHT_G` + packaging/shipment engines |
| Supplier document matrix | Supplier + document links |
| Evidence matrix | `DOCUMENT_LINK` on components / TF |
| Recyclability assessment | `RECYCLABILITY_SUMMARY` + linked docs |
| Packaging minimisation assessment | Linked assessment docs + notes |
| Empty space assessment | Linked assessment docs + notes |
| PFAS / restricted substances | `SUBSTANCE_OF_CONCERN_NOTES` + evidence |
| Labelling assessment | Linked docs + notes |
| Revision history | `REVISION_NO`, dates, `SUPERSEDES` on config |

---

## 4. Commercial Scenario rules

Commercial Scenario **only** determines:

1. **Declaration of Conformity variant**  
2. **Shipment document package** (which DoC / statement pack travels with the shipment)

Typical Incoterms variants: **EXW, FCA, FOB, CIP, DAP, DDP** (via `INCOTERM_ID` on the scenario).

It does **not** rewrite packaging recipes or Technical Files.

---

## 5. Shipment rules

### User enters (minimum)

| Field | Ops meaning | Schema mapping (no new columns) |
|-------|-------------|----------------------------------|
| Customer | Who receives | Via `COMMERCIAL_SCENARIO.CUSTOMER_ID` |
| Destination Country | Market | Via scenario (optional shipment override) |
| Incoterms | Trade term | Via `COMMERCIAL_SCENARIO.INCOTERM_ID` |
| Packaging Configuration | Physical arrangement | `SHIPMENT.PACKAGING_CONFIGURATION_ID` |
| Quantity | Product units shipped | `SHIPMENT.QTY_PRODUCT_UNITS` |
| Lot Number | Production / batch lot | `SHIPMENT.EXTERNAL_REF` (**Lot Number**) |

### System determines automatically

- Technical File (from Packaging Configuration)  
- Declaration of Conformity variant (from Commercial Scenario)  
- Statement content (from shipment freeze lines)  
- Required documents (links + package checklist)

---

## 6. Battery business modes (İnci Akü)

| Mode | Packaging Configuration pattern |
|------|----------------------------------|
| Starter batteries | Pallet-based multi-battery recipes |
| Industrial batteries | Typically one battery per pallet |
| Container loading materials | Managed as components / transport lines; own configs as needed |

---

## 7. Non-goals

- No entity renames  
- No relationship redesign  
- No new tables for Lot Number, assessments, or document matrices  
- Product remains an ERP reference, not the PPWR primary object  
