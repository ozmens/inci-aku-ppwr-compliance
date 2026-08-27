# Implementation Changes — Phase 11 PPWR Operational Model

| Document | Version | Date |
|----------|---------|------|
| IMPLEMENTATION_CHANGES.md | 1.0.0 | 2026-08-02 |

## Scope

| Allowed | Not allowed |
|---------|-------------|
| Workflow / Dashboard wizard | New entities |
| Navigation & UX copy | Relationship redesign |
| Validation rules (Excel engines) | Renaming tables/columns |
| Documentation | Generic PIMS redesign |
| Mapping Lot Number → existing `EXTERNAL_REF` | New Lot table |

Schema remains **1.0.0** / `FINAL_DATABASE.md`.

---

## Documents added

| File | Purpose |
|------|---------|
| `PPWR_OPERATIONAL_MODEL.md` | Binding ops model (primary object, hierarchy, TF/scenario/shipment rules) |
| `USER_WORKFLOW.md` | Daily 7-step operator guide |
| `IMPLEMENTATION_CHANGES.md` | This change log for Phase 11 |

---

## Workbook changes

### Dashboard (Command Center)

- Replaced product-first 8-step wizard with **PPWR 7-step** wizard:
  1. Packaging Configuration  
  2. Commercial Scenario (Incoterms / customer / country)  
  3. Shipment info (Qty + **Lot Number** + Shipment Number)  
  4. Shipment Statement  
  5. Declaration of Conformity  
  6. Technical File reference  
  7. Archive shipment  
- **NEXT** banner rewritten for this sequence.  
- Product ID removed as a wizard input; product resolves from scenario (green).  
- Lot Number yellow input → copy to `SHIPMENT.EXTERNAL_REF`.  
- Auto package checklist: TF ← config, DoC ← scenario, Statement ← shipment.

### Screen tips & panels

- Product Master: ERP link only; not primary.  
- Packaging Configuration: primary object + TF ownership.  
- Technical File: content checklist aligned to PPWR dossier (docs via `DOCUMENT_LINK`).  
- Shipment: customer/country/incoterms via scenario; lot via `EXTERNAL_REF`.  
- DoC: variant from scenario; TF from packaging configuration.

### Validations

| Rule | Meaning |
|------|---------|
| `V-TF-INCI-01` | Technical File must have `PACKAGING_CONFIGURATION_ID` |
| `V-SHP-INCI-01` | Shipment must have scenario + packaging configuration |
| `V-SHP-INCI-02` | Warn if Lot Number (`EXTERNAL_REF`) blank on shipments |

### Navigation

- Visible sheet order unchanged in spirit: Dashboard → Packaging Configuration first among masters.  
- HOME / Dashboard / BACK → Dashboard NEXT banner.  
- GO buttons follow the 7-step path.

---

## Explicit non-changes

- No new Excel tables for assessments, lot master, or document matrices.  
- Drawings/photos/evidence = `DOCUMENT_LIBRARY` + `DOCUMENT_LINK`.  
- Incoterms remain `LKP_INCOTERM` via Commercial Scenario.  
- Product table retained for ERP reference / scenario FK only.

---

## Operator impact

| Before (Phase 10) | After (Phase 11) |
|-------------------|------------------|
| Start with Product | Start with Packaging Configuration |
| 8 steps incl. product pick | 7 steps; product automatic from scenario |
| Lot not guided | Lot Number required in wizard → `EXTERNAL_REF` |
| TF step near end of package | TF is reference to config-owned dossier |
