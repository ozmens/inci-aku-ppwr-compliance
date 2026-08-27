# User Manual  
## İnci Akü PPWR Packaging Management System (Rev00)

| Item | Value |
|------|-------|
| Product | Inci_Aku_PPWR_Packaging_Management_System_Rev00 |
| Workbook | `Inci_Aku_PPWR_Packaging_Management_System_Rev00.xlsx` |
| Schema | 1.0.0 (frozen) |
| Platform | Microsoft Excel (Microsoft 365 recommended) |
| Automation | Python 3.12 generators (no VBA in workbook) |

---

## 1. Purpose

This system is the company packaging database for PPWR compliance. It manages:

- Packaging components and multi-material composition  
- Packaging and transport configurations  
- Commercial scenarios  
- Shipments and frozen shipment packaging lines  
- Statements, Technical Files, Declarations of Conformity  
- Documents and evidence links  

---

## 2. Getting Started

1. Open the workbook — you land on **Dashboard (PPWR Command Center)**.  
2. Follow the **NEXT** banner (7-step PPWR daily workflow).  
3. Yellow = you select/enter · Green = automatic.  
4. Use **GO** only when the step needs that sheet.  
5. Full operator guide: [`USER_WORKFLOW.md`](USER_WORKFLOW.md) · model: [`PPWR_OPERATIONAL_MODEL.md`](PPWR_OPERATIONAL_MODEL.md)

### Daily Workflow (always start on Dashboard)

| Step | You do | System does |
|------|--------|-------------|
| 1 | Select **Packaging Configuration** | Resolves arrangement; finds Tech File |
| 2 | Select **Commercial Scenario** | Customer / Country / Incoterms |
| 3 | Enter Qty + **Lot Number** → create Shipment | Draft values; auto TF / DoC / docs |
| 4 | Generate Shipment Statement | Composition from freeze lines |
| 5 | Generate Declaration of Conformity | Variant from scenario |
| 6 | Reference Technical File | Owned by Packaging Configuration only |
| 7 | Archive Shipment | Export when checklist = READY |

**Lot Number** is entered on the wizard and stored in `SHIPMENT.EXTERNAL_REF` (no new database column).

### Visible screens only (8)

1. Dashboard  
2. Packaging Configuration (primary object — header + recipe lines)  
3. Component Master  
4. Product Master (ERP reference only)  
5. Shipment  
6. Statements  
7. Technical File  
8. Declaration of Conformity  

Every screen has **HOME / Dashboard / BACK** buttons.

### Passwords

| Area | Password | Notes |
|------|----------|-------|
| UI / data sheets | `PIMS_UI` | Unlocks structure changes; yellow cells are already editable |
| Technical `SYS_*` sheets | `PIMS_TECH` | Hidden system sheets — do not edit casually |

---

## 3. Cell Legend

| Color | Meaning |
|-------|---------|
| Dark blue header | Column titles (protected) |
| Yellow | User input |
| Blue | Calculation guidance / helpers |
| Green | Automatic outputs / KPIs |
| Gray | Protected / system values |

---

## 4. Recommended Data Entry Order

1. Review lookups (`LKP_*`) — pre-seeded for production  
2. Organization: `LEGAL_ENTITY`, `PERSON`, `SUPPLIER`, `CUSTOMER`, `PLANT`  
3. `COMPONENT` (enter **WEIGHT_G** here only) + `COMPONENT_MATERIAL`  
4. **`PACKAGING_CONFIGURATION` + `_LINE`** ← primary master you manage daily  
5. `TRANSPORT_CONFIGURATION` + `_LINE` (pallet/container materials as lines)  
6. `TECHNICAL_FILE` — set **only** `PACKAGING_CONFIGURATION_ID` (1 TF ↔ 1 config)  
7. `PRODUCT` (ERP code/name reference — not the managed object)  
8. `COMMERCIAL_SCENARIO` — market/customer/incoterms + which DoC/document package variant  
9. `DECLARATION_OF_CONFORMITY` — variant for scenario; links the config’s Technical File  
10. `SHIPMENT` — select **Product + Packaging Configuration + Commercial Scenario** (+ qty)  
11. Confirm → `SHIPMENT_LINE` freeze; system assembles the shipment document package  
12. `STATEMENT` / `STATEMENT_SHIPMENT` / `STATEMENT_LINE` as needed

---

## 5. Engines (read-only)

| Sheet | What it calculates |
|-------|--------------------|
| ENG_MATERIAL_WEIGHT | Material share weights from components |
| ENG_PACKAGING_WEIGHT | Packaging BOM line weights |
| ENG_TRANSPORT_WEIGHT | Transport-unit line weights |
| ENG_SHIPMENT_WEIGHT | Shipment packaging totals |
| ENG_PLASTIC/PAPER/WOOD_SUMMARY | Family totals |
| ENG_STATEMENT | Statement reconcile vs shipment lines |
| ENG_TECHNICAL_FILE | Technical file completeness |
| ENG_DECLARATION | DoC completeness |
| ENG_IMPACT_ANALYSIS | Component change impact |
| ENG_VALIDATION | Rule results (OK / ERROR / WARN / PENDING) |

---

## 6. Critical Business Rules

See `PPWR_OPERATIONAL_MODEL.md` for the full model. Summary:

1. **Packaging Configuration** is the primary object (not Product).  
2. Hierarchy: Component → Packaging Config → Technical File → Scenario (Incoterms) → DoC → Statement.  
3. Technical File is **never** for a Product; version when components change.  
4. Commercial Scenario only selects DoC variant + shipment package.  
5. Shipment minimum: Customer/Country/Incoterms (via scenario) + Packaging Config + Qty + Lot Number.  
6. Component owns `WEIGHT_G`; confirmed `SHIPMENT_LINE` is immutable.

---

## 7. Rebuild from Python (IT / Admin)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python build.py --check
python release.py
```

Outputs:

- `output/Inci_Aku_PPWR_Packaging_Management_System_Rev00.xlsx`  
- `export/Inci_Aku_PPWR_Packaging_Management_System_Rev00/`  

---

## 8. Support Roles

| Role | Responsibilities |
|------|------------------|
| Packaging Engineering | Components, configs |
| Sales Ops | Commercial scenarios |
| Logistics | Shipments |
| Compliance | Statements, Tech Files, DoCs |
| IT / Architecture | Rebuild, schema changes |

---

## 9. Versioning

- Workbook revision: **Rev00**  
- Schema: **1.0.0**  
- Configuration changes after shipment confirm require a **new revision**, not silent edits.  
