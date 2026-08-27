# User Workflow — İnci Akü PPWR Daily Operations

| Document | Version | Status | Date |
|----------|---------|--------|------|
| USER_WORKFLOW.md | 1.0.0 | Binding (operations) | 2026-08-02 |

**Authority:** `PPWR_OPERATIONAL_MODEL.md`  
**Schema:** frozen (`FINAL_DATABASE.md`) — no new tables.

Every operation begins on **Dashboard (Command Center)**.  
Follow the blue **NEXT** banner. Yellow = you type/select. Green = automatic.

---

## Daily workflow (7 steps)

### Step 1 — Select Packaging Configuration

1. On Dashboard, open the yellow **Packaging Configuration** dropdown.  
2. Choose the physical arrangement (e.g. Starter 18/Euro, Industrial 1/pallet, 40HC loading).  
3. Green cell shows config code/name.  
4. System looks up the linked **Technical File** for that config.

You are managing packaging arrangements — not products.

---

### Step 2 — Select Commercial Scenario

1. Select **Commercial Scenario** (yellow dropdown).  
2. System shows Incoterms, customer, destination country, transport config (green).  
3. Scenario must fit the market (EXW / FCA / FOB / CIP / DAP / DDP, etc.).  

**Rule:** This step never changes the Technical File.  
It only selects the **DoC variant** and **shipment document package**.

---

### Step 3 — Enter Shipment Information

Minimum entry:

| You enter | Where |
|-----------|--------|
| Packaging Configuration | Already from Step 1 (draft) |
| Commercial Scenario | Already from Step 2 (brings Customer / Country / Incoterms) |
| Quantity | Yellow qty on Dashboard |
| Lot Number | Yellow lot on Dashboard → stored as `SHIPMENT.EXTERNAL_REF` |

1. Open **Shipment** via GO.  
2. Paste draft values from Dashboard.  
3. Confirm shipment.  
4. Paste **Shipment Number** back into Dashboard.

System then determines automatically:

- Technical File  
- Declaration variant  
- Statement readiness  
- Required documents  

---

### Step 4 — Generate Shipment Statement

1. GO → **Statements**.  
2. Link the shipment from Step 3.  
3. Composition, materials, weights, suppliers assemble from freeze lines.

---

### Step 5 — Generate Declaration of Conformity

1. GO → **Declaration of Conformity**.  
2. Use the variant implied by the Commercial Scenario (Incoterms / market).  
3. Link the **Technical File of the Packaging Configuration** (not a product TF).  

---

### Step 6 — Reference Technical File

1. Confirm Dashboard shows the TF code for the Packaging Configuration.  
2. If missing → open **Technical File**, set only `PACKAGING_CONFIGURATION_ID`, attach evidence/docs.  
3. Do not create a TF for a product.

---

### Step 7 — Archive Shipment

1. When package checklist = **READY**, export/print Statement + DoC + TF reference pack.  
2. Leave shipment status confirmed; do not edit frozen `SHIPMENT_LINE` rows.  
3. Clear or keep wizard selections for the next job.

---

## Master-data workflow (occasional)

Use when recipes or components change — not every shipment:

1. Maintain **Components** (weight, supplier, evidence).  
2. Maintain **Packaging Configuration** recipe lines.  
3. If components/recipe change → **new Technical File revision**.  
4. Link ERP **Products** only so scenarios can resolve; products are references.  
5. Maintain **Commercial Scenarios** for Incoterms / customers / markets.

---

## Cell colours

| Colour | Meaning |
|--------|---------|
| Yellow | User input |
| Blue | Guidance / calc helper |
| Green | Automatic output |
| Gray | Protected |

---

## Passwords

| Area | Password |
|------|----------|
| UI / ops sheets | `PIMS_UI` |
| Technical `SYS_*` | `PIMS_TECH` |
