# Import Guide — İnci Akü PPWR

| Document | Version | Date |
|----------|---------|------|
| IMPORT_GUIDE.md | 1.0.0 | 2026-08-02 |

**Template file:** `templates/import/Import_Templates.xlsx`  
**Regenerate:** `python release.py` (builds templates automatically)

---

## 1. Template sheets

| Sheet | Target table | Notes |
|-------|--------------|-------|
| Component_Master | `COMPONENT` | `WEIGHT_G` mandatory |
| Packaging_Configuration | `PACKAGING_CONFIGURATION` | Primary object |
| Packaging_Configuration_Lines | `PACKAGING_CONFIGURATION_LINE` | Recipe quantities |
| Product_Master | `PRODUCT` | ERP reference only |
| Commercial_Scenarios | `COMMERCIAL_SCENARIO` | Incoterms / customer / country |
| Shipment_Import | `SHIPMENT` | `EXTERNAL_REF` = **Lot Number** |
| Supplier_Documents | `DOCUMENT_LIBRARY` (+ link hints) | Then create `DOCUMENT_LINK` |

---

## 2. How to import (operators)

1. Open `Import_Templates.xlsx`.  
2. Fill rows under the dark-blue header (do not rename columns).  
3. Use numeric FK IDs from workbook lookups (`LKP_*`, suppliers, etc.).  
4. Open the production workbook.  
5. Paste values into the matching Excel Table (visible sheet or unhide master if admin).  
6. Dashboard validation KPIs / `ENG_VALIDATION` will flag issues.

**Recommended order**

1. Components  
2. Packaging Configuration + Lines  
3. Products (if needed)  
4. Commercial Scenarios (needs Product + Transport Configuration)  
5. Supplier Documents + DOCUMENT_LINK  
6. Shipments (after scenarios & configs exist)

Technical Files and DoCs are normally created in the workbook (not via these templates) so ownership rules stay clear.

---

## 3. Shipment import fields

| Column | Meaning |
|--------|---------|
| COMMERCIAL_SCENARIO_ID | Brings Customer, Country, Incoterms |
| PACKAGING_CONFIGURATION_ID | Physical arrangement |
| QTY_PRODUCT_UNITS | Quantity |
| EXTERNAL_REF | **Lot Number** |
| TRANSPORT_CONFIGURATION_ID | Usually from scenario |

---

## 4. Supplier documents

`Supplier_Documents` columns map to `DOCUMENT_LIBRARY`.  
Helper columns `LINK_*` are for your tracking — after paste into `DOCUMENT_LIBRARY`, create rows in `DOCUMENT_LINK` to attach documents to Component / Packaging Configuration / Technical File.

---

## 5. Rules

- Do not invent new columns.  
- Do not import packaging weights onto Product or Configuration headers.  
- Never attach a Technical File to a Product.  
- Keep IDs unique across the workbook.  
