# Excel Recovery — Root Cause Analysis

- **Timestamp:** 2026-08-10T19:29:54.429206+00:00
- **Corrupt file:** `output/INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION.xlsx`
- **Phase E blank:** `output/INCI_AKU_PPWR_PIMS_Rev00.xlsx`

## Native Microsoft Excel status (before rebuild)

Both the Phase E blank template and the production workbook **fail** `Workbooks.Open`
with Excel error (Turkish UI): workbook cannot be opened / repaired.

openpyxl can still read both files (49 sheets). ZIP integrity (`testzip`) is clean.
Therefore the defect is **structural OOXML**, not a truncated ZIP.

## Confirmed technical cause

**Worksheet-level `autoFilter` written together with an Excel Table (ListObject).**

Reproduction (isolated):

| Construct | Excel open |
|-----------|------------|
| Headers only | PASS |
| Styled cells | PASS |
| Excel Table only | PASS |
| Excel Table + `ws.auto_filter.ref = …` | **FAIL** |
| `write_database_sheet` (had both) | **FAIL** |

Phase E `db_sheets.write_database_sheet` / `_write_populated` and Phase F
`promoter._replace_table_rows` set `ws.auto_filter.ref` **after** `ws.add_table(...)`.

Excel Tables already embed their own `<autoFilter>` inside `xl/tables/tableN.xml`.
Writing a second AutoFilter on the worksheet produces invalid dual AutoFilter markup.
Microsoft Excel then refuses to open the package (no usable repair path).

OOXML dual-AutoFilter sheet count (pre-fix artifacts):

- Phase E blank: **43** sheets with tablePart + worksheet autoFilter
- Production: **43** sheets with tablePart + worksheet autoFilter

## Non-causes (checked)

- ZIP CRC / truncated package: OK (`testzip` None)
- Missing `[Content_Types].xml` / theme / styles: present
- `#REF!` defined names: 0
- Duplicate Excel table displayNames: not observed as the open-blocker
- Absolute relationship Targets (`/xl/...`): also present in tiny workbooks that **do** open
- External links: none
- Shared strings missing: not required (inline strings); not the blocker

## Obsolete prototype

`Inci_Aku_PPWR_Packaging_Management_System_Rev00.xlsx` (if present under export/legacy)
is **OBSOLETE / NOT FOR DELIVERY** and must not be used as production PIMS.

## Fix applied

1. Remove worksheet `auto_filter.ref` wherever an Excel Table is created/resized.
2. In promoter: delete ListObject **before** `delete_rows`, then recreate table after data write.
3. Rebuild Phase E blank from Schema 1.0.0.
4. Re-promote Phase F normalized data into a new Excel-safe workbook.
5. Validate with native Microsoft Excel COM (open + Save As round-trip).
