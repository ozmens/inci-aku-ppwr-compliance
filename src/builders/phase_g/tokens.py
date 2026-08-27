"""Centralized merge token names for Phase G Golden Word integration."""

from __future__ import annotations

# Logical tokens (without braces)
TOKEN_NAMES = (
    "CONFIG_SET_CODE",
    "CONFIG_ID",
    "SOURCE_CONFIG_ID",
    "VARIANT_BASIS_TR",
    "VARIANT_BASIS_EN",
    "VARIANT_BASIS_PAIR",  # TR / EN on one line for compact cells
    "PACKAGING_DESCRIPTION_TR",
    "PACKAGING_DESCRIPTION_EN",
    "PACKAGING_DESCRIPTION_PAIR",
    "NOMINAL_LOAD_TR",
    "NOMINAL_LOAD_EN",
    "NOMINAL_LOAD_PAIR",
    "PRODUCT_QTY_PAIR",
    "TOTAL_TARE_KG",
    "TOTAL_TARE_KG_PLAIN",
    "TF_ID",
    "DOC_ID",
    "LABEL_ID",
    "STM_ID",
    "REVISION",
    "REVISION_DATE",
    "REVISION_PAIR",
    "LEGAL_NAME",
    "LEGAL_ADDRESS",
    "LEGAL_EMAIL",
    "LEGAL_WEBSITE",
    "LEGAL_PHONE",
    "ARTICLE5_BASIS",
    "ANNEX_DRAWINGS_STATUS",
    "TRACEABILITY_KEY",
    "PRODUCT_LINES",  # multi-line product list for simple cells
    "SHIPMENT_NO",
    "SHIPMENT_DATE",
    "CUSTOMER_OEM",
    "DESTINATION",
    "INCOTERM",
    "PACKING_LIST_REF",
    "PACKAGING_LOT",
)

# Dynamic table markers
BOM_TABLE_MARKER = "{{BOM_TABLE}}"
PRODUCT_TABLE_MARKER = "{{PRODUCT_TABLE}}"
MATERIAL_SUMMARY_MARKER = "{{MATERIAL_SUMMARY_TABLE}}"


def token(name: str) -> str:
    if name not in TOKEN_NAMES and name not in {
        "BOM_CODE",
        "BOM_DESC",
        "BOM_QTY",
        "BOM_UOM",
        "BOM_UNIT_WT",
        "BOM_LINE_WT",
        "PRODUCT_CODE",
        "PRODUCT_DESC",
    }:
        raise KeyError(name)
    return "{{" + name + "}}"


# ST-012-EUR-01 sample literals → tokens (longest first)
# NOTE: YS/D/0020–0023 QMS type numbers are FIXED — do not tokenize.
SAMPLE_LITERAL_MAP: list[tuple[str, str]] = [
    ("IA-PPWR-TF-ST-012-EUR-01-R00", "{{TF_ID}}"),
    ("IA-PPWR-DOC-ST-012-EUR-01-R00", "{{DOC_ID}}"),
    ("IA-PPWR-LBL-ST-012-EUR-01-R00", "{{LABEL_ID}}"),
    ("IA-PPWR-STM-ST-012-EUR-01-R00", "{{STM_ID}}"),
    ("IA-ST-012-EUR-01", "{{CONFIG_ID}}"),
    ("IA-ST-CFG-0004", "{{SOURCE_CONFIG_ID}}"),
    ("ST-012-EUR-01", "{{CONFIG_SET_CODE}}"),
    ("Rev.00 / 08.08.2026", "{{REVISION_PAIR}}"),
    (
        "Starter 12-unit | Euro Pallet 800×1200×152 mm | Variant 01",
        "{{PACKAGING_DESCRIPTION_EN}}",
    ),
    (
        "Starter 12’li - Euro Palet 800×1200×152 mm - Varyant 01",
        "{{PACKAGING_DESCRIPTION_TR}}",
    ),
    (
        "Starter 12'li - Euro Palet 800×1200×152 mm - Varyant 01",
        "{{PACKAGING_DESCRIPTION_TR}}",
    ),
    (
        "Starter 12-unit - Euro Pallet 800×1200×152 mm - Variant 01",
        "{{PACKAGING_DESCRIPTION_EN}}",
    ),
    (
        "C | Palet (4000038): 1 adet | Köşebent (4000590): 4 adet",
        "{{VARIANT_BASIS_TR}}",
    ),
    (
        "C | Pallet (4000038): 1 pc | Edge protector (4000590): 4 pcs",
        "{{VARIANT_BASIS_EN}}",
    ),
    ("12 starter akü / palet", "{{NOMINAL_LOAD_TR}}"),
    ("12 starter batteries / pallet", "{{NOMINAL_LOAD_EN}}"),
    ("12 adet", "{{PRODUCT_QTY_PAIR}}"),  # careful - may over-replace; scoped in cells
    ("12 pcs", "{{PRODUCT_QTY_PAIR}}"),
    ("29.3260 kg", "{{TOTAL_TARE_KG}}"),
    ("29.3260", "{{TOTAL_TARE_KG_PLAIN}}"),
    ("1015169 • C • 220 Ah • AGM", "{{PRODUCT_LINES}}"),
    ("1015169", "{{PRODUCT_LINES}}"),
]

FORBIDDEN_SAMPLE_LEAKS = (
    "ST-012-EUR-01",
    "IA-ST-012-EUR-01",
    "IA-ST-CFG-0004",
    "1015169",
    "29.3260",
)

FORBIDDEN_CONTENT = (
    "OPEN",
    "BLOCKED",
    "PRE-RELEASE",
    "NOT RELEASED",
    "CANDIDATE",
    "PENDING DOCUMENT IMPORT",
    "SELECT FROM PIMS",
    "DEPENDS ON VARIANT",
    "ANKA",
    "CMS",
    "RED BULL",
    "TOPRAK",
)

ALLOWED_PENDING = "PENDING – DRAWINGS / PHOTOGRAPHS"
ALLOWED_PENDING_ALT = (
    "PENDING - DRAWINGS / PHOTOGRAPHS",
    "PENDING - DRAWING",
    "PENDING - PHOTOGRAPHS",
)

GOLDEN_FILES = {
    "TECHNICAL_FILE": "01_Technical_File_GOLDEN.docx",
    "DOC": "02_EU_DoC_GOLDEN.docx",
    "LABEL": "03_Label_GOLDEN.docx",
    "STATEMENT": "04_Shipment_Statement_GOLDEN.docx",
}

RUNTIME_FILES = {
    "TECHNICAL_FILE": "01_Technical_File_RUNTIME.docx",
    "DOC": "02_EU_DoC_RUNTIME.docx",
    "LABEL": "03_Label_RUNTIME.docx",
    "STATEMENT": "04_Shipment_Statement_RUNTIME.docx",
}
