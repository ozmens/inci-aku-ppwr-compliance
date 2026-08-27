"""
Frozen table registry content for schema 1.0.0.

Column lists are architectural contracts for future Excel/SQL builders.
Authority: FINAL_DATABASE.md
"""

from __future__ import annotations

from .table_definition import ColumnDefinition, TableDefinition


def _col(
    name: str,
    data_type: str,
    *,
    required: bool = True,
    pk: bool = False,
    fk_table: str | None = None,
    description: str = "",
) -> ColumnDefinition:
    return ColumnDefinition(
        name=name,
        data_type=data_type,
        required=required,
        is_pk=pk,
        is_fk=fk_table is not None,
        fk_table=fk_table,
        description=description,
    )


def _table(
    name: str,
    pk: str,
    group: str,
    purpose: str,
    owner: str,
    frequency: str,
    columns: list[ColumnDefinition],
) -> TableDefinition:
    fks = tuple(c.name for c in columns if c.is_fk)
    return TableDefinition(
        name=name,
        primary_key=pk,
        group=group,
        business_purpose=purpose,
        data_owner=owner,
        update_frequency=frequency,
        columns=tuple(columns),
        foreign_keys=fks,
    )


# ---------------------------------------------------------------------------
# System
# ---------------------------------------------------------------------------

SYS_WORKBOOK_INFO = _table(
    "SYS_WORKBOOK_INFO",
    "INFO_KEY",
    "SYSTEM",
    "Schema version and workbook metadata",
    "IT / Solution Architecture",
    "On schema release only",
    [
        _col("INFO_KEY", "TEXT", pk=True),
        _col("INFO_VALUE", "TEXT"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

SYS_PARAMETER = _table(
    "SYS_PARAMETER",
    "PARAMETER_ID",
    "SYSTEM",
    "System behavior parameters",
    "Compliance + Architecture",
    "Rare",
    [
        _col("PARAMETER_ID", "INT", pk=True),
        _col("PARAMETER_CODE", "TEXT"),
        _col("PARAMETER_VALUE", "TEXT"),
        _col("DESCRIPTION", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Lookups (abbreviated columns; codes + names standard)
# ---------------------------------------------------------------------------


def _lkp(
    name: str,
    pk: str,
    purpose: str,
    extra: list[ColumnDefinition] | None = None,
) -> TableDefinition:
    code = pk.replace("_ID", "_CODE")
    display = pk.replace("_ID", "_NAME")
    cols = [
        _col(pk, "INT", pk=True),
        _col(code, "TEXT"),
        _col(display, "TEXT"),
    ]
    if extra:
        cols.extend(extra)
    return _table(name, pk, "LOOKUP", purpose, "MDM + Compliance", "Rare", cols)


LKP_STATUS = _table(
    "LKP_STATUS",
    "STATUS_ID",
    "LOOKUP",
    "Lifecycle/approval statuses by domain",
    "Architecture",
    "Rare",
    [
        _col("STATUS_ID", "INT", pk=True),
        _col("STATUS_CODE", "TEXT"),
        _col("STATUS_NAME", "TEXT"),
        _col("STATUS_DOMAIN", "TEXT"),
        _col("IS_EDITABLE", "BOOL"),
        _col("SORT_ORDER", "INT"),
    ],
)

LKP_UOM = _lkp("LKP_UOM", "UOM_ID", "Units of measure", [_col("UOM_DIMENSION", "TEXT")])
LKP_PACKAGING_LEVEL = _lkp(
    "LKP_PACKAGING_LEVEL",
    "PACKAGING_LEVEL_ID",
    "Primary / Secondary / Tertiary",
    [_col("SORT_ORDER", "INT")],
)
LKP_PACKAGING_FUNCTION = _lkp(
    "LKP_PACKAGING_FUNCTION", "PACKAGING_FUNCTION_ID", "Sales / Grouped / Transport"
)
LKP_COMPONENT_TYPE = _table(
    "LKP_COMPONENT_TYPE",
    "COMPONENT_TYPE_ID",
    "LOOKUP",
    "Component type vocabulary incl. container materials",
    "MDM + Compliance",
    "Rare",
    [
        _col("COMPONENT_TYPE_ID", "INT", pk=True),
        _col("COMPONENT_TYPE_CODE", "TEXT"),
        _col("COMPONENT_TYPE_NAME", "TEXT"),
        _col("DEFAULT_PACKAGING_LEVEL_ID", "INT", required=False, fk_table="LKP_PACKAGING_LEVEL"),
    ],
)
LKP_MATERIAL_FAMILY = _lkp("LKP_MATERIAL_FAMILY", "MATERIAL_FAMILY_ID", "Material families")
LKP_PPWR_MATERIAL_CATEGORY = _lkp(
    "LKP_PPWR_MATERIAL_CATEGORY", "PPWR_MATERIAL_CATEGORY_ID", "PPWR material categories"
)
LKP_MATERIAL = _table(
    "LKP_MATERIAL",
    "MATERIAL_ID",
    "LOOKUP",
    "Controlled material vocabulary",
    "MDM + Compliance",
    "Rare",
    [
        _col("MATERIAL_ID", "INT", pk=True),
        _col("MATERIAL_CODE", "TEXT"),
        _col("MATERIAL_NAME", "TEXT"),
        _col("MATERIAL_FAMILY_ID", "INT", fk_table="LKP_MATERIAL_FAMILY"),
        _col("PPWR_MATERIAL_CATEGORY_ID", "INT", required=False, fk_table="LKP_PPWR_MATERIAL_CATEGORY"),
        _col("IS_COMPOSITE", "BOOL"),
        _col("NOTES", "TEXT", required=False),
    ],
)
LKP_RECYCLABILITY_CLASS = _lkp(
    "LKP_RECYCLABILITY_CLASS", "RECYCLABILITY_CLASS_ID", "Recyclability class"
)
LKP_OWNERSHIP_TYPE = _lkp(
    "LKP_OWNERSHIP_TYPE",
    "OWNERSHIP_TYPE_ID",
    "Disposable / returnable / pool (returnable architecture hook)",
)
LKP_TRANSPORT_UNIT_TYPE = _lkp(
    "LKP_TRANSPORT_UNIT_TYPE", "TRANSPORT_UNIT_TYPE_ID", "Pallet / container / truck / …"
)
LKP_LINE_ROLE = _lkp("LKP_LINE_ROLE", "LINE_ROLE_ID", "Configuration line roles")
LKP_PRODUCT_CATEGORY = _lkp(
    "LKP_PRODUCT_CATEGORY", "PRODUCT_CATEGORY_ID", "Starter / industrial battery category"
)
LKP_COUNTRY = _table(
    "LKP_COUNTRY",
    "COUNTRY_ID",
    "LOOKUP",
    "ISO countries + EU market flag",
    "MDM + Compliance",
    "Rare",
    [
        _col("COUNTRY_ID", "INT", pk=True),
        _col("ISO2", "TEXT"),
        _col("ISO3", "TEXT"),
        _col("COUNTRY_NAME", "TEXT"),
        _col("IS_EU_MARKET", "BOOL"),
    ],
)
LKP_TRANSPORT_MODE = _lkp("LKP_TRANSPORT_MODE", "TRANSPORT_MODE_ID", "Transport modes")
LKP_INCOTERM = _lkp("LKP_INCOTERM", "INCOTERM_ID", "Incoterms")
LKP_STATEMENT_TYPE = _lkp("LKP_STATEMENT_TYPE", "STATEMENT_TYPE_ID", "Statement types")
LKP_DOCUMENT_TYPE = _lkp("LKP_DOCUMENT_TYPE", "DOCUMENT_TYPE_ID", "Document types")
LKP_SCENARIO_TYPE = _lkp("LKP_SCENARIO_TYPE", "SCENARIO_TYPE_ID", "Commercial scenario types")

# ---------------------------------------------------------------------------
# Organization
# ---------------------------------------------------------------------------

LEGAL_ENTITY = _table(
    "LEGAL_ENTITY",
    "LEGAL_ENTITY_ID",
    "ORGANIZATION",
    "Company / economic operator issuing DoCs",
    "Finance / Legal / MDM",
    "Rare",
    [
        _col("LEGAL_ENTITY_ID", "INT", pk=True),
        _col("LEGAL_ENTITY_CODE", "TEXT"),
        _col("LEGAL_ENTITY_NAME", "TEXT"),
        _col("COUNTRY_ID", "INT", fk_table="LKP_COUNTRY"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("NOTES", "TEXT", required=False),
    ],
)

PERSON = _table(
    "PERSON",
    "PERSON_ID",
    "ORGANIZATION",
    "Approvers and DoC responsible persons",
    "HR / Compliance",
    "Occasional",
    [
        _col("PERSON_ID", "INT", pk=True),
        _col("PERSON_CODE", "TEXT"),
        _col("FULL_NAME", "TEXT"),
        _col("EMAIL", "TEXT", required=False),
        _col("JOB_TITLE", "TEXT", required=False),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
    ],
)

SUPPLIER = _table(
    "SUPPLIER",
    "SUPPLIER_ID",
    "ORGANIZATION",
    "Packaging suppliers",
    "Procurement",
    "Occasional",
    [
        _col("SUPPLIER_ID", "INT", pk=True),
        _col("SUPPLIER_CODE", "TEXT"),
        _col("SUPPLIER_NAME", "TEXT"),
        _col("COUNTRY_ID", "INT", required=False, fk_table="LKP_COUNTRY"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EXTERNAL_REF", "TEXT", required=False),
        _col("NOTES", "TEXT", required=False),
    ],
)

CUSTOMER = _table(
    "CUSTOMER",
    "CUSTOMER_ID",
    "ORGANIZATION",
    "Sold-to party for commercial scenarios",
    "Sales MDM",
    "Occasional",
    [
        _col("CUSTOMER_ID", "INT", pk=True),
        _col("CUSTOMER_CODE", "TEXT"),
        _col("CUSTOMER_NAME", "TEXT"),
        _col("COUNTRY_ID", "INT", required=False, fk_table="LKP_COUNTRY"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EXTERNAL_REF", "TEXT", required=False),
        _col("NOTES", "TEXT", required=False),
    ],
)

PLANT = _table(
    "PLANT",
    "PLANT_ID",
    "ORGANIZATION",
    "Ship-from / producing location",
    "Operations MDM",
    "Rare",
    [
        _col("PLANT_ID", "INT", pk=True),
        _col("PLANT_CODE", "TEXT"),
        _col("PLANT_NAME", "TEXT"),
        _col("LEGAL_ENTITY_ID", "INT", fk_table="LEGAL_ENTITY"),
        _col("COUNTRY_ID", "INT", fk_table="LKP_COUNTRY"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EXTERNAL_REF", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Masters
# ---------------------------------------------------------------------------

PRODUCT = _table(
    "PRODUCT",
    "PRODUCT_ID",
    "MASTER",
    "Battery SKU master; product net weight only",
    "Product MDM / Engineering",
    "Occasional",
    [
        _col("PRODUCT_ID", "INT", pk=True),
        _col("PRODUCT_CODE", "TEXT"),
        _col("PRODUCT_NAME", "TEXT"),
        _col("PRODUCT_CATEGORY_ID", "INT", fk_table="LKP_PRODUCT_CATEGORY"),
        _col("NET_WEIGHT_G", "DECIMAL", description="Battery net weight, not packaging"),
        _col("LENGTH_MM", "DECIMAL", required=False),
        _col("WIDTH_MM", "DECIMAL", required=False),
        _col("HEIGHT_MM", "DECIMAL", required=False),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EFFECTIVE_FROM", "DATE"),
        _col("EFFECTIVE_TO", "DATE", required=False),
        _col("EXTERNAL_REF", "TEXT", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

COMPONENT = _table(
    "COMPONENT",
    "COMPONENT_ID",
    "MASTER",
    "Atomic packaging item; owns unit weight",
    "Packaging Engineering",
    "Frequent",
    [
        _col("COMPONENT_ID", "INT", pk=True),
        _col("COMPONENT_CODE", "TEXT"),
        _col("COMPONENT_NAME", "TEXT"),
        _col("COMPONENT_TYPE_ID", "INT", fk_table="LKP_COMPONENT_TYPE"),
        _col("PACKAGING_LEVEL_ID", "INT", fk_table="LKP_PACKAGING_LEVEL"),
        _col("PACKAGING_FUNCTION_ID", "INT", fk_table="LKP_PACKAGING_FUNCTION"),
        _col("OWNERSHIP_TYPE_ID", "INT", fk_table="LKP_OWNERSHIP_TYPE"),
        _col("SUPPLIER_ID", "INT", required=False, fk_table="SUPPLIER"),
        _col("WEIGHT_G", "DECIMAL", description="Unit weight source of truth"),
        _col("LENGTH_MM", "DECIMAL", required=False),
        _col("WIDTH_MM", "DECIMAL", required=False),
        _col("HEIGHT_MM", "DECIMAL", required=False),
        _col("RECYCLED_CONTENT_PCT", "DECIMAL", required=False),
        _col("RECYCLABILITY_CLASS_ID", "INT", required=False, fk_table="LKP_RECYCLABILITY_CLASS"),
        _col("REUSE_CYCLE_TARGET", "INT", required=False),
        _col("SPEC_REF", "TEXT", required=False),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EFFECTIVE_FROM", "DATE"),
        _col("EFFECTIVE_TO", "DATE", required=False),
        _col("EXTERNAL_REF", "TEXT", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

COMPONENT_MATERIAL = _table(
    "COMPONENT_MATERIAL",
    "COMPONENT_MATERIAL_ID",
    "MASTER",
    "Multi-material composition shares for PPWR reporting",
    "Packaging Engineering",
    "Occasional",
    [
        _col("COMPONENT_MATERIAL_ID", "INT", pk=True),
        _col("COMPONENT_ID", "INT", fk_table="COMPONENT"),
        _col("MATERIAL_ID", "INT", fk_table="LKP_MATERIAL"),
        _col("SHARE_PCT", "DECIMAL"),
        _col("SORT_ORDER", "INT"),
        _col("NOTES", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Configurations
# ---------------------------------------------------------------------------

PACKAGING_CONFIGURATION = _table(
    "PACKAGING_CONFIGURATION",
    "PACKAGING_CONFIGURATION_ID",
    "CONFIGURATION",
    "Packed-unit packaging BOM header (revisionable)",
    "Packaging Engineering",
    "Occasional",
    [
        _col("PACKAGING_CONFIGURATION_ID", "INT", pk=True),
        _col("CONFIG_GROUP_CODE", "TEXT"),
        _col("REVISION_NO", "INT"),
        _col("PACKAGING_CONFIGURATION_NAME", "TEXT"),
        _col("DESCRIPTION", "TEXT", required=False),
        _col("SUPERSEDES_ID", "INT", required=False, fk_table="PACKAGING_CONFIGURATION"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EFFECTIVE_FROM", "DATE"),
        _col("EFFECTIVE_TO", "DATE", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

PACKAGING_CONFIGURATION_LINE = _table(
    "PACKAGING_CONFIGURATION_LINE",
    "PACKAGING_CONFIGURATION_LINE_ID",
    "CONFIGURATION",
    "BOM lines: components per packed product unit",
    "Packaging Engineering",
    "With configuration revision",
    [
        _col("PACKAGING_CONFIGURATION_LINE_ID", "INT", pk=True),
        _col("PACKAGING_CONFIGURATION_ID", "INT", fk_table="PACKAGING_CONFIGURATION"),
        _col("COMPONENT_ID", "INT", fk_table="COMPONENT"),
        _col("QUANTITY", "DECIMAL"),
        _col("LINE_ROLE_ID", "INT", fk_table="LKP_LINE_ROLE"),
        _col("SORT_ORDER", "INT"),
        _col("IS_OPTIONAL", "BOOL"),
        _col("NOTES", "TEXT", required=False),
    ],
)

TRANSPORT_CONFIGURATION = _table(
    "TRANSPORT_CONFIGURATION",
    "TRANSPORT_CONFIGURATION_ID",
    "CONFIGURATION",
    "Transport unitization (pallet/container) header",
    "Packaging / Logistics Engineering",
    "Occasional",
    [
        _col("TRANSPORT_CONFIGURATION_ID", "INT", pk=True),
        _col("CONFIG_GROUP_CODE", "TEXT"),
        _col("REVISION_NO", "INT"),
        _col("TRANSPORT_CONFIGURATION_NAME", "TEXT"),
        _col("PACKAGING_CONFIGURATION_ID", "INT", fk_table="PACKAGING_CONFIGURATION"),
        _col("TRANSPORT_UNIT_TYPE_ID", "INT", fk_table="LKP_TRANSPORT_UNIT_TYPE"),
        _col("UNITS_PER_LAYER", "INT", required=False),
        _col("LAYERS_PER_UNIT", "INT", required=False),
        _col("CONTAINER_PAYLOAD_UNITS", "INT", required=False),
        _col("MAX_GROSS_WEIGHT_KG", "DECIMAL", required=False),
        _col("SUPERSEDES_ID", "INT", required=False, fk_table="TRANSPORT_CONFIGURATION"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EFFECTIVE_FROM", "DATE"),
        _col("EFFECTIVE_TO", "DATE", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

TRANSPORT_CONFIGURATION_LINE = _table(
    "TRANSPORT_CONFIGURATION_LINE",
    "TRANSPORT_CONFIGURATION_LINE_ID",
    "CONFIGURATION",
    "Transport-level components per transport unit",
    "Packaging / Logistics Engineering",
    "With transport revision",
    [
        _col("TRANSPORT_CONFIGURATION_LINE_ID", "INT", pk=True),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", fk_table="TRANSPORT_CONFIGURATION"),
        _col("COMPONENT_ID", "INT", fk_table="COMPONENT"),
        _col("QUANTITY_PER_TRANSPORT_UNIT", "DECIMAL"),
        _col("LINE_ROLE_ID", "INT", fk_table="LKP_LINE_ROLE"),
        _col("SORT_ORDER", "INT"),
        _col("NOTES", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Commercial & logistics
# ---------------------------------------------------------------------------

COMMERCIAL_SCENARIO = _table(
    "COMMERCIAL_SCENARIO",
    "COMMERCIAL_SCENARIO_ID",
    "COMMERCIAL",
    "Product + transport + market/customer commercial variant",
    "Sales Ops + Packaging Engineering",
    "Occasional",
    [
        _col("COMMERCIAL_SCENARIO_ID", "INT", pk=True),
        _col("COMMERCIAL_SCENARIO_CODE", "TEXT"),
        _col("COMMERCIAL_SCENARIO_NAME", "TEXT"),
        _col("SCENARIO_TYPE_ID", "INT", fk_table="LKP_SCENARIO_TYPE"),
        _col("PRODUCT_ID", "INT", fk_table="PRODUCT"),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", fk_table="TRANSPORT_CONFIGURATION"),
        _col("CUSTOMER_ID", "INT", required=False, fk_table="CUSTOMER"),
        _col("DESTINATION_COUNTRY_ID", "INT", fk_table="LKP_COUNTRY"),
        _col("INCOTERM_ID", "INT", required=False, fk_table="LKP_INCOTERM"),
        _col("TRANSPORT_MODE_ID", "INT", required=False, fk_table="LKP_TRANSPORT_MODE"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("VALID_FROM", "DATE"),
        _col("VALID_TO", "DATE", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

SHIPMENT = _table(
    "SHIPMENT",
    "SHIPMENT_ID",
    "LOGISTICS",
    "Operational shipment quantity fact with pinned config revisions",
    "Logistics",
    "Frequent",
    [
        _col("SHIPMENT_ID", "INT", pk=True),
        _col("SHIPMENT_NUMBER", "TEXT"),
        _col("COMMERCIAL_SCENARIO_ID", "INT", fk_table="COMMERCIAL_SCENARIO"),
        _col("PLANT_ID", "INT", fk_table="PLANT"),
        _col("SHIP_DATE", "DATE"),
        _col("QTY_PRODUCT_UNITS", "DECIMAL"),
        _col("PACKAGING_CONFIGURATION_ID", "INT", fk_table="PACKAGING_CONFIGURATION"),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", fk_table="TRANSPORT_CONFIGURATION"),
        _col("DESTINATION_COUNTRY_ID", "INT", required=False, fk_table="LKP_COUNTRY"),
        _col("TRANSPORT_MODE_ID", "INT", required=False, fk_table="LKP_TRANSPORT_MODE"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EXTERNAL_REF", "TEXT", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CONFIRMED_AT", "DATETIME", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

SHIPMENT_LINE = _table(
    "SHIPMENT_LINE",
    "SHIPMENT_LINE_ID",
    "LOGISTICS",
    "Frozen packaging composition for a confirmed shipment",
    "System (generated) / Compliance",
    "Written once on CONFIRM",
    [
        _col("SHIPMENT_LINE_ID", "INT", pk=True),
        _col("SHIPMENT_ID", "INT", fk_table="SHIPMENT"),
        _col("COMPONENT_ID", "INT", fk_table="COMPONENT"),
        _col("MATERIAL_ID", "INT", fk_table="LKP_MATERIAL"),
        _col("PACKAGING_LEVEL_ID", "INT", fk_table="LKP_PACKAGING_LEVEL"),
        _col("PACKAGING_FUNCTION_ID", "INT", fk_table="LKP_PACKAGING_FUNCTION"),
        _col("OWNERSHIP_TYPE_ID", "INT", fk_table="LKP_OWNERSHIP_TYPE"),
        _col("COMPONENT_QTY", "DECIMAL"),
        _col("WEIGHT_G", "DECIMAL"),
        _col("RECYCLED_CONTENT_PCT", "DECIMAL", required=False),
        _col("SOURCE_LAYER", "TEXT"),
        _col("NOTES", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Compliance
# ---------------------------------------------------------------------------

TECHNICAL_FILE = _table(
    "TECHNICAL_FILE",
    "TECHNICAL_FILE_ID",
    "COMPLIANCE",
    "PPWR technical documentation package",
    "Compliance / Packaging Engineering",
    "Occasional",
    [
        _col("TECHNICAL_FILE_ID", "INT", pk=True),
        _col("TECHNICAL_FILE_CODE", "TEXT"),
        _col("TITLE", "TEXT"),
        _col("COMPONENT_ID", "INT", required=False, fk_table="COMPONENT"),
        _col("PACKAGING_CONFIGURATION_ID", "INT", required=False, fk_table="PACKAGING_CONFIGURATION"),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", required=False, fk_table="TRANSPORT_CONFIGURATION"),
        _col("REVISION_NO", "INT"),
        _col("ASSESSMENT_DATE", "DATE", required=False),
        _col("RECYCLABILITY_SUMMARY", "TEXT", required=False),
        _col("SUBSTANCE_OF_CONCERN_NOTES", "TEXT", required=False),
        _col("DESIGN_FOR_RECYCLING_NOTES", "TEXT", required=False),
        _col("OWNER_PERSON_ID", "INT", required=False, fk_table="PERSON"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("EFFECTIVE_FROM", "DATE"),
        _col("EFFECTIVE_TO", "DATE", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

DECLARATION_OF_CONFORMITY = _table(
    "DECLARATION_OF_CONFORMITY",
    "DECLARATION_OF_CONFORMITY_ID",
    "COMPLIANCE",
    "Formal PPWR Declaration of Conformity",
    "Compliance",
    "Occasional",
    [
        _col("DECLARATION_OF_CONFORMITY_ID", "INT", pk=True),
        _col("DOC_NUMBER", "TEXT"),
        _col("TITLE", "TEXT"),
        _col("LEGAL_ENTITY_ID", "INT", fk_table="LEGAL_ENTITY"),
        _col("PRODUCT_ID", "INT", required=False, fk_table="PRODUCT"),
        _col("PACKAGING_CONFIGURATION_ID", "INT", required=False, fk_table="PACKAGING_CONFIGURATION"),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", required=False, fk_table="TRANSPORT_CONFIGURATION"),
        _col("TECHNICAL_FILE_ID", "INT", fk_table="TECHNICAL_FILE"),
        _col("RESPONSIBLE_PERSON_ID", "INT", fk_table="PERSON"),
        _col("REGULATION_REFERENCE", "TEXT"),
        _col("CONFORMITY_STATEMENT", "TEXT"),
        _col("ISSUE_DATE", "DATE"),
        _col("VALID_UNTIL", "DATE", required=False),
        _col("REVISION_NO", "INT"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("APPROVED_AT", "DATETIME", required=False),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

STATEMENT = _table(
    "STATEMENT",
    "STATEMENT_ID",
    "COMPLIANCE",
    "Placed-on-market statement header",
    "Compliance",
    "Periodic",
    [
        _col("STATEMENT_ID", "INT", pk=True),
        _col("STATEMENT_CODE", "TEXT"),
        _col("STATEMENT_TYPE_ID", "INT", fk_table="LKP_STATEMENT_TYPE"),
        _col("LEGAL_ENTITY_ID", "INT", fk_table="LEGAL_ENTITY"),
        _col("COUNTRY_ID", "INT", fk_table="LKP_COUNTRY"),
        _col("PERIOD_YEAR", "INT"),
        _col("PERIOD_MONTH", "INT", required=False),
        _col("PERIOD_FROM", "DATE"),
        _col("PERIOD_TO", "DATE"),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("GENERATED_AT", "DATETIME", required=False),
        _col("APPROVED_BY_PERSON_ID", "INT", required=False, fk_table="PERSON"),
        _col("APPROVED_AT", "DATETIME", required=False),
        _col("NOTES", "TEXT", required=False),
    ],
)

STATEMENT_SHIPMENT = _table(
    "STATEMENT_SHIPMENT",
    "STATEMENT_SHIPMENT_ID",
    "COMPLIANCE",
    "Bridge: shipments included in a statement",
    "Compliance",
    "During statement draft",
    [
        _col("STATEMENT_SHIPMENT_ID", "INT", pk=True),
        _col("STATEMENT_ID", "INT", fk_table="STATEMENT"),
        _col("SHIPMENT_ID", "INT", fk_table="SHIPMENT"),
        _col("INCLUDED_AT", "DATETIME"),
    ],
)

STATEMENT_LINE = _table(
    "STATEMENT_LINE",
    "STATEMENT_LINE_ID",
    "COMPLIANCE",
    "Frozen aggregate packaging weights for regulatory reporting",
    "Compliance",
    "On generate/approve",
    [
        _col("STATEMENT_LINE_ID", "INT", pk=True),
        _col("STATEMENT_ID", "INT", fk_table="STATEMENT"),
        _col("MATERIAL_ID", "INT", fk_table="LKP_MATERIAL"),
        _col("PACKAGING_LEVEL_ID", "INT", fk_table="LKP_PACKAGING_LEVEL"),
        _col("OWNERSHIP_TYPE_ID", "INT", fk_table="LKP_OWNERSHIP_TYPE"),
        _col("TOTAL_WEIGHT_KG", "DECIMAL"),
        _col("RECYCLED_CONTENT_WEIGHT_KG", "DECIMAL", required=False),
        _col("SOURCE_SHIPMENT_COUNT", "INT", required=False),
        _col("NOTES", "TEXT", required=False),
    ],
)

DOCUMENT_LIBRARY = _table(
    "DOCUMENT_LIBRARY",
    "DOCUMENT_ID",
    "COMPLIANCE",
    "Central catalog of evidence files",
    "Document Control",
    "Occasional",
    [
        _col("DOCUMENT_ID", "INT", pk=True),
        _col("DOCUMENT_CODE", "TEXT"),
        _col("DOCUMENT_TITLE", "TEXT"),
        _col("DOCUMENT_TYPE_ID", "INT", fk_table="LKP_DOCUMENT_TYPE"),
        _col("FILE_URI", "TEXT"),
        _col("FILE_HASH", "TEXT", required=False),
        _col("ISSUE_DATE", "DATE", required=False),
        _col("STATUS_ID", "INT", fk_table="LKP_STATUS"),
        _col("NOTES", "TEXT", required=False),
        _col("CREATED_AT", "DATETIME"),
        _col("UPDATED_AT", "DATETIME"),
    ],
)

DOCUMENT_LINK = _table(
    "DOCUMENT_LINK",
    "DOCUMENT_LINK_ID",
    "COMPLIANCE",
    "Links a library document to exactly one business record",
    "Compliance / Engineering",
    "Occasional",
    [
        _col("DOCUMENT_LINK_ID", "INT", pk=True),
        _col("DOCUMENT_ID", "INT", fk_table="DOCUMENT_LIBRARY"),
        _col("COMPONENT_ID", "INT", required=False, fk_table="COMPONENT"),
        _col("PRODUCT_ID", "INT", required=False, fk_table="PRODUCT"),
        _col("PACKAGING_CONFIGURATION_ID", "INT", required=False, fk_table="PACKAGING_CONFIGURATION"),
        _col("TRANSPORT_CONFIGURATION_ID", "INT", required=False, fk_table="TRANSPORT_CONFIGURATION"),
        _col("TECHNICAL_FILE_ID", "INT", required=False, fk_table="TECHNICAL_FILE"),
        _col("DECLARATION_OF_CONFORMITY_ID", "INT", required=False, fk_table="DECLARATION_OF_CONFORMITY"),
        _col("STATEMENT_ID", "INT", required=False, fk_table="STATEMENT"),
        _col("SORT_ORDER", "INT"),
        _col("NOTES", "TEXT", required=False),
    ],
)

# ---------------------------------------------------------------------------
# Ordered registry list (sheet creation order)
# ---------------------------------------------------------------------------

ALL_TABLES: tuple[TableDefinition, ...] = (
    SYS_WORKBOOK_INFO,
    SYS_PARAMETER,
    LKP_STATUS,
    LKP_UOM,
    LKP_PACKAGING_LEVEL,
    LKP_PACKAGING_FUNCTION,
    LKP_COMPONENT_TYPE,
    LKP_MATERIAL_FAMILY,
    LKP_PPWR_MATERIAL_CATEGORY,
    LKP_MATERIAL,
    LKP_RECYCLABILITY_CLASS,
    LKP_OWNERSHIP_TYPE,
    LKP_TRANSPORT_UNIT_TYPE,
    LKP_LINE_ROLE,
    LKP_PRODUCT_CATEGORY,
    LKP_COUNTRY,
    LKP_TRANSPORT_MODE,
    LKP_INCOTERM,
    LKP_STATEMENT_TYPE,
    LKP_DOCUMENT_TYPE,
    LKP_SCENARIO_TYPE,
    LEGAL_ENTITY,
    PERSON,
    SUPPLIER,
    CUSTOMER,
    PLANT,
    PRODUCT,
    COMPONENT,
    COMPONENT_MATERIAL,
    PACKAGING_CONFIGURATION,
    PACKAGING_CONFIGURATION_LINE,
    TRANSPORT_CONFIGURATION,
    TRANSPORT_CONFIGURATION_LINE,
    COMMERCIAL_SCENARIO,
    SHIPMENT,
    SHIPMENT_LINE,
    TECHNICAL_FILE,
    DECLARATION_OF_CONFORMITY,
    STATEMENT,
    STATEMENT_SHIPMENT,
    STATEMENT_LINE,
    DOCUMENT_LIBRARY,
    DOCUMENT_LINK,
)

EXPECTED_TABLE_COUNT = 43
