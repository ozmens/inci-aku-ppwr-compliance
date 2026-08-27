"""
Business rule catalog (V-* codes from FINAL_DATABASE.md).

Architecture phase stores rule metadata only; data execution comes later.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ValidationRule:
    rule_id: str
    description: str
    severity: str  # ERROR | WARN
    entity: str


RULES: tuple[ValidationRule, ...] = (
    ValidationRule("V-PK-01", "All PKs unique and non-null", "ERROR", "*"),
    ValidationRule("V-FK-01", "All non-null FKs resolve to parent PKs", "ERROR", "*"),
    ValidationRule("V-WT-01", "COMPONENT.WEIGHT_G > 0", "ERROR", "COMPONENT"),
    ValidationRule(
        "V-WT-02",
        "No editable total weight on config/product/shipment headers",
        "ERROR",
        "CONFIGURATION",
    ),
    ValidationRule(
        "V-MAT-01",
        "COMPONENT_MATERIAL shares sum to ~100%",
        "ERROR",
        "COMPONENT_MATERIAL",
    ),
    ValidationRule(
        "V-MAT-02",
        "ACTIVE component has >=1 COMPONENT_MATERIAL row",
        "ERROR",
        "COMPONENT",
    ),
    ValidationRule(
        "V-CFG-01",
        "ACTIVE packaging configuration has >=1 line",
        "ERROR",
        "PACKAGING_CONFIGURATION",
    ),
    ValidationRule(
        "V-TRN-01",
        "Pallet only as TRANSPORT_CONFIGURATION_LINE role PALLET",
        "ERROR",
        "TRANSPORT_CONFIGURATION",
    ),
    ValidationRule(
        "V-TRN-02",
        "PALLET requires layer fields; CONTAINER requires payload units",
        "ERROR",
        "TRANSPORT_CONFIGURATION",
    ),
    ValidationRule(
        "V-SCN-01",
        "CUSTOMER scenario type requires CUSTOMER_ID",
        "ERROR",
        "COMMERCIAL_SCENARIO",
    ),
    ValidationRule(
        "V-SHP-01",
        "Confirm requires pinned configs and >=1 SHIPMENT_LINE",
        "ERROR",
        "SHIPMENT",
    ),
    ValidationRule(
        "V-SHP-02",
        "Confirmed shipment and lines immutable",
        "ERROR",
        "SHIPMENT",
    ),
    ValidationRule(
        "V-STM-01",
        "Approved statement lines/links immutable",
        "ERROR",
        "STATEMENT",
    ),
    ValidationRule(
        "V-STM-02",
        "Statement aggregates reconcile to included SHIPMENT_LINE",
        "ERROR",
        "STATEMENT_LINE",
    ),
    ValidationRule(
        "V-TF-01",
        "Technical file has exactly one subject FK",
        "ERROR",
        "TECHNICAL_FILE",
    ),
    ValidationRule(
        "V-DOC-01",
        "Document link has exactly one target FK",
        "ERROR",
        "DOCUMENT_LINK",
    ),
    ValidationRule(
        "V-DoC-01",
        "DoC requires legal entity, person, technical file, >=1 scope FK",
        "ERROR",
        "DECLARATION_OF_CONFORMITY",
    ),
)
