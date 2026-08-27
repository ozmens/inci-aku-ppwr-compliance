"""Traceability path helper (architecture documentation in code)."""

from __future__ import annotations

TRACEABILITY_CHAIN: tuple[str, ...] = (
    "PRODUCT",
    "COMMERCIAL_SCENARIO",
    "TRANSPORT_CONFIGURATION",
    "PACKAGING_CONFIGURATION",
    "COMPONENT",
    "COMPONENT_MATERIAL",
    "SHIPMENT",
    "SHIPMENT_LINE",
    "STATEMENT_SHIPMENT",
    "STATEMENT_LINE",
)


def describe_traceability() -> str:
    return " → ".join(TRACEABILITY_CHAIN)
