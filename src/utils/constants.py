"""Shared constants for Golden Variant / document architecture."""

from __future__ import annotations

from enum import Enum


class Severity(str, Enum):
    ERROR = "ERROR"
    WARNING = "WARNING"
    INFO = "INFO"


class DocumentKind(str, Enum):
    TECHNICAL_FILE = "TECHNICAL_FILE"
    DECLARATION = "DECLARATION_OF_CONFORMITY"
    LABEL = "PACKAGING_IDENTIFICATION_LABEL"
    STATEMENT = "SHIPMENT_PACKAGING_INFORMATION_STATEMENT"


# Customer-facing / commercial fields forbidden on Technical File context
TECHNICAL_FILE_FORBIDDEN_FIELDS: frozenset[str] = frozenset(
    {
        "customer_name",
        "customer_id",
        "oem_name",
        "customer_market",
        "destination_country",
        "incoterm",
        "commercial_scenario_id",
    }
)

ANNEX_DRAWINGS_PENDING = "PENDING – DRAWINGS / PHOTOGRAPHS"
ARTICLE5_BASIS_LABEL = "REV00 CURRENT EVIDENCE BASIS — ARTICLE 5 ASSESSMENT BASIS"

# Variant basis description priority (higher first)
VARIANT_BASIS_ROLE_PRIORITY: tuple[str, ...] = (
    "BATTERY_TYPE",
    "OUTER_CARTON",
    "UPPER_COVER",
    "SEPARATOR",
    "EDGE_PROTECTOR",
    "CORNER",
    "SHRINK",
    "STRETCH",
    "STRAP",
    "OTHER",
)

# Map common line-role codes to basis priority buckets
LINE_ROLE_TO_BASIS: dict[str, str] = {
    "OUTER_CARTON": "OUTER_CARTON",
    "BASE": "OUTER_CARTON",
    "UPPER_COVER": "UPPER_COVER",
    "SEPARATOR": "SEPARATOR",
    "EDGE_PROTECTOR": "EDGE_PROTECTOR",
    "CORNER": "EDGE_PROTECTOR",
    "SHRINK": "SHRINK",
    "STRETCH": "STRETCH",
    "WRAP": "STRETCH",
    "STRAP": "STRAP",
    "FILM": "STRETCH",
    "LABEL": "OTHER",
    "PALLET": "OTHER",
    "OTHER": "OTHER",
}
