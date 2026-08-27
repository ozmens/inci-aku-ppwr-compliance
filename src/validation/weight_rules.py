"""Weight ownership invariants (documentation + future enforcement)."""

from __future__ import annotations

WEIGHT_OWNER_TABLE = "COMPONENT"
WEIGHT_OWNER_COLUMN = "WEIGHT_G"

FORBIDDEN_EDITABLE_TOTAL_WEIGHT_TABLES = frozenset(
    {
        "PACKAGING_CONFIGURATION",
        "TRANSPORT_CONFIGURATION",
        "PRODUCT",
        "COMMERCIAL_SCENARIO",
        "SHIPMENT",
    }
)

FREEZE_WEIGHT_TABLES = frozenset({"SHIPMENT_LINE", "STATEMENT_LINE"})


def describe_weight_chain() -> str:
    return (
        "COMPONENT.WEIGHT_G → configuration lines → "
        "SHIPMENT_LINE (confirm freeze) → STATEMENT_LINE (approve freeze)"
    )
