"""Validation runner — architecture checks now; data checks later."""

from __future__ import annotations

from models.registry import SchemaRegistry
from models.tables import EXPECTED_TABLE_COUNT

from .rules import RULES


class ValidationRunner:
    def __init__(self, registry: SchemaRegistry) -> None:
        self.registry = registry

    def check_architecture(self) -> list[str]:
        """Non-data checks safe to run without Excel."""
        issues: list[str] = []
        issues.extend(self.registry.validate_count())

        if not RULES:
            issues.append("Validation rule catalog is empty")

        # Mandatory entities from FINAL_DATABASE / Phase C freeze
        mandatory = {
            "COMPONENT",
            "COMPONENT_MATERIAL",
            "PRODUCT",
            "PACKAGING_CONFIGURATION",
            "PACKAGING_CONFIGURATION_LINE",
            "TRANSPORT_CONFIGURATION",
            "TRANSPORT_CONFIGURATION_LINE",
            "COMMERCIAL_SCENARIO",
            "SHIPMENT",
            "SHIPMENT_LINE",
            "STATEMENT",
            "STATEMENT_LINE",
            "TECHNICAL_FILE",
            "DECLARATION_OF_CONFORMITY",
            "DOCUMENT_LIBRARY",
            "DOCUMENT_LINK",
        }
        names = set(self.registry.names())
        missing = sorted(mandatory - names)
        if missing:
            issues.append(f"Missing mandatory tables: {', '.join(missing)}")

        # Naming freeze smoke checks
        forbidden = {
            "PACKAGING_CONFIG",
            "LOADING_CONFIG",
            "SHIPMENT_PACKAGING_LINE",
            "COMPONENT_MATERIAL_SHARE",
        }
        present_forbidden = sorted(forbidden & names)
        if present_forbidden:
            issues.append(f"Deprecated table names present: {', '.join(present_forbidden)}")

        if self.registry.table_count != EXPECTED_TABLE_COUNT:
            # already reported via validate_count; keep explicit
            pass

        # Ensure every table has PK metadata
        for table in self.registry.tables:
            if not table.primary_key:
                issues.append(f"{table.name} missing primary_key metadata")
            if table.primary_key not in table.column_names:
                issues.append(
                    f"{table.name} PK {table.primary_key} not in column list"
                )

        return issues

    def validate_dataframe_bundle(self, frames: dict) -> list[str]:
        """
        Future: accept {table_name: pandas.DataFrame} and execute V-* rules.

        Not implemented in architecture phase.
        """
        raise NotImplementedError(
            "Data validation against pandas DataFrames is reserved for a later phase."
        )
