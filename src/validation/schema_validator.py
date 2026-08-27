"""Schema architecture checks against frozen registry."""

from __future__ import annotations

from models.registry import SchemaRegistry
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class SchemaValidator:
    def __init__(self, registry: SchemaRegistry, expected_version: str = "1.0.0") -> None:
        self.registry = registry
        self.expected_version = expected_version

    def validate(self) -> ValidationResult:
        result = ValidationResult()
        if self.registry.schema_version != self.expected_version:
            result.add(
                ValidationIssue(
                    code="SCH-01",
                    severity=Severity.ERROR,
                    object_type="SCHEMA",
                    object_id=self.registry.schema_version,
                    message=f"Schema version mismatch; expected {self.expected_version}",
                )
            )
        if self.registry.table_count != 43:
            result.add(
                ValidationIssue(
                    code="SCH-02",
                    severity=Severity.ERROR,
                    object_type="SCHEMA",
                    object_id=str(self.registry.table_count),
                    message="Expected 43 frozen tables",
                )
            )
        return result
