"""BOM validation."""

from __future__ import annotations

from models.packaging_configuration import PackagingConfiguration
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class BomValidator:
    def validate(self, config: PackagingConfiguration) -> ValidationResult:
        result = ValidationResult()
        if not config.lines:
            result.add(
                ValidationIssue(
                    code="BOM-01",
                    severity=Severity.ERROR,
                    object_type="PACKAGING_CONFIGURATION",
                    object_id=config.final_configuration_id,
                    message="BOM missing — exact packaging lines required",
                    field="lines",
                )
            )
            return result
        for line in config.lines:
            if not (line.component_id or line.component_erp_code):
                result.add(
                    ValidationIssue(
                        code="BOM-02",
                        severity=Severity.ERROR,
                        object_type="PACKAGING_CONFIGURATION_LINE",
                        object_id=config.final_configuration_id,
                        message="BOM component missing identity",
                        field="component_id",
                    )
                )
            if line.quantity is None or float(line.quantity) <= 0:
                result.add(
                    ValidationIssue(
                        code="BOM-03",
                        severity=Severity.ERROR,
                        object_type="PACKAGING_CONFIGURATION_LINE",
                        object_id=config.final_configuration_id,
                        message="BOM quantity must be > 0",
                        field="quantity",
                    )
                )
        return result
