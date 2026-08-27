"""Weight validation."""

from __future__ import annotations

from models.packaging_configuration import PackagingConfiguration
from services.weight_service import WeightService
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class WeightValidator:
    def __init__(self, weight_service: WeightService | None = None) -> None:
        self.weights = weight_service or WeightService()

    def validate(self, config: PackagingConfiguration) -> ValidationResult:
        result = ValidationResult()
        wr = self.weights.calculate_tare(config.lines)
        if wr.missing_weight_component_ids:
            result.add(
                ValidationIssue(
                    code="WT-01",
                    severity=Severity.ERROR,
                    object_type="PACKAGING_CONFIGURATION",
                    object_id=config.final_configuration_id,
                    message=f"Weight missing for components: {', '.join(wr.missing_weight_component_ids)}",
                    field="weight_g",
                )
            )
        claimed = config.extras.get("claimed_tare_g")
        if claimed is not None and not self.weights.assert_consistent(
            calculated_tare_g=wr.total_tare_g, claimed_tare_g=float(claimed)
        ):
            result.add(
                ValidationIssue(
                    code="WT-02",
                    severity=Severity.ERROR,
                    object_type="PACKAGING_CONFIGURATION",
                    object_id=config.final_configuration_id,
                    message="Calculated tare mismatch vs claimed tare",
                    field="total_tare_g",
                )
            )
        return result
