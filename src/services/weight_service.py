"""
Weight ownership — component WEIGHT_G is master; configuration tare is DERIVED.

All TF / DoC / Label / Statement tares must consume this service.
"""

from __future__ import annotations

from dataclasses import dataclass

from models.packaging_configuration import PackagingConfigurationLine


@dataclass(frozen=True, slots=True)
class WeightResult:
    total_tare_g: float
    line_weights_g: tuple[float, ...]
    missing_weight_component_ids: tuple[str, ...]


class WeightService:
    """Derive packaging tare from exact BOM × component weights."""

    def calculate_tare(self, lines: list[PackagingConfigurationLine]) -> WeightResult:
        totals: list[float] = []
        missing: list[str] = []
        for line in lines:
            if line.weight_g is None:
                missing.append(line.component_id or line.component_erp_code)
                totals.append(0.0)
                continue
            totals.append(float(line.weight_g) * float(line.quantity))
        return WeightResult(
            total_tare_g=round(sum(totals), 6),
            line_weights_g=tuple(round(v, 6) for v in totals),
            missing_weight_component_ids=tuple(missing),
        )

    def assert_consistent(
        self,
        *,
        calculated_tare_g: float,
        claimed_tare_g: float,
        tolerance_g: float = 0.01,
    ) -> bool:
        return abs(calculated_tare_g - claimed_tare_g) <= tolerance_g
