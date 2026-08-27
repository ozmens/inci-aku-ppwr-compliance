"""Transport configuration domain model."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class TransportConfigurationLine:
    component_id: str
    quantity_per_transport_unit: float
    line_role: str = "OTHER"
    sort_order: int = 0


@dataclass(slots=True)
class TransportConfiguration:
    transport_configuration_id: str
    name: str
    packaging_configuration_id: str
    transport_unit_type: str = "PALLET"
    lines: list[TransportConfigurationLine] = field(default_factory=list)
    revision_no: int = 0
    status: str = "ACTIVE"
