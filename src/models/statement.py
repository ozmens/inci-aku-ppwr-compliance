"""Statement domain model."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


@dataclass(slots=True)
class StatementLine:
    component_id: str
    material_id: str | None
    weight_g: float
    component_qty: float
    notes: str | None = None


@dataclass(slots=True)
class Statement:
    statement_id: str
    statement_number: str
    title: str
    revision_no: int = 0
    period_from: date | None = None
    period_to: date | None = None
    shipment_ids: list[str] = field(default_factory=list)
    lines: list[StatementLine] = field(default_factory=list)
    status: str = "DRAFT"
