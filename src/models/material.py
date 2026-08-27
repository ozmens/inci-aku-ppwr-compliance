"""Material domain model."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(slots=True)
class Material:
    material_id: str
    material_code: str
    material_name: str
    material_family: str
    ppwr_category: str | None = None
    is_composite: bool = False
