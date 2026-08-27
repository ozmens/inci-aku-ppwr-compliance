"""Component domain model."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class ComponentMaterialShare:
    material_code: str
    share_pct: float
    material_family: str | None = None


@dataclass(slots=True)
class Component:
    component_id: str
    component_code: str
    component_name: str
    weight_g: float
    component_type: str | None = None
    packaging_level: str | None = None
    packaging_function: str | None = None
    supplier_id: str | None = None
    materials: list[ComponentMaterialShare] = field(default_factory=list)
    name_tr: str | None = None
    name_en: str | None = None
    status: str = "ACTIVE"
