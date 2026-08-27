"""Product / SKU domain model — ERP reference for Battery DPP chain."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class Product:
    product_id: str
    product_code: str
    product_name: str
    product_category: str | None = None
    net_weight_g: float | None = None
    status: str = "ACTIVE"
    # Traceability to packaging (non-destructive mapping)
    linked_final_configuration_ids: list[str] = field(default_factory=list)
    linked_source_configuration_ids: list[str] = field(default_factory=list)
