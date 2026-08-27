"""In-memory staging structures for Phase F migration."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class StageLineage:
    source_file: str
    source_sheet: str
    source_row: int | None = None
    source_reference: str | None = None
    source_configuration_id: str | None = None
    source_hash: str | None = None


@dataclass
class StagedConfiguration:
    family: str
    packaging_set_code: str
    final_configuration_id: str
    source_configuration_id: str
    configuration_name: str
    packaging_mass_kg: float
    variant_basis_tr: str
    variant_basis_en: str
    technical_file_id: str
    doc_id: str
    label_id: str
    statement_id: str
    revision: str
    evidence_basis: str
    nominal_product_qty: float | None = None
    pallet_class: str | None = None
    lineage: StageLineage | None = None
    extras: dict[str, Any] = field(default_factory=dict)


@dataclass
class StagedBomLine:
    packaging_set_code: str
    final_configuration_id: str
    source_configuration_id: str
    component_code: str
    component_group: str | None
    erp_description: str
    quantity: float
    uom: str
    unit_weight_kg: float | None
    line_weight_kg: float
    weight_basis: str
    lineage: StageLineage


@dataclass
class StagedProductMap:
    product_code: str
    product_description: str
    battery_type: str | None
    customer_market: str | None
    nominal_qty: float | None
    source_configuration_id: str | None
    final_set_code: str | None
    final_configuration_id: str | None
    status: str
    lineage: StageLineage


@dataclass
class StagedComponent:
    erp_code: str
    description: str
    component_group: str | None = None
    unit_weight_kg: float | None = None
    material_family: str | None = None
    aliases: list[str] = field(default_factory=list)
    lineage: StageLineage | None = None


@dataclass
class StagingBundle:
    configurations: list[StagedConfiguration] = field(default_factory=list)
    bom_lines: list[StagedBomLine] = field(default_factory=list)
    product_maps: list[StagedProductMap] = field(default_factory=list)
    components: dict[str, StagedComponent] = field(default_factory=dict)
    discrepancies: list[dict[str, Any]] = field(default_factory=list)
    domestic_source_configs: list[dict[str, Any]] = field(default_factory=list)
    level2_meta: dict[str, Any] = field(default_factory=dict)
