"""Domain models — packaging configuration / Golden Variant."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class PackagingConfigurationLine:
    """Exact BOM line for one final packaging configuration."""

    component_id: str
    component_erp_code: str
    component_name: str
    quantity: float
    uom: str = "PCS"
    line_role: str = "OTHER"
    sort_order: int = 0
    is_optional: bool = False
    notes: str | None = None
    weight_g: float | None = None  # from component master; not stored as independent truth
    name_tr: str | None = None
    name_en: str | None = None


@dataclass(slots=True)
class SourceLineage:
    """Non-destructive source-data lineage for a final configuration."""

    source_configuration_id: str | None = None
    source_file: str | None = None
    source_sheet: str | None = None
    source_row: int | None = None
    source_reference: str | None = None


@dataclass(slots=True)
class PackagingConfiguration:
    """
    One physically distinct packaging BOM = one final Packaging Configuration.

    Identifiers are owned by IdService — do not invent ad-hoc IDs here.
    """

    packaging_set_code: str
    final_configuration_id: str
    parent_family_code: str
    variant_number: int
    name: str
    revision_no: int = 0
    revision_code: str = "R00"
    description: str | None = None
    variant_basis_tr: str = ""
    variant_basis_en: str = ""
    bom_signature: str | None = None
    lines: list[PackagingConfigurationLine] = field(default_factory=list)
    lineage: SourceLineage = field(default_factory=SourceLineage)
    status: str = "DRAFT"
    extras: dict[str, Any] = field(default_factory=dict)

    @property
    def has_bom(self) -> bool:
        return bool(self.lines)
