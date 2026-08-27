"""Named ranges for PK columns (FK data-validation sources)."""

from __future__ import annotations

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook

from models.registry import SchemaRegistry
from models.table_definition import TableDefinition

from .sheet_builder import SheetBuilder


def pk_named_range_name(primary_key: str) -> str:
    return f"NR_{primary_key}"


def plan_named_ranges(registry: SchemaRegistry) -> dict[str, str]:
    """Map NR_<PK> → structured table reference description."""
    ranges: dict[str, str] = {}
    for table in registry.tables:
        ranges[pk_named_range_name(table.primary_key)] = (
            f"T_{table.name}[{table.primary_key}]"
        )
    return ranges


def apply_pk_named_ranges(
    workbook: Workbook,
    registry: SchemaRegistry,
    sheet_builder: SheetBuilder,
) -> list[str]:
    """Register workbook-level defined names for every table PK column."""
    created: list[str] = []
    for table in registry.tables:
        name = pk_named_range_name(table.primary_key)
        # Avoid duplicate defined names if two tables shared a PK name (should not)
        if name in workbook.defined_names:
            continue
        defn = DefinedName(
            name=name,
            attr_text=sheet_builder.pk_range_formula(table),
        )
        workbook.defined_names.add(defn)
        created.append(name)
    return created


def parent_pk_for_fk(column: object, registry: SchemaRegistry) -> str | None:
    """Resolve FK column → parent table PK named range name."""
    fk_table = getattr(column, "fk_table", None)
    if not fk_table:
        return None
    try:
        parent: TableDefinition = registry.get(fk_table)
    except KeyError:
        return None
    return pk_named_range_name(parent.primary_key)
