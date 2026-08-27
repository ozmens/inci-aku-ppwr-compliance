"""Excel data validations — lookups, booleans, dates (performance-safe)."""

from __future__ import annotations

from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from models.registry import SchemaRegistry
from models.table_definition import TableDefinition

from builders.named_ranges import pk_named_range_name
from builders.sheet_builder import excel_col_letter

# Only low-volume lookup FK targets get dropdown lists.
LOOKUP_FK_PREFIX = "LKP_"


def apply_pk_named_ranges(workbook: Workbook, registry: SchemaRegistry) -> list[str]:
    created: list[str] = []
    for table in registry.tables:
        name = pk_named_range_name(table.primary_key)
        if name in workbook.defined_names:
            continue
        col_idx = list(table.column_names).index(table.primary_key) + 1
        col = excel_col_letter(col_idx)
        # Bounded range for performance (not full column)
        formula = f"'{table.name}'!${col}$2:${col}$5000"
        workbook.defined_names.add(DefinedName(name=name, attr_text=formula))
        created.append(name)
    return created


def apply_validations(workbook: Workbook, registry: SchemaRegistry) -> int:
    """Add targeted validations. Returns validation count."""
    count = 0
    for table in registry.tables:
        ws = workbook[table.name]
        for col in table.columns:
            col_letter = excel_col_letter(list(table.column_names).index(col.name) + 1)
            rng = f"{col_letter}2:{col_letter}5000"

            if col.data_type == "BOOL":
                dv = DataValidation(
                    type="list",
                    formula1='"TRUE,FALSE"',
                    allow_blank=True,
                    showErrorMessage=True,
                    errorTitle="Boolean",
                    error="Enter TRUE or FALSE",
                )
                dv.add(rng)
                ws.add_data_validation(dv)
                count += 1
                continue

            if col.data_type == "DATE":
                dv = DataValidation(
                    type="date",
                    operator="greaterThan",
                    formula1="DATE(2000,1,1)",
                    allow_blank=True,
                    showErrorMessage=True,
                    errorTitle="Date",
                    error="Enter a valid date (DD.MM.YYYY)",
                )
                dv.add(rng)
                ws.add_data_validation(dv)
                count += 1
                continue

            if col.is_fk and col.fk_table and col.fk_table.startswith(LOOKUP_FK_PREFIX):
                nr = pk_named_range_name(_parent_pk(registry, col.fk_table))
                dv = DataValidation(
                    type="list",
                    formula1=f"={nr}",
                    allow_blank=True,
                    showErrorMessage=True,
                    errorTitle="Lookup FK",
                    error=f"Value must exist in {col.fk_table}",
                )
                dv.add(rng)
                ws.add_data_validation(dv)
                count += 1
    return count


def _parent_pk(registry: SchemaRegistry, table_name: str) -> str:
    parent: TableDefinition = registry.get(table_name)
    return parent.primary_key
