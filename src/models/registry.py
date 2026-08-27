"""Schema registry facade over frozen table definitions."""

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass

from .schema_version import SCHEMA_VERSION
from .table_definition import TableDefinition
from .tables import ALL_TABLES, EXPECTED_TABLE_COUNT


@dataclass(slots=True)
class SchemaRegistry:
    schema_version: str
    tables: tuple[TableDefinition, ...]

    @classmethod
    def load(cls) -> SchemaRegistry:
        return cls(schema_version=SCHEMA_VERSION, tables=ALL_TABLES)

    @property
    def table_count(self) -> int:
        return len(self.tables)

    def get(self, name: str) -> TableDefinition:
        for table in self.tables:
            if table.name == name:
                return table
        raise KeyError(f"Unknown table: {name}")

    def names(self) -> list[str]:
        return [t.name for t in self.tables]

    def grouped_tables(self) -> OrderedDict[str, list[TableDefinition]]:
        groups: OrderedDict[str, list[TableDefinition]] = OrderedDict()
        for table in self.tables:
            groups.setdefault(table.group, []).append(table)
        return groups

    def validate_count(self) -> list[str]:
        issues: list[str] = []
        if self.table_count != EXPECTED_TABLE_COUNT:
            issues.append(
                f"Expected {EXPECTED_TABLE_COUNT} tables, found {self.table_count}"
            )
        names = self.names()
        if len(names) != len(set(names)):
            issues.append("Duplicate table names in registry")
        return issues


# Re-export for importers that expect TableDefinition from registry
__all__ = ["SchemaRegistry", "TableDefinition"]
