"""Immutable table metadata used by builders and validators."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True, slots=True)
class ColumnDefinition:
    name: str
    data_type: str
    required: bool = True
    is_pk: bool = False
    is_fk: bool = False
    fk_table: str | None = None
    description: str = ""


@dataclass(frozen=True, slots=True)
class TableDefinition:
    name: str
    primary_key: str
    group: str
    business_purpose: str
    data_owner: str
    update_frequency: str
    columns: tuple[ColumnDefinition, ...] = field(default_factory=tuple)
    foreign_keys: tuple[str, ...] = field(default_factory=tuple)

    @property
    def column_names(self) -> tuple[str, ...]:
        return tuple(c.name for c in self.columns)
