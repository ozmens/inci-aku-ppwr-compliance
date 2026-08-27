"""Document library / link domain models."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class Document:
    document_id: str
    document_code: str
    title: str
    document_type: str
    file_uri: str | None = None
    issue_date: date | None = None
    status: str = "ACTIVE"


@dataclass(slots=True)
class DocumentLink:
    document_link_id: str
    document_id: str
    packaging_configuration_id: str | None = None
    component_id: str | None = None
    technical_file_id: str | None = None
    declaration_id: str | None = None
    statement_id: str | None = None
    product_id: str | None = None
    sort_order: int = 0
