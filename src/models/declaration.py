"""Declaration of Conformity domain model."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date


@dataclass(slots=True)
class Declaration:
    declaration_id: str
    doc_number: str
    title: str
    technical_file_id: str
    packaging_configuration_id: str
    legal_entity_id: str | None = None
    product_id: str | None = None
    regulation_reference: str | None = None
    conformity_statement: str | None = None
    issue_date: date | None = None
    revision_no: int = 0
    revision_code: str = "R00"
    status: str = "DRAFT"
