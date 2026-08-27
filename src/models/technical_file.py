"""Technical File domain model."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date


@dataclass(slots=True)
class Article5Assessment:
    """
    Assessment-basis structure — never auto-promote evidence existence
    into an unsupported numerical compliance claim.
    """

    basis_label: str
    evidence_references: list[str] = field(default_factory=list)
    numerical_claim: str | None = None
    notes: str | None = None


@dataclass(slots=True)
class TechnicalFile:
    technical_file_id: str
    code: str
    title: str
    packaging_configuration_id: str
    revision_no: int = 0
    revision_code: str = "R00"
    assessment_date: date | None = None
    recyclability_summary: str | None = None
    substance_of_concern_notes: str | None = None
    design_for_recycling_notes: str | None = None
    article5: Article5Assessment | None = None
    annex_drawings_status: str = "PENDING – DRAWINGS / PHOTOGRAPHS"
    status: str = "DRAFT"
