"""
Centralized deterministic ID generation (Golden Variant).

All document/config IDs MUST be created here — never inside report builders.
"""

from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class GeneratedIds:
    parent_family_code: str
    variant_number: int
    packaging_set_code: str
    final_configuration_id: str
    technical_file_id: str
    doc_id: str
    label_id: str
    statement_id: str
    revision_code: str


class IdService:
    """Deterministic identifiers for packaging sets and PPWR documents."""

    def __init__(
        self,
        *,
        company_prefix: str = "IA",
        ppwr_prefix: str = "IA-PPWR",
        variant_width: int = 2,
        default_revision_code: str = "R00",
    ) -> None:
        self.company_prefix = company_prefix
        self.ppwr_prefix = ppwr_prefix
        self.variant_width = variant_width
        self.default_revision_code = default_revision_code

    def normalize_family_code(self, family_or_set: str) -> str:
        """Strip trailing -NN variant suffix → parent family (e.g. ST-051-STD)."""
        text = family_or_set.strip().upper().replace(" ", "")
        m = re.match(r"^(.*)-(\d{2})$", text)
        if m:
            return m.group(1)
        return text

    def packaging_set_code(self, parent_family_code: str, variant_number: int) -> str:
        if variant_number < 1:
            raise ValueError("variant_number must be >= 1 (use 1 for single-BOM families)")
        family = self.normalize_family_code(parent_family_code)
        return f"{family}-{variant_number:0{self.variant_width}d}"

    def final_configuration_id(self, packaging_set_code: str) -> str:
        code = packaging_set_code.strip().upper()
        if code.startswith(f"{self.company_prefix}-"):
            return code
        return f"{self.company_prefix}-{code}"

    def technical_file_id(self, packaging_set_code: str, revision_code: str | None = None) -> str:
        rev = revision_code or self.default_revision_code
        return f"{self.ppwr_prefix}-TF-{packaging_set_code}-{rev}"

    def doc_id(self, packaging_set_code: str, revision_code: str | None = None) -> str:
        rev = revision_code or self.default_revision_code
        return f"{self.ppwr_prefix}-DOC-{packaging_set_code}-{rev}"

    def label_id(self, packaging_set_code: str, revision_code: str | None = None) -> str:
        rev = revision_code or self.default_revision_code
        return f"{self.ppwr_prefix}-LBL-{packaging_set_code}-{rev}"

    def statement_id(self, packaging_set_code: str, revision_code: str | None = None) -> str:
        rev = revision_code or self.default_revision_code
        return f"{self.ppwr_prefix}-STM-{packaging_set_code}-{rev}"

    def generate(
        self,
        parent_family_code: str,
        variant_number: int = 1,
        *,
        revision_code: str | None = None,
    ) -> GeneratedIds:
        """
        Examples:
          ST-051-STD + 1 → ST-051-STD-01 → IA-ST-051-STD-01 → docs …-R00
          ST-028-STD + 1 → ST-028-STD-01 (single-variant families still get -01)
        """
        rev = revision_code or self.default_revision_code
        family = self.normalize_family_code(parent_family_code)
        set_code = self.packaging_set_code(family, variant_number)
        return GeneratedIds(
            parent_family_code=family,
            variant_number=variant_number,
            packaging_set_code=set_code,
            final_configuration_id=self.final_configuration_id(set_code),
            technical_file_id=self.technical_file_id(set_code, rev),
            doc_id=self.doc_id(set_code, rev),
            label_id=self.label_id(set_code, rev),
            statement_id=self.statement_id(set_code, rev),
            revision_code=rev,
        )
