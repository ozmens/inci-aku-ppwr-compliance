"""Variant Basis uniqueness within parent family."""

from __future__ import annotations

from collections import defaultdict

from models.packaging_configuration import PackagingConfiguration
from services.variant_basis_service import VariantBasisService
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class VariantValidator:
    def __init__(self, basis_service: VariantBasisService | None = None) -> None:
        self.basis = basis_service or VariantBasisService()

    def validate(self, configs: list[PackagingConfiguration]) -> ValidationResult:
        result = ValidationResult()
        by_family: dict[str, list[PackagingConfiguration]] = defaultdict(list)
        for cfg in configs:
            by_family[cfg.parent_family_code.upper()].append(cfg)

        for family, members in by_family.items():
            active = [m for m in members if (m.status or "").upper() in {"ACTIVE", "DRAFT", ""}]
            seen: list[str] = []
            for cfg in active:
                basis = cfg.variant_basis_tr or ""
                if not basis:
                    result.add(
                        ValidationIssue(
                            code="VAR-01",
                            severity=Severity.WARNING,
                            object_type="PACKAGING_CONFIGURATION",
                            object_id=cfg.final_configuration_id,
                            message="Variant Basis TR is empty",
                            field="variant_basis_tr",
                        )
                    )
                    continue
                if not self.basis.is_unique_within_family(basis, seen):
                    result.add(
                        ValidationIssue(
                            code="VAR-02",
                            severity=Severity.ERROR,
                            object_type="PACKAGING_CONFIGURATION",
                            object_id=cfg.final_configuration_id,
                            message=f"Duplicate Variant Basis within parent family {family}",
                            field="variant_basis_tr",
                        )
                    )
                seen.append(basis)
        return result
