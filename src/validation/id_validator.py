"""Identifier uniqueness / format validation."""

from __future__ import annotations

import re

from models.packaging_configuration import PackagingConfiguration
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult

_SET_RE = re.compile(r"^.+-\d{2}$")


class IdValidator:
    def validate_set(
        self,
        configs: list[PackagingConfiguration],
    ) -> ValidationResult:
        result = ValidationResult()
        seen_set: set[str] = set()
        seen_final: set[str] = set()
        for cfg in configs:
            if not _SET_RE.match(cfg.packaging_set_code or ""):
                result.add(
                    ValidationIssue(
                        code="ID-01",
                        severity=Severity.ERROR,
                        object_type="PACKAGING_CONFIGURATION",
                        object_id=cfg.final_configuration_id or cfg.packaging_set_code,
                        message="Missing variant suffix — Packaging Set Code must end with -NN (e.g. -01)",
                        field="packaging_set_code",
                    )
                )
            key = (cfg.packaging_set_code or "").upper()
            if key in seen_set:
                result.add(
                    ValidationIssue(
                        code="ID-02",
                        severity=Severity.ERROR,
                        object_type="PACKAGING_CONFIGURATION",
                        object_id=key,
                        message="Duplicate Packaging Set Code",
                        field="packaging_set_code",
                    )
                )
            seen_set.add(key)
            fid = (cfg.final_configuration_id or "").upper()
            if fid in seen_final:
                result.add(
                    ValidationIssue(
                        code="ID-03",
                        severity=Severity.ERROR,
                        object_type="PACKAGING_CONFIGURATION",
                        object_id=fid,
                        message="Duplicate Final Configuration ID",
                        field="final_configuration_id",
                    )
                )
            seen_final.add(fid)
            if not cfg.lineage.source_configuration_id and not cfg.lineage.source_reference:
                result.add(
                    ValidationIssue(
                        code="ID-04",
                        severity=Severity.WARNING,
                        object_type="PACKAGING_CONFIGURATION",
                        object_id=fid,
                        message="Unresolved source lineage",
                        field="lineage",
                    )
                )
        return result
