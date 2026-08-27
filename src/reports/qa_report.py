"""QA report architecture stub — no production batch yet."""

from __future__ import annotations

from dataclasses import dataclass

from validation.base import ValidationResult


@dataclass(slots=True)
class QaReport:
    title: str
    validation: ValidationResult

    def summary(self) -> str:
        return (
            f"{self.title}: errors={len(self.validation.errors)} "
            f"warnings={len(self.validation.warnings)} "
            f"blocks_release={self.validation.blocks_release}"
        )
