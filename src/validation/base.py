"""Validation result primitives."""

from __future__ import annotations

from dataclasses import dataclass, field

from utils.constants import Severity


@dataclass(slots=True)
class ValidationIssue:
    code: str
    severity: Severity
    object_type: str
    object_id: str
    message: str
    field: str | None = None
    source_reference: str | None = None


@dataclass(slots=True)
class ValidationResult:
    issues: list[ValidationIssue] = field(default_factory=list)

    def add(self, issue: ValidationIssue) -> None:
        self.issues.append(issue)

    @property
    def errors(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.severity == Severity.ERROR]

    @property
    def warnings(self) -> list[ValidationIssue]:
        return [i for i in self.issues if i.severity == Severity.WARNING]

    @property
    def blocks_release(self) -> bool:
        return bool(self.errors)

    def merge(self, other: ValidationResult) -> ValidationResult:
        self.issues.extend(other.issues)
        return self
