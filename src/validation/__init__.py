"""Validation package — architecture rules and Phase D validators."""

from .base import ValidationIssue, ValidationResult
from .runner import ValidationRunner
from .schema_validator import SchemaValidator
from .bom_validator import BomValidator
from .weight_validator import WeightValidator
from .mapping_validator import MappingValidator
from .id_validator import IdValidator
from .variant_validator import VariantValidator
from .document_validator import DocumentValidator

# Optional workbook validator (Phase E+)
try:
    from .workbook_validator import WorkbookValidator, ValidationReport
except ImportError:  # pragma: no cover
    WorkbookValidator = None  # type: ignore
    ValidationReport = None  # type: ignore

__all__ = [
    "ValidationIssue",
    "ValidationResult",
    "ValidationRunner",
    "SchemaValidator",
    "BomValidator",
    "WeightValidator",
    "MappingValidator",
    "IdValidator",
    "VariantValidator",
    "DocumentValidator",
    "WorkbookValidator",
    "ValidationReport",
]
