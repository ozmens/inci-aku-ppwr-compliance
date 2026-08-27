"""Product mapping validation."""

from __future__ import annotations

from models.product import Product
from services.product_mapping_service import ProductMappingService
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class MappingValidator:
    def __init__(self, mapping: ProductMappingService) -> None:
        self.mapping = mapping

    def validate(self, products: list[Product]) -> ValidationResult:
        result = ValidationResult()
        for orphan in self.mapping.orphan_products(products):
            result.add(
                ValidationIssue(
                    code="MAP-01",
                    severity=Severity.WARNING,
                    object_type="PRODUCT",
                    object_id=orphan.product_code,
                    message="Orphan product mapping — no final packaging configuration linked",
                    field="linked_final_configuration_ids",
                )
            )
        return result
