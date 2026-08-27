"""Document context / link validation."""

from __future__ import annotations

from models.document_context import DocumentContext
from services.document_link_service import DocumentLinkService
from utils.constants import Severity
from .base import ValidationIssue, ValidationResult


class DocumentValidator:
    def __init__(self, link_service: DocumentLinkService | None = None) -> None:
        self.links = link_service or DocumentLinkService()

    def validate_technical_file_privacy(self, ctx: DocumentContext) -> ValidationResult:
        result = ValidationResult()
        tf_ctx = ctx.for_technical_file()
        for field_name in ("customer_name", "oem_name", "customer_market"):
            if getattr(tf_ctx, field_name, None):
                result.add(
                    ValidationIssue(
                        code="DOC-01",
                        severity=Severity.ERROR,
                        object_type="TECHNICAL_FILE",
                        object_id=ctx.document_ids.technical_file_id,
                        message=f"Forbidden customer field present in Technical File context: {field_name}",
                        field=field_name,
                    )
                )
        return result

    def validate_links(self) -> ValidationResult:
        result = ValidationResult()
        for broken in self.links.broken_links():
            result.add(
                ValidationIssue(
                    code="DOC-02",
                    severity=Severity.ERROR,
                    object_type="DOCUMENT_LINK",
                    object_id=broken.document_link_id,
                    message="Broken document link — DOCUMENT_ID not in library",
                    field="document_id",
                )
            )
        return result
