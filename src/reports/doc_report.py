"""Declaration of Conformity report stub."""

from __future__ import annotations

from .base import BaseReport


class DeclarationOfConformityReport(BaseReport):
    report_name = "DECLARATION_OF_CONFORMITY"

    def plan(self) -> dict:
        return {
            "report": self.report_name,
            "sources": [
                "DECLARATION_OF_CONFORMITY",
                "TECHNICAL_FILE",
                "LEGAL_ENTITY",
                "PERSON",
                "DOCUMENT_LINK",
            ],
            "status": "architecture_only",
        }
