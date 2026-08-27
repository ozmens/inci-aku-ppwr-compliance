"""Technical File package report stub."""

from __future__ import annotations

from .base import BaseReport


class TechnicalFileReport(BaseReport):
    report_name = "TECHNICAL_FILE"

    def plan(self) -> dict:
        return {
            "report": self.report_name,
            "sources": ["TECHNICAL_FILE", "DOCUMENT_LINK", "DOCUMENT_LIBRARY"],
            "subject_xor": [
                "COMPONENT_ID",
                "PACKAGING_CONFIGURATION_ID",
                "TRANSPORT_CONFIGURATION_ID",
            ],
            "status": "architecture_only",
        }
