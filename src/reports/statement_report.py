"""Shipment statement report stub."""

from __future__ import annotations

from .base import BaseReport


class StatementReport(BaseReport):
    report_name = "STATEMENT"

    def plan(self) -> dict:
        return {
            "report": self.report_name,
            "sources": ["STATEMENT", "STATEMENT_LINE", "STATEMENT_SHIPMENT", "SHIPMENT_LINE"],
            "grain": "STATEMENT × MATERIAL × PACKAGING_LEVEL × OWNERSHIP_TYPE",
            "status": "architecture_only",
        }
