"""
Export Engine — shipment package manifest (no Word generation).

Prepares folder layout + JSON manifest listing required package members
for a shipment archive / customer document pack.
"""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PACKAGE_MEMBERS = (
    {
        "code": "SHIPMENT_STATEMENT",
        "required": True,
        "description": "Shipment Statement (composition / materials / weights)",
        "source": "STATEMENT + STATEMENT_LINE + SHIPMENT_LINE",
    },
    {
        "code": "DECLARATION_OF_CONFORMITY",
        "required": True,
        "description": "DoC variant selected by Commercial Scenario (Incoterms)",
        "source": "DECLARATION_OF_CONFORMITY linked to Packaging Configuration TF",
    },
    {
        "code": "TECHNICAL_FILE_REFERENCE",
        "required": True,
        "description": "Technical File reference for Packaging Configuration (not Product)",
        "source": "TECHNICAL_FILE.PACKAGING_CONFIGURATION_ID",
    },
    {
        "code": "SHIPMENT_FREEZE",
        "required": True,
        "description": "Confirmed shipment freeze lines",
        "source": "SHIPMENT_LINE",
    },
    {
        "code": "EVIDENCE_PACK",
        "required": False,
        "description": "Supplier / evidence documents linked to TF or components",
        "source": "DOCUMENT_LIBRARY + DOCUMENT_LINK",
    },
)


class ExportEngine:
    """Build export manifests for shipment document packages."""

    def manifest(
        self,
        *,
        shipment_number: str,
        packaging_configuration_id: int | None = None,
        commercial_scenario_id: int | None = None,
        lot_number: str | None = None,
    ) -> dict[str, Any]:
        return {
            "engine": "Inci_Aku_PPWR_ExportEngine",
            "generated_at_utc": datetime.now(timezone.utc).isoformat(),
            "shipment_number": shipment_number,
            "lot_number": lot_number,
            "packaging_configuration_id": packaging_configuration_id,
            "commercial_scenario_id": commercial_scenario_id,
            "rules": {
                "primary_object": "PACKAGING_CONFIGURATION",
                "technical_file_owner": "PACKAGING_CONFIGURATION",
                "scenario_role": "DoC variant + shipment package only",
            },
            "package_members": list(PACKAGE_MEMBERS),
            "folder_layout": [
                f"shipments/{shipment_number}/manifest.json",
                f"shipments/{shipment_number}/statement/",
                f"shipments/{shipment_number}/declaration/",
                f"shipments/{shipment_number}/technical_file/",
                f"shipments/{shipment_number}/evidence/",
            ],
            "status": "READY_FOR_TEMPLATE_RENDER",
            "word_generation": False,
        }

    def write_manifest(self, output_dir: Path, shipment_number: str = "SAMPLE") -> Path:
        dest = output_dir / "shipments" / shipment_number
        dest.mkdir(parents=True, exist_ok=True)
        for sub in ("statement", "declaration", "technical_file", "evidence"):
            (dest / sub).mkdir(exist_ok=True)
            (dest / sub / ".keep").write_text("", encoding="utf-8")
        path = dest / "manifest.json"
        path.write_text(
            json.dumps(self.manifest(shipment_number=shipment_number), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
        return path


def write_export_engine_stub(output_dir: Path) -> Path:
    """Write sample export package layout under templates/export/."""
    engine = ExportEngine()
    return engine.write_manifest(output_dir, shipment_number="SAMPLE")
