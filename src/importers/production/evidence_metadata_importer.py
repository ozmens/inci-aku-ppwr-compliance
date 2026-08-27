"""Supplier evidence archive metadata importer (no binary mutation)."""

from __future__ import annotations

from pathlib import Path

from .source_inventory import sha256_file
from .staging import StagingBundle


def import_evidence_metadata(path: Path, bundle: StagingBundle) -> StagingBundle:
    if not path.exists():
        return bundle
    digest = sha256_file(path)
    bundle.discrepancies.append(
        {
            "severity": "INFO",
            "object_type": "EVIDENCE_ARCHIVE",
            "object_id": path.name,
            "field": "sha256",
            "golden_value": "",
            "secondary_value": digest,
            "source_file": path.name,
            "source_reference": "LEVEL_4_EVIDENCE_ARCHIVE",
            "resolution": "INVENTORIED_ONLY",
            "resolution_basis": (
                "Archive not modified. Component-to-document links not guessed. "
                "Drawings/photos remain PENDING."
            ),
            "blocking": False,
        }
    )
    bundle.level2_meta.setdefault("evidence_archives", []).append(
        {"file": path.name, "sha256": digest, "size": path.stat().st_size}
    )
    return bundle
