"""Level-3 operational source importers — lineage inventory only (no identity overwrite)."""

from __future__ import annotations

from pathlib import Path

from .source_inventory import sha256_file
from .staging import StagingBundle


def _note(bundle: StagingBundle, path: Path, role: str) -> None:
    if not path.exists():
        return
    bundle.discrepancies.append(
        {
            "severity": "INFO",
            "object_type": "SOURCE_FILE",
            "object_id": path.name,
            "field": "role",
            "golden_value": "LEVEL_1_AUTHORITATIVE",
            "secondary_value": role,
            "source_file": path.name,
            "source_reference": sha256_file(path)[:16],
            "resolution": "LINEAGE_ONLY",
            "resolution_basis": "Operational source retained for lineage; not used to redefine final configs",
            "blocking": False,
        }
    )


def import_starter_source(path: Path, bundle: StagingBundle) -> StagingBundle:
    _note(bundle, path, "LEVEL_3_STARTER_OPERATIONAL")
    return bundle


def import_industrial_source(path: Path, bundle: StagingBundle) -> StagingBundle:
    _note(bundle, path, "LEVEL_3_INDUSTRIAL_OPERATIONAL")
    return bundle


def import_container_source(path: Path, bundle: StagingBundle) -> StagingBundle:
    _note(bundle, path, "LEVEL_3_CONTAINER_OPERATIONAL")
    return bundle
