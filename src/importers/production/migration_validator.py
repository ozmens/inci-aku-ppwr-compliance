"""Migration validation and discrepancy writers."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook

from .normalizer import NormalizedStore
from .staging import StagingBundle
from .variant_description_codec import VariantDescriptionCodec


def validate_promotion_ready(store: NormalizedStore, bundle: StagingBundle) -> list[str]:
    errors = list(store.blocking_errors)
    if store.stats.get("configurations") != 247:
        errors.append(f"Expected 247 configs promoted, got {store.stats.get('configurations')}")
    if store.stats.get("starter") != 240:
        errors.append(f"Starter count {store.stats.get('starter')} != 240")
    if store.stats.get("industrial") != 3:
        errors.append(f"Industrial count {store.stats.get('industrial')} != 3")
    if store.stats.get("container") != 4:
        errors.append(f"Container count {store.stats.get('container')} != 4")
    if store.stats.get("bom_lines", 0) < 247:
        errors.append("BOM line count suspiciously low")
    # codec roundtrip sample
    codec = VariantDescriptionCodec()
    for cfg in bundle.configurations[:5]:
        if not codec.roundtrip_ok(cfg.variant_basis_tr, cfg.variant_basis_en):
            errors.append(f"Codec roundtrip failed for {cfg.packaging_set_code}")
    return errors


def write_discrepancies(
    discrepancies: list[dict[str, Any]],
    xlsx_path: Path,
    md_path: Path,
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "DISCREPANCIES"
    headers = [
        "Discrepancy ID",
        "Severity",
        "Object Type",
        "Object ID",
        "Field",
        "Golden Value",
        "Secondary Source Value",
        "Source File",
        "Source Reference",
        "Resolution",
        "Resolution Basis",
        "Blocking?",
    ]
    ws.append(headers)
    for i, d in enumerate(discrepancies, start=1):
        ws.append(
            [
                f"DISC-{i:04d}",
                d.get("severity"),
                d.get("object_type"),
                d.get("object_id"),
                d.get("field"),
                d.get("golden_value"),
                d.get("secondary_value"),
                d.get("source_file"),
                d.get("source_reference"),
                d.get("resolution"),
                d.get("resolution_basis"),
                "Yes" if d.get("blocking") else "No",
            ]
        )
    if not discrepancies:
        ws.append(["DISC-0000", "INFO", "NONE", "", "", "", "", "", "", "N/A", "No discrepancies logged", "No"])
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(xlsx_path)
    wb.close()

    lines = ["# Phase F Migration Discrepancies", "", f"Count: {len(discrepancies)}", ""]
    for i, d in enumerate(discrepancies, start=1):
        lines.append(
            f"- DISC-{i:04d} [{d.get('severity')}] {d.get('object_type')} `{d.get('object_id')}` "
            f"{d.get('field')}: golden={d.get('golden_value')!r} secondary={d.get('secondary_value')!r} "
            f"→ {d.get('resolution')} ({d.get('resolution_basis')})"
        )
    if not discrepancies:
        lines.append("- None")
    lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")
