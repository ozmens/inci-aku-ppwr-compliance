"""
ProductionMigrationService — Phase F orchestration.

SOURCE → inventory → qualify → stage → normalize → validate → promote → QA
"""

from __future__ import annotations

import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from models.registry import SchemaRegistry

from .container_source_importer import import_container_source
from .evidence_metadata_importer import import_evidence_metadata
from .golden_register_importer import import_golden_register
from .industrial_source_importer import import_industrial_source
from .legacy_pims_importer import import_legacy_pims
from .migration_audit import dpp_traversal_check, write_migration_qa
from .migration_validator import validate_promotion_ready, write_discrepancies
from .normalizer import normalize_bundle
from .promoter import promote_to_workbook
from .qualify import qualify_golden_register
from .source_inventory import build_inventory, write_inventory_csv, write_inventory_md
from .source_reader import (
    find_evidence_archive,
    find_level1_golden,
    find_level2,
    find_level3,
    production_dir,
)
from .starter_source_importer import import_starter_source
from .staging import StagingBundle


@dataclass
class MigrationResult:
    success: bool
    run_id: str
    messages: list[str]
    production_workbook: Path | None = None


class ProductionMigrationService:
    def __init__(self, project_root: Path, settings) -> None:
        self.root = project_root
        self.settings = settings
        self.output = project_root / "output"
        self.production = production_dir(project_root)

    def run(self, *, run_tests: bool = True) -> MigrationResult:
        run_id = datetime.now(timezone.utc).strftime("PF-%Y%m%dT%H%M%SZ")
        messages: list[str] = []

        level1 = find_level1_golden(self.production)
        if level1 is None:
            return MigrationResult(False, run_id, ["Level-1 Golden Register not found under input/production/"])

        # Prefer explicit GOLDEN_VARIANTS_FINAL name when present
        preferred = self.production / "INCI_AKU_PPWR_Final_Configuration_Register_Rev00_GOLDEN_VARIANTS_FINAL.xlsx"
        if preferred.exists():
            level1 = preferred

        messages.append(f"Level-1 candidate: {level1.name}")
        qualification = qualify_golden_register(level1)
        messages.extend(qualification.summary_lines())
        if not qualification.passed:
            # still write inventory for diagnostics
            inv = build_inventory(self.production, qualification_pass=False)
            write_inventory_csv(inv, self.output / "PHASE_F_SOURCE_INVENTORY.csv")
            write_inventory_md(inv, self.output / "PHASE_F_SOURCE_INVENTORY.md", title="Phase F Source Inventory")
            return MigrationResult(False, run_id, messages + ["STOP: Level-1 qualification FAILED"])

        if not (
            qualification.total_configurations == 247
            and qualification.starter_count == 240
            and qualification.industrial_count == 3
            and qualification.container_count == 4
        ):
            return MigrationResult(
                False,
                run_id,
                messages + ["STOP: count gate failed (expected 247 / 240 / 3 / 4)"],
            )

        inv = build_inventory(
            self.production,
            qualification_pass=True,
            record_count_hints={
                level1.name: f"configs={qualification.total_configurations}; bom={qualification.exact_bom_rows}; products={qualification.product_map_rows}"
            },
        )
        # mark level1 role explicitly
        for row in inv:
            if row.file_name == level1.name:
                row.source_role = "LEVEL_1_GOLDEN_REGISTER"
                row.source_priority = 1
                row.notes = "Content-qualified PASS (247/240/3/4)"
                row.migration_status = "QUALIFIED"
        write_inventory_csv(inv, self.output / "PHASE_F_SOURCE_INVENTORY.csv")
        write_inventory_md(inv, self.output / "PHASE_F_SOURCE_INVENTORY.md", title="Phase F Source Inventory")

        bundle = StagingBundle()
        import_golden_register(level1, bundle)

        level2 = find_level2(self.production)
        if level2:
            import_legacy_pims(level2, bundle)
            messages.append(f"Level-2: {level2.name}")

        l3 = find_level3(self.production)
        if l3.get("starter"):
            import_starter_source(l3["starter"], bundle)
        if l3.get("industrial"):
            import_industrial_source(l3["industrial"], bundle)
        if l3.get("container"):
            import_container_source(l3["container"], bundle)

        evidence = find_evidence_archive(self.production)
        if evidence:
            import_evidence_metadata(evidence, bundle)

        store = normalize_bundle(bundle)
        errors = validate_promotion_ready(store, bundle)
        write_discrepancies(
            bundle.discrepancies,
            self.output / "PHASE_F_MIGRATION_DISCREPANCIES.xlsx",
            self.output / "PHASE_F_MIGRATION_DISCREPANCIES.md",
        )

        if errors:
            messages.extend(errors)
            write_migration_qa(
                self.output / "PHASE_F_MIGRATION_QA.md",
                run_id=run_id,
                qualification=qualification,
                inventory_rows=inv,
                store=store,
                bundle=bundle,
                production_path=self.output / "INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION.xlsx",
                test_rc=-1,
                dpp=dpp_traversal_check(bundle, store),
            )
            return MigrationResult(False, run_id, messages + ["STOP: normalization/validation blocking errors"])

        template = self.output / "INCI_AKU_PPWR_PIMS_Rev00.xlsx"
        if not template.exists():
            return MigrationResult(False, run_id, messages + [f"Missing Phase E template: {template}"])

        production_path = self.output / "INCI_AKU_PPWR_PIMS_Rev00_PRODUCTION.xlsx"
        registry = SchemaRegistry.load()
        promote_to_workbook(
            template_path=template,
            output_path=production_path,
            store=store,
            registry=registry,
        )
        messages.append(f"Production workbook: {production_path}")

        # Ensure blank template untouched size check
        if not template.exists():
            messages.append("WARNING: Phase E template missing after promote")

        dpp = dpp_traversal_check(bundle, store)
        messages.append(f"DPP traversal: {dpp}")

        test_rc = 0
        if run_tests:
            test_rc = subprocess.call(
                [sys.executable, "-m", "unittest", "discover", "-s", "tests", "-v"],
                cwd=str(self.root),
            )
            messages.append(f"Tests exit code: {test_rc}")

        write_migration_qa(
            self.output / "PHASE_F_MIGRATION_QA.md",
            run_id=run_id,
            qualification=qualification,
            inventory_rows=inv,
            store=store,
            bundle=bundle,
            production_path=production_path,
            test_rc=test_rc,
            dpp=dpp,
        )

        success = (
            test_rc == 0
            and dpp.get("ok") is True
            and store.stats.get("configurations") == 247
            and not store.blocking_errors
        )
        return MigrationResult(success, run_id, messages, production_path)
