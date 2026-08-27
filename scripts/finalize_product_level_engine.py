"""Finalize Document Engine after PDF resume completes."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))

from generate_product_level_delivery import (  # noqa: E402
    CONTROL,
    MASTER,
    PRODUCT_SETS,
    build_engine,
    list_controlled_products,
)
from ppwr_engine.starter_loader import StarterMasterLoader  # noqa: E402

STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]


def main() -> None:
    folders = [p for p in PRODUCT_SETS.iterdir() if p.is_dir()]
    missing = 0
    for folder in folders:
        for stem in STEMS:
            pdf = folder / f"{stem}.pdf"
            if (not pdf.exists()) or pdf.stat().st_size == 0:
                missing += 1
    loader = StarterMasterLoader(MASTER)
    loader.open()
    products = list_controlled_products(loader)
    sets = loader.list_controlled_set_codes()
    loader.close()
    build_engine(products, set_count=len(sets))
    word_n = len(list(PRODUCT_SETS.rglob("*.docx")))
    pdf_n = len([p for p in PRODUCT_SETS.rglob("*.pdf") if p.stat().st_size > 0])
    report = {
        "products": len(folders),
        "sets": len(sets),
        "word": word_n,
        "pdf": pdf_n,
        "missing_pdf": missing,
        "expected_files": len(folders) * 4,
    }
    CONTROL.mkdir(parents=True, exist_ok=True)
    (CONTROL / "FINALIZE_REPORT.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, indent=2), flush=True)


if __name__ == "__main__":
    main()
