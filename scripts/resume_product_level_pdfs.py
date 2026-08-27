"""Resume missing product-level PDFs only (Word already generated)."""

from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))

from builders.phase_i.render_batch import render_docx_batch
from generate_product_level_delivery import PRODUCT_SETS, CONTROL, build_engine, list_controlled_products
from ppwr_engine.starter_loader import StarterMasterLoader
from generate_product_level_delivery import MASTER

STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]
PDF_CHUNK = 30


def main() -> None:
    jobs: list[tuple[Path, Path]] = []
    folders = sorted([p for p in PRODUCT_SETS.iterdir() if p.is_dir()], key=lambda p: p.name)
    print(f"Folders: {len(folders)}", flush=True)
    for folder in folders:
        for stem in STEMS:
            docx = folder / f"{stem}.docx"
            pdf = folder / f"{stem}.pdf"
            if not docx.exists():
                print(f"MISSING DOCX {docx}", flush=True)
                continue
            if (not pdf.exists()) or pdf.stat().st_size == 0:
                jobs.append((docx, pdf))
    print(f"Missing/zero PDFs to render: {len(jobs)}", flush=True)
    if not jobs:
        print("Nothing to do", flush=True)
        return

    ok = 0
    fail = 0
    for i in range(0, len(jobs), PDF_CHUNK):
        chunk = jobs[i : i + PDF_CHUNK]
        print(f"PDF resume chunk {i // PDF_CHUNK + 1}/{(len(jobs)+PDF_CHUNK-1)//PDF_CHUNK} size={len(chunk)}", flush=True)
        results = render_docx_batch(chunk, progress_every=10, log=[])
        for r in results:
            p = Path(r["pdf"])
            if p.exists() and p.stat().st_size > 0:
                ok += 1
            else:
                fail += 1
                print(f"FAIL {r.get('docx')} err={r.get('error')}", flush=True)
        # hard kill word leftovers between chunks to avoid hang accumulation
        try:
            import subprocess

            subprocess.run(
                ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                capture_output=True,
                check=False,
            )
        except Exception:
            pass
        print(f"  cumulative ok={ok} fail={fail}", flush=True)

    # rebuild engine after PDFs ready
    loader = StarterMasterLoader(MASTER)
    loader.open()
    products = list_controlled_products(loader)
    sets = loader.list_controlled_set_codes()
    loader.close()
    print("Rebuilding Document Engine…", flush=True)
    build_engine(products, set_count=len(sets))

    word_n = len(list(PRODUCT_SETS.rglob("*.docx")))
    pdf_n = len([p for p in PRODUCT_SETS.rglob("*.pdf") if p.stat().st_size > 0])
    report = {
        "products": len(folders),
        "word": word_n,
        "pdf": pdf_n,
        "resume_ok": ok,
        "resume_fail": fail,
    }
    CONTROL.mkdir(parents=True, exist_ok=True)
    (CONTROL / "PDF_RESUME_REPORT.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, indent=2), flush=True)


if __name__ == "__main__":
    main()
