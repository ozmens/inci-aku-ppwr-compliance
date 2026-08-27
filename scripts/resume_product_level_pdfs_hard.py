"""Hardened PDF resume: small batches, per-batch Word restart, retries."""

from __future__ import annotations

import json
import subprocess
import sys
import time
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
BATCH = 12
MAX_PASSES = 6


def kill_word() -> None:
    subprocess.run(
        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
        capture_output=True,
        check=False,
    )
    time.sleep(2)


def missing_jobs() -> list[tuple[Path, Path]]:
    jobs = []
    for folder in sorted(PRODUCT_SETS.iterdir(), key=lambda p: p.name):
        if not folder.is_dir():
            continue
        for stem in STEMS:
            docx = folder / f"{stem}.docx"
            pdf = folder / f"{stem}.pdf"
            if docx.exists() and ((not pdf.exists()) or pdf.stat().st_size == 0):
                jobs.append((docx, pdf))
    return jobs


def render_one(word, docx: Path, pdf: Path) -> tuple[bool, str | None]:
    doc = None
    try:
        pdf.parent.mkdir(parents=True, exist_ok=True)
        if pdf.exists():
            try:
                pdf.unlink()
            except Exception:
                pass
        doc = word.Documents.Open(str(docx.resolve()), ReadOnly=True)
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf.resolve()),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=0,
        )
        ok = pdf.exists() and pdf.stat().st_size > 0
        return ok, None if ok else "empty_pdf"
    except Exception as exc:  # noqa: BLE001
        return False, str(exc)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass


def render_batch(jobs: list[tuple[Path, Path]]) -> tuple[int, int]:
    import pythoncom
    import win32com.client  # type: ignore

    ok = 0
    fail = 0
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for docx, pdf in jobs:
            good, err = render_one(word, docx, pdf)
            if good:
                ok += 1
            else:
                fail += 1
                print(f"FAIL {docx.name} @ {docx.parent.name}: {err}", flush=True)
                # Word COM often dies after RPC — recreate
                if err and ("-2147023174" in err or "RPC" in err.upper()):
                    try:
                        word.Quit()
                    except Exception:
                        pass
                    kill_word()
                    time.sleep(2)
                    word = win32com.client.DispatchEx("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    good2, err2 = render_one(word, docx, pdf)
                    if good2:
                        ok += 1
                        fail -= 1
                        print(f"  RETRY OK {docx.parent.name}/{docx.name}", flush=True)
                    else:
                        print(f"  RETRY FAIL {err2}", flush=True)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        kill_word()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return ok, fail


def main() -> None:
    kill_word()
    total_ok = 0
    total_fail = 0
    for pass_no in range(1, MAX_PASSES + 1):
        jobs = missing_jobs()
        print(f"PASS {pass_no}: missing={len(jobs)}", flush=True)
        if not jobs:
            break
        for i in range(0, len(jobs), BATCH):
            chunk = jobs[i : i + BATCH]
            n = (len(jobs) + BATCH - 1) // BATCH
            print(f"  batch {i // BATCH + 1}/{n} size={len(chunk)}", flush=True)
            ok, fail = render_batch(chunk)
            total_ok += ok
            total_fail += fail
            print(f"  batch ok={ok} fail={fail} cum_ok={total_ok}", flush=True)
            time.sleep(1)
        # if a pass made no progress, stop
        left = len(missing_jobs())
        print(f"PASS {pass_no} done; still missing={left}", flush=True)
        if left == len(jobs):
            print("No progress this pass — stopping.", flush=True)
            break

    left = missing_jobs()
    print(f"FINAL missing PDFs: {len(left)}", flush=True)

    loader = StarterMasterLoader(MASTER)
    loader.open()
    products = list_controlled_products(loader)
    sets = loader.list_controlled_set_codes()
    loader.close()
    print("Building Document Engine…", flush=True)
    build_engine(products, set_count=len(sets))

    word_n = len(list(PRODUCT_SETS.rglob("*.docx")))
    pdf_n = len([p for p in PRODUCT_SETS.rglob("*.pdf") if p.stat().st_size > 0])
    report = {
        "products": len([p for p in PRODUCT_SETS.iterdir() if p.is_dir()]),
        "word": word_n,
        "pdf": pdf_n,
        "missing_pdf": len(left),
        "resume_ok": total_ok,
        "resume_fail_events": total_fail,
    }
    CONTROL.mkdir(parents=True, exist_ok=True)
    (CONTROL / "PDF_RESUME_REPORT.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    print(json.dumps(report, indent=2), flush=True)


if __name__ == "__main__":
    main()
