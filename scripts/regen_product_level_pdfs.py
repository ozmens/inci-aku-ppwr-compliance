"""Force-regenerate all product-level PDFs from fixed DOCX (OpenAndRepair).

Uses Word COM only for ExportAsFixedFormat — no interactive UI intended.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import pythoncom
import win32com.client  # type: ignore

ROOT = Path(__file__).resolve().parents[1]
PRODUCT_SETS = (
    ROOT
    / "output"
    / "INCI_AKU_PPWR_STARTER_PRODUCT_LEVEL_CUSTOMER_DELIVERY_REV00_CANDIDATE"
    / "01_PRODUCT_DOCUMENT_SETS"
)
LOG = ROOT / "output" / "_product_level_pdf_regen.log"
STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]


def kill() -> None:
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, check=False)
    time.sleep(2)


def jobs() -> list[tuple[Path, Path]]:
    out: list[tuple[Path, Path]] = []
    for folder in sorted(p for p in PRODUCT_SETS.iterdir() if p.is_dir()):
        for stem in STEMS:
            docx = folder / f"{stem}.docx"
            pdf = folder / f"{stem}.pdf"
            if docx.exists():
                out.append((docx, pdf))
    return out


def convert_one(docx: Path, pdf: Path) -> bool:
    tmpdir = Path(tempfile.mkdtemp(prefix="ppwr_regen_"))
    tmp_docx = tmpdir / "in.docx"
    tmp_pdf = tmpdir / "out.pdf"
    word = None
    doc = None
    try:
        shutil.copy2(docx, tmp_docx)
        if pdf.exists():
            try:
                pdf.unlink()
            except Exception:
                pass
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(tmp_docx),
            False,
            True,
            False,
            "",
            "",
            False,
            "",
            "",
            0,
            0,
            True,
            True,  # OpenAndRepair
        )
        doc.ExportAsFixedFormat(
            OutputFileName=str(tmp_pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=0,
        )
        if tmp_pdf.exists() and tmp_pdf.stat().st_size > 0:
            shutil.copy2(tmp_pdf, pdf)
            return True
        return False
    finally:
        try:
            if doc is not None:
                doc.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)


def main() -> int:
    LOG.write_text("", encoding="utf-8")
    all_jobs = jobs()
    print(f"total={len(all_jobs)}", flush=True)
    ok = fail = 0
    batch = 0
    kill()
    for i, (docx, pdf) in enumerate(all_jobs, 1):
        try:
            if convert_one(docx, pdf):
                ok += 1
            else:
                fail += 1
                LOG.open("a", encoding="utf-8").write(f"EMPTY {docx}\n")
        except Exception as exc:  # noqa: BLE001
            fail += 1
            LOG.open("a", encoding="utf-8").write(f"FAIL {docx}: {exc}\n")
            kill()
            time.sleep(3)
        if i % 16 == 0:
            batch += 1
            print(f"progress {i}/{len(all_jobs)} ok={ok} fail={fail}", flush=True)
            kill()
            time.sleep(2)
    kill()
    print(f"DONE ok={ok} fail={fail}", flush=True)
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
