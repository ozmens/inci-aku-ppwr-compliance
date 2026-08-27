"""Retry last PDFs via short-path copy + OpenAndRepair."""

from __future__ import annotations

import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import pythoncom
import win32com.client  # type: ignore

ROOT = Path(__file__).resolve().parents[1]
SETS = (
    ROOT
    / "output"
    / "INCI_AKU_PPWR_STARTER_PRODUCT_LEVEL_CUSTOMER_DELIVERY_REV00_CANDIDATE"
    / "01_PRODUCT_DOCUMENT_SETS"
)
JOBS = [
    "1014048/04_Shipment_Statement.docx",
    "1014616/03_Label.docx",
    "1014631/02_EU_DoC.docx",
    "1014850/01_Technical_File.docx",
    "1015016/01_Technical_File.docx",
    "1015018/04_Shipment_Statement.docx",
    "1015132/02_EU_DoC.docx",
    "1015132/03_Label.docx",
    "1015334/03_Label.docx",
    "1015336/02_EU_DoC.docx",
]


def kill() -> None:
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True, check=False)
    time.sleep(4)


def one(rel: str) -> bool:
    src = SETS / rel.replace("/", "\\")
    dest_pdf = src.with_suffix(".pdf")
    print(f"TRY {rel}", flush=True)
    for attempt in range(1, 4):
        kill()
        pythoncom.CoInitialize()
        word = None
        doc = None
        tmpdir = Path(tempfile.mkdtemp(prefix="ppwrpdf_"))
        tmp_docx = tmpdir / "in.docx"
        tmp_pdf = tmpdir / "out.pdf"
        try:
            shutil.copy2(src, tmp_docx)
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            time.sleep(2)
            # OpenAndRepair=True
            doc = word.Documents.Open(
                str(tmp_docx),
                False,  # ConfirmConversions
                True,  # ReadOnly
                False,  # AddToRecentFiles
                "",  # PasswordDocument
                "",  # PasswordTemplate
                False,  # Revert
                "",  # WritePasswordDocument
                "",  # WritePasswordTemplate
                0,  # Format
                0,  # Encoding
                True,  # Visible (doc window - keep false via app)
                True,  # OpenAndRepair
            )
            time.sleep(1.5)
            doc.ExportAsFixedFormat(
                OutputFileName=str(tmp_pdf),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                CreateBookmarks=0,
            )
            if tmp_pdf.exists() and tmp_pdf.stat().st_size > 0:
                if dest_pdf.exists():
                    dest_pdf.unlink()
                shutil.copy2(tmp_pdf, dest_pdf)
                print(f"  attempt {attempt}: OK size={dest_pdf.stat().st_size}", flush=True)
                return True
            print(f"  attempt {attempt}: EMPTY", flush=True)
        except Exception as e:
            print(f"  attempt {attempt} ERR: {e}", flush=True)
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
            kill()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            shutil.rmtree(tmpdir, ignore_errors=True)
            time.sleep(5)
    return False


def main() -> int:
    ok = fail = 0
    for rel in JOBS:
        if one(rel):
            ok += 1
        else:
            fail += 1
    print(f"FINAL ok={ok} fail={fail}", flush=True)
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
