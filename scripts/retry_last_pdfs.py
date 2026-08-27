"""Careful one-by-one retry for last hard-fail product-level PDFs."""

from __future__ import annotations

import subprocess
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
    time.sleep(4.0)


def one(rel: str) -> bool:
    docx = SETS / rel.replace("/", "\\")
    pdf = docx.with_suffix(".pdf")
    print(f"TRY {rel}", flush=True)
    for attempt in range(1, 4):
        kill()
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            time.sleep(2.0)
            if pdf.exists():
                try:
                    pdf.unlink()
                except Exception:
                    pass
            doc = word.Documents.Open(
                str(docx.resolve()),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                NoEncodingDialog=True,
            )
            time.sleep(1.5)
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf.resolve()),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    CreateBookmarks=0,
                )
            except Exception as e1:
                print(f"  export fail a{attempt}: {e1}", flush=True)
                doc.SaveAs2(str(pdf.resolve()), FileFormat=17)
            ok = pdf.exists() and pdf.stat().st_size > 0
            print(
                f"  attempt {attempt}: {'OK' if ok else 'EMPTY'} "
                f"size={pdf.stat().st_size if pdf.exists() else 0}",
                flush=True,
            )
            if ok:
                return True
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
