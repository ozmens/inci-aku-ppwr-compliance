"""Render one component packaging PDF then exit (no builders package import)."""
from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    args = ap.parse_args()
    docx = Path(args.docx).resolve()
    pdf = docx.with_suffix(".pdf")
    subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
    time.sleep(1.5)
    if pdf.exists():
        try:
            pdf.unlink()
        except Exception:
            pass

    import pythoncom
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    word = None
    doc = None
    ok = False
    err = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(docx), ReadOnly=True)
        doc.ExportAsFixedFormat(
            OutputFileName=str(pdf),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            CreateBookmarks=0,
        )
        ok = pdf.exists() and pdf.stat().st_size > 0
    except Exception as exc:  # noqa: BLE001
        err = str(exc)
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    print("OK" if ok else f"FAIL {err}", flush=True)
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
