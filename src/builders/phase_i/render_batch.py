"""Efficient Word COM batch renderer for Phase I QA evidence."""

from __future__ import annotations

from pathlib import Path
from typing import Any


def render_docx_batch(
    jobs: list[tuple[Path, Path]],
    *,
    progress_every: int = 25,
    log: list[str] | None = None,
) -> list[dict[str, Any]]:
    """
    Open each DOCX once in a shared Word instance, export PDF, capture page count.

    jobs: list of (docx_path, pdf_out_path)
    """
    import pythoncom
    import win32com.client  # type: ignore

    results: list[dict[str, Any]] = []
    if not jobs:
        return results

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for i, (docx_path, pdf_out) in enumerate(jobs, start=1):
            row: dict[str, Any] = {
                "docx": str(docx_path),
                "pdf": str(pdf_out),
                "render_ok": False,
                "page_count": 0,
                "error": None,
            }
            doc = None
            try:
                pdf_out.parent.mkdir(parents=True, exist_ok=True)
                doc = word.Documents.Open(str(docx_path.resolve()), ReadOnly=True)
                row["page_count"] = int(doc.ComputeStatistics(2))
                doc.ExportAsFixedFormat(
                    OutputFileName=str(pdf_out.resolve()),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    CreateBookmarks=0,
                )
                row["render_ok"] = pdf_out.exists() and pdf_out.stat().st_size > 0
            except Exception as exc:  # noqa: BLE001
                row["error"] = str(exc)
                row["render_ok"] = False
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
            results.append(row)
            if log is not None and (i % progress_every == 0 or i == len(jobs)):
                ok = sum(1 for r in results if r["render_ok"])
                msg = f"Render progress {i}/{len(jobs)} ok={ok}"
                log.append(msg)
                print(msg, flush=True)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return results
