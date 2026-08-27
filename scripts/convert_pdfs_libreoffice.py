"""Batch DOCX -> PDF conversion.

Prefers LibreOffice headless when installed (fast, no visible UI). Falls back
to Microsoft Word via COM automation when LibreOffice is not present on the
machine (common on Windows workstations that only have Office installed).
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MSWORD_MARKER = Path("MSWORD_COM")
SOFFICE_CANDIDATES = [
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
]
STEMS_ALL = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]


def _word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except OSError:
        return False


def find_soffice() -> Path:
    """Locate a usable DOCX->PDF converter. Returns MSWORD_MARKER for Word COM."""
    for p in SOFFICE_CANDIDATES:
        if p.exists():
            return p
    w = shutil.which("soffice") or shutil.which("soffice.exe")
    if w:
        return Path(w)
    if _word_available():
        return MSWORD_MARKER
    raise FileNotFoundError(
        "PDF dönüştürücü bulunamadı: LibreOffice (soffice.exe) veya Microsoft Word gerekli"
    )


def _convert_batch_word(batch: list[Path]) -> tuple[int, int]:
    """Convert DOCX -> PDF in place using Microsoft Word COM automation."""
    import pythoncom
    import win32com.client

    ok = fail = 0
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        for docx in batch:
            pdf = docx.with_suffix(".pdf")
            doc = None
            try:
                doc = word.Documents.Open(str(docx), ReadOnly=True)
                doc.SaveAs(str(pdf), FileFormat=17)  # wdFormatPDF
                ok += 1
            except Exception as e:  # noqa: BLE001
                fail += 1
                print(f"FAIL(Word) {docx.name}: {e}", flush=True)
            finally:
                if doc is not None:
                    doc.Close(False)
    finally:
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()
    return ok, fail


def needs_refresh(docx: Path, pdf: Path) -> bool:
    if not docx.exists():
        return False
    if not pdf.exists() or pdf.stat().st_size < 500:
        return True
    return docx.stat().st_mtime > pdf.stat().st_mtime + 1.0


def collect_docx(root: Path, stems: list[str], stale_only: bool) -> list[Path]:
    files: list[Path] = []
    if root.is_file() and root.suffix.lower() == ".docx":
        pdf = root.with_suffix(".pdf")
        if (not stale_only) or needs_refresh(root, pdf):
            return [root]
        return []
    for folder in sorted(p for p in root.iterdir() if p.is_dir()):
        for stem in stems:
            docx = folder / f"{stem}.docx"
            if not docx.exists():
                continue
            pdf = folder / f"{stem}.pdf"
            if stale_only and not needs_refresh(docx, pdf):
                continue
            files.append(docx)
        # also one level deeper (configs)
        for sub in sorted(p for p in folder.iterdir() if p.is_dir()):
            for stem in stems:
                docx = sub / f"{stem}.docx"
                if not docx.exists():
                    continue
                pdf = sub / f"{stem}.pdf"
                if stale_only and not needs_refresh(docx, pdf):
                    continue
                files.append(docx)
    # if root itself holds packs
    for stem in stems:
        docx = root / f"{stem}.docx"
        if docx.exists():
            pdf = root / f"{stem}.pdf"
            if (not stale_only) or needs_refresh(docx, pdf):
                files.append(docx)
    # dedupe
    seen = set()
    out = []
    for f in files:
        k = str(f.resolve()).lower()
        if k in seen:
            continue
        seen.add(k)
        out.append(f)
    return out


def convert_batch(soffice: Path, batch: list[Path], profile: Path) -> tuple[int, int]:
    """Convert a batch; returns (ok, fail). PDFs written next to each docx."""
    if not batch:
        return 0, 0
    # Group by parent dir — LO outdir is per-call; convert per-folder for correct placement
    by_dir: dict[Path, list[Path]] = {}
    for p in batch:
        by_dir.setdefault(p.parent, []).append(p)

    ok = fail = 0
    profile.mkdir(parents=True, exist_ok=True)
    uri = "file:///" + profile.resolve().as_posix()

    for outdir, files in by_dir.items():
        cmd = [
            str(soffice),
            f"-env:UserInstallation={uri}",
            "--headless",
            "--nologo",
            "--nolockcheck",
            "--nodefault",
            "--norestore",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(outdir),
            *[str(f) for f in files],
        ]
        try:
            r = subprocess.run(cmd, capture_output=True, text=True, timeout=max(120, 30 * len(files)))
        except subprocess.TimeoutExpired:
            fail += len(files)
            print(f"TIMEOUT dir={outdir} n={len(files)}", flush=True)
            continue
        for docx in files:
            pdf = docx.with_suffix(".pdf")
            if pdf.exists() and pdf.stat().st_size > 0 and (
                not needs_refresh(docx, pdf) or pdf.stat().st_mtime >= docx.stat().st_mtime - 2
            ):
                # Accept if pdf exists and is not clearly stale relative to this conversion
                # After convert, mtime should be new
                if pdf.stat().st_mtime + 5 >= docx.stat().st_mtime or pdf.stat().st_size > 500:
                    # re-check: if still older than docx by >2s, fail
                    if pdf.stat().st_mtime + 2 < docx.stat().st_mtime:
                        fail += 1
                    else:
                        ok += 1
                else:
                    fail += 1
            else:
                fail += 1
                if r.returncode != 0 and fail <= 3:
                    print(f"FAIL {docx.name} rc={r.returncode} {r.stderr[-300:]}", flush=True)
    return ok, fail


def convert_batch_via_temp(soffice: Path, batch: list[Path], profile: Path, chunk: int = 40) -> tuple[int, int]:
    """Copy chunk to temp, convert once, move PDFs back — fewer soffice startups."""
    if not batch:
        return 0, 0
    if soffice == MSWORD_MARKER:
        ok = fail = 0
        for i in range(0, len(batch), chunk):
            part = batch[i : i + chunk]
            o, f = _convert_batch_word(part)
            ok += o
            fail += f
            print(f"chunk {min(i+chunk,len(batch))}/{len(batch)} ok={ok} fail={fail}", flush=True)
        return ok, fail

    ok = fail = 0
    profile.mkdir(parents=True, exist_ok=True)
    uri = "file:///" + profile.resolve().as_posix()

    for i in range(0, len(batch), chunk):
        part = batch[i : i + chunk]
        with tempfile.TemporaryDirectory(prefix="lo_batch_") as tmp:
            tmp_dir = Path(tmp)
            mapping: list[tuple[Path, Path]] = []  # temp_docx, final_pdf
            for src in part:
                local = tmp_dir / f"{src.parent.name}__{src.name}"
                shutil.copy2(src, local)
                mapping.append((local, src.with_suffix(".pdf")))
            cmd = [
                str(soffice),
                f"-env:UserInstallation={uri}",
                "--headless",
                "--nologo",
                "--nolockcheck",
                "--nodefault",
                "--norestore",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                str(tmp_dir),
                *[str(m[0]) for m in mapping],
            ]
            try:
                subprocess.run(cmd, capture_output=True, text=True, timeout=max(180, 20 * len(part)))
            except subprocess.TimeoutExpired:
                fail += len(part)
                print(f"TIMEOUT chunk@{i} n={len(part)}", flush=True)
                continue
            for local, final_pdf in mapping:
                produced = local.with_suffix(".pdf")
                if not produced.exists():
                    fail += 1
                    continue
                final_pdf.parent.mkdir(parents=True, exist_ok=True)
                if final_pdf.exists():
                    try:
                        final_pdf.unlink()
                    except OSError:
                        pass
                shutil.move(str(produced), str(final_pdf))
                if final_pdf.exists() and final_pdf.stat().st_size > 0:
                    ok += 1
                else:
                    fail += 1
        print(f"chunk {min(i+chunk,len(batch))}/{len(batch)} ok={ok} fail={fail}", flush=True)
    return ok, fail


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, required=True)
    ap.add_argument("--stale-only", action="store_true")
    ap.add_argument("--tf-only", action="store_true")
    ap.add_argument("--chunk", type=int, default=40)
    ap.add_argument("--limit", type=int, default=0)
    args = ap.parse_args()

    soffice = find_soffice()
    stems = ["01_Technical_File"] if args.tf_only else STEMS_ALL
    files = collect_docx(args.root, stems, args.stale_only)
    if args.limit:
        files = files[: args.limit]
    print(f"SOFFICE={soffice}", flush=True)
    print(f"JOBS={len(files)} chunk={args.chunk} stale_only={args.stale_only} tf_only={args.tf_only}", flush=True)
    if not files:
        print("DONE ok=0 fail=0 (nothing to do)", flush=True)
        return 0

    profile = ROOT / "output" / "_lo_profile_pdf"
    t0 = time.time()
    ok, fail = convert_batch_via_temp(soffice, files, profile, chunk=args.chunk)
    print(f"DONE ok={ok} fail={fail} elapsed={time.time()-t0:.0f}s", flush=True)
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
