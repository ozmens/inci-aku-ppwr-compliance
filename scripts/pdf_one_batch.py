"""Render up to N missing product-level PDFs then exit cleanly."""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

PRODUCT_SETS = (
    ROOT
    / "output"
    / "INCI_AKU_PPWR_STARTER_PRODUCT_LEVEL_CUSTOMER_DELIVERY_REV00_CANDIDATE"
    / "01_PRODUCT_DOCUMENT_SETS"
)
STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]
SKIP_FILE = ROOT / "output" / "_pdf_skip_paths.txt"
FAIL_COUNTS = ROOT / "output" / "_pdf_fail_counts.txt"


def kill_word() -> None:
    subprocess.run(
        ["taskkill", "/F", "/IM", "WINWORD.EXE"],
        capture_output=True,
        check=False,
    )
    time.sleep(1.5)


def _rel_key(docx: Path) -> str:
    try:
        return str(docx.resolve().relative_to(PRODUCT_SETS.resolve())).replace("\\", "/")
    except Exception:
        return f"{docx.parent.name}/{docx.name}"


def load_skips() -> set[str]:
    if not SKIP_FILE.exists():
        return set()
    return {
        ln.strip().replace("\\", "/")
        for ln in SKIP_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
        if ln.strip() and not ln.strip().startswith("#")
    }


def load_fail_counts() -> dict[str, int]:
    counts: dict[str, int] = {}
    if not FAIL_COUNTS.exists():
        return counts
    for ln in FAIL_COUNTS.read_text(encoding="utf-8", errors="replace").splitlines():
        ln = ln.strip()
        if not ln or "|" not in ln:
            continue
        key, raw = ln.split("|", 1)
        try:
            counts[key.replace("\\", "/")] = int(raw)
        except ValueError:
            continue
    return counts


def save_fail_counts(counts: dict[str, int]) -> None:
    FAIL_COUNTS.parent.mkdir(parents=True, exist_ok=True)
    lines = [f"{k}|{counts[k]}" for k in sorted(counts)]
    FAIL_COUNTS.write_text("\n".join(lines) + ("\n" if lines else ""), encoding="utf-8")


def add_skip(key: str) -> None:
    skips = load_skips()
    if key in skips:
        return
    SKIP_FILE.parent.mkdir(parents=True, exist_ok=True)
    with SKIP_FILE.open("a", encoding="utf-8") as f:
        f.write(key + "\n")
    print(f"SKIP_ADD {key}", flush=True)


def all_missing() -> list[tuple[Path, Path]]:
    jobs: list[tuple[Path, Path]] = []
    if not PRODUCT_SETS.exists():
        return jobs
    for folder in sorted(PRODUCT_SETS.iterdir(), key=lambda p: p.name):
        if not folder.is_dir():
            continue
        for stem in STEMS:
            docx = folder / f"{stem}.docx"
            pdf = folder / f"{stem}.pdf"
            if docx.exists() and ((not pdf.exists()) or pdf.stat().st_size == 0):
                jobs.append((docx, pdf))
    return jobs


def missing(limit: int | None = None, include_skipped: bool = False) -> list[tuple[Path, Path]]:
    skips = load_skips()
    jobs = all_missing()
    if not include_skipped and skips:
        primary = [(d, p) for d, p in jobs if _rel_key(d) not in skips]
        deferred = [(d, p) for d, p in jobs if _rel_key(d) in skips]
        # Keep skipped at the very end so normal work proceeds
        jobs = primary + deferred
    if limit is not None:
        # Prefer non-skipped within the batch window
        if skips:
            primary = [(d, p) for d, p in jobs if _rel_key(d) not in skips]
            if primary:
                return primary[:limit]
        return jobs[:limit]
    return jobs


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=20)
    ap.add_argument(
        "--include-skipped",
        action="store_true",
        help="Also attempt previously skipped hard-fail docs",
    )
    args = ap.parse_args()

    skips = load_skips()
    total_all = len(all_missing())
    total_active = len([1 for d, _ in all_missing() if _rel_key(d) not in skips])
    jobs = missing(args.limit, include_skipped=args.include_skipped)
    # If only skipped remain and not include_skipped, try them once at end
    if not jobs and total_all > 0 and not args.include_skipped:
        jobs = missing(args.limit, include_skipped=True)
        print("RETRY_SKIPPED_BATCH", flush=True)

    print(
        f"missing_total~{total_active} missing_all~{total_all} this_batch={len(jobs)}",
        flush=True,
    )
    if not jobs:
        print("DONE", flush=True)
        return 0

    import pythoncom
    import win32com.client  # type: ignore

    kill_word()
    pythoncom.CoInitialize()
    word = None
    ok = 0
    fail = 0
    fail_counts = load_fail_counts()

    def new_word():
        kill_word()
        time.sleep(2)
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        return app

    try:
        for attempt in range(5):
            try:
                word = new_word()
                break
            except Exception as exc:  # noqa: BLE001
                print(f"WORD_START_FAIL attempt={attempt+1}: {exc}", flush=True)
                time.sleep(3)
        if word is None:
            print("batch_ok=0 batch_fail=0 WORD_UNAVAILABLE", flush=True)
            return 1
        for docx, pdf in jobs:
            doc = None
            key = _rel_key(docx)
            try:
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
                if pdf.exists() and pdf.stat().st_size > 0:
                    ok += 1
                    if key in fail_counts:
                        fail_counts.pop(key, None)
                else:
                    fail += 1
                    print(f"FAIL empty {key}", flush=True)
                    fail_counts[key] = fail_counts.get(key, 0) + 1
                    if fail_counts[key] >= 3:
                        add_skip(key)
            except Exception as exc:  # noqa: BLE001
                fail += 1
                print(f"FAIL {key}: {exc}", flush=True)
                fail_counts[key] = fail_counts.get(key, 0) + 1
                if fail_counts[key] >= 3:
                    add_skip(key)
                try:
                    if doc is not None:
                        doc.Close(False)
                        doc = None
                except Exception:
                    pass
                try:
                    word.Quit()
                except Exception:
                    pass
                word = None
                for attempt in range(4):
                    try:
                        word = new_word()
                        break
                    except Exception as exc2:  # noqa: BLE001
                        print(f"WORD_RESTART_FAIL attempt={attempt+1}: {exc2}", flush=True)
                        time.sleep(3)
                if word is None:
                    print("WORD_UNAVAILABLE abort_batch", flush=True)
                    break
                continue
            finally:
                if doc is not None:
                    try:
                        doc.Close(False)
                    except Exception:
                        pass
    finally:
        save_fail_counts(fail_counts)
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

    print(f"batch_ok={ok} batch_fail={fail}", flush=True)
    return 0 if ok > 0 or fail == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
