"""Resilient product-level PDF supervisor (ASCII log, no Tee-Object)."""

from __future__ import annotations

import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LOG = ROOT / "output" / "_pdf_supervisor.log"
BATCH = ROOT / "scripts" / "pdf_one_batch.py"
FINALIZE = ROOT / "scripts" / "finalize_product_level_engine.py"


def log(msg: str) -> None:
    line = f"[{datetime.now().isoformat(timespec='seconds')}] {msg}"
    LOG.parent.mkdir(parents=True, exist_ok=True)
    with LOG.open("a", encoding="utf-8", errors="replace") as f:
        f.write(line + "\n")
    print(line, flush=True)


def run_py(script: Path, *args: str) -> str:
    p = subprocess.run(
        [sys.executable, "-u", str(script), *args],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    out = ((p.stdout or "") + (p.stderr or "")).strip()
    if out:
        with LOG.open("a", encoding="utf-8", errors="replace") as f:
            f.write(out + "\n")
        print(out, flush=True)
    return out


def kill_cms_pdf_jobs() -> None:
    """CMS PDF resume fights Word COM; stop it while product-level runs."""
    subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-Command",
            (
                "Get-CimInstance Win32_Process | "
                "Where-Object { $_.CommandLine -and "
                "($_.CommandLine -like '*resume_cms_pdfs*' -or $_.CommandLine -like '*continue_cms_pdfs*') } | "
                "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
            ),
        ],
        capture_output=True,
        check=False,
    )


def main() -> int:
    log("supervisor start (python loop)")
    idle = 0
    batches = 0
    while True:
        # Only sweep CMS every 5 batches to keep COM free without heavy overhead
        if batches % 5 == 0:
            kill_cms_pdf_jobs()
        batches += 1
        try:
            text = run_py(BATCH, "--limit", "16")
        except Exception as exc:  # noqa: BLE001
            log(f"batch exception: {exc!r}")
            time.sleep(8)
            continue

        if "WORD_UNAVAILABLE" in text:
            log("word unavailable — cooling down 20s")
            time.sleep(20)
            continue

        if "DONE" in text:
            log("all pdfs complete")
            run_py(FINALIZE)
            break

        if "batch_ok=0" in text:
            idle += 1
        else:
            idle = 0

        if idle >= 8:
            log("stopping after idle batches")
            run_py(FINALIZE)
            break

        time.sleep(2)

    log("supervisor end")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
