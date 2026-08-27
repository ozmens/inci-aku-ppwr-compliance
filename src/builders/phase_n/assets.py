"""Phase N — extract approved İnci Akü logo from Golden DOCX (read-only)."""

from __future__ import annotations

import zipfile
from pathlib import Path


def extract_inci_aku_logo(project_root: Path, dest_dir: Path) -> Path:
    """Copy logo bytes from Golden Technical File media only. Never modifies Golden."""
    golden = (
        project_root
        / "templates"
        / "word_golden"
        / "01_Technical_File_GOLDEN.docx"
    )
    if not golden.exists():
        raise FileNotFoundError(f"Golden DOCX missing: {golden}")
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / "inci_aku_logo.png"
    with zipfile.ZipFile(golden) as zf:
        data = zf.read("word/media/image1.png")
    dest.write_bytes(data)
    return dest
