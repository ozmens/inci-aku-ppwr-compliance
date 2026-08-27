"""Excel utility placeholders (Phase D — no workbook generation)."""

from __future__ import annotations

from pathlib import Path


class ExcelGenerationNotEnabledError(RuntimeError):
    """Raised when Excel generation is attempted during architecture phase."""


def assert_excel_allowed(enabled: bool) -> None:
    if not enabled:
        raise ExcelGenerationNotEnabledError(
            "Excel access is disabled. "
            "Use Phase E `python build.py --excel-template` or enable the repository flag."
        )


def ensure_parent(path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    return path
