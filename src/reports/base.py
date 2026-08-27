"""Shared report base class."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class BaseReport(ABC):
    report_name: str = "BASE"

    @abstractmethod
    def plan(self) -> dict:
        """Return metadata describing the future report output."""

    def render(self, output_dir: Path) -> Path:
        raise NotImplementedError(
            f"{self.report_name} rendering is reserved for a later phase "
            "(ENABLE_REPORT_EXPORT)."
        )
