"""Revision control helpers (customer delivery R00 / Rev.00 until changed)."""

from __future__ import annotations

from dataclasses import dataclass
import re


@dataclass(frozen=True, slots=True)
class Revision:
    number: int
    code: str  # R00
    display: str  # Rev.00


class RevisionService:
    def __init__(self, default_number: int = 0) -> None:
        self.default_number = default_number

    def from_number(self, number: int) -> Revision:
        if number < 0:
            raise ValueError("revision number must be >= 0")
        code = f"R{number:02d}"
        display = f"Rev.{number:02d}"
        return Revision(number=number, code=code, display=display)

    def default(self) -> Revision:
        return self.from_number(self.default_number)

    def bump(self, current: Revision) -> Revision:
        return self.from_number(current.number + 1)

    def parse(self, text: str) -> Revision:
        m = re.search(r"(\d+)", text)
        if not m:
            return self.default()
        return self.from_number(int(m.group(1)))
