"""Text / filename helpers (Windows-safe)."""

from __future__ import annotations

import re

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def windows_safe_filename(name: str, max_len: int = 180) -> str:
    cleaned = _UNSAFE.sub("_", name.strip())
    cleaned = cleaned.replace(" ", "_")
    return cleaned[:max_len] or "unnamed"


def normalize_ws(text: str | None) -> str:
    if not text:
        return ""
    return " ".join(str(text).split())
