"""Shared web/desktop mode detection for backend modules."""

from __future__ import annotations

import os


def is_web_mode() -> bool:
    return os.environ.get("INCI_PPWR_WEB", "").strip().lower() in {"1", "true", "yes"} or bool(
        os.environ.get("RENDER")
    )
