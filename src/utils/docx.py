"""
DOCX utility placeholders for future Golden Word templates.

Phase D: style configuration only — no document generation.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class DocumentStyleConfig:
    """Centralized style for future Tahoma / TR-EN Golden templates."""

    font_family: str = "Tahoma"
    body_pt: int = 10
    paragraph_spacing_pt: float = 6.0
    after_table_spacing_pt: float = 6.0
    turkish_first: bool = True
    english_italic: bool = True
    header_fill: str = "#1F4E79"
    header_font_color: str = "#FFFFFF"
    body_text_color: str = "#1A1A1A"
    white_text_only_on_dark: bool = True


class WordGenerationNotEnabledError(RuntimeError):
    """Raised when Word generation is attempted during architecture phase."""


def assert_word_allowed(enabled: bool) -> None:
    if not enabled:
        raise WordGenerationNotEnabledError(
            "Word generation is disabled. "
            "Enable ENABLE_WORD_PILOT_GENERATION for Phase G pilots only."
        )


def assert_word_pilot_allowed(*, pilot: bool, batch: bool) -> None:
    if batch:
        raise WordGenerationNotEnabledError("Word batch generation is disabled.")
    if not pilot:
        raise WordGenerationNotEnabledError("Word pilot generation is disabled.")


def default_style() -> DocumentStyleConfig:
    return DocumentStyleConfig()
