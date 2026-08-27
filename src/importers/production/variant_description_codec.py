"""Deterministic Variant Basis TR/EN ↔ PACKAGING_CONFIGURATION.DESCRIPTION codec."""

from __future__ import annotations

import re
from dataclasses import dataclass


class VariantDescriptionCodecError(ValueError):
    """Raised when DESCRIPTION is not a valid codec payload."""


@dataclass(frozen=True, slots=True)
class VariantDescriptionCodec:
    """
    Lossless serialize/deserialize for Schema 1.0.0 DESCRIPTION field.

    Format (delimiter-safe):
      TR: <text> | EN: <text>

    If TR/EN text contains ' | EN: ' or leading 'TR: ', those sequences are
    escaped with a private-use marker during serialize and restored on deserialize.
    """

    _SEP = " | EN: "
    _TR_PREFIX = "TR: "
    _ESC_SEP = "\uE000ENSEP\uE001"
    _ESC_TR = "\uE000TRPFX\uE001"

    def serialize(self, tr: str, en: str) -> str:
        tr_s = self._escape(str(tr or "").strip())
        en_s = self._escape(str(en or "").strip())
        if not tr_s and not en_s:
            raise VariantDescriptionCodecError("Variant Basis TR and EN are both empty")
        return f"{self._TR_PREFIX}{tr_s}{self._SEP}{en_s}"

    def deserialize(self, description: str | None) -> tuple[str, str]:
        text = str(description or "").strip()
        if not text.startswith(self._TR_PREFIX):
            raise VariantDescriptionCodecError(
                f"DESCRIPTION must start with {self._TR_PREFIX!r}"
            )
        body = text[len(self._TR_PREFIX) :]
        if self._SEP not in body:
            raise VariantDescriptionCodecError(
                f"DESCRIPTION missing separator {self._SEP!r}"
            )
        # Split on last occurrence of separator after TR body —
        # escaped payloads never contain raw separator.
        tr_esc, en_esc = body.split(self._SEP, 1)
        return self._unescape(tr_esc), self._unescape(en_esc)

    def roundtrip_ok(self, tr: str, en: str) -> bool:
        return self.deserialize(self.serialize(tr, en)) == (tr.strip(), en.strip())

    def _escape(self, value: str) -> str:
        return (
            value.replace(self._SEP, self._ESC_SEP).replace(self._TR_PREFIX, self._ESC_TR)
        )

    def _unescape(self, value: str) -> str:
        return value.replace(self._ESC_SEP, self._SEP).replace(self._ESC_TR, self._TR_PREFIX)


_CODEC = VariantDescriptionCodec()


def serialize_variant_description(tr: str, en: str) -> str:
    return _CODEC.serialize(tr, en)


def deserialize_variant_description(description: str | None) -> tuple[str, str]:
    return _CODEC.deserialize(description)
