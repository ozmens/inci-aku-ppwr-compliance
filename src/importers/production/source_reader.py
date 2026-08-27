"""Thin source reader helpers."""

from __future__ import annotations

from pathlib import Path


def production_dir(project_root: Path) -> Path:
    return project_root / "input" / "production"


def find_level1_golden(production: Path) -> Path | None:
    preferred = production / "INCI_AKU_PPWR_Final_Configuration_Register_Rev00_GOLDEN_VARIANTS_FINAL.xlsx"
    if preferred.exists():
        return preferred
    for p in production.glob("*Final_Configuration_Register*.xlsx"):
        if "GOLDEN" in p.name.upper() or "FINAL" in p.name.upper():
            return p
    matches = list(production.glob("*Final_Configuration_Register*.xlsx"))
    return matches[0] if matches else None


def find_level2(production: Path) -> Path | None:
    for pat in ("*BOM_CLOSED*.xlsx", "*OEM_ART5*.xlsx", "*PIMS_Data_Package*.xlsx"):
        hits = list(production.glob(pat))
        if hits:
            return hits[0]
    return None


def find_level3(production: Path) -> dict[str, Path | None]:
    def one(*pats: str) -> Path | None:
        for pat in pats:
            hits = list(production.glob(pat))
            if hits:
                return hits[0]
        return None

    return {
        "starter": one("*Mamul Ambalaj*", "*Mamul*"),
        "industrial": one("*End*striyel Ambalaj*", "*Endustriyel*"),
        "container": one("*Y*klemede*", "*Yukleme*"),
    }


def find_evidence_archive(production: Path) -> Path | None:
    hits = list(production.glob("*.7z")) + list(production.glob("*TEDAR*.7z"))
    return hits[0] if hits else None
