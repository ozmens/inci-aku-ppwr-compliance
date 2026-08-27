"""Rebuild final customer delivery ZIP."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path

ROOT = Path(r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS")
FINAL = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL"
ZIP = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL.zip"
SHA = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL_SHA256.txt"
ENG = FINAL / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
DOC = FINAL / "01_DOCUMENT_SETS"


def main() -> None:
    assert ENG.exists(), "engine missing"
    sets = [p for p in DOC.iterdir() if p.is_dir()]
    words = [p for p in DOC.rglob("*.docx") if not p.name.startswith("~$")]
    pdfs = [p for p in DOC.rglob("*.pdf") if p.stat().st_size > 0]
    assert len(sets) == 287, len(sets)
    assert len(words) == 1148, len(words)
    assert len(pdfs) == 1148, len(pdfs)

    (FINAL / "00_AC_DOCUMENT_ENGINE.cmd").write_text(
        "@echo off\r\n"
        "cd /d \"%~dp0\"\r\n"
        "start \"\" \"%~dp000_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\"\r\n",
        encoding="utf-8",
    )
    (FINAL / "00_READ_ME_NASIL_ACILIR.txt").write_text(
        "INCI AKU PPWR STARTER MUSTERI PAKETI Rev.00\r\n"
        "========================================\r\n"
        "1) Bu ZIP dosyasina CIFT TIKLAMAYIN.\r\n"
        "2) Sag tik > Tumunu ayikla / Extract All.\r\n"
        "3) 00_AC_DOCUMENT_ENGINE.cmd dosyasina cift tiklayin\r\n"
        "   VEYA acin: 00_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\r\n"
        "4) HOME / SEARCH / DOCUMENT CENTER kullanin.\r\n"
        "5) Salt Okunur acildiysa linkler calismaz — kapatip Extract ile acin.\r\n"
        "\r\n"
        "Icerik:\r\n"
        "- 00_CONTROL : Document Engine\r\n"
        "- 01_DOCUMENT_SETS : 287 set x 4 Word + 4 PDF\r\n"
        "- 02_OPTIONAL_EVIDENCE\r\n"
        "- 03_ARCHIVE\r\n"
        "- 04_DOMESTIC_42_DATA_GAP : belgesi uretilmeyen 42 urun\r\n",
        encoding="utf-8",
    )

    for p in (FINAL / "00_CONTROL").glob("*"):
        if p.name.startswith("~$"):
            p.unlink(missing_ok=True)
        elif p.name != "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx" and p.suffix.lower() in {
            ".xlsx",
            ".tmp",
        }:
            if any(x in p.name.upper() for x in ("QMS", "CANDIDATE", "DEBUG", "SMOKE", "PREV", "FIXED", "REGEN")):
                p.unlink(missing_ok=True)
                print("removed", p.name)

    if ZIP.exists():
        ZIP.unlink()
    with zipfile.ZipFile(ZIP, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in FINAL.rglob("*"):
            if p.is_file() and not p.name.startswith("~$"):
                zf.write(p, p.relative_to(FINAL).as_posix())

    h = hashlib.sha256()
    with ZIP.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    digest = h.hexdigest()
    SHA.write_text(digest + "\n", encoding="utf-8")
    size_mb = ZIP.stat().st_size / (1024 * 1024)
    print("SETS", len(sets))
    print("WORD", len(words))
    print("PDF", len(pdfs))
    print("ZIP", ZIP)
    print("SIZE_MB", round(size_mb, 1))
    print("SHA256", digest)


if __name__ == "__main__":
    main()
