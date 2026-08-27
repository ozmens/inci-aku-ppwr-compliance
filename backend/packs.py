"""Build candidate PPWR packs (DOCX + PDF) under candidates/ only.

Never writes into frozen Inci_Aku_PPWR_PIMS delivery folders.
PDF via LibreOffice headless only — never WINWORD.
"""

from __future__ import annotations

import json
import os
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from fastapi import HTTPException

from masters import get_bom, get_product
from pipeline import photo_root
from web_mode import is_web_mode

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from bilingual_format import translate_component, translate_product  # noqa: E402
from convert_pdfs_libreoffice import convert_batch_via_temp, find_soffice  # noqa: E402
from premium_pack_from_runtime import build_premium_pack  # noqa: E402
import photo_annex as pa  # noqa: E402

CANDIDATES = ROOT / "candidates"
PACKS_ROOT = CANDIDATES / "packs"
INDEX_FILE = CANDIDATES / "packs_index.json"
# LibreOffice UserInstallation must be ASCII-safe (Unicode project path breaks soffice)
LO_PROFILE = Path(tempfile.gettempdir()) / "inci_ppwr_lo_profile_packs"

FROZEN = (
    "01_STARTER_INDIVIDUAL_DELIVERY_REV00",
    "02_INDUSTRIAL_DELIVERY_REV00",
    "03_CONTAINER_DELIVERY_REV00",
    "04_COMPONENT_SPARE_DELIVERY_REV00",
)

STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]

SCOPE_API = {
    "starter": "STARTER",
    "industrial": "INDUSTRIAL",
    "container": "CONTAINER",
    "component": "COMPONENT",
}


def _assert_candidates_only(path: Path) -> None:
    resolved = str(path.resolve()).replace("/", "\\").upper()
    for marker in FROZEN:
        if marker.upper() in resolved:
            raise HTTPException(403, f"Refusing write into frozen delivery: {marker}")
    cand = str(CANDIDATES.resolve()).replace("/", "\\").upper()
    if cand not in resolved:
        raise HTTPException(403, "Packs may only be written under candidates/")


def _safe_key(code: str) -> str:
    raw = (code or "").strip()
    if not raw or not re.fullmatch(r"[A-Za-z0-9._\-]+", raw):
        raise HTTPException(400, "Invalid product_code (use letters, digits, ._- only)")
    return raw


def _prep_photo_root() -> None:
    root = photo_root()
    pa.PHOTO_ROOT = root
    pa.MAPPING = root / "PHOTO_MAPPING_INDEX.csv"
    pa.clear_mapping_cache()
    # avoid writing missing log into frozen PIMS output
    pa.LOG = CANDIDATES / "_photo_annex_mapping.log"


def _bom_for_set(set_code: str) -> tuple[list[dict], float | None, dict]:
    bom = get_bom("starter", set_code)
    lines = []
    for line in bom.get("lines") or []:
        tr = str(line.get("description") or "")
        if "\n" in tr:
            tr = tr.split("\n", 1)[0].strip()
        lines.append(
            {
                "component_code": str(line.get("component_code") or "").strip(),
                "description": tr,
                "name_tr": tr,
                "name_en": translate_component(tr),
                "qty": line.get("qty"),
                "uom": str(line.get("uom") or "ADT").split("/")[0].strip() or "ADT",
                "unit_weight": line.get("unit_weight"),
                "line_weight": line.get("line_weight"),
            }
        )
    meta = bom.get("meta") or {}
    tare = meta.get("tare_kg")
    try:
        tare_f = float(tare) if tare is not None else None
    except (TypeError, ValueError):
        tare_f = None
    return lines, tare_f, meta


def _file_inventory(folder: Path) -> list[dict]:
    out = []
    for stem in STEMS:
        for ext, kind in (("docx", "WORD"), ("pdf", "PDF")):
            path = folder / f"{stem}.{ext}"
            out.append(
                {
                    "stem": stem,
                    "kind": kind,
                    "name": path.name,
                    "exists": path.exists() and path.stat().st_size > 200,
                    "size": path.stat().st_size if path.exists() else 0,
                }
            )
    return out


def _load_index() -> list[dict]:
    CANDIDATES.mkdir(parents=True, exist_ok=True)
    if not INDEX_FILE.exists():
        return []
    data = json.loads(INDEX_FILE.read_text(encoding="utf-8"))
    return list(data.get("packs") or [])


def _save_index(items: list[dict]) -> None:
    _assert_candidates_only(INDEX_FILE)
    payload = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "note": "Candidate packs only — frozen deliveries untouched",
        "packs": items,
    }
    INDEX_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def list_packs() -> dict:
    PACKS_ROOT.mkdir(parents=True, exist_ok=True)
    items = []
    for folder in sorted(p for p in PACKS_ROOT.iterdir() if p.is_dir()):
        meta_path = folder / "PACK_META.json"
        meta = {}
        if meta_path.exists():
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
        inv = _file_inventory(folder)
        items.append(
            {
                "product_code": folder.name,
                "path": str(folder),
                "meta": meta,
                "files": inv,
                "complete": all(f["exists"] for f in inv),
                "docx": sum(1 for f in inv if f["kind"] == "WORD" and f["exists"]),
                "pdf": sum(1 for f in inv if f["kind"] == "PDF" and f["exists"]),
            }
        )
    return {
        "packs_root": str(PACKS_ROOT),
        "count": len(items),
        "packs": items,
        "write_policy": "candidates_only",
        "pdf_engine": "LibreOffice headless",
    }


def build_candidate_pack(
    *,
    product_code: str,
    description: str | None = None,
    set_code: str | None = None,
    scope: str = "starter",
    skip_pdf: bool = False,
) -> dict:
    key = _safe_key(product_code)
    desc = (description or "").strip()
    sc = (set_code or "").strip()

    if not desc or not sc:
        # try master
        try:
            prod = get_product("starter", key)
            desc = desc or str(prod.get("description") or "")
            sc = sc or str(prod.get("set_code") or "")
        except Exception:
            pass
    if not desc:
        raise HTTPException(400, "description required (or known master product)")
    if not sc:
        raise HTTPException(400, "set_code required (or known master product with set)")

    lines, tare, meta = _bom_for_set(sc)
    if not lines:
        raise HTTPException(400, f"No BOM lines for set {sc}")

    dest = PACKS_ROOT / key
    _assert_candidates_only(dest)
    PACKS_ROOT.mkdir(parents=True, exist_ok=True)
    _prep_photo_root()

    scope_key = SCOPE_API.get(scope.lower(), scope.upper())
    final_id = str(meta.get("final_id") or f"IA-{sc}")

    info = build_premium_pack(
        dest,
        key=key,
        description_tr=desc,
        description_en=translate_product(desc),
        set_code=sc,
        bom_lines=lines,
        scope=scope_key,
        total_tare_kg=tare,
        config_id=final_id,
    )

    pdf_ok = pdf_fail = 0
    soffice = None
    if not skip_pdf:
        try:
            soffice = str(find_soffice())
        except FileNotFoundError as e:
            raise HTTPException(500, f"LibreOffice not found: {e}") from e
        docxs = [dest / f"{stem}.docx" for stem in STEMS]
        for d in docxs:
            if not d.exists():
                raise HTTPException(500, f"DOCX missing after build: {d.name}")
        LO_PROFILE.mkdir(parents=True, exist_ok=True)
        pdf_ok, pdf_fail = convert_batch_via_temp(
            find_soffice(), docxs, LO_PROFILE, chunk=4
        )

    inv = _file_inventory(dest)
    pack_meta = {
        "product_code": key,
        "description": desc,
        "set_code": sc,
        "final_id": final_id,
        "scope": scope_key,
        "tare_kg": tare,
        "bom_lines": len(lines),
        "photos": info.get("photos"),
        "pdf_ok": pdf_ok,
        "pdf_fail": pdf_fail,
        "soffice": soffice,
        "built_at": datetime.now(timezone.utc).isoformat(),
        "status": "CANDIDATE",
        "note": "Not part of frozen delivery Rev.00",
    }
    meta_path = dest / "PACK_META.json"
    _assert_candidates_only(meta_path)
    meta_path.write_text(json.dumps(pack_meta, ensure_ascii=False, indent=2), encoding="utf-8")

    # friendly open launcher (folder only — no frozen engine)
    launcher = dest / "00_OPEN_CANDIDATE_PACK.cmd"
    launcher.write_text(
        "@echo off\r\n"
        "cd /d \"%~dp0\"\r\n"
        "echo Candidate pack — not a frozen delivery.\r\n"
        "start \"\" explorer \"%cd%\"\r\n",
        encoding="utf-8",
    )

    # update index
    items = [p for p in _load_index() if p.get("product_code") != key]
    items.append(
        {
            "product_code": key,
            "set_code": sc,
            "path": str(dest),
            "built_at": pack_meta["built_at"],
            "complete": all(f["exists"] for f in inv),
        }
    )
    items.sort(key=lambda x: x.get("product_code") or "")
    _save_index(items)

    return {
        "pack": pack_meta,
        "folder": str(dest),
        "files": inv,
        "complete": all(f["exists"] for f in inv),
        "write_policy": "candidates_only",
    }


def get_pack(product_code: str) -> dict:
    key = _safe_key(product_code)
    folder = PACKS_ROOT / key
    if not folder.is_dir():
        raise HTTPException(404, f"Candidate pack not found: {key}")
    meta = {}
    mp = folder / "PACK_META.json"
    if mp.exists():
        meta = json.loads(mp.read_text(encoding="utf-8"))
    inv = _file_inventory(folder)
    return {
        "product_code": key,
        "folder": str(folder),
        "meta": meta,
        "files": inv,
        "complete": all(f["exists"] for f in inv),
    }


def open_pack_file(product_code: str, file_name: str) -> dict:
    path = resolve_pack_file(product_code, file_name)
    if is_web_mode() or not hasattr(os, "startfile"):
        key = _safe_key(product_code)
        return {
            "opened": str(path),
            "download": True,
            "download_url": (
                f"/api/packs/file?product_code={key}&file={file_name}"
            ),
        }
    os.startfile(str(path))
    return {"opened": str(path), "download": False}


def resolve_pack_file(product_code: str, file_name: str) -> Path:
    key = _safe_key(product_code)
    folder = (PACKS_ROOT / key).resolve()
    path = (folder / file_name).resolve()
    try:
        path.relative_to(folder)
    except ValueError as e:
        raise HTTPException(400, "Invalid path") from e
    if path.suffix.lower() not in {".docx", ".pdf"}:
        raise HTTPException(400, "Only DOCX/PDF open allowed")
    if not path.exists():
        raise HTTPException(404, f"Missing: {file_name}")
    return path


def open_packs_folder() -> dict:
    if is_web_mode() or not hasattr(os, "startfile"):
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    PACKS_ROOT.mkdir(parents=True, exist_ok=True)
    os.startfile(str(PACKS_ROOT))
    return {"opened": str(PACKS_ROOT)}


def delete_pack(product_code: str) -> dict:
    import shutil

    key = _safe_key(product_code)
    folder = PACKS_ROOT / key
    if not folder.is_dir():
        raise HTTPException(404, "Pack not found")
    _assert_candidates_only(folder)
    shutil.rmtree(folder)
    items = [p for p in _load_index() if p.get("product_code") != key]
    _save_index(items)
    return {"deleted": key, "count": len(items)}

