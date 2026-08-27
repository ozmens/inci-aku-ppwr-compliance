"""Software workspace — source of truth for packs & revisions.

Never writes frozen Inci_Aku_PPWR_PIMS deliveries.
Revision labels: Rev.00, Rev.01, … Status: DRAFT | ISSUED | SUPERSEDED.
Desktop multi-SKU export = ZIP only.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import tempfile
import zipfile
from datetime import datetime, timezone
from pathlib import Path

from fastapi import HTTPException

from masters import get_bom, get_product
from packs import (  # reuse BOM/photo/build helpers
    LO_PROFILE,
    STEMS,
    _bom_for_set,
    _file_inventory,
    _prep_photo_root,
    _safe_key,
)
import sys

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

from bilingual_format import translate_product  # noqa: E402
from convert_pdfs_libreoffice import convert_batch_via_temp, find_soffice  # noqa: E402
from premium_pack_from_runtime import build_premium_pack  # noqa: E402

# Writable data root (Render disk / local). Default: repo/workspace
WS_ROOT = Path(os.environ.get("INCI_PPWR_WORKSPACE_ROOT", str(ROOT / "workspace"))).resolve()
WORKSPACE = WS_ROOT
PRODUCTS = WS_ROOT / "products"
DROPS_LOG = WS_ROOT / "drops"
EXPORTS = WS_ROOT / "exports"
ACTIVITY_LOG = WS_ROOT / "activity.jsonl"
DESKTOP = Path(os.environ.get("INCI_PPWR_DESKTOP", str(Path.home() / "Desktop")))
from web_mode import is_web_mode

WEB_MODE = is_web_mode()

REV_RE = re.compile(r"^Rev\.(\d{2})$")
STATUSES = ("DRAFT", "ISSUED", "SUPERSEDED")
SCOPE_API = {
    "starter": "STARTER",
    "industrial": "INDUSTRIAL",
    "container": "CONTAINER",
    "component": "COMPONENT",
}

_REASON_MAP = {
    "bulk smoke": "Toplu oluşturma",
    "initial issue": "İlk yayın",
    "label correction smoke test": "Etiket düzeltmesi",
    "bulk import from master": "Master’dan toplu oluşturma",
}


def _public_reason(reason: str) -> str:
    raw = (reason or "").strip()
    if not raw:
        return ""
    mapped = _REASON_MAP.get(raw.lower())
    if mapped:
        return mapped
    if re.search(r"\bsmoke\b", raw, re.I):
        cleaned = re.sub(r"\bsmoke\s*(test)?\b", "", raw, flags=re.I)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" -·–—")
        return cleaned or "Toplu oluşturma"
    m = re.match(r"^ensure packs for customer\s*(.*)$", raw, re.I)
    if m:
        name = (m.group(1) or "").strip()
        return f"Müşteri Paketi: {name}" if name else "Müşteri Paketi"
    return raw


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _ensure_ws() -> None:
    PRODUCTS.mkdir(parents=True, exist_ok=True)
    DROPS_LOG.mkdir(parents=True, exist_ok=True)
    EXPORTS.mkdir(parents=True, exist_ok=True)
    readme = WORKSPACE / "README.txt"
    if not readme.exists():
        readme.write_text(
            "İnci Akü PPWR Yazılımı — workspace (source of truth)\n"
            "Frozen Rev.00 deliveries under Inci_Aku_PPWR_PIMS are NEVER modified.\n"
            "Revisions: Rev.00, Rev.01, …  Status: DRAFT | ISSUED | SUPERSEDED\n",
            encoding="utf-8",
        )


def log_activity(action: str, **detail) -> None:
    _ensure_ws()
    entry = {"at": _now(), "action": action, **detail}
    with ACTIVITY_LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def recent_activity(limit: int = 30) -> dict:
    _ensure_ws()
    if not ACTIVITY_LOG.exists():
        return {"count": 0, "events": []}
    lines = ACTIVITY_LOG.read_text(encoding="utf-8").splitlines()
    events = []
    for line in lines[-max(1, min(limit, 200)) :]:
        line = line.strip()
        if not line:
            continue
        try:
            events.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    events.reverse()
    return {"count": len(events), "events": events}


FROZEN = (
    "01_STARTER_INDIVIDUAL_DELIVERY_REV00",
    "02_INDUSTRIAL_DELIVERY_REV00",
    "03_CONTAINER_DELIVERY_REV00",
    "04_COMPONENT_SPARE_DELIVERY_REV00",
)


def _assert_workspace(path: Path) -> None:
    resolved = str(path.resolve()).replace("/", "\\").upper()
    for marker in FROZEN:
        if marker.upper() in resolved:
            raise HTTPException(403, f"Refusing write into frozen delivery: {marker}")
    try:
        path.resolve().relative_to(WORKSPACE.resolve())
    except ValueError as e:
        raise HTTPException(403, "Workspace writes only under workspace/") from e


def _product_dir(code: str) -> Path:
    return PRODUCTS / _safe_key(code)


def _rev_dir(code: str, rev: str) -> Path:
    return _product_dir(code) / "revisions" / rev


def _parse_rev_num(rev: str) -> int:
    m = REV_RE.match(rev)
    if not m:
        raise HTTPException(400, f"Invalid revision label: {rev} (use Rev.00)")
    return int(m.group(1))


def _fmt_rev(n: int) -> str:
    if n < 0 or n > 99:
        raise HTTPException(400, "Revision out of range")
    return f"Rev.{n:02d}"


def _next_rev(code: str) -> str:
    revs = list_revisions(code)
    if not revs:
        return "Rev.00"
    nums = [_parse_rev_num(r["revision"]) for r in revs]
    return _fmt_rev(max(nums) + 1)


def _load_product(code: str) -> dict:
    path = _product_dir(code) / "PRODUCT.json"
    if not path.exists():
        raise HTTPException(404, f"Product not in workspace: {code}")
    return json.loads(path.read_text(encoding="utf-8"))


def _save_product(code: str, data: dict) -> None:
    folder = _product_dir(code)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "PRODUCT.json"
    _assert_workspace(path)
    data["updated_at"] = _now()
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_rev_meta(code: str, rev: str) -> dict:
    path = _rev_dir(code, rev) / "meta.json"
    if not path.exists():
        raise HTTPException(404, f"Revision missing: {code} {rev}")
    return json.loads(path.read_text(encoding="utf-8"))


def _save_rev_meta(code: str, rev: str, meta: dict) -> None:
    folder = _rev_dir(code, rev)
    folder.mkdir(parents=True, exist_ok=True)
    path = folder / "meta.json"
    _assert_workspace(path)
    path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


def workspace_status() -> dict:
    _ensure_ws()
    products = list_products()
    return {
        "workspace": str(WORKSPACE),
        "products": len(products),
        "issued": sum(1 for p in products if p.get("status") == "ISSUED"),
        "revision_scheme": "Rev.00 / Rev.01 / …",
        "desktop_export": "ZIP",
        "engine_role": "background_optional_export_later",
        "write_policy": "workspace_only",
    }


def list_products() -> list[dict]:
    _ensure_ws()
    out = []
    if not PRODUCTS.exists():
        return out
    for folder in sorted(p for p in PRODUCTS.iterdir() if p.is_dir()):
        pj = folder / "PRODUCT.json"
        if not pj.exists():
            continue
        data = json.loads(pj.read_text(encoding="utf-8"))
        cur = data.get("current_revision")
        complete = False
        if cur:
            inv = _file_inventory(_rev_dir(folder.name, cur))
            complete = all(f["exists"] for f in inv)
        out.append(
            {
                "product_code": data.get("product_code") or folder.name,
                "description": data.get("description") or "",
                "set_code": data.get("set_code") or "",
                "current_revision": cur,
                "status": data.get("status") or "DRAFT",
                "revision_count": len(list_revisions(folder.name)),
                "complete": complete,
            }
        )
    return out


def search_workspace(q: str = "", limit: int = 100) -> dict:
    """Search workspace products by code, description, or set code."""
    needle = (q or "").strip().lower()
    items = list_products()
    if needle:
        items = [
            p
            for p in items
            if needle in (p.get("product_code") or "").lower()
            or needle in (p.get("description") or "").lower()
            or needle in (p.get("set_code") or "").lower()
            or needle in (p.get("status") or "").lower()
            or needle in (p.get("current_revision") or "").lower()
        ]
    items = items[: max(1, min(limit, 500))]
    return {
        "source": "workspace",
        "q": q,
        "total": len(items),
        "products": items,
    }


def list_revisions(code: str) -> list[dict]:
    key = _safe_key(code)
    root = _product_dir(key) / "revisions"
    if not root.exists():
        return []
    items = []
    for folder in sorted(p for p in root.iterdir() if p.is_dir() and REV_RE.match(p.name)):
        meta = {}
        mp = folder / "meta.json"
        if mp.exists():
            meta = json.loads(mp.read_text(encoding="utf-8"))
        inv = _file_inventory(folder)
        items.append(
            {
                "revision": folder.name,
                "status": meta.get("status") or "DRAFT",
                "reason": _public_reason(str(meta.get("reason") or "")),
                "built_at": meta.get("built_at"),
                "set_code": meta.get("set_code"),
                "complete": all(f["exists"] for f in inv),
                "files": inv,
            }
        )
    items.sort(key=lambda x: _parse_rev_num(x["revision"]))
    return items


def get_workspace_product(code: str) -> dict:
    key = _safe_key(code)
    data = _load_product(key)
    revs = list_revisions(key)
    current = data.get("current_revision")
    files = _file_inventory(_rev_dir(key, current)) if current else []
    return {
        "product": data,
        "revisions": revs,
        "current_files": files,
        "folder": str(_product_dir(key)),
    }


def _build_into(dest: Path, *, key: str, desc: str, set_code: str, scope: str, skip_pdf: bool) -> dict:
    dest.parent.mkdir(parents=True, exist_ok=True)
    _assert_workspace(dest.parent)
    lines, tare, meta = _bom_for_set(set_code)
    if not lines:
        raise HTTPException(400, f"No BOM for set {set_code}")
    _prep_photo_root()
    scope_key = SCOPE_API.get(scope.lower(), scope.upper())
    final_id = str(meta.get("final_id") or f"IA-{set_code}")
    info = build_premium_pack(
        dest,
        key=key,
        description_tr=desc,
        description_en=translate_product(desc),
        set_code=set_code,
        bom_lines=lines,
        scope=scope_key,
        total_tare_kg=tare,
        config_id=final_id,
    )
    pdf_ok = pdf_fail = 0
    soffice = None
    if not skip_pdf:
        soffice = str(find_soffice())
        docxs = [dest / f"{stem}.docx" for stem in STEMS]
        for d in docxs:
            if not d.exists():
                raise HTTPException(500, f"DOCX missing: {d.name}")
        LO_PROFILE.mkdir(parents=True, exist_ok=True)
        pdf_ok, pdf_fail = convert_batch_via_temp(find_soffice(), docxs, LO_PROFILE, chunk=4)
    inv = _file_inventory(dest)
    complete = all(f["exists"] for f in inv)
    return {
        "photos": info.get("photos"),
        "pdf_ok": pdf_ok,
        "pdf_fail": pdf_fail,
        "soffice": soffice,
        "tare_kg": tare,
        "final_id": final_id,
        "bom_lines": len(lines),
        "files": inv,
        "complete": complete,
        "scope": scope_key,
    }


def create_variant(
    *,
    product_code: str,
    description: str = "",
    set_code: str = "",
    scope: str = "starter",
    reason: str = "İlk yayın",
    skip_pdf: bool = False,
    issue: bool = True,
) -> dict:
    """Create product with Rev.00. Status DRAFT or ISSUED."""
    _ensure_ws()
    key = _safe_key(product_code)
    if (_product_dir(key) / "PRODUCT.json").exists():
        raise HTTPException(409, f"Product already exists — use revise: {key}")

    desc = (description or "").strip()
    sc = (set_code or "").strip()
    if not desc or not sc:
        try:
            prod = get_product("starter", key)
            desc = desc or str(prod.get("description") or "")
            sc = sc or str(prod.get("set_code") or "")
        except Exception:
            pass
    if not desc:
        raise HTTPException(400, "description required")
    if not sc:
        raise HTTPException(400, "set_code required")

    rev = "Rev.00"
    dest = _rev_dir(key, rev)
    built = _build_into(dest, key=key, desc=desc, set_code=sc, scope=scope, skip_pdf=skip_pdf)
    status = "ISSUED" if (issue and built["complete"]) else "DRAFT"
    if issue and not built["complete"]:
        status = "DRAFT"

    rev_meta = {
        "product_code": key,
        "revision": rev,
        "status": status,
        "reason": _public_reason(reason or "İlk yayın"),
        "description": desc,
        "set_code": sc,
        "scope": built["scope"],
        "final_id": built["final_id"],
        "tare_kg": built["tare_kg"],
        "bom_lines": built["bom_lines"],
        "photos": built["photos"],
        "pdf_ok": built["pdf_ok"],
        "pdf_fail": built["pdf_fail"],
        "built_at": _now(),
        "supersedes": None,
    }
    _save_rev_meta(key, rev, rev_meta)
    product = {
        "product_code": key,
        "description": desc,
        "set_code": sc,
        "scope": built["scope"],
        "current_revision": rev,
        "status": status,
        "created_at": _now(),
    }
    _save_product(key, product)
    log_activity(
        "create_variant",
        product_code=key,
        revision=rev,
        status=status,
        complete=built["complete"],
    )
    return {
        "product": product,
        "revision": rev_meta,
        "files": built["files"],
        "complete": built["complete"],
    }


def bulk_create_from_codes(
    *,
    codes: list[str],
    scope: str = "starter",
    reason: str = "Master’dan toplu oluşturma",
    skip_existing: bool = True,
    skip_pdf: bool = False,
) -> dict:
    """Create Rev.00 for each code using master data. Skip existing by default."""
    _ensure_ws()
    raw: list[str] = []
    for c in codes:
        for part in re.split(r"[\s,;]+", str(c or "").strip()):
            if part:
                raw.append(part)
    seen: set[str] = set()
    keys: list[str] = []
    for c in raw:
        try:
            k = _safe_key(c)
        except HTTPException:
            continue
        if k not in seen:
            seen.add(k)
            keys.append(k)
    if not keys:
        raise HTTPException(400, "No valid product codes")

    created = []
    skipped = []
    failed = []
    for key in keys:
        if (_product_dir(key) / "PRODUCT.json").exists():
            if skip_existing:
                skipped.append({"product_code": key, "reason": "zaten kayıtlı"})
                continue
            failed.append({"product_code": key, "error": "exists — use revise"})
            continue
        try:
            result = create_variant(
                product_code=key,
                description="",
                set_code="",
                scope=scope,
                reason=reason,
                skip_pdf=skip_pdf,
                issue=True,
            )
            created.append(
                {
                    "product_code": key,
                    "revision": result["revision"]["revision"],
                    "status": result["revision"]["status"],
                    "complete": result["complete"],
                    "set_code": result["product"].get("set_code"),
                }
            )
        except HTTPException as e:
            failed.append({"product_code": key, "error": e.detail if isinstance(e.detail, str) else str(e.detail)})
        except Exception as e:
            failed.append({"product_code": key, "error": str(e)})

    out = {
        "requested": len(keys),
        "created": created,
        "skipped": skipped,
        "failed": failed,
        "count_created": len(created),
        "count_skipped": len(skipped),
        "count_failed": len(failed),
    }
    log_activity(
        "bulk_create",
        count_created=len(created),
        count_skipped=len(skipped),
        count_failed=len(failed),
    )
    return out


def complete_pdfs(product_code: str, revision: str | None = None) -> dict:
    """Generate missing PDFs via LibreOffice; promote DRAFT→ISSUED when complete."""
    _ensure_ws()
    key = _safe_key(product_code)
    product = _load_product(key)
    rev = revision or product.get("current_revision")
    if not rev:
        raise HTTPException(404, "No revision")
    folder = _rev_dir(key, rev)
    if not folder.is_dir():
        raise HTTPException(404, f"Revision folder missing: {rev}")

    missing_docx = [stem for stem in STEMS if not (folder / f"{stem}.docx").exists()]
    if missing_docx:
        raise HTTPException(400, f"DOCX missing: {', '.join(missing_docx)}")

    need = [
        folder / f"{stem}.docx"
        for stem in STEMS
        if not (folder / f"{stem}.pdf").exists() or (folder / f"{stem}.pdf").stat().st_size < 200
    ]
    pdf_ok = pdf_fail = 0
    if need:
        LO_PROFILE.mkdir(parents=True, exist_ok=True)
        pdf_ok, pdf_fail = convert_batch_via_temp(find_soffice(), need, LO_PROFILE, chunk=4)

    inv = _file_inventory(folder)
    complete = all(f["exists"] for f in inv)
    meta = _load_rev_meta(key, rev)
    if complete:
        meta["status"] = "ISSUED"
        product["status"] = "ISSUED"
    meta["pdf_ok"] = pdf_ok
    meta["pdf_fail"] = pdf_fail
    meta["pdf_completed_at"] = _now()
    _save_rev_meta(key, rev, meta)
    if product.get("current_revision") == rev:
        _save_product(key, product)

    log_activity(
        "complete_pdfs",
        product_code=key,
        revision=rev,
        pdf_ok=pdf_ok,
        pdf_fail=pdf_fail,
        complete=complete,
        status=meta.get("status"),
    )
    return {
        "product_code": key,
        "revision": rev,
        "pdf_ok": pdf_ok,
        "pdf_fail": pdf_fail,
        "complete": complete,
        "status": meta.get("status"),
        "files": inv,
    }


def complete_all_incomplete(*, limit: int = 50) -> dict:
    """Complete PDFs for current incomplete packs."""
    _ensure_ws()
    results = []
    for p in list_products():
        if p.get("complete"):
            continue
        if len(results) >= limit:
            break
        try:
            results.append(complete_pdfs(p["product_code"]))
        except HTTPException as e:
            results.append(
                {
                    "product_code": p["product_code"],
                    "error": e.detail if isinstance(e.detail, str) else str(e.detail),
                }
            )
    log_activity("complete_all_incomplete", count=len(results))
    return {
        "count": len(results),
        "results": results,
        "completed": sum(1 for r in results if r.get("complete")),
    }


def revise_product(
    *,
    product_code: str,
    reason: str,
    description: str | None = None,
    set_code: str | None = None,
    scope: str | None = None,
    skip_pdf: bool = False,
) -> dict:
    """Bump to next Rev.xx; archive previous as SUPERSEDED. Reason required."""
    _ensure_ws()
    key = _safe_key(product_code)
    reason = _public_reason(reason)
    if len(reason) < 3:
        raise HTTPException(400, "Revizyon gerekçesi gerekli (en az 3 karakter)")

    product = _load_product(key)
    prev_rev = product.get("current_revision")
    if not prev_rev:
        raise HTTPException(400, "No current revision to supersede")

    prev_meta = _load_rev_meta(key, prev_rev)
    if prev_meta.get("status") == "DRAFT":
        # rebuild same draft revision in place
        dest = _rev_dir(key, prev_rev)
        desc = (description if description is not None else product.get("description") or "").strip()
        sc = (set_code if set_code is not None else product.get("set_code") or "").strip()
        scp = scope or "starter"
        built = _build_into(dest, key=key, desc=desc, set_code=sc, scope=scp, skip_pdf=skip_pdf)
        status = "ISSUED" if built["complete"] else "DRAFT"
        prev_meta.update(
            {
                "status": status,
                "reason": reason,
                "description": desc,
                "set_code": sc,
                "built_at": _now(),
                "photos": built["photos"],
                "pdf_ok": built["pdf_ok"],
                "pdf_fail": built["pdf_fail"],
                "tare_kg": built["tare_kg"],
                "bom_lines": built["bom_lines"],
            }
        )
        _save_rev_meta(key, prev_rev, prev_meta)
        product.update(
            {
                "description": desc,
                "set_code": sc,
                "status": status,
            }
        )
        _save_product(key, product)
        log_activity(
            "revise_draft_inplace",
            product_code=key,
            revision=prev_rev,
            status=status,
            complete=built["complete"],
            reason=reason,
        )
        return {
            "product": product,
            "revision": prev_meta,
            "files": built["files"],
            "complete": built["complete"],
            "note": "DRAFT rebuilt in place (no Rev bump)",
        }

    # supersede previous
    prev_meta["status"] = "SUPERSEDED"
    prev_meta["superseded_at"] = _now()
    _save_rev_meta(key, prev_rev, prev_meta)

    new_rev = _next_rev(key)
    desc = (description if description is not None else product.get("description") or "").strip()
    sc = (set_code if set_code is not None else product.get("set_code") or "").strip()
    scp = scope or "starter"
    if not desc or not sc:
        raise HTTPException(400, "description and set_code required")

    dest = _rev_dir(key, new_rev)
    built = _build_into(dest, key=key, desc=desc, set_code=sc, scope=scp, skip_pdf=skip_pdf)
    status = "ISSUED" if built["complete"] else "DRAFT"
    rev_meta = {
        "product_code": key,
        "revision": new_rev,
        "status": status,
        "reason": reason,
        "description": desc,
        "set_code": sc,
        "scope": built["scope"],
        "final_id": built["final_id"],
        "tare_kg": built["tare_kg"],
        "bom_lines": built["bom_lines"],
        "photos": built["photos"],
        "pdf_ok": built["pdf_ok"],
        "pdf_fail": built["pdf_fail"],
        "built_at": _now(),
        "supersedes": prev_rev,
    }
    _save_rev_meta(key, new_rev, rev_meta)
    product.update(
        {
            "description": desc,
            "set_code": sc,
            "current_revision": new_rev,
            "status": status,
        }
    )
    _save_product(key, product)
    log_activity(
        "revise",
        product_code=key,
        revision=new_rev,
        superseded=prev_rev,
        status=status,
        complete=built["complete"],
        reason=reason,
    )
    return {
        "product": product,
        "revision": rev_meta,
        "files": built["files"],
        "complete": built["complete"],
        "superseded": prev_rev,
    }


def open_workspace_file(product_code: str, file_name: str, revision: str | None = None) -> dict:
    key = _safe_key(product_code)
    product = _load_product(key)
    rev = revision or product.get("current_revision")
    if not rev:
        raise HTTPException(404, "No revision")
    folder = _rev_dir(key, rev).resolve()
    path = (folder / file_name).resolve()
    try:
        path.relative_to(folder)
    except ValueError as e:
        raise HTTPException(400, "Invalid path") from e
    if path.suffix.lower() not in {".docx", ".pdf"}:
        raise HTTPException(400, "Only DOCX/PDF")
    if not path.exists():
        raise HTTPException(404, f"Missing: {file_name}")
    if WEB_MODE or not hasattr(os, "startfile"):
        return {
            "opened": str(path),
            "revision": rev,
            "download": True,
            "download_url": (
                f"/api/workspace/file?product_code={key}"
                f"&file={file_name}&revision={rev}"
            ),
        }
    os.startfile(str(path))
    return {"opened": str(path), "revision": rev, "download": False}


def resolve_workspace_file(product_code: str, file_name: str, revision: str | None = None) -> Path:
    key = _safe_key(product_code)
    product = _load_product(key)
    rev = revision or product.get("current_revision")
    if not rev:
        raise HTTPException(404, "No revision")
    folder = _rev_dir(key, rev).resolve()
    path = (folder / file_name).resolve()
    try:
        path.relative_to(folder)
    except ValueError as e:
        raise HTTPException(400, "Invalid path") from e
    if path.suffix.lower() not in {".docx", ".pdf"}:
        raise HTTPException(400, "Only DOCX/PDF")
    if not path.exists():
        raise HTTPException(404, f"Missing: {file_name}")
    return path


CUSTOMER_ZIP_PDFS = ("01_Technical_File", "02_EU_DoC")
CUSTOMER_ZIP_NAMES = frozenset({"01_Technical_File.pdf", "02_EU_DoC.pdf", "_MANIFEST.json"})


def _strip_zip_to_names(zip_path: Path, allowed_names: frozenset[str]) -> list[str]:
    """Rewrite ZIP so only allowed basenames remain. Returns kept member paths."""
    tmp = zip_path.with_name(zip_path.stem + ".__filter__.zip")
    kept: list[str] = []
    with zipfile.ZipFile(zip_path, "r") as src:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                base = Path(info.filename).name
                if base not in allowed_names:
                    continue
                dst.writestr(info, src.read(info.filename))
                kept.append(info.filename)
    os.replace(tmp, zip_path)
    return kept


def finalize_customer_zip(zip_path: Path) -> list[str]:
    """Hard guarantee: customer ZIP contains only Technical File + EU DoC PDFs."""
    kept = _strip_zip_to_names(Path(zip_path), CUSTOMER_ZIP_NAMES)
    pdfs = [n for n in kept if n.lower().endswith(".pdf")]
    extra = [n for n in kept if Path(n).name not in CUSTOMER_ZIP_NAMES]
    if extra:
        raise HTTPException(500, f"Customer ZIP leaked extra files: {extra[:8]}")
    if any(not n.lower().endswith(".pdf") and Path(n).name != "_MANIFEST.json" for n in kept):
        raise HTTPException(500, "Customer ZIP must be PDF-only (plus manifest)")
    return pdfs


def desktop_zip_drop(
    *,
    codes: list[str],
    label: str = "MULTI",
    revision_policy: str = "current",  # current = each product's current ISSUED/DRAFT
    pack: str = "full",  # full = 4 docs × WORD+PDF; customer = Technical File + EU DoC PDFs only
) -> dict:
    """Build ZIP with 4-doc packs. Local → Desktop; web/Render → workspace/exports + download."""
    _ensure_ws()
    customer_only = (pack or "full").strip().lower() == "customer"
    raw_codes = []
    for c in codes:
        for part in re.split(r"[\s,;]+", str(c or "").strip()):
            if part:
                raw_codes.append(part)
    # unique preserve order
    seen = set()
    keys = []
    for c in raw_codes:
        k = _safe_key(c)
        if k not in seen:
            seen.add(k)
            keys.append(k)
    if not keys:
        raise HTTPException(400, "Paste at least one product code")

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    safe_label = re.sub(r"[^\w\-]+", "_", (label or "MULTI").strip())[:40] or "MULTI"
    zip_name = f"PPWR_{safe_label}_{stamp}.zip"
    if customer_only:
        zip_name = f"PPWR_{safe_label}_TF_DoC_{stamp}.zip"
    out_dir = EXPORTS if WEB_MODE else DESKTOP
    out_dir.mkdir(parents=True, exist_ok=True)
    zip_path = out_dir / zip_name

    included = []
    missing = []
    expected = 2 if customer_only else 8
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for key in keys:
            try:
                product = _load_product(key)
            except HTTPException:
                missing.append({"product_code": key, "error": "not in workspace"})
                continue
            rev = product.get("current_revision")
            if not rev:
                missing.append({"product_code": key, "error": "no revision"})
                continue
            folder = _rev_dir(key, rev)
            if not folder.is_dir():
                missing.append({"product_code": key, "error": f"rev folder missing {rev}"})
                continue
            pack_files = 0
            if customer_only:
                for stem in CUSTOMER_ZIP_PDFS:
                    src = folder / f"{stem}.pdf"
                    if src.exists() and src.stat().st_size > 200:
                        zf.write(src, arcname=f"{key}/{rev}/{src.name}")
                        pack_files += 1
            else:
                for stem in STEMS:
                    for ext in ("docx", "pdf"):
                        src = folder / f"{stem}.{ext}"
                        if src.exists() and src.stat().st_size > 200:
                            zf.write(src, arcname=f"{key}/{rev}/{src.name}")
                            pack_files += 1
                meta = folder / "meta.json"
                if meta.exists():
                    zf.write(meta, arcname=f"{key}/{rev}/meta.json")
            if pack_files < expected:
                missing.append(
                    {"product_code": key, "error": f"incomplete files ({pack_files}/{expected})", "revision": rev}
                )
            else:
                included.append({"product_code": key, "revision": rev, "files": pack_files})

        manifest = {
            "created_at": _now(),
            "label": label,
            "zip": zip_name,
            "pack": "customer" if customer_only else "full",
            "included": included,
            "missing": missing,
            "note": (
                "Customer ZIP — Technical File + EU DoC PDFs only"
                if customer_only
                else "Generated by İnci PPWR Yazılımı — workspace source of truth"
            ),
        }
        zf.writestr("_MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2))

    if customer_only:
        finalize_customer_zip(zip_path)

    # log drop
    log_path = DROPS_LOG / f"{stamp}_{safe_label}.json"
    log_path.write_text(json.dumps({"zip": str(zip_path), **manifest}, ensure_ascii=False, indent=2), encoding="utf-8")

    result = {
        "zip": str(zip_path),
        "zip_name": zip_name,
        "desktop": str(out_dir),
        "included": included,
        "missing": missing,
        "count_ok": len(included),
        "count_missing": len(missing),
        "download_url": f"/api/workspace/zip-download?name={zip_name}",
    }
    log_activity(
        "customer_zip" if customer_only else "desktop_zip",
        label=label,
        zip=zip_name,
        count_ok=len(included),
        count_missing=len(missing),
    )
    return result


def open_workspace_folder(product_code: str | None = None) -> dict:
    _ensure_ws()
    if WEB_MODE or not hasattr(os, "startfile"):
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    path = _product_dir(product_code) if product_code else WORKSPACE
    if product_code:
        path = _product_dir(product_code)
        if not path.exists():
            raise HTTPException(404, "Product folder missing")
    os.startfile(str(path))
    return {"opened": str(path)}
