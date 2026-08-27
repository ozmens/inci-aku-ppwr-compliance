"""Customer cards — saved multi-SKU lists for desktop ZIP drops.

Stored under workspace/customers/ only. Never touches frozen deliveries.
"""

from __future__ import annotations

import json
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from fastapi import HTTPException

from workspace_store import (
    WORKSPACE,
    _assert_workspace,
    desktop_zip_drop,
    finalize_customer_zip,
    _safe_key,
    list_products,
)

CUSTOMERS = WORKSPACE / "customers"


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _ensure() -> None:
    CUSTOMERS.mkdir(parents=True, exist_ok=True)


def _slug(name: str) -> str:
    s = re.sub(r"[^\w\-]+", "_", (name or "").strip(), flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("_")
    return (s or "MUSTERI")[:48]


def _parse_codes(codes: list[str] | str) -> list[str]:
    raw: list[str] = []
    if isinstance(codes, str):
        text = codes
    else:
        text = "\n".join(str(c) for c in codes)
    for ln in text.replace(",", "\n").replace(";", "\n").splitlines():
        part = ln.strip()
        if part:
            raw.append(part)
    seen: set[str] = set()
    out: list[str] = []
    for c in raw:
        try:
            k = _safe_key(c)
        except HTTPException:
            continue
        if k not in seen:
            seen.add(k)
            out.append(k)
    return out


def _path(customer_id: str) -> Path:
    return CUSTOMERS / f"{customer_id}.json"


def list_customers() -> dict:
    _ensure()
    items = []
    for p in sorted(CUSTOMERS.glob("*.json")):
        data = json.loads(p.read_text(encoding="utf-8"))
        items.append(
            {
                "id": data.get("id") or p.stem,
                "name": data.get("name") or p.stem,
                "code_count": len(data.get("product_codes") or []),
                "product_codes": data.get("product_codes") or [],
                "note": data.get("note") or "",
                "updated_at": data.get("updated_at"),
            }
        )
    items.sort(key=lambda x: (x.get("name") or "").lower())
    return {"customers_root": str(CUSTOMERS), "count": len(items), "customers": items}


def get_customer(customer_id: str) -> dict:
    _ensure()
    path = _path(customer_id)
    if not path.exists():
        raise HTTPException(404, f"Customer not found: {customer_id}")
    return json.loads(path.read_text(encoding="utf-8"))


def save_customer(
    *,
    name: str,
    codes_text: str = "",
    product_codes: list[str] | None = None,
    note: str = "",
    customer_id: str | None = None,
) -> dict:
    _ensure()
    name = (name or "").strip()
    if len(name) < 2:
        raise HTTPException(400, "Customer name required")
    codes = _parse_codes(product_codes if product_codes is not None else codes_text)
    if not codes:
        raise HTTPException(400, "At least one product code required")

    cid = (customer_id or "").strip() or f"{_slug(name)}_{uuid.uuid4().hex[:6]}"
    # sanitize id
    cid = re.sub(r"[^\w\-]+", "_", cid)[:64]
    path = _path(cid)
    _assert_workspace(CUSTOMERS)

    prev = {}
    if path.exists():
        prev = json.loads(path.read_text(encoding="utf-8"))

    data = {
        "id": cid,
        "name": name,
        "product_codes": codes,
        "note": (note or "").strip(),
        "created_at": prev.get("created_at") or _now(),
        "updated_at": _now(),
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return data


def delete_customer(customer_id: str) -> dict:
    path = _path(customer_id)
    if not path.exists():
        raise HTTPException(404, "Customer not found")
    _assert_workspace(path)
    path.unlink()
    return {"deleted": customer_id}


def zip_from_customer(customer_id: str, label: str | None = None) -> dict:
    data = get_customer(customer_id)
    codes = data.get("product_codes") or []
    if not codes:
        raise HTTPException(400, "Customer has no product codes")
    zip_label = (label or data.get("name") or "MUSTERI").strip()
    result = desktop_zip_drop(codes=codes, label=zip_label, pack="customer")
    zip_path = result.get("zip")
    if zip_path:
        pdfs = finalize_customer_zip(zip_path)
        result["pack"] = "customer"
        result["pdf_files"] = pdfs
        result["note"] = "Technical File + EU DoC PDF only"
    result["customer_id"] = data.get("id")
    result["customer_name"] = data.get("name")
    return result


def customer_coverage(customer_id: str) -> dict:
    """Per-code readiness vs workspace (missing / incomplete / ready)."""
    data = get_customer(customer_id)
    codes = data.get("product_codes") or []
    by_code = {p["product_code"]: p for p in list_products()}
    rows = []
    missing = incomplete = ready = 0
    for code in codes:
        p = by_code.get(code)
        if not p:
            state = "missing"
            missing += 1
            rows.append(
                {
                    "product_code": code,
                    "state": state,
                    "in_workspace": False,
                    "complete": False,
                    "status": None,
                    "revision": None,
                    "description": "",
                    "set_code": "",
                }
            )
            continue
        complete = bool(p.get("complete"))
        if complete:
            state = "ready"
            ready += 1
        else:
            state = "incomplete"
            incomplete += 1
        rows.append(
            {
                "product_code": code,
                "state": state,
                "in_workspace": True,
                "complete": complete,
                "status": p.get("status"),
                "revision": p.get("current_revision"),
                "description": p.get("description") or "",
                "set_code": p.get("set_code") or "",
            }
        )
    return {
        "customer_id": data.get("id"),
        "customer_name": data.get("name"),
        "total": len(codes),
        "ready": ready,
        "incomplete": incomplete,
        "missing": missing,
        "zip_ready": missing == 0 and incomplete == 0 and len(codes) > 0,
        "rows": rows,
    }


def ensure_customer_packs(
    customer_id: str,
    *,
    scope: str = "starter",
    skip_pdf: bool = False,
    then_zip: bool = True,
    label: str | None = None,
) -> dict:
    """Create missing workspace packs from master, then optional desktop ZIP."""
    from workspace_store import bulk_create_from_codes

    data = get_customer(customer_id)
    codes = data.get("product_codes") or []
    if not codes:
        raise HTTPException(400, "Customer has no product codes")
    bulk = bulk_create_from_codes(
        codes=codes,
        scope=scope,
        reason=f"Müşteri Paketi: {data.get('name')}",
        skip_existing=True,
        skip_pdf=skip_pdf,
    )
    out: dict = {"customer_id": data.get("id"), "customer_name": data.get("name"), "bulk": bulk}
    if then_zip:
        out["zip"] = zip_from_customer(customer_id, label=label)
    return out
