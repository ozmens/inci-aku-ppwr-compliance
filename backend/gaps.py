"""Gap wizard — suggest packaging sets; write ONLY to candidates workspace.

Never writes into frozen delivery folders or PIMS master workbooks.
"""

from __future__ import annotations

import json
import uuid
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import HTTPException
from openpyxl import Workbook

from masters import get_bom, _sheet_rows, _hi
from web_mode import is_web_mode

ROOT = Path(__file__).resolve().parents[1]
CANDIDATES = ROOT / "candidates"
ASSIGNMENTS_FILE = CANDIDATES / "gap_assignments.json"
FROZEN_MARKERS = (
    "01_STARTER_INDIVIDUAL_DELIVERY_REV00",
    "02_INDUSTRIAL_DELIVERY_REV00",
    "03_CONTAINER_DELIVERY_REV00",
    "04_COMPONENT_SPARE_DELIVERY_REV00",
)

FORM_TOKENS = [
    "TRACTORL",
    "TRACTOR",
    "C-IGYASMF",
    "C-IGYSMF",
    "C-IGYA",
    "B-IGYSMF",
    "B-IGYA",
    "C-FLAT",
    "B-FLAT",
    "A-IGYA",
    "A-FLAT",
    "D31",
    "D26",
    "D23",
    "L6",
    "L5",
    "L3",
    "L2",
    "L1",
    "AGM",
]
FORM_ALIAS = {
    "C-IGYSMF": "C-IGYA",
    "C-IGYASMF": "C-IGYASMF",
}


def _norm(s: str) -> str:
    return (
        (s or "")
        .upper()
        .replace("İ", "I")
        .replace("Ş", "S")
        .replace("Ğ", "G")
        .replace("Ü", "U")
        .replace("Ö", "O")
        .replace("Ç", "C")
    )


def form_key(desc: str) -> str:
    d = _norm(desc)
    for tok in FORM_TOKENS:
        if tok in d:
            return FORM_ALIAS.get(tok, tok)
    return "OTHER"


def ensure_candidates_dir() -> Path:
    CANDIDATES.mkdir(parents=True, exist_ok=True)
    readme = CANDIDATES / "README.txt"
    if not readme.exists():
        readme.write_text(
            "İnci Akü PPWR — Gap candidates workspace\n"
            "Assignments here are drafts for NEW packs only.\n"
            "Frozen deliveries under Inci_Aku_PPWR_PIMS\\output\\0*_DELIVERY_REV00 are NEVER modified.\n",
            encoding="utf-8",
        )
    return CANDIDATES


def assert_not_frozen_write(path: Path) -> None:
    resolved = str(path.resolve()).replace("/", "\\").upper()
    for marker in FROZEN_MARKERS:
        if marker.upper() in resolved:
            raise HTTPException(403, f"Refusing write into frozen delivery: {marker}")


def _load_assignments() -> list[dict]:
    ensure_candidates_dir()
    if not ASSIGNMENTS_FILE.exists():
        return []
    data = json.loads(ASSIGNMENTS_FILE.read_text(encoding="utf-8"))
    return list(data.get("assignments") or [])


def _save_assignments(items: list[dict]) -> None:
    ensure_candidates_dir()
    assert_not_frozen_write(ASSIGNMENTS_FILE)
    payload = {
        "updated_at": datetime.now(timezone.utc).isoformat(),
        "note": "Draft gap assignments — not applied to frozen deliveries",
        "assignments": items,
    }
    ASSIGNMENTS_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _form_set_index(kind: str = "starter") -> dict[str, Counter]:
    headers, rows = _sheet_rows(kind, "PRODUCT_MASTER")
    hi = _hi(headers)
    status_i = hi.get("Physical Packaging Status")
    set_i = hi.get("Packaging Set Code")
    desc_i = hi.get("Technical Description")
    form_sets: dict[str, Counter] = defaultdict(Counter)
    if status_i is None or set_i is None or desc_i is None:
        return form_sets
    for row in rows:
        st = str(row[status_i] or "")
        if "CONTROLLED" not in st.upper():
            continue
        sc = str(row[set_i] or "").strip()
        if not sc or "NOT" in sc.upper() or "DATA" in sc.upper():
            continue
        fk = form_key(str(row[desc_i] or ""))
        form_sets[fk][sc] += 1
    return form_sets


def scan_gaps(kind: str = "starter", limit: int = 200) -> dict:
    headers, rows = _sheet_rows(kind, "PRODUCT_MASTER")
    hi = _hi(headers)
    code_i = hi.get("Product Code") or 0
    desc_i = hi.get("Technical Description")
    status_i = hi.get("Physical Packaging Status")
    set_i = hi.get("Packaging Set Code")
    scope_i = hi.get("Scope Status")
    gaps = []
    for row in rows:
        if not row or not row[code_i]:
            continue
        st = str(row[status_i] or "") if status_i is not None else ""
        sc = str(row[set_i] or "") if set_i is not None else ""
        scope = str(row[scope_i] or "") if scope_i is not None else ""
        blob = f"{st} {sc} {scope}".upper()
        is_gap = (
            "DATA REQUIRED" in blob
            or "DATA GAP" in blob
            or "NOT ASSIGNED" in blob
            or ("GAP" in blob and "CONTROLLED" not in st.upper())
        )
        if not is_gap and st and "CONTROLLED" not in st.upper():
            is_gap = True
        if not is_gap:
            continue
        desc = str(row[desc_i] or "") if desc_i is not None else ""
        gaps.append(
            {
                "product_code": str(row[code_i]).strip(),
                "description": desc,
                "status": st,
                "set_code": sc,
                "form": form_key(desc),
            }
        )
        if len(gaps) >= limit:
            break
    return {
        "kind": kind,
        "count": len(gaps),
        "gaps": gaps,
        "note": "Controlled starter master may already be closed; use manual suggest for new products.",
    }


def suggest(description: str, product_code: str | None = None, top_n: int = 5) -> dict:
    fk = form_key(description)
    index = _form_set_index("starter")
    counter = index.get(fk) or index.get("OTHER") or Counter()
    if not counter:
        # fallback: any most common set
        merged: Counter = Counter()
        for c in index.values():
            merged.update(c)
        counter = merged
    suggestions = []
    for set_code, n in counter.most_common(top_n):
        bom = get_bom("starter", set_code)
        suggestions.append(
            {
                "set_code": set_code,
                "peer_products": n,
                "tare_kg": bom.get("meta", {}).get("tare_kg"),
                "bom_lines": len(bom.get("lines") or []),
                "final_id": bom.get("meta", {}).get("final_id"),
                "description": bom.get("meta", {}).get("description"),
            }
        )
    return {
        "product_code": product_code or "",
        "description": description,
        "form": fk,
        "suggestions": suggestions,
    }


def list_assignments() -> dict:
    items = _load_assignments()
    return {
        "candidates_root": str(CANDIDATES),
        "count": len(items),
        "assignments": items,
    }


def save_assignment(body: dict[str, Any]) -> dict:
    pc = str(body.get("product_code") or "").strip()
    set_code = str(body.get("set_code") or "").strip()
    desc = str(body.get("description") or "").strip()
    if not pc or not set_code:
        raise HTTPException(400, "product_code and set_code required")
    # validate set exists
    bom = get_bom("starter", set_code)
    if not bom.get("lines"):
        raise HTTPException(400, f"Unknown or empty BOM for set: {set_code}")

    items = _load_assignments()
    # upsert by product_code
    now = datetime.now(timezone.utc).isoformat()
    entry = {
        "id": str(uuid.uuid4()),
        "product_code": pc,
        "description": desc,
        "form": form_key(desc) if desc else body.get("form") or "OTHER",
        "set_code": set_code,
        "final_id": bom.get("meta", {}).get("final_id"),
        "tare_kg": bom.get("meta", {}).get("tare_kg"),
        "bom_lines": len(bom.get("lines") or []),
        "note": str(body.get("note") or ""),
        "status": "DRAFT_CANDIDATE",
        "created_at": now,
        "updated_at": now,
    }
    kept = [a for a in items if a.get("product_code") != pc]
    # preserve id if updating
    prev = next((a for a in items if a.get("product_code") == pc), None)
    if prev:
        entry["id"] = prev["id"]
        entry["created_at"] = prev.get("created_at", now)
    kept.append(entry)
    kept.sort(key=lambda a: a.get("product_code") or "")
    _save_assignments(kept)
    _export_xlsx(kept)
    return {"saved": entry, "count": len(kept)}


def delete_assignment(assignment_id: str) -> dict:
    items = _load_assignments()
    kept = [a for a in items if a.get("id") != assignment_id]
    if len(kept) == len(items):
        raise HTTPException(404, "Assignment not found")
    _save_assignments(kept)
    _export_xlsx(kept)
    return {"deleted": assignment_id, "count": len(kept)}


def _export_xlsx(items: list[dict]) -> Path:
    ensure_candidates_dir()
    path = CANDIDATES / "GAP_CANDIDATES_Rev00.xlsx"
    assert_not_frozen_write(path)
    wb = Workbook()
    ws = wb.active
    ws.title = "GAP_CANDIDATES"
    headers = [
        "Product Code",
        "Description",
        "Form",
        "Packaging Set Code",
        "Final Configuration ID",
        "Tare kg",
        "BOM Lines",
        "Status",
        "Note",
        "Updated At",
    ]
    ws.append(headers)
    for a in items:
        ws.append(
            [
                a.get("product_code"),
                a.get("description"),
                a.get("form"),
                a.get("set_code"),
                a.get("final_id"),
                a.get("tare_kg"),
                a.get("bom_lines"),
                a.get("status"),
                a.get("note"),
                a.get("updated_at"),
            ]
        )
    wb.save(path)
    return path


def workspace_status() -> dict:
    ensure_candidates_dir()
    items = _load_assignments()
    return {
        "candidates_root": str(CANDIDATES),
        "assignments_file": str(ASSIGNMENTS_FILE),
        "xlsx": str(CANDIDATES / "GAP_CANDIDATES_Rev00.xlsx"),
        "count": len(items),
        "frozen_deliveries_writable": False,
        "write_policy": "candidates_only",
    }


def open_candidates_folder() -> dict:
    ensure_candidates_dir()
    import os

    if is_web_mode() or not hasattr(os, "startfile"):
        raise HTTPException(400, "Klasör açma yalnızca masaüstü kurulumunda")
    os.startfile(str(CANDIDATES))
    return {"opened": str(CANDIDATES)}
