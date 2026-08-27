"""FULL Industrial production: 2736 product-level 4-doc packs.

- Source: Endustriyel_ambalaj_FINAL_ADET_VE_AGIRLIKLAR.xlsx
- Date: 11.08.2026
- PDF: LibreOffice only (never Word)
- Resume-safe
"""

from __future__ import annotations

import hashlib
import json
import shutil
import sys
import time
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import rebuild_industrial_from_excel as ind  # noqa: E402
from convert_pdfs_libreoffice import convert_batch_via_temp, find_soffice  # noqa: E402
from photo_annex import append_photo_annex, resolve_photos, _load_mapping  # noqa: E402
from ppwr_engine_builder import build_document_engine, verify_links  # noqa: E402

OUT = ROOT / "output"
CAND = OUT / "02_INDUSTRIAL_INDIVIDUAL_DELIVERY_REV00_CANDIDATE"
FINAL = OUT / "02_INDUSTRIAL_DELIVERY_REV00"
DATE = "11.08.2026"
STEMS = [
    "01_Technical_File",
    "02_EU_DoC",
    "03_Label",
    "04_Shipment_Statement",
]
PROGRESS = OUT / "_INDUSTRIAL_FULL_PROGRESS.json"


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def pack_complete_docx(folder: Path) -> bool:
    if not folder.exists():
        return False
    for stem in STEMS:
        if not (folder / f"{stem}.docx").exists():
            return False
    # annex present on TF
    try:
        with zipfile.ZipFile(folder / "01_Technical_File.docx", "r") as zf:
            xml = zf.read("word/document.xml")
            if b"Representative Packaging Component Photos" not in xml:
                return False
    except Exception:
        return False
    return True


def pack_complete_pdf(folder: Path) -> bool:
    for stem in STEMS:
        pdf = folder / f"{stem}.pdf"
        if not pdf.exists() or pdf.stat().st_size < 500:
            return False
        docx = folder / f"{stem}.docx"
        if docx.exists() and pdf.stat().st_mtime + 2 < docx.stat().st_mtime:
            return False
    return True


def write_progress(**kwargs) -> None:
    data = {}
    if PROGRESS.exists():
        try:
            data = json.loads(PROGRESS.read_text(encoding="utf-8"))
        except Exception:
            data = {}
    data.update(kwargs)
    data["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")
    PROGRESS.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def build_engine(control: Path, products: list[dict]) -> Path:
    delivery_root = control.parent
    records = [
        {
            "key": p["product_code"],
            "label": p["description"],
            "voltage": p["voltage"],
        }
        for p in products
    ]
    return build_document_engine(
        delivery_root=delivery_root,
        engine_filename="INCI_AKU_PPWR_INDUSTRIAL_ENGINE_Rev00.xlsx",
        title="İNCI AKÜ PPWR — INDUSTRIAL INDIVIDUAL ENGINE Rev00",
        docs_subdir="01_PRODUCTS",
        records=records,
        extra_home={
            "PUBLISH DATE": DATE,
            "SIGNATORY": "Numan Alver — Operations Director",
            "QMS": "TF YS/D/0020 · DoC YS/D/0021 · Label YS/D/0022 · STM YS/D/0023",
            "MODEL": "1 Product Code = 1 customer-facing 4-doc set (DOCX+PDF)",
        },
        extra_field="voltage",
    )


def main() -> int:
    print("INDUSTRIAL FULL RUN start", flush=True)
    # warm mapping cache
    _load_mapping()

    print("Parsing Excel…", flush=True)
    _, products = ind.parse_industrial(ind.SRC)
    print(f"products={len(products)}", flush=True)
    assert len(products) >= 2700, f"unexpected product count {len(products)}"

    control = CAND / "00_CONTROL"
    docs = CAND / "01_PRODUCTS"
    control.mkdir(parents=True, exist_ok=True)
    docs.mkdir(parents=True, exist_ok=True)
    shutil.copy2(ind.SRC, control / ind.SRC.name)
    if ind.MASTER.exists():
        shutil.copy2(ind.MASTER, control / ind.MASTER.name)
    else:
        ind.write_master(products, ind.MASTER)
        shutil.copy2(ind.MASTER, control / ind.MASTER.name)

    # Phase 1: DOCX
    t0 = time.time()
    made = skipped = 0
    for i, p in enumerate(products, 1):
        dest = docs / p["product_code"]
        if pack_complete_docx(dest):
            skipped += 1
        else:
            ind.generate_pack(p, dest)
            made += 1
        if i <= 5 or i % 100 == 0 or i == len(products):
            print(
                f"DOCX {i}/{len(products)} made={made} skipped={skipped} elapsed={time.time()-t0:.0f}s",
                flush=True,
            )
            write_progress(phase="docx", i=i, total=len(products), made=made, skipped=skipped)

    # Phase 2: PDF via LibreOffice
    print("Collecting stale PDFs…", flush=True)
    pdf_jobs: list[Path] = []
    for p in products:
        folder = docs / p["product_code"]
        if not folder.exists():
            continue
        if pack_complete_pdf(folder):
            continue
        for stem in STEMS:
            docx = folder / f"{stem}.docx"
            pdf = folder / f"{stem}.pdf"
            if not docx.exists():
                continue
            if (not pdf.exists()) or pdf.stat().st_size < 500 or docx.stat().st_mtime > pdf.stat().st_mtime + 1:
                pdf_jobs.append(docx)

    print(f"PDF jobs={len(pdf_jobs)}", flush=True)
    write_progress(phase="pdf", jobs=len(pdf_jobs))
    ok = fail = 0
    if pdf_jobs:
        soffice = find_soffice()
        # process in chunks; convert_batch_via_temp already chunks
        ok, fail = convert_batch_via_temp(
            soffice, pdf_jobs, OUT / "_lo_profile_pdf_industrial", chunk=40
        )
    print(f"PDF done ok={ok} fail={fail}", flush=True)
    write_progress(phase="pdf_done", ok=ok, fail=fail)

    # Engine
    print("Building engine…", flush=True)
    eng = build_engine(control, products)
    print("ENGINE", eng, flush=True)

    # QA counts
    n_folders = sum(1 for p in products if (docs / p["product_code"]).is_dir())
    n_docx = sum(
        1
        for p in products
        for stem in STEMS
        if (docs / p["product_code"] / f"{stem}.docx").exists()
    )
    n_pdf = sum(
        1
        for p in products
        for stem in STEMS
        if (docs / p["product_code"] / f"{stem}.pdf").exists()
        and (docs / p["product_code"] / f"{stem}.pdf").stat().st_size > 500
    )
    sample_codes = [products[0]["product_code"], products[-1]["product_code"]]
    # also a full-BOM sample if available
    full = next((p for p in products if p["bom_slots"] >= 10), None)
    if full:
        sample_codes.append(full["product_code"])

    sample_ok = True
    for code in sample_codes:
        folder = docs / code
        if not pack_complete_docx(folder) or not pack_complete_pdf(folder):
            sample_ok = False

    gate = (
        n_folders == len(products)
        and n_docx == len(products) * 4
        and n_pdf == len(products) * 4
        and fail == 0
        and sample_ok
    )
    report = {
        "products": len(products),
        "folders": n_folders,
        "docx": n_docx,
        "pdf": n_pdf,
        "pdf_fail": fail,
        "sample_ok": sample_ok,
        "sample_codes": sample_codes,
        "GATE": "PASS" if gate else "FAIL",
        "candidate": str(CAND),
        "engine": str(eng),
    }
    (control / "QA_FULL_INDUSTRIAL.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    OUT.joinpath("_INDUSTRIAL_FULL_QA.json").write_text(
        json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8"
    )
    print(json.dumps(report, indent=2), flush=True)

    if not gate:
        print("STOP — GATE FAIL, not promoting", flush=True)
        return 1

    # Promote
    if FINAL.exists():
        bak = OUT / f"_BACKUP_02_INDUSTRIAL_DELIVERY_REV00_{time.strftime('%Y%m%d_%H%M%S')}"
        FINAL.rename(bak)
        print("backed up", bak, flush=True)
    shutil.copytree(CAND, FINAL)

    zip_path = OUT / "02_INDUSTRIAL_DELIVERY_REV00.zip"
    if zip_path.exists():
        zip_path.unlink()
    print("Zipping…", flush=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as zf:
        for f in FINAL.rglob("*"):
            if f.is_file() and not f.name.startswith("~$"):
                zf.write(f, f.relative_to(FINAL).as_posix())
    digest = sha256_file(zip_path)
    (OUT / "02_INDUSTRIAL_DELIVERY_REV00_SHA256.txt").write_text(digest + "\n", encoding="utf-8")
    print("PROMOTED", FINAL, flush=True)
    print("ZIP", zip_path, flush=True)
    print("SHA256", digest, flush=True)
    write_progress(phase="done", GATE="PASS", sha256=digest)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
