"""Full regen: exact Desktop YS_D templates → all 287 sets. No invented QMS table."""

from __future__ import annotations

import hashlib
import shutil
import sys
import zipfile
from collections import Counter
from pathlib import Path

ROOT = Path(r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS")
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "src"))
sys.path.insert(0, str(ROOT / "tools"))

from docx import Document

from builders.phase_g.merge_engine import merge_document
from builders.phase_g.runtime_template_builder import build_runtime_templates
from builders.phase_i.render_batch import render_docx_batch
from models.technical_file import Article5Assessment
from services.document_context_factory import DocumentContextFactory
from utils.constants import ARTICLE5_BASIS_LABEL
from ppwr_engine.starter_loader import StarterMasterLoader
import generate_ppwr_documents as gen

DESK = Path(r"c:\Users\burcu\Desktop\İNCİ AKÜ PPWR\DOKÜMAN SETİ")
GOLDEN = ROOT / "templates" / "word_golden"
LOCKED = ROOT / "templates" / "ppwr_rev00_locked"
RUNTIME = LOCKED / "runtime"
MASTER = ROOT / "output" / "INCI_AKU_PPWR_STARTER_MASTER_Rev00.xlsx"
FINAL = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL"
DOC_SETS = FINAL / "01_DOCUMENT_SETS"
ENG = FINAL / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
SIG = ROOT / "assets" / "signatory" / "numan_alver_signature_transparent.png"
ZIP_PATH = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL.zip"
SHA_PATH = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL_SHA256.txt"

MAP = {
    "YS_D_0020 - Technical File.docx": "01_Technical_File_GOLDEN.docx",
    "YS_D_0021 - EU DoC.docx": "02_EU_DoC_GOLDEN.docx",
    "YS_D_0022 - Label.docx": "03_Label_GOLDEN.docx",
    "YS_D_0023 - Shipment Statement.docx": "04_Shipment_Statement_GOLDEN.docx",
}
SPECS = [
    ("TECHNICAL_FILE", "01_Technical_File", True, "YS/D/0020"),
    ("DOC", "02_EU_DoC", False, "YS/D/0021"),
    ("LABEL", "03_Label", False, "YS/D/0022"),
    ("STATEMENT", "04_Shipment_Statement", False, "YS/D/0023"),
]


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def install_exact_templates() -> None:
    LOCKED.mkdir(parents=True, exist_ok=True)
    for src_name, dst_name in MAP.items():
        src = DESK / src_name
        assert src.exists(), src
        shutil.copy2(src, GOLDEN / dst_name)
        shutil.copy2(src, LOCKED / dst_name)  # exact — no scrub/patch
    build_runtime_templates(LOCKED, RUNTIME)


def fix_absolute_links() -> None:
    import pythoncom
    import win32com.client as win32

    if not ENG.exists():
        return
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(ENG.resolve()))
        dc = wb.Worksheets("DOCUMENT_CENTER")
        r = 5
        while dc.Cells(r, 2).Value:
            sc = str(dc.Cells(r, 2).Value).strip()
            for col, fname in [
                (6, "01_Technical_File.docx"),
                (7, "01_Technical_File.pdf"),
                (9, "02_EU_DoC.docx"),
                (10, "02_EU_DoC.pdf"),
                (12, "03_Label.docx"),
                (13, "03_Label.pdf"),
                (15, "04_Shipment_Statement.docx"),
                (16, "04_Shipment_Statement.pdf"),
            ]:
                cell = dc.Cells(r, col)
                abs_path = str((DOC_SETS / sc / fname).resolve())
                label = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                try:
                    if cell.Hyperlinks.Count:
                        cell.Hyperlinks.Delete()
                except Exception:
                    pass
                cell.Value = label
                dc.Hyperlinks.Add(Anchor=cell, Address=abs_path, TextToDisplay=label)
            r += 1
            if r > 400:
                break
        search = wb.Worksheets("SEARCH")
        search.Range("Z1").Value = str(FINAL.resolve()) + "\\"
        search.Columns("Z").Hidden = True
        issued = (
            'AND($B$8<>"",$A$8<>"NOT FOUND",ISNUMBER(SEARCH("ISSUED",$G$8)),'
            'NOT(ISNUMBER(SEARCH("NOT ISSUED",$G$8))),NOT(ISNUMBER(SEARCH("YURT",$G$8))))'
        )
        domestic = 'OR(ISNUMBER(SEARCH("YURT",$G$8)),ISNUMBER(SEARCH("NOT ISSUED",$G$8)))'
        empty = 'OR($B$4="",$A$8="",$A$8="NOT FOUND")'

        def af(stem: str, word: bool) -> str:
            label = "OPEN WORD" if word else "OPEN PDF"
            ext = "docx" if word else "pdf"
            path = f'$Z$1&"01_DOCUMENT_SETS\\"&$B$8&"\\{stem}.{ext}"'
            return (
                f'=IF({empty},"",IF({domestic},"DOCUMENTS NOT ISSUED",'
                f'HYPERLINK(IF({issued},{path},""),IF({issued},"{label}",""))))'
            )

        for row, stem in [
            (13, "01_Technical_File"),
            (15, "02_EU_DoC"),
            (17, "03_Label"),
            (19, "04_Shipment_Statement"),
        ]:
            search.Cells(row, 1).Formula = af(stem, True)
            search.Cells(row, 2).Formula = af(stem, False)
        wb.Save()
        wb.Close(False)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> None:
    assert SIG.exists()
    print("1) Exact Desktop templates…", flush=True)
    install_exact_templates()
    runtime = {
        "TECHNICAL_FILE": RUNTIME / "01_Technical_File_RUNTIME.docx",
        "DOC": RUNTIME / "02_EU_DoC_RUNTIME.docx",
        "LABEL": RUNTIME / "03_Label_RUNTIME.docx",
        "STATEMENT": RUNTIME / "04_Shipment_Statement_RUNTIME.docx",
    }

    if DOC_SETS.exists():
        shutil.rmtree(DOC_SETS)
    DOC_SETS.mkdir(parents=True)

    print("2) Word 287×4…", flush=True)
    loader = StarterMasterLoader(MASTER)
    loader.open()
    codes = loader.list_controlled_set_codes()
    assert len(codes) == 287, len(codes)
    factory = DocumentContextFactory()
    sig_ok = 0
    pdf_jobs: list[tuple[Path, Path]] = []

    for i, sc in enumerate(codes, 1):
        bundle = loader.load_set(sc)
        ctx = factory.build(
            bundle.configuration,
            products=bundle.products,
            article5=Article5Assessment(basis_label=ARTICLE5_BASIS_LABEL),
        )
        ctx.total_tare_g = float(bundle.packaging_tare_kg) * 1000.0
        ctx.annex_drawings_status = "OPTIONAL EVIDENCE — NOT INCLUDED IN REV.00"
        ctx.document_ids.technical_file_id = bundle.doc_ids["tf"]
        ctx.document_ids.doc_id = bundle.doc_ids["doc"]
        ctx.document_ids.label_id = bundle.doc_ids["label"]
        ctx.document_ids.statement_id = bundle.doc_ids["stm"]
        out_dir = DOC_SETS / sc
        out_dir.mkdir(parents=True)
        for dtype, stem, is_tf, _qms in SPECS:
            out = out_dir / f"{stem}.docx"
            merge_document(runtime[dtype], out, ctx, for_technical_file=is_tf)
            if dtype == "DOC":
                if gen.embed_signature(out, SIG):
                    sig_ok += 1
            pdf_jobs.append((out, out_dir / f"{stem}.pdf"))
        if i % 25 == 0 or i == len(codes):
            print(f"  Word {i}/{len(codes)}", flush=True)
    loader.close()
    print({"sig_ok": sig_ok})

    print("3) PDF…", flush=True)
    for i in range(0, len(pdf_jobs), 40):
        render_docx_batch(pdf_jobs[i : i + 40], progress_every=40, log=[])
        print(f"  PDF {min(i+40,len(pdf_jobs))}/{len(pdf_jobs)}", flush=True)
    retry = [(d, p) for d, p in pdf_jobs if not p.exists() or p.stat().st_size == 0]
    if retry:
        for i in range(0, len(retry), 20):
            render_docx_batch(retry[i : i + 20], progress_every=20, log=[])

    print("4) QA…", flush=True)
    words = [p for p in DOC_SETS.rglob("*.docx") if not p.name.startswith("~$")]
    pdfs = [p for p in DOC_SETS.rglob("*.pdf") if p.stat().st_size > 0]
    qms = Counter()
    numan = marker = 0
    for p in words:
        doc = Document(str(p))
        blob = "\n".join(x.text for x in doc.paragraphs)
        for _dtype, stem, _is_tf, code in SPECS:
            if p.name == f"{stem}.docx" and f"Doküman No/Doc. Nr.: {code}" in blob:
                qms[code] += 1
        if p.name == "02_EU_DoC.docx":
            if "Numan Alver" in blob:
                numan += 1
            if "SIGNATORY_SIGNATURE" in blob or "[[SIGNATORY" in blob:
                marker += 1

    print(
        {
            "sets": len([p for p in DOC_SETS.iterdir() if p.is_dir()]),
            "word": len(words),
            "pdf": len(pdfs),
            "qms": dict(qms),
            "numan": numan,
            "marker": marker,
            "sig_ok": sig_ok,
        },
        flush=True,
    )

    print("5) Absolute links…", flush=True)
    fix_absolute_links()

    print("6) ZIP…", flush=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in FINAL.rglob("*"):
            if p.is_file() and not p.name.startswith("~$"):
                zf.write(p, p.relative_to(FINAL).as_posix())
    digest = sha256_file(ZIP_PATH)
    SHA_PATH.write_text(digest + "\n", encoding="utf-8")

    gate = (
        len(words) == 1148
        and len(pdfs) == 1148
        and all(qms.get(c, 0) == 287 for c in ("YS/D/0020", "YS/D/0021", "YS/D/0022", "YS/D/0023"))
        and sig_ok == 287
        and marker == 0
    )
    print(f"SHA256 {digest}")
    print(f"GATE {'PASS' if gate else 'FAIL'}")
    if not gate:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
