"""Urgent: fix Document Engine links for customer package.

- DC: Windows relative ..\01_DOCUMENT_SETS\<SET>\<FILE>
- SEARCH: restore Z1 product-key + HYPERLINK formulas
- Verify by path resolve only (NO FollowHyperlink / NO Word open)
- Rebuild ZIP + prove extracted package resolves
"""

from __future__ import annotations

import hashlib
import shutil
import zipfile
from pathlib import Path

import pythoncom
import win32com.client as win32

ROOT = Path(r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS")
FINAL = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL"
ENG = FINAL / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
DOC = FINAL / "01_DOCUMENT_SETS"
ZIP_PATH = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL.zip"
SHA_PATH = ROOT / "output" / "INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL_SHA256.txt"
SMOKE = Path(r"C:\Users\burcu\Desktop\PPWR_LINK_SMOKE_EXTRACT")

COLS = [
    (6, "01_Technical_File.docx"),
    (7, "01_Technical_File.pdf"),
    (9, "02_EU_DoC.docx"),
    (10, "02_EU_DoC.pdf"),
    (12, "03_Label.docx"),
    (13, "03_Label.pdf"),
    (15, "04_Shipment_Statement.docx"),
    (16, "04_Shipment_Statement.pdf"),
]


def sha256_file(p: Path) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def win_rel(sc: str, fname: str) -> str:
    return f"..\\01_DOCUMENT_SETS\\{sc}\\{fname}"


def apply_links() -> dict:
    pythoncom.CoInitialize()
    excel = None
    stats = {"dc": 0, "missing": 0}
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(ENG.resolve()))

        # HOME reminder
        try:
            home = wb.Worksheets("00_HOME")
            home.Range("B28").Value = (
                "LINKLER: Bu dosyayi YALNIZCA 00_CONTROL klasorunden acin "
                "veya 00_AC_DOCUMENT_ENGINE.cmd kullanin."
            )
        except Exception:
            pass

        dc = wb.Worksheets("DOCUMENT_CENTER")
        r = 5
        while dc.Cells(r, 2).Value:
            sc = str(dc.Cells(r, 2).Value).strip()
            for col, fname in COLS:
                cell = dc.Cells(r, col)
                label = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                addr = win_rel(sc, fname)
                try:
                    if cell.Hyperlinks.Count:
                        cell.Hyperlinks.Delete()
                except Exception:
                    pass
                cell.Value = label
                dc.Hyperlinks.Add(Anchor=cell, Address=addr, TextToDisplay=label)
                stats["dc"] += 1
                if not (ENG.parent / addr).resolve().exists():
                    stats["missing"] += 1
            r += 1
            if r > 400:
                break

        # SEARCH helpers + action formulas
        search = wb.Worksheets("SEARCH")
        search.Columns("Z").Hidden = False
        search.Columns("AA").Hidden = False
        search.Range("Z1").NumberFormat = "General"
        search.Range("Z1").ClearContents()
        search.Range("Z1").Formula = '=IF($B$4="","",TRIM($B$4&""))'
        search.Range("AA1").NumberFormat = "General"
        search.Range("AA1").ClearContents()
        search.Range("AA1").Formula = (
            '=IF($Z$1="","",IFERROR(MATCH($Z$1,SEARCH_DATA!H:H,0),'
            "IFERROR(MATCH($Z$1,SEARCH_DATA!A:A,0),"
            "IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!A:A,0),"
            'IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!H:H,0),"NOT FOUND")))))'
        )
        search.Columns("Z").Hidden = True
        search.Columns("AA").Hidden = True

        issued = (
            'AND($B$8<>"",$A$8<>"NOT FOUND",ISNUMBER(SEARCH("ISSUED",$G$8)),'
            'NOT(ISNUMBER(SEARCH("NOT ISSUED",$G$8))),NOT(ISNUMBER(SEARCH("YURT",$G$8))))'
        )
        domestic = 'OR(ISNUMBER(SEARCH("YURT",$G$8)),ISNUMBER(SEARCH("NOT ISSUED",$G$8)))'
        empty = 'OR($B$4="",$A$8="",$A$8="NOT FOUND")'

        def af(stem: str, word: bool) -> str:
            label = "OPEN WORD" if word else "OPEN PDF"
            ext = "docx" if word else "pdf"
            path = f'"..\\01_DOCUMENT_SETS\\"&$B$8&"\\{stem}.{ext}"'
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
    return stats


def prove_resolve(engine: Path) -> dict:
    """Resolve hyperlink targets without opening Word/PDF."""
    pythoncom.CoInitialize()
    excel = None
    out = {"dc": 0, "search": 0, "fail": []}
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(engine.resolve()), ReadOnly=True)
        try:
            excel.CalculateFullRebuild()
        except Exception:
            excel.CalculateFull()

        dc = wb.Worksheets("DOCUMENT_CENTER")
        r = 5
        n = 0
        while dc.Cells(r, 2).Value and n < 40:
            sc = str(dc.Cells(r, 2).Value).strip()
            for col, fname in COLS:
                cell = dc.Cells(r, col)
                if cell.Hyperlinks.Count < 1:
                    out["fail"].append(f"DC no-hl {sc}/{fname}")
                    n += 1
                    continue
                addr = str(cell.Hyperlinks(1).Address or "")
                target = (engine.parent / addr).resolve()
                expect = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                if (
                    str(cell.Value) != expect
                    or not target.exists()
                    or not addr.startswith("..\\01_DOCUMENT_SETS\\")
                ):
                    out["fail"].append(
                        f"DC {sc}/{fname} val={cell.Value!r} addr={addr!r} exists={target.exists()}"
                    )
                else:
                    out["dc"] += 1
                n += 1
            r += 1

        search = wb.Worksheets("SEARCH")
        sd = wb.Worksheets("SEARCH_DATA")
        products = ["1015169"]
        for i in range(2, min(int(sd.UsedRange.Rows.Count), 300) + 1):
            st = str(sd.Cells(i, 7).Value or "")
            if "ISSUED" in st and "NOT ISSUED" not in st and "YURT" not in st:
                pc = str(sd.Cells(i, 1).Value).strip().replace(".0", "")
                if pc and pc not in products:
                    products.append(pc)
            if len(products) >= 5:
                break

        for pc in products[:5]:
            search.Range("B4").Value = str(pc)
            try:
                excel.CalculateFullRebuild()
            except Exception:
                excel.CalculateFull()
            set_code = str(search.Range("B8").Value or "").strip()
            a8 = str(search.Range("A8").Value or "")
            if not set_code or a8 == "NOT FOUND":
                out["fail"].append(f"SEARCH notfound {pc} A8={a8!r}")
                continue
            for row, stem in [
                (13, "01_Technical_File"),
                (15, "02_EU_DoC"),
                (17, "03_Label"),
                (19, "04_Shipment_Statement"),
            ]:
                for col, ext in ((1, "docx"), (2, "pdf")):
                    disp = str(search.Cells(row, col).Value or "")
                    label = "OPEN WORD" if ext == "docx" else "OPEN PDF"
                    addr = f"..\\01_DOCUMENT_SETS\\{set_code}\\{stem}.{ext}"
                    target = (engine.parent / addr).resolve()
                    if disp != label or not target.exists():
                        out["fail"].append(
                            f"SEARCH {pc}/{stem}.{ext} disp={disp!r} exists={target.exists()}"
                        )
                    else:
                        out["search"] += 1
        wb.Close(False)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return out


def main() -> None:
    assert ENG.exists(), ENG
    assert DOC.exists(), DOC

    (FINAL / "00_AC_DOCUMENT_ENGINE.cmd").write_text(
        "@echo off\r\n"
        "cd /d \"%~dp0\"\r\n"
        "echo Opening Document Engine from 00_CONTROL...\r\n"
        "start \"\" \"%~dp000_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\"\r\n",
        encoding="utf-8",
    )

    print("1) Apply links…", flush=True)
    stats = apply_links()
    print(stats, flush=True)
    if stats["missing"]:
        raise SystemExit("missing document targets")

    print("2) Prove from 00_CONTROL (path resolve only)…", flush=True)
    a = prove_resolve(ENG)
    print(a, flush=True)
    if a["dc"] != 40 or a["search"] != 40 or a["fail"]:
        raise SystemExit("CTRL prove FAIL")

    print("3) Rebuild ZIP…", flush=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in FINAL.rglob("*"):
            if p.is_file() and not p.name.startswith("~$"):
                zf.write(p, p.relative_to(FINAL).as_posix())
    digest = sha256_file(ZIP_PATH)
    SHA_PATH.write_text(digest + "\n", encoding="utf-8")

    print("4) Extract + prove…", flush=True)
    if SMOKE.exists():
        shutil.rmtree(SMOKE, ignore_errors=True)
    SMOKE.mkdir(parents=True)
    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        zf.extractall(SMOKE)
    b = prove_resolve(SMOKE / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx")
    print(b, flush=True)
    if b["dc"] != 40 or b["search"] != 40 or b["fail"]:
        raise SystemExit("EXTRACT prove FAIL")

    print("SHA256", digest)
    print("OPEN_THIS", ENG)
    print("OR_CMD", FINAL / "00_AC_DOCUMENT_ENGINE.cmd")
    print("GATE PASS 80/80 (resolve)")


if __name__ == "__main__":
    main()
