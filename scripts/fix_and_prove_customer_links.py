"""Make ZIP-portable Windows links and prove 80/80 opens from extracted package."""

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
SMOKE = Path(r"C:\Users\burcu\Desktop\PPWR_LINK_TEST_EXTRACT")

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


def apply_links(engine: Path) -> int:
    pythoncom.CoInitialize()
    excel = None
    n = 0
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(engine.resolve()))
        dc = wb.Worksheets("DOCUMENT_CENTER")
        r = 5
        while dc.Cells(r, 2).Value:
            sc = str(dc.Cells(r, 2).Value).strip()
            for col, fname in COLS:
                cell = dc.Cells(r, col)
                label = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                try:
                    if cell.Hyperlinks.Count:
                        cell.Hyperlinks.Delete()
                except Exception:
                    pass
                cell.Value = label
                dc.Hyperlinks.Add(Anchor=cell, Address=win_rel(sc, fname), TextToDisplay=label)
                n += 1
            r += 1
            if r > 400:
                break

        search = wb.Worksheets("SEARCH")
        search.Columns("Z").Hidden = False
        search.Range("Z1").NumberFormat = "General"
        search.Range("Z1").ClearContents()
        search.Range("Z1").Formula = '=IF($B$4="","",TRIM($B$4&""))'
        search.Range("AA1").NumberFormat = "General"
        search.Range("AA1").ClearContents()
        search.Range("AA1").Formula = (
            '=IF($Z$1="","",IFERROR(MATCH($Z$1,SEARCH_DATA!H:H,0),'
            'IFERROR(MATCH($Z$1,SEARCH_DATA!A:A,0),'
            'IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!A:A,0),'
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
    return n


def prove(engine: Path) -> dict:
    pythoncom.CoInitialize()
    excel = None
    out = {"dc": 0, "search": 0, "fail": []}
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(engine.resolve()), ReadOnly=False)
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
                addr = cell.Hyperlinks(1).Address if cell.Hyperlinks.Count else ""
                target = (engine.parent / addr).resolve()
                expect = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                if str(cell.Value) != expect or not target.exists():
                    out["fail"].append(f"DC {sc}/{fname} val={cell.Value!r} addr={addr!r} exists={target.exists()}")
                else:
                    try:
                        wb.FollowHyperlink(Address=addr)
                        out["dc"] += 1
                    except Exception as e:
                        out["fail"].append(f"DC follow {sc}/{fname}: {e}")
                n += 1
            r += 1

        search = wb.Worksheets("SEARCH")
        sd = wb.Worksheets("SEARCH_DATA")
        products = ["1015169", "1000069"]
        for i in range(2, min(sd.UsedRange.Rows.Count, 400) + 1):
            st = str(sd.Cells(i, 7).Value or "")
            if "ISSUED" in st and "NOT ISSUED" not in st and "YURT" not in st:
                pc = str(sd.Cells(i, 1).Value).strip().replace(".0", "")
                if pc not in products:
                    products.append(pc)
            if len(products) >= 5:
                break
        for pc in products:
            search.Range("B4").Value = pc
            excel.CalculateFull()
            set_code = str(search.Range("B8").Value or "").strip()
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
                        out["fail"].append(f"SEARCH {pc}/{stem}.{ext} disp={disp!r}")
                    else:
                        try:
                            wb.FollowHyperlink(Address=addr)
                            out["search"] += 1
                        except Exception as e:
                            out["fail"].append(f"SEARCH follow {pc}/{stem}.{ext}: {e}")
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
    (FINAL / "00_AC_DOCUMENT_ENGINE.cmd").write_text(
        "@echo off\r\n"
        "cd /d \"%~dp0\"\r\n"
        "start \"\" \"%~dp000_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\"\r\n",
        encoding="utf-8",
    )
    (FINAL / "00_READ_ME_NASIL_ACILIR.txt").write_text(
        "NASIL ACILIR / HOW TO OPEN\r\n"
        "1) ZIP'i bir klasore cikarin (Extract).\r\n"
        "2) 00_AC_DOCUMENT_ENGINE.cmd dosyasina cift tiklayin.\r\n"
        "   veya acin: 00_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\r\n"
        "3) Baska bir Excel kopyasini acmayin. Linkler bu dosyaya gore calisir.\r\n",
        encoding="utf-8",
    )

    print("1) Apply Windows relative links…", flush=True)
    n = apply_links(ENG)
    print("DC links", n, flush=True)

    print("2) Prove from 00_CONTROL…", flush=True)
    a = prove(ENG)
    print(a, flush=True)
    if a["dc"] + a["search"] != 80 or a["fail"]:
        raise SystemExit("CTRL prove FAIL")

    print("3) ZIP…", flush=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in FINAL.rglob("*"):
            if p.is_file() and not p.name.startswith("~$"):
                zf.write(p, p.relative_to(FINAL).as_posix())
    digest = sha256_file(ZIP_PATH)
    SHA_PATH.write_text(digest + "\n", encoding="utf-8")

    print("4) Extract to Desktop and prove…", flush=True)
    if SMOKE.exists():
        shutil.rmtree(SMOKE, ignore_errors=True)
    SMOKE.mkdir(parents=True)
    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        zf.extractall(SMOKE)
    smoke_eng = SMOKE / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
    b = prove(smoke_eng)
    print(b, flush=True)
    if b["dc"] + b["search"] != 80 or b["fail"]:
        raise SystemExit("EXTRACT prove FAIL")

    print("SHA256", digest)
    print("EXTRACT_TEST", SMOKE)
    print("GATE PASS 80/80")


if __name__ == "__main__":
    main()
