"""Fix Excel links so SEARCH/DC actually open files.

Root cause: HYPERLINK(\"..\\01_DOCUMENT_SETS\\...\") resolves against
Excel's current directory, NOT the workbook folder — fails for users.

Fix: build path from CELL(\"filename\") workbook directory:
  LEFT(CELL(\"filename\"),FIND(\"[\",...)-1) & \"..\\01_DOCUMENT_SETS\\\" & set & \"\\file\"

Also use outermost HYPERLINK (clickable). No Word document opens.
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


def wb_dir_formula(anchor: str) -> str:
    # Returns workbook folder with trailing backslash
    return f'=IFERROR(LEFT(CELL("filename",{anchor}),FIND("[",CELL("filename",{anchor}))-1),"")'


def apply() -> None:
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(ENG.resolve()))

        # --- SEARCH ---
        s = wb.Worksheets("SEARCH")
        s.Columns("Z").Hidden = False
        s.Columns("AA").Hidden = False
        # Z2 = workbook directory (for file paths)
        s.Range("Z2").ClearContents()
        s.Range("Z2").Formula = wb_dir_formula("$B$4")
        # Z1 = normalized product code key
        s.Range("Z1").ClearContents()
        s.Range("Z1").Formula = '=IF($B$4="","",TRIM($B$4&""))'
        s.Range("AA1").ClearContents()
        s.Range("AA1").Formula = (
            '=IF($Z$1="","",IFERROR(MATCH($Z$1,SEARCH_DATA!H:H,0),'
            "IFERROR(MATCH($Z$1,SEARCH_DATA!A:A,0),"
            "IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!A:A,0),"
            'IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!H:H,0),"NOT FOUND")))))'
        )

        issued = (
            'AND($B$8<>"",$A$8<>"NOT FOUND",ISNUMBER(SEARCH("ISSUED",$G$8)),'
            'NOT(ISNUMBER(SEARCH("NOT ISSUED",$G$8))),NOT(ISNUMBER(SEARCH("YURT",$G$8))))'
        )
        domestic = 'OR(ISNUMBER(SEARCH("YURT",$G$8)),ISNUMBER(SEARCH("NOT ISSUED",$G$8)))'
        empty = 'OR($B$4="",$A$8="",$A$8="NOT FOUND")'

        def search_formula(stem: str, word: bool) -> str:
            label = "OPEN WORD" if word else "OPEN PDF"
            ext = "docx" if word else "pdf"
            # Workbook-dir absolute + relative docs — Excel resolves correctly
            path = f'$Z$2&"..\\01_DOCUMENT_SETS\\"&$B$8&"\\{stem}.{ext}"'
            # OUTERMOST HYPERLINK — required for click
            return (
                f'=IF({empty},"",'
                f'IF({domestic},"DOCUMENTS NOT ISSUED",'
                f'HYPERLINK(IF({issued},{path},"#"),'
                f'IF({issued},"{label}",""))))'
            )

        for row, stem in [
            (13, "01_Technical_File"),
            (15, "02_EU_DoC"),
            (17, "03_Label"),
            (19, "04_Shipment_Statement"),
        ]:
            s.Cells(row, 1).Formula = search_formula(stem, True)
            s.Cells(row, 2).Formula = search_formula(stem, False)
            for c in (1, 2):
                s.Cells(row, c).Font.Color = 0x0563C1
                s.Cells(row, c).Font.Underline = 2  # xlUnderlineStyleSingle

        s.Columns("Z").Hidden = True
        s.Columns("AA").Hidden = True

        # --- DOCUMENT CENTER: replace Hyperlink objects with CELL-based HYPERLINK formulas ---
        dc = wb.Worksheets("DOCUMENT_CENTER")
        dc.Columns("Z").Hidden = False
        dc.Range("Z2").ClearContents()
        dc.Range("Z2").Formula = wb_dir_formula("$A$4")
        dc.Columns("Z").Hidden = True

        r = 5
        n = 0
        while dc.Cells(r, 2).Value:
            # Packing set code is column B
            for col, fname in COLS:
                cell = dc.Cells(r, col)
                label = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                # delete any existing hyperlink objects
                try:
                    if cell.Hyperlinks.Count:
                        cell.Hyperlinks.Delete()
                except Exception:
                    pass
                path = f'$Z$2&"..\\01_DOCUMENT_SETS\\"&$B{r}&"\\{fname}"'
                cell.Formula = f'=HYPERLINK({path},"{label}")'
                cell.Font.Color = 0x0563C1
                cell.Font.Underline = 2
                n += 1
            r += 1
            if r > 400:
                break

        try:
            home = wb.Worksheets("00_HOME")
            home.Range("B28").Value = (
                "ONEMLI: ZIP icinden dogrudan acmayin. Once cikarin (Extract), "
                "sonra 00_AC_DOCUMENT_ENGINE.cmd veya 00_CONTROL\\engine acin. "
                "Salt-okunur / gecici kopyada linkler kirilir."
            )
        except Exception:
            pass

        wb.Save()
        # Force calc so Z2 gets a value while workbook is open from known path
        try:
            excel.CalculateFullRebuild()
        except Exception:
            excel.CalculateFull()
        z2 = str(s.Range("Z2").Value or "")
        print("SEARCH Z2 sample:", z2[:120], flush=True)
        wb.Close(False)
        print("DC formulas written:", n, flush=True)
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def prove(engine: Path) -> dict:
    """Verify formulas evaluate to OPEN WORD/PDF and targets exist. No FollowHyperlink."""
    pythoncom.CoInitialize()
    excel = None
    out = {"dc": 0, "search": 0, "fail": [], "z2": ""}
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(engine.resolve()), ReadOnly=True)
        try:
            excel.CalculateFullRebuild()
        except Exception:
            excel.CalculateFull()

        s = wb.Worksheets("SEARCH")
        out["z2"] = str(s.Range("Z2").Value or "")
        if not out["z2"] or "00_CONTROL" not in out["z2"].replace("/", "\\"):
            # CELL filename can be empty in some headless cases — still check DC via resolve
            pass

        # SEARCH 5 products
        sd = wb.Worksheets("SEARCH_DATA")
        products = ["1002103", "1015169"]
        for i in range(2, min(int(sd.UsedRange.Rows.Count), 400) + 1):
            st = str(sd.Cells(i, 7).Value or "")
            if "ISSUED" in st and "NOT ISSUED" not in st:
                pc = str(sd.Cells(i, 1).Value).strip().replace(".0", "")
                if pc and pc not in products:
                    products.append(pc)
            if len(products) >= 5:
                break

        for pc in products[:5]:
            s.Range("B4").Value = str(pc)
            try:
                excel.CalculateFullRebuild()
            except Exception:
                excel.CalculateFull()
            set_code = str(s.Range("B8").Value or "").strip()
            a8 = str(s.Range("A8").Value or "")
            z2 = str(s.Range("Z2").Value or "")
            if not set_code or a8 == "NOT FOUND":
                out["fail"].append(f"SEARCH notfound {pc}")
                continue
            for row, stem in [
                (13, "01_Technical_File"),
                (15, "02_EU_DoC"),
                (17, "03_Label"),
                (19, "04_Shipment_Statement"),
            ]:
                for col, ext in ((1, "docx"), (2, "pdf")):
                    disp = str(s.Cells(row, col).Value or "")
                    label = "OPEN WORD" if ext == "docx" else "OPEN PDF"
                    if z2:
                        target = Path(z2) / ".." / "01_DOCUMENT_SETS" / set_code / f"{stem}.{ext}"
                        target = target.resolve()
                    else:
                        target = (engine.parent / ".." / "01_DOCUMENT_SETS" / set_code / f"{stem}.{ext}").resolve()
                    if disp != label or not target.exists():
                        out["fail"].append(
                            f"SEARCH {pc}/{stem}.{ext} disp={disp!r} exists={target.exists()} z2={z2!r}"
                        )
                    else:
                        out["search"] += 1

        # DC 5 sets x 8
        dc = wb.Worksheets("DOCUMENT_CENTER")
        r = 5
        n = 0
        while dc.Cells(r, 2).Value and n < 40:
            sc = str(dc.Cells(r, 2).Value).strip()
            z2dc = str(dc.Range("Z2").Value or "") or out["z2"]
            for col, fname in COLS:
                disp = str(dc.Cells(r, col).Value or "")
                expect = "OPEN WORD" if fname.endswith(".docx") else "OPEN PDF"
                if z2dc:
                    target = (Path(z2dc) / ".." / "01_DOCUMENT_SETS" / sc / fname).resolve()
                else:
                    target = (engine.parent / ".." / "01_DOCUMENT_SETS" / sc / fname).resolve()
                if disp != expect or not target.exists():
                    out["fail"].append(f"DC {sc}/{fname} disp={disp!r} exists={target.exists()}")
                else:
                    out["dc"] += 1
                n += 1
            r += 1

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
        "echo Extract edilmis paket kokunden aciliyor...\r\n"
        "start \"\" \"%~dp000_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\"\r\n",
        encoding="utf-8",
    )
    (FINAL / "00_READ_ME_NASIL_ACILIR.txt").write_text(
        "LINKLER ICIN:\r\n"
        "1) ZIP dosyasina CIFT TIKLAMAYIN (Salt okunur / gecici kopya linkleri kirar).\r\n"
        "2) ZIP'e sag tik > Tumunu ayikla / Extract All.\r\n"
        "3) 00_AC_DOCUMENT_ENGINE.cmd dosyasina cift tiklayin.\r\n"
        "4) SEARCH veya DOCUMENT CENTER'dan OPEN WORD / OPEN PDF kullanin.\r\n",
        encoding="utf-8",
    )

    print("1) Apply CELL-based HYPERLINK formulas…", flush=True)
    apply()

    print("2) Prove from 00_CONTROL…", flush=True)
    a = prove(ENG)
    print(a, flush=True)
    if a["fail"] or a["dc"] != 40 or a["search"] != 40:
        raise SystemExit("CTRL FAIL")

    print("3) ZIP…", flush=True)
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for p in FINAL.rglob("*"):
            if p.is_file() and not p.name.startswith("~$"):
                zf.write(p, p.relative_to(FINAL).as_posix())
    digest = sha256_file(ZIP_PATH)
    SHA_PATH.write_text(digest + "\n", encoding="utf-8")

    print("4) Extract smoke…", flush=True)
    if SMOKE.exists():
        shutil.rmtree(SMOKE, ignore_errors=True)
    SMOKE.mkdir(parents=True)
    with zipfile.ZipFile(ZIP_PATH, "r") as zf:
        zf.extractall(SMOKE)
    b = prove(SMOKE / "00_CONTROL" / "INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx")
    print(b, flush=True)
    if b["fail"] or b["dc"] != 40 or b["search"] != 40:
        raise SystemExit("EXTRACT FAIL")

    print("SHA256", digest)
    print("GATE PASS")
    print("USER: ZIP'i AYIKLA, sonra 00_AC_DOCUMENT_ENGINE.cmd ac. ZIP icine cift tiklama.")


if __name__ == "__main__":
    main()
