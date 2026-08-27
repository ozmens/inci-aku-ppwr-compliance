"""Fix SEARCH hyperlinks using workbook-folder path via CELL(\"filename\").

Document Center keeps Hyperlink objects with ..\\ relative (workbook-relative).
SEARCH uses outermost HYPERLINK + CELL filename (formula-relative is broken).
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
DESK_LNK = Path(r"C:\Users\burcu\Desktop\INCI_AKU_PPWR_DOCUMENT_ENGINE.lnk")

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


def search_action_formula(stem: str, word: bool) -> str:
    label = "OPEN WORD" if word else "OPEN PDF"
    ext = "docx" if word else "pdf"
    # Inline workbook directory — evaluated when clicked
    folder = 'LEFT(CELL("filename",$B$4),FIND("[",CELL("filename",$B$4))-1)'
    path = f'IFERROR({folder},"")&"..\\01_DOCUMENT_SETS\\"&$B$8&"\\{stem}.{ext}"'
    issued = (
        'AND($B$8<>"",$A$8<>"NOT FOUND",ISNUMBER(SEARCH("ISSUED",$G$8)),'
        'NOT(ISNUMBER(SEARCH("NOT ISSUED",$G$8))),NOT(ISNUMBER(SEARCH("YURT",$G$8))))'
    )
    domestic = 'OR(ISNUMBER(SEARCH("YURT",$G$8)),ISNUMBER(SEARCH("NOT ISSUED",$G$8)))'
    empty = 'OR($B$4="",$A$8="",$A$8="NOT FOUND")'
    # Outermost HYPERLINK; "#" when inactive avoids broken empty links
    return (
        f'=IF({empty},"",'
        f'IF({domestic},"DOCUMENTS NOT ISSUED",'
        f'HYPERLINK(IF({issued},{path},"#"),IF({issued},"{label}",""))))'
    )


def apply() -> int:
    pythoncom.CoInitialize()
    excel = None
    n = 0
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(ENG.resolve()))

        # SEARCH
        s = wb.Worksheets("SEARCH")
        s.Columns("Z").Hidden = False
        s.Range("Z1").ClearContents()
        s.Range("Z1").Formula = '=IF($B$4="","",TRIM($B$4&""))'
        s.Range("AA1").ClearContents()
        s.Range("AA1").Formula = (
            '=IF($Z$1="","",IFERROR(MATCH($Z$1,SEARCH_DATA!H:H,0),'
            "IFERROR(MATCH($Z$1,SEARCH_DATA!A:A,0),"
            "IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!A:A,0),"
            'IFERROR(MATCH(VALUE($Z$1),SEARCH_DATA!H:H,0),"NOT FOUND")))))'
        )
        for row, stem in [
            (13, "01_Technical_File"),
            (15, "02_EU_DoC"),
            (17, "03_Label"),
            (19, "04_Shipment_Statement"),
        ]:
            s.Cells(row, 1).Formula = search_action_formula(stem, True)
            s.Cells(row, 2).Formula = search_action_formula(stem, False)
            for c in (1, 2):
                s.Cells(row, c).Font.Color = 0x0563C1
                s.Cells(row, c).Font.Underline = 2
        s.Columns("Z").Hidden = True
        s.Columns("AA").Hidden = True

        # DOCUMENT CENTER — Hyperlink objects (workbook-relative)
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
                # clear formula if previous fix left HYPERLINK formulas
                cell.Value = label
                dc.Hyperlinks.Add(Anchor=cell, Address=win_rel(sc, fname), TextToDisplay=label)
                cell.Font.Color = 0x0563C1
                cell.Font.Underline = 2
                n += 1
            r += 1
            if r > 400:
                break

        try:
            home = wb.Worksheets("00_HOME")
            home.Range("B28").Value = (
                "KRITIK: ZIP icine cift tiklamayin (Salt Okunur). "
                "Once Extract All, sonra 00_AC_DOCUMENT_ENGINE.cmd acin."
            )
        except Exception:
            pass

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
                    out["fail"].append(f"DC no hl {sc}/{fname}")
                else:
                    addr = str(cell.Hyperlinks(1).Address)
                    target = (engine.parent / addr).resolve()
                    if not addr.startswith("..\\01_DOCUMENT_SETS\\") or not target.exists():
                        out["fail"].append(f"DC {sc}/{fname} {addr}")
                    else:
                        out["dc"] += 1
                n += 1
            r += 1

        s = wb.Worksheets("SEARCH")
        # Must show OPEN PDF for 1002103 / ST-045-STD-09
        for pc in ["1002103", "1015169", "1000069", "1000441", "1000442"]:
            s.Range("B4").Value = pc
            try:
                excel.CalculateFullRebuild()
            except Exception:
                excel.CalculateFull()
            set_code = str(s.Range("B8").Value or "").strip()
            if not set_code or str(s.Range("A8").Value) == "NOT FOUND":
                out["fail"].append(f"SEARCH miss {pc}")
                continue
            # Formula must contain CELL("filename") 
            f15 = str(s.Range("B15").Formula or "")
            if 'CELL("filename"' not in f15 and "CELL(\"filename\"" not in f15:
                # Turkish Excel may store differently; check filename
                if "filename" not in f15.lower():
                    out["fail"].append(f"SEARCH formula missing CELL {pc}")
            for row, stem in [
                (13, "01_Technical_File"),
                (15, "02_EU_DoC"),
                (17, "03_Label"),
                (19, "04_Shipment_Statement"),
            ]:
                for col, ext in ((1, "docx"), (2, "pdf")):
                    disp = str(s.Cells(row, col).Value or "")
                    label = "OPEN WORD" if ext == "docx" else "OPEN PDF"
                    target = (
                        engine.parent / ".." / "01_DOCUMENT_SETS" / set_code / f"{stem}.{ext}"
                    ).resolve()
                    if disp != label or not target.exists():
                        out["fail"].append(f"SEARCH {pc}/{stem}.{ext} disp={disp!r}")
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


def desktop_shortcut() -> None:
    try:
        shell = win32.Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(str(DESK_LNK))
        sc.Targetpath = str(ENG.resolve())
        sc.WorkingDirectory = str(ENG.parent.resolve())
        sc.Description = "PPWR Document Engine — 00_CONTROL only"
        sc.save()
        print("shortcut", DESK_LNK, flush=True)
    except Exception as e:
        print("shortcut skip", e, flush=True)


def main() -> None:
    (FINAL / "00_AC_DOCUMENT_ENGINE.cmd").write_text(
        "@echo off\r\n"
        "cd /d \"%~dp0\"\r\n"
        "start \"\" \"%~dp000_CONTROL\\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx\"\r\n",
        encoding="utf-8",
    )
    (FINAL / "00_READ_ME_NASIL_ACILIR.txt").write_text(
        "1) ZIP dosyasina CIFT TIKLAMAYIN.\r\n"
        "2) Sag tik > Tumunu ayikla / Extract All.\r\n"
        "3) 00_AC_DOCUMENT_ENGINE.cmd calistirin.\r\n"
        "4) SEARCH'te urun kodu girip OPEN WORD / OPEN PDF tiklayin.\r\n"
        "Salt Okunur (Read Only) acildiysa linkler CALISMAZ — kapatip Extract ile acin.\r\n",
        encoding="utf-8",
    )

    print("1) Apply…", flush=True)
    print("DC links", apply(), flush=True)
    desktop_shortcut()

    print("2) Prove CTRL…", flush=True)
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

    print("4) Extract prove…", flush=True)
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
    print("CLOSE salt-okunur Excel. Open Desktop shortcut or 00_AC_DOCUMENT_ENGINE.cmd")


if __name__ == "__main__":
    main()
