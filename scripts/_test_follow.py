import pythoncom
import win32com.client as win32
from pathlib import Path

p = Path(
    r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output"
    r"\INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL\00_CONTROL"
    r"\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
)
root = Path(
    r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output"
    r"\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
)

pythoncom.CoInitialize()
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

for label, path in [("CTRL", p), ("ROOT", root)]:
    wb = excel.Workbooks.Open(str(path.resolve()))
    print("===", label, "Path=", wb.Path)
    try:
        print("HyperlinkBase=", repr(wb.BuiltinDocumentProperties("Hyperlink Base").Value))
    except Exception as e:
        print("HyperlinkBase err", e)
    dc = wb.Worksheets("DOCUMENT_CENTER")
    dc.Activate()
    cell = dc.Cells(5, 6)
    print("Value", cell.Value, "HL", cell.Hyperlinks(1).Address if cell.Hyperlinks.Count else None)
    try:
        cell.Hyperlinks(1).Follow(True, False)  # new window? 
        print("Follow OK")
    except Exception as e:
        print("Follow FAIL", e)
    wb.Close(False)

excel.Quit()
pythoncom.CoUninitialize()
