import pythoncom
import win32com.client as win32
from pathlib import Path

p = Path(
    r"C:\Users\burcu\Documents\YAZILIM\Inci_Aku_PPWR_PIMS\output"
    r"\INCI_AKU_PPWR_STARTER_CUSTOMER_DELIVERY_REV00_FINAL\00_CONTROL"
    r"\INCI_AKU_PPWR_DOCUMENT_ENGINE_Rev00.xlsx"
)
pythoncom.CoInitialize()
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(str(p.resolve()))
s = wb.Worksheets("SEARCH")
s.Range("B4").Value = "1015169"
excel.CalculateFull()

formulas = {
    "Z10": '=HYPERLINK("..\\01_DOCUMENT_SETS\\ST-012-EUR-01\\01_Technical_File.docx","OPEN WORD")',
    "Z11": '=IF(TRUE,HYPERLINK("..\\01_DOCUMENT_SETS\\ST-012-EUR-01\\01_Technical_File.docx","OPEN WORD"),"X")',
    "Z12": '=HYPERLINK(IF(TRUE,"..\\01_DOCUMENT_SETS\\ST-012-EUR-01\\01_Technical_File.docx",""),IF(TRUE,"OPEN WORD","NO"))',
}
for a, f in formulas.items():
    s.Range(a).Formula = f
excel.CalculateFull()
for a in formulas:
    try:
        c = s.Range(a).Hyperlinks.Count
    except Exception as e:
        c = str(e)
    print(a, "val", s.Range(a).Value, "hlcount", c)

dc = wb.Worksheets("DOCUMENT_CENTER")
print("DC F5 hlcount", dc.Cells(5, 6).Hyperlinks.Count, "addr", dc.Cells(5, 6).Hyperlinks(1).Address)

wb.Close(False)
excel.Quit()
pythoncom.CoUninitialize()
