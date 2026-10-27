from pathlib import Path

import pythoncom
import win32com.client

path = Path(r"D:\MERND\output\Salary_Attendance_August_2026.xlsx").resolve()

pythoncom.CoInitialize()
excel = None
try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(str(path))
    try:
        ws = wb.Worksheets("FIXED")
        ws.Range("A5:AT934").AutoFilter(
            Field=46,
            Criteria1=["AO-Finishing", "Packing"],
            Operator=7,  # xlFilterValues
        )
        excel.Calculate()
        print("A4 after multi-filter:", ws.Range("A4").Value)
        print("visible:", ws.Range("A5:AT934").SpecialCells(12).Rows.Count)
        ws.ShowAllData()
        excel.Calculate()
        print("A4 after clear:", ws.Range("A4").Value)
    finally:
        wb.Close(SaveChanges=False)
finally:
    if excel is not None:
        excel.Quit()
    pythoncom.CoUninitialize()
