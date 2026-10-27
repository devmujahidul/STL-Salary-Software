from pathlib import Path

import pythoncom
import win32com.client

path = Path(r"D:\MERND\output\Salary_Attendance_August_2026.xlsm").resolve()

pythoncom.CoInitialize()
excel = None
try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityLow: enable macros
    wb = excel.Workbooks.Open(str(path))
    try:
        ws = wb.Worksheets("FIXED")
        print("before A4:", ws.Range("A4").Value)
        print("before filter mode:", ws.AutoFilterMode)
        ws.Range("A4").Value = "AO-Finishing"
        excel.Calculate()
        print("after select A4:", ws.Range("A4").Value)
        print("filter mode:", ws.AutoFilterMode, "filter on:", ws.AutoFilter.FilterMode)
        print("criteria:", ws.AutoFilter.Filters(46).Criteria1)
        visible = 0
        for r in range(6, ws.UsedRange.Rows.Count + 1):
            if not ws.Rows(r).Hidden:
                visible += 1
        print("visible data rows:", visible)

        ws.Range("A4").Value = "All"
        excel.Calculate()
        print("after All A4:", ws.Range("A4").Value)
        print("filter on after All:", ws.AutoFilter.FilterMode)
        visible = 0
        for r in range(6, ws.UsedRange.Rows.Count + 1):
            if not ws.Rows(r).Hidden:
                visible += 1
        print("visible data rows after All:", visible)
    finally:
        wb.Close(SaveChanges=False)
finally:
    if excel is not None:
        excel.Quit()
    pythoncom.CoUninitialize()
