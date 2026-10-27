import sys
from pathlib import Path

import pythoncom
import win32com.client

xlsx = Path(sys.argv[1]).resolve()
pdf = Path(sys.argv[2]).resolve()
pdf.parent.mkdir(parents=True, exist_ok=True)

pythoncom.CoInitialize()
excel = None
try:
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = excel.Workbooks.Open(str(xlsx))
    wb.ExportAsFixedFormat(0, str(pdf))
    wb.Close(SaveChanges=False)
    print(f"Exported {pdf}")
finally:
    if excel is not None:
        excel.Quit()
    pythoncom.CoUninitialize()
