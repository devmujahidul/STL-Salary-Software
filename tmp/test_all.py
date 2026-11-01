"""Nail down the All-restore failure: event firing + protection interference."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = r"D:\STL-Salary-Software\output\Salary_Attendance_September_2026.xlsm"
COPY = r"D:\STL-Salary-Software\tmp\all_test.xlsm"

for _ in range(5):
    try:
        Path(COPY).unlink(missing_ok=True)
        break
    except PermissionError:
        time.sleep(1)
shutil.copyfile(SRC, COPY)

pythoncom.CoInitialize()
ex = win32com.client.DispatchEx("Excel.Application")
ex.Visible = False
ex.DisplayAlerts = False
wb = None
try:
    wb = ex.Workbooks.Open(COPY)
    ws = wb.Worksheets("FIXED")
    sec_col = 45

    def visible():
        vis = Counter()
        r = 6
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                vis[str(ws.Cells(r, sec_col).Value or "")] += 1
            r += 3
        return dict(vis)

    def pick(name):
        ex.EnableEvents = True  # make sure the handler can fire
        print(f"-- pick {name!r} (events={ex.EnableEvents})")
        try:
            ws.Range("A4").Value = name
        except Exception as e:
            print("   A4 set raised:", e)
        time.sleep(0.4)
        print(f"   A4={ws.Range('A4').Value!r} events={ex.EnableEvents} filterMode={ws.FilterMode} visible={visible()}")

    pick("AO-Knitting")
    pick("AO Barteck")   # no match -> everything hidden?
    pick("All")          # does this restore?
    pick("All")          # and a second time?
finally:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    ex.Quit()
    pythoncom.CoUninitialize()
    for _ in range(5):
        try:
            Path(COPY).unlink(missing_ok=True)
            break
        except PermissionError:
            time.sleep(1)
