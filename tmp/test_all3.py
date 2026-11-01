"""Definitive: after a zero-match pick, do later event picks still run?"""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = r"D:\STL-Salary-Software\output\Salary_Attendance_September_2026.xlsm"
COPY = r"D:\STL-Salary-Software\tmp\all_test3.xlsm"

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

    def state(tag):
        try:
            crit = ws.AutoFilter.Filters(45).Criteria1
        except Exception:
            crit = "(none)"
        print(f"{tag}: A4.Value={ws.Range('A4').Value!r} A4.Formula={str(ws.Range('A4').Formula)[:60]!r} "
              f"filterMode={ws.FilterMode} events={ex.EnableEvents} visible={visible()} crit={crit}")

    state("initial")

    ws.Range("A4").Value = "AO-Knitting"
    time.sleep(0.4)
    state("after AO-Knitting")

    ws.Range("A4").Value = "AO Barteck"
    time.sleep(0.4)
    state("after AO Barteck (zero match)")

    ws.Range("A4").Value = "All"
    time.sleep(0.4)
    state("after All (event)")

    ws.Range("A4").Value = "AO-Finishing"
    time.sleep(0.4)
    state("after AO-Finishing (event)")
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
