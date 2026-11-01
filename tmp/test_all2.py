"""Isolate the All-restore failure: protection vs logic vs event path."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = r"D:\STL-Salary-Software\output\Salary_Attendance_September_2026.xlsm"
COPY = r"D:\STL-Salary-Software\tmp\all_test2.xlsm"

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
        except Exception as e:
            crit = f"(n/a: {e})"
        print(f"{tag}: filterMode={ws.FilterMode} visible={visible()} crit45={crit}")

    # 1. filter to one subsection via the normal event path
    ws.Range("A4").Value = "AO-Knitting"
    time.sleep(0.4)
    state("after pick AO-Knitting")

    # 2. try python-direct ShowAllData (is it protection-blocked?)
    try:
        ws.ShowAllData()
        print("ShowAllData: OK")
    except Exception as e:
        print("ShowAllData FAILED:", e)
    state("after ShowAllData")

    # 3. filter again via event, then call the VBA sub directly with A4='All'
    ws.Range("A4").Value = "AO-Knitting"
    time.sleep(0.4)
    state("re-filtered AO-Knitting")
    ws.Range("A4").Value = "All"
    try:
        ex.Run("ApplySubSectionFilter", ws)
        print("direct ApplySubSectionFilter: ran without exception")
    except Exception as e:
        print("direct ApplySubSectionFilter FAILED:", e)
    state("after direct call with A4=All")
    print("A4 content:", repr(ws.Range("A4").Value))
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
