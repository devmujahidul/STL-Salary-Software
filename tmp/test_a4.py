"""Reproduce the A4 subsection picker behavior on the current build."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = r"D:\STL-Salary-Software\output\Salary_Attendance_September_2026.xlsm"
COPY = r"D:\STL-Salary-Software\tmp\a4_test.xlsm"

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

    # row-5 titles right of the last day column
    titles = []
    for c in range(33, 47):
        titles.append((c, str(ws.Cells(5, c).Value)))
    print("row-5 titles:", titles)

    # the A4 dropdown source
    try:
        print("A4 validation formula:", ws.Range("A4").Validation.Formula1)
    except Exception as e:
        print("A4 validation: n/a", e)

    lists = wb.Worksheets("Lists")
    lst = []
    r = 2
    while str(lists.Cells(r, 49).Value or "") != "":
        lst.append(str(lists.Cells(r, 49).Value))
        r += 1
    print("Lists (col AX):", lst)

    def visible_state(label):
        vis = Counter()
        r = 6
        sec_col = 45
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                vis[str(ws.Cells(r, sec_col).Value or "")] += 1
            r += 3
        af = ws.AutoFilterMode
        fm = ws.FilterMode
        print(f"  A4 = {ws.Range('A4').Value!r}  | filterMode={fm} autofilterMode={af} | visible by sub: {dict(vis)}")

    print("initial state:")
    visible_state("init")

    for pick in ("AO-Knitting", "AO-Finishing", "AO Barteck", "All"):
        print(f"pick {pick!r}:")
        ws.Range("A4").Value = pick
        time.sleep(0.4)
        visible_state(pick)
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
