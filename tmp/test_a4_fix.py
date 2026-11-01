"""Full A4 picker flow after the fix (DisplayAlerts=False suppresses MsgBox)."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = r"D:\STL-Salary-Software\output\Salary_Attendance_September_2026.xlsm"
COPY = r"D:\STL-Salary-Software\tmp\a4_fix_test.xlsm"

ok = True
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

    def pick(name, expect_vis, expect_label_contains):
        global ok
        ws.Range("A4").Value = name
        time.sleep(0.4)
        got = visible()
        a4 = str(ws.Range("A4").Value)
        good = got == expect_vis and expect_label_contains in a4
        print(f"pick {name!r:24} -> A4={a4!r}  visible={got}  {'PASS' if good else 'FAIL'}")
        if not good:
            ok = False

    all_vis = {"AO-Finishing": 22, "AO-Knitting": 28}

    pick("AO-Knitting", {"AO-Knitting": 28}, "AO-Knitting")
    pick("All", all_vis, "All")                       # previously broken
    pick("AO-Finishing", {"AO-Finishing": 22}, "AO-Finishing")
    pick("AO Barteck", all_vis, "All")                # no match -> reset + message suppressed
    pick("AO-Knitting", {"AO-Knitting": 28}, "AO-Knitting")

    # multi-select via the '+' form (what the checkbox flow produces)
    ex.Run("ApplySubSectionFilter", ws)
    ws.Range("A4").Value = "AO-Knitting + AO-Finishing"
    time.sleep(0.4)
    got = visible()
    a4 = str(ws.Range("A4").Value)
    good = got == all_vis and "AO-Knitting" in a4 and "AO-Finishing" in a4
    print(f"multi 'AO-Knitting + AO-Finishing' -> A4={a4!r} visible={got} {'PASS' if good else 'FAIL'}")
    if not good:
        ok = False

    pick("All", all_vis, "All")                       # recover from multi
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

print()
print("RESULT:", "ALL PASS" if ok else "FAILURES ABOVE")
