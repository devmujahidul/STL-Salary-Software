"""Stress-test repeated AutoFilter applications on the same field, with and
without ShowAllData in between. Look for the state where filtering silently
leaves everything visible (the user-reported symptom).
"""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = Path(r"D:\STL-Salary-Software\tmp\user_sim.xlsm")
COPY = Path(r"D:\STL-Salary-Software\tmp\user_sim_stress.xlsm")

try:
    COPY.unlink(missing_ok=True)
except PermissionError:
    time.sleep(2)
    COPY.unlink(missing_ok=True)
shutil.copyfile(SRC, COPY)

pythoncom.CoInitialize()
ex = win32com.client.DispatchEx("Excel.Application")
ex.Visible = False
ex.DisplayAlerts = False
ex.AutomationSecurity = 1
wb = None
try:
    wb = ex.Workbooks.Open(str(COPY))
    ex.EnableEvents = False
    ws = wb.Worksheets("FIXED")
    ws.Activate()
    col = ws.Cells(5, ws.Columns.Count).End(-4159).Column
    last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    rng = ws.Range(ws.Cells(5, 1), ws.Cells(last, col))
    print(f"col={col} last={last} rng={rng.Address} autofilter={ws.AutoFilter.Range.Address}", flush=True)

    def visible():
        vis = Counter()
        r = 6
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                vis[str(ws.Cells(r, col).Value or "")] += 1
            r += 3
        return dict(vis)

    picks = ["AO Barteck", "AO Finishing", "AO Hole Button"] * 3
    mode = sys.argv[1] if len(sys.argv) > 1 else "showall"
    print("mode:", mode, flush=True)

    fails = 0
    for i, pick in enumerate(picks):
        if mode == "showall" and ws.FilterMode:
            ws.ShowAllData()
        rng.AutoFilter(Field=col, Criteria1=pick)
        time.sleep(0.2)
        vis = visible()
        ok = set(vis) == {pick}
        if not ok:
            fails += 1
        print(f"  #{i:02d} {pick!r:18} filterMode={ws.FilterMode} visible={vis} {'OK' if ok else 'MISMATCH'}", flush=True)
    print("mismatches:", fails, flush=True)

    # also check Subtotal right after a mismatch
finally:
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    ex.Quit()
    pythoncom.CoUninitialize()
    time.sleep(1)
    try:
        COPY.unlink(missing_ok=True)
    except PermissionError:
        pass
