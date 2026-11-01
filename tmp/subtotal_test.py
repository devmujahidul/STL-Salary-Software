"""Check WorksheetFunction.Subtotal vs actual visible rows after filtering."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = Path(r"D:\STL-Salary-Software\tmp\user_sim.xlsm")
COPY = Path(r"D:\STL-Salary-Software\tmp\user_sim_sub.xlsm")

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
    for sheet in ("FIXED", "LO_PO"):
        ws = wb.Worksheets(sheet)
        col = ws.Cells(5, ws.Columns.Count).End(-4159).Column
        last = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
        rng = ws.Range(ws.Cells(5, 1), ws.Cells(last, col))
        print(f"\n=== {sheet} col={col} last={last} range={rng.Address}", flush=True)
        ws.Activate()
        for pick in ("AO Barteck", "AO Finishing", "AO Hole Button"):
            if ws.FilterMode:
                ws.ShowAllData()
            rng.AutoFilter(Field=col, Criteria1=pick)
            vis = Counter()
            r = 6
            while ws.Cells(r, 1).Value not in (None, ""):
                if not ws.Rows(r).Hidden:
                    vis[str(ws.Cells(r, col).Value)] += 1
                r += 3
            sub103 = ex.WorksheetFunction.Subtotal(103, ws.Range(ws.Cells(6, 1), ws.Cells(last, 1)))
            sub3 = ex.WorksheetFunction.Subtotal(3, ws.Range(ws.Cells(6, 1), ws.Cells(last, 1)))
            countif = ex.WorksheetFunction.CountIf(ws.Range(ws.Cells(6, col), ws.Cells(last, col)), pick)
            print(f"  {pick!r}: visible={dict(vis)} subtotal103={sub103} subtotal3={sub3} countif={countif}", flush=True)
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
