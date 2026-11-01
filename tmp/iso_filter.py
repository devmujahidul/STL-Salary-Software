"""Isolate raw AutoFilter behavior on the FIXED sheet, no VBA events."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = Path(r"D:\STL-Salary-Software\tmp\user_sim.xlsm")
COPY = Path(r"D:\STL-Salary-Software\tmp\user_sim_iso.xlsm")

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
    ws = wb.Worksheets("FIXED")
    ex.EnableEvents = False

    col = ws.Cells(5, ws.Columns.Count).End(-4159).Column
    rng = ws.Range(ws.Cells(5, 1), ws.Cells(153, col))
    print("range:", rng.Address, "| col:", col, flush=True)

    def visible():
        vis = Counter()
        r = 6
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                vis[str(ws.Cells(r, col).Value)] += 1
            r += 3
        return dict(vis)

    def state(tag):
        try:
            crit = ws.AutoFilter.Filters(col).Criteria1
        except Exception as exc:
            crit = f"n/a({exc})"
        print(f"  {tag}: A4={ws.Range('A4').Value!r} filterMode={ws.FilterMode} crit={crit!r} visible={visible()}", flush=True)

    picks = ["AO Hole Button", "AO Barteck", "AO Finishing", "AO Hole Button", "AO Barteck"]
    for pick in picks:
        print(f"\n>>> raw autofilter {pick!r}", flush=True)
        if ws.FilterMode:
            ws.ShowAllData()
        rng.AutoFilter(Field=col, Criteria1=pick)
        state("after autofilter")
        ws.ShowAllData()
        state("after showall")

    # same but via the VBA routine with events disabled
    print("\n=== direct ApplySubSectionFilter calls ===", flush=True)
    for pick in ("AO Hole Button", "AO Barteck", "AO Finishing"):
        ws.Range("A4").Value = pick
        ex.Run("ApplySubSectionFilter", ws)
        state(f"after ApplySubSectionFilter {pick!r}")
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
