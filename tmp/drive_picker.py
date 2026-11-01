"""Drive the A4 sub-section picker on every attendance sheet and report."""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

SRC = Path(r"D:\STL-Salary-Software\tmp\user_sim.xlsm")
COPY = Path(r"D:\STL-Salary-Software\tmp\user_sim_run.xlsm")

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
wb = None
try:
    wb = ex.Workbooks.Open(str(COPY))

    def filter_col(ws):
        # last header cell in row 5
        return ws.Cells(5, ws.Columns.Count).End(-4159).Column  # xlToLeft

    def visible_counts(ws):
        col = filter_col(ws)
        vis = Counter()
        r = 6
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                vis[str(ws.Cells(r, col).Value or "")] += 1
            r += 3
        return dict(vis)

    def subs_for(ws):
        col = filter_col(ws)
        s = set()
        r = 6
        while ws.Cells(r, 1).Value not in (None, ""):
            s.add(str(ws.Cells(r, col).Value or ""))
            r += 3
        return sorted(s)

    for sheet in ("FIXED", "LO_PO"):
        ws = wb.Worksheets(sheet)
        names = subs_for(ws)
        print(f"\n=== {sheet}: subsections {names} ===")
        for pick in names + ["All"]:
            ws.Range("A4").Value = pick
            time.sleep(0.5)
            a4 = str(ws.Range("A4").Value)
            vis = visible_counts(ws)
            try:
                crit = ws.AutoFilter.Filters(filter_col(ws)).Criteria1
            except Exception as exc:
                crit = f"(n/a {exc})"
            print(f"pick {pick!r:20} -> A4={a4!r:55} visible={vis} crit={crit!r} filterMode={ws.FilterMode}")
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
