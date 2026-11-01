"""Inspect the FIXED sheet: filter cell values (repr) and pick behavior, and
auto-dismiss any VBA message box so the automation never hangs.
"""
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

import win32con
import win32gui

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


def dismiss_dialogs():
    """Close every Excel modal dialog, returning its text."""
    seen = []

    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        if win32gui.GetClassName(hwnd) != "#32770":
            return
        title = win32gui.GetWindowText(hwnd)
        if "Sub Section" not in title and "Excel" not in title and "Microsoft" not in title:
            return
        texts = []
        child = win32gui.FindWindowEx(hwnd, 0, None, None)
        while child:
            t = win32gui.GetWindowText(child)
            if t:
                texts.append(t)
            child = win32gui.FindWindowEx(hwnd, child, None, None)
        seen.append((title, texts))
        win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)  # click OK
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)

    win32gui.EnumWindows(cb, None)
    return seen


pythoncom.CoInitialize()
ex = win32com.client.DispatchEx("Excel.Application")
ex.Visible = False
ex.DisplayAlerts = False
ex.AutomationSecurity = 1  # force-enable macros
wb = None
try:
    wb = ex.Workbooks.Open(str(COPY))
    ws = wb.Worksheets("FIXED")

    col = ws.Cells(5, ws.Columns.Count).End(-4159).Column
    print("filter column:", col, "autofilter range:", ws.AutoFilter.Range.Address)
    print("last row (col A):", ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
    r = 6
    shown = 0
    while ws.Cells(r, 1).Value not in (None, "") and shown < 8:
        print(f"  row {r}: {ws.Cells(r, col).Value!r}")
        r += 3
        shown += 1

    for pick in ("AO Barteck", "AO Finishing", "AO Hole Button", "All"):
        ws.Range("A4").Value = pick
        time.sleep(0.6)
        dialogs = dismiss_dialogs()
        time.sleep(0.3)
        print(f"\npick {pick!r}: dialogs={dialogs}")
        print(f"  A4={ws.Range('A4').Value!r}")
        print(f"  filterMode={ws.FilterMode}")
        try:
            print(f"  crit={ws.AutoFilter.Filters(col).Criteria1!r}")
        except Exception as exc:
            print(f"  crit n/a: {exc}")
        vis = Counter()
        r2 = 6
        while ws.Cells(r2, 1).Value not in (None, ""):
            if not ws.Rows(r2).Hidden:
                vis[str(ws.Cells(r2, col).Value or "")] += 1
            r2 += 3
        print(f"  visible={dict(vis)}")
finally:
    dismiss_dialogs()
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
