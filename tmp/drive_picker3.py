"""Dump every filter-column value (repr) and drive picks while a background
thread dismisses the VBA message box."""
import shutil
import sys
import threading
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

stop = threading.Event()
dialog_log = []


def dialog_watch():
    while not stop.is_set():
        found = []

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if win32gui.GetClassName(hwnd) != "#32770":
                return
            title = win32gui.GetWindowText(hwnd)
            if "Sub Section" not in title:
                return
            texts = []
            child = win32gui.FindWindowEx(hwnd, 0, None, None)
            while child:
                t = win32gui.GetWindowText(child)
                if t:
                    texts.append(t)
                child = win32gui.FindWindowEx(hwnd, child, None, None)
            found.append((title, texts))
            win32gui.PostMessage(hwnd, win32con.WM_COMMAND, 1, 0)

        win32gui.EnumWindows(cb, None)
        for f in found:
            print("DIALOG:", f, flush=True)
        time.sleep(0.2)


threading.Thread(target=dialog_watch, daemon=True).start()

pythoncom.CoInitialize()
ex = win32com.client.DispatchEx("Excel.Application")
ex.Visible = False
ex.DisplayAlerts = False
ex.AutomationSecurity = 1
wb = None
try:
    wb = ex.Workbooks.Open(str(COPY))
    ws = wb.Worksheets("FIXED")

    col = ws.Cells(5, ws.Columns.Count).End(-4159).Column
    print("filter col:", col, "| autofilter:", ws.AutoFilter.Range.Address, flush=True)
    values = Counter()
    r = 6
    while ws.Cells(r, 1).Value not in (None, ""):
        values[repr(ws.Cells(r, col).Value)] += 1
        r += 3
    print("distinct filter values:", dict(values), flush=True)

    for pick in ("AO Finishing", "AO Barteck", "AO Hole Button", "All"):
        print(f"\n>>> pick {pick!r}", flush=True)
        ws.Range("A4").Value = pick
        time.sleep(0.5)
        print(f"  A4={ws.Range('A4').Value!r} filterMode={ws.FilterMode}", flush=True)
        try:
            print(f"  crit={ws.AutoFilter.Filters(col).Criteria1!r}", flush=True)
        except Exception as exc:
            print(f"  crit n/a: {exc}", flush=True)
        vis = Counter()
        r2 = 6
        while ws.Cells(r2, 1).Value not in (None, ""):
            if not ws.Rows(r2).Hidden:
                vis[repr(ws.Cells(r2, col).Value)] += 1
            r2 += 3
        print(f"  visible={dict(vis)}", flush=True)
finally:
    stop.set()
    time.sleep(0.4)
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
