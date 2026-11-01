"""Inject a logging version of the filter routine into a copy of the workbook
and drive the picks, capturing what the VBA actually does.
"""
import shutil
import sys
import threading
import time
from pathlib import Path

import win32con
import win32gui

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pythoncom
import win32com.client

ROOT = Path(r"D:\STL-Salary-Software")
SRC = ROOT / "tmp" / "user_sim.xlsm"
COPY = ROOT / "tmp" / "user_debug.xlsm"
LOG = ROOT / "tmp" / "filter_log.txt"

try:
    COPY.unlink(missing_ok=True)
except PermissionError:
    time.sleep(2)
    COPY.unlink(missing_ok=True)
shutil.copyfile(SRC, COPY)
LOG.unlink(missing_ok=True)

DEBUG_CODE = r'''
Public Sub FilterDebug(ByVal ws As Worksheet)
    Dim chosen As String
    Dim lastCol As Long
    Dim lastRow As Long
    Dim filterCol As Long
    Dim rng As Range
    Dim visCount As Long
    Dim crit As String
    Dim afRange As String

    On Error GoTo Failed
    chosen = CStr(ws.Range("A4").Value)
    LogLine "---- pick=" & "[" & chosen & "]"
    lastCol = ws.Cells(5, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    afRange = "(none)"
    On Error Resume Next
    afRange = ws.AutoFilter.Range.Address
    On Error GoTo Failed
    LogLine "lastCol=" & lastCol & " lastRow=" & lastRow & " autofilter=" & afRange
    filterCol = lastCol
    Set rng = ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol))
    LogLine "rng=" & rng.Address & " filterCol=" & filterCol
    On Error Resume Next
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo Failed
    rng.AutoFilter Field:=filterCol, Criteria1:=chosen
    crit = "(n/a)"
    On Error Resume Next
    crit = CStr(ws.AutoFilter.Filters(filterCol).Criteria1)
    On Error GoTo Failed
    LogLine "after AutoFilter: filterMode=" & ws.FilterMode & " crit=" & crit
    visCount = Application.WorksheetFunction.Subtotal(103, ws.Range(ws.Cells(6, 1), ws.Cells(lastRow, 1)))
    LogLine "visCount=" & visCount
    Exit Sub
Failed:
    LogLine "ERROR " & Err.Number & ": " & Err.Description
End Sub

Public Sub LogLine(ByVal text As String)
    Dim f As Integer
    f = FreeFile
    Open "D:\STL-Salary-Software\tmp\filter_log.txt" For Append As #f
    Print #f, text
    Close #f
End Sub
'''

stop = threading.Event()


def dialog_watch():
    while not stop.is_set():
        found = []

        def cb(hwnd, _):
            if not win32gui.IsWindowVisible(hwnd):
                return
            if win32gui.GetClassName(hwnd) != "#32770":
                return
            title = win32gui.GetWindowText(hwnd)
            if "Sub Section" not in title and "Microsoft" not in title:
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
    project = wb.VBProject
    comp = project.VBComponents.Add(1)  # vbext_ct_StdModule
    comp.Name = "DebugFilterMod"
    comp.CodeModule.AddFromString(DEBUG_CODE)
    time.sleep(1)

    ex.EnableEvents = False
    ws = wb.Worksheets("FIXED")
    ws.Activate()

    picks = ["AO Barteck", "AO Finishing", "AO Hole Button", "AO Barteck", "All"]
    for pick in picks:
        print(f"\n>>> pick {pick!r}", flush=True)
        ws.Range("A4").Value = pick
        ex.Run("FilterDebug", ws)
        time.sleep(0.3)
        print(f"  A4 now={ws.Range('A4').Value!r}", flush=True)
        r = 6
        vis = {}
        while ws.Cells(r, 1).Value not in (None, ""):
            if not ws.Rows(r).Hidden:
                v = str(ws.Cells(r, 45).Value)
                vis[v] = vis.get(v, 0) + 1
            r += 3
        print(f"  visible={vis} formula={ws.Range('A4').Formula!r}", flush=True)
finally:
    stop.set()
    try:
        if wb is not None:
            wb.Close(SaveChanges=False)
    except Exception:
        pass
    ex.Quit()
    pythoncom.CoUninitialize()

print("\n=== filter_log.txt ===")
print(LOG.read_text(encoding="utf-8", errors="replace"))
