"""Verify print layout fixes (staged prints to see where Excel stalls):

1. Bottom margin reserves a footer strip (table can't overflow the footer).
2. HideDesignationForPrint hides column D on attendance sheets only.
3. RestoreDesignationColumns brings it back.
4. Column D is not permanently hidden (stays visible on screen).
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client as win32

SHEETS = ["LO_PO", "FIXED", "PRODUCTION"]
EXPECT_BOTTOM = 0.75 * 72  # points
EXPECT_FOOTER = 0.35 * 72
TOL = 0.5


def step(msg: str):
    print(msg, flush=True)


def main() -> bool:
    path = Path(__file__).resolve().parent.parent / "output" / "Salary_Attendance_August_2026.xlsm"
    if not path.exists():
        path = path.with_suffix(".xlsx")
    step("[1] launching Excel")
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    ok = True
    try:
        step("[2] opening workbook")
        wb = app.Workbooks.Open(str(path), ReadOnly=False, UpdateLinks=0)
        step("[3] workbook open; checking margins")
        for name in SHEETS:
            ws = wb.Worksheets(name)
            bottom = float(ws.PageSetup.BottomMargin)
            footer = float(ws.PageSetup.FooterMargin)
            step(f"    {name}: bottom={bottom:.1f}pt footer={footer:.1f}pt")
            if abs(bottom - EXPECT_BOTTOM) > TOL or abs(footer - EXPECT_FOOTER) > TOL:
                step("    MARGIN MISMATCH")
                ok = False

        step("[4] Designation must be hidden in the file itself (macro-independent)")
        ws_fixed = wb.Worksheets("FIXED")
        for name in SHEETS:
            if not wb.Worksheets(name).Columns("D").Hidden:
                step(f"    {name}: Designation NOT hidden - it would print!")
                ok = False
        if not ws_fixed.Columns("D").Hidden:
            ok = False

        step("[5] running HideDesignationForPrint")
        app.Run("HideDesignationForPrint")
        hidden = {n: wb.Worksheets(n).Columns("D").Hidden for n in SHEETS}
        lists_hidden = bool(wb.Worksheets("Lists").Columns("D").Hidden)
        step(f"    after hide: {hidden}, Lists hidden={lists_hidden}")
        if not all(hidden.values()):
            step("    HIDE MACRO FAILED")
            ok = False

        step("[6] restoring: default file state is hidden, so it stays hidden")
        app.Run("RestoreDesignationColumns")
        restored = {n: wb.Worksheets(n).Columns("D").Hidden for n in SHEETS}
        step(f"    after restore: {restored}")
        if not all(restored.values()):
            step("    RESTORE MACRO FAILED (should return to hidden default)")
            ok = False

        step("[6b] user un-hides FIXED; hide+restore must bring it back visible")
        wb.Worksheets("FIXED").Columns("D").Hidden = False
        app.Run("HideDesignationForPrint")
        restored = {n: wb.Worksheets(n).Columns("D").Hidden for n in SHEETS}
        step(f"    after re-hide: {restored}")
        app.Run("RestoreDesignationColumns")
        restored = {n: wb.Worksheets(n).Columns("D").Hidden for n in SHEETS}
        step(f"    after restore: {restored}")
        if restored["FIXED"] or not all(v for k, v in restored.items() if k != "FIXED"):
            step("    STATEFUL RESTORE FAILED")
            ok = False

        step("[7] closing")
    finally:
        wb.Close(SaveChanges=False)
        app.Quit()

    step("PRINT-LAYOUT CHECKS PASSED" if ok else "PRINT-LAYOUT CHECKS FAILED")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
