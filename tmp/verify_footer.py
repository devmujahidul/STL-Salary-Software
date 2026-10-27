"""Verify the dynamic signature footer and the salary-month label in Excel:

1. Each printed page footer shows the names stored on the Lists sheet.
2. Editing a name on Lists + running ApplyFooterSignatures refreshes footers.
3. Row 3 label is 'ATTENDANCE FOR <prev month> - <mentioned month> yyyy'.
"""
from __future__ import annotations

import sys
from pathlib import Path

import win32com.client as win32

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from attendance_export.constants import FOOTER_SIGNATURES, LIST_SHEET

SHEETS = ["LO_PO", "FIXED", "PRODUCTION"]


def main() -> bool:
    path = Path(__file__).resolve().parent.parent / "output" / "Salary_Attendance_August_2026.xlsm"
    if not path.exists():
        path = path.with_suffix(".xlsx")
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    ok = True
    try:
        wb = app.Workbooks.Open(str(path), ReadOnly=False)

        # 1) Lists sheet stores the names
        lst = wb.Worksheets(LIST_SHEET)
        stored = []
        r = 2
        while True:
            v = lst.Cells(r, 1).Value
            if v in (None, ""):
                break
            stored.append(str(v))
            r += 1
        print(f"Lists sheet names: {stored}")
        if stored != list(FOOTER_SIGNATURES):
            print("LISTS CONTENT MISMATCH")
            ok = False

        # 2) footers of every attendance sheet carry those names
        for name in SHEETS:
            ws = wb.Worksheets(name)
            lf, cf, rf = ws.PageSetup.LeftFooter, ws.PageSetup.CenterFooter, ws.PageSetup.RightFooter
            print(f"{name} footer: L={lf!r} C={cf!r} R={rf!r}")
            for zone, want in zip((lf, cf, rf), FOOTER_SIGNATURES):
                if str(want) not in str(zone):
                    print(f"  MISSING {want!r} in zone {zone!r}")
                    ok = False

        # 3) month label on row 3
        ws = wb.Worksheets("FIXED")
        label = str(ws.Range("A3").Value)
        print(f"Period label: {label!r}")
        if label != "ATTENDANCE FOR July - August 2026":
            print("LABEL MISMATCH")
            ok = False

        # 4) dynamic refresh: change a Lists name, run the macro, footer follows
        lst.Cells(2, 1).Value = "Prepared by (edited)"
        app.Run("ApplyFooterSignatures")
        lf = wb.Worksheets("FIXED").PageSetup.LeftFooter
        print(f"after edit + macro: L={lf!r}")
        if "Prepared by (edited)" not in lf:
            print("DYNAMIC FOOTER REFRESH FAILED")
            ok = False
        else:
            print("DYNAMIC FOOTER REFRESH OK")
    finally:
        wb.Close(SaveChanges=False)
        app.Quit()

    print("FOOTER/LABEL CHECKS PASSED" if ok else "FOOTER/LABEL CHECKS FAILED")
    return ok


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
