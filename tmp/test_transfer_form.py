"""End-to-end test of the duty transfer feature (form + engine + undo).

Runs the generated .xlsm headlessly through COM and checks:
  1. The UserForm and its code behind exist and initialise (combos filled).
  2. A weekday transfer moves P-day + OT/EOT, source becomes A/0/0, monthly
     P/A summary formulas shift by +1/-1, and other workers are untouched.
  3. The count in the result message matches the expected scan counts.
  4. A weekend transfer applies OT = full-hours - 8 (max 2, rest to EOT),
     zeroes the source W day, keeps the W stamp.
  5. UndoLastTransfer restores every value of the last batch exactly.
"""
from __future__ import annotations

import re
from pathlib import Path

XLSM = Path("output/Salary_Attendance_August_2026.xlsm").resolve()
ok = True


def check(label: str, condition: bool, detail: str = "") -> None:
    global ok
    status = "OK " if condition else "FAIL"
    print(f"[{status}] {label}{(' - ' + detail) if detail else ''}")
    if not condition:
        ok = False


def main() -> None:
    import pythoncom
    import win32com.client as win32

    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = 1  # enable macros

    wb = excel.Workbooks.Open(str(XLSM))
    try:
        # ---------- 1. form + module exist ----------
        names = [c.Name for c in wb.VBProject.VBComponents]
        check("DutyTransferHelpers module present", "DutyTransferHelpers" in names)
        check("frmDutyTransfer present", "frmDutyTransfer" in names)
        form_code = wb.VBProject.VBComponents("frmDutyTransfer").CodeModule.Lines(1, 200)
        check("form code behind has button handlers",
              "btnGo_Click" in form_code and "btnUndo_Click" in form_code)

        # ---------- pick sheets / scan layout ----------
        ws = wb.Worksheets("FIXED")
        # date columns: row 5 headers from col 5, contiguous numeric days
        col = 5
        headers: dict[int, str] = {}
        while True:
            v = ws.Cells(5, col).Value
            if v is None or not str(v).strip().isdigit():
                break
            headers[col] = str(v).strip()
            col += 1
        ndates = len(headers)
        # continue rightwards to the first blank header; the last non-blank
        # header cell is the SubSection Filter column
        summary_a = None
        while True:
            v = str(ws.Cells(5, col).Value or "").strip()
            if v == "":
                break
            if v == "P" and summary_a is None:
                summary_p = col
                summary_a = col + 1
            col += 1
        filter_col = col - 1
        print(f"layout: {ndates} date columns, filter col {filter_col}, P col {summary_p}")

        # scan worker blocks: r = 6, 9, ...
        block_rows = []
        r = 6
        last_row = ws.Cells(ws.Rows.Count, filter_col).End(-4162).Row  # xlUp
        while r <= last_row:
            block_rows.append(r)
            r += 3

        def code_at(r: int, c: int) -> str:
            return str(ws.Cells(r, c).Value or "").strip().upper()

        def num_at(r: int, c: int) -> float:
            val = ws.Cells(r, c).Value
            try:
                return float(val)
            except (TypeError, ValueError):
                return 0.0

        # find weekday & weekend test cases within one sub section
        weekday = None
        weekend = None
        for r in block_rows:
            sec = str(ws.Cells(r, filter_col).Value or "").strip()
            for cf in headers:
                if code_at(r, cf) == "P":
                    for ct in headers:
                        if code_at(r, ct) == "A":
                            weekday = dict(sec=sec, r=r, from_col=cf, to_col=ct)
                            break
                    if weekday:
                        break
            if weekday:
                break
        for r in block_rows:
            sec = str(ws.Cells(r, filter_col).Value or "").strip()
            for cf in headers:
                if code_at(r, cf) == "W" and num_at(r + 2, cf) > 0:
                    for ct in headers:
                        if code_at(r, ct) == "A" and ct != cf:
                            weekend = dict(sec=sec, r=r, from_col=cf, to_col=ct)
                            break
                    if weekend:
                        break
            if weekend:
                break
        check("found a weekday test case", weekday is not None)
        check("found a weekend test case", weekend is not None)
        if not (weekday and weekend):
            return

        all_sheets = [wb.Worksheets(n) for n in ("LO_PO", "FIXED", "PRODUCTION")]

        def expected_counts(sec: str, fd_day: str, td_day: str) -> tuple[int, int]:
            """The engine runs the subsection across every attendance sheet."""
            moved = skipped = 0
            for sheet in all_sheets:
                # per-sheet column map (identical layout, still resolve safely)
                day_col = {str(sheet.Cells(5, c).Value).strip(): c
                           for c in headers}
                s_fcol = sheet.Cells(5, sheet.Columns.Count).End(-4159).Column  # xlToLeft
                for rr in range(6, last_row + 1, 3):
                    block_sec = str(sheet.Cells(rr, s_fcol).Value or "").strip()
                    if block_sec.lower() != sec.lower():
                        continue
                    src = str(sheet.Cells(rr, day_col[fd_day]).Value or "").strip().upper()
                    ot2 = sheet.Cells(rr + 2, day_col[fd_day]).Value or 0
                    present = src == "P" or (src == "W" and float(ot2) > 0)
                    if not present:
                        continue
                    tgt = str(sheet.Cells(rr, day_col[td_day]).Value or "").strip().upper()
                    if tgt == "A":
                        moved += 1
                    else:
                        skipped += 1
            return moved, skipped

        def run_case(case: dict, weekend_case: bool) -> None:
            nonlocal_ok = True
            r, cf, ct = case["r"], case["from_col"], case["to_col"]
            sec, fd, td = case["sec"], headers[cf], headers[ct]

            base = [ws.Cells(rr, cc).Value for rr, cc in
                    ((r, cf), (r + 1, cf), (r + 2, cf), (r, ct), (r + 1, ct), (r + 2, ct))]
            p_before = ws.Cells(r, summary_p).Value
            a_before = ws.Cells(r, summary_a).Value
            # an untouched neighbouring cell in another block
            nbr_r = 6 if r != 6 else 9
            nbr_cell = (nbr_r, cf)
            nbr_before = ws.Cells(*nbr_cell).Value

            exp_moved, exp_skipped = expected_counts(sec, fd, td)
            label = "weekend" if weekend_case else "weekday"
            res = excel.Run("ExecuteDutyTransfer", sec, fd, td)
            print(f"    engine said: {res!r}")

            m = re.search(r"(\d+) worker\(s\) moved.*?(\d+) worker\(s\) skipped", res, re.S)
            got = (int(m.group(1)), int(m.group(2))) if m else None
            check(f"[{label}] result counts match scan",
                  bool(m) and got == (exp_moved, exp_skipped),
                  f"engine=" + (str(got) if got else "unparsed") + f" expected=({exp_moved},{exp_skipped})")

            if weekend_case:
                full = float(base[2])
                rest = max(full - 8, 0)
                exp_ot = min(rest, 2)
                exp_eot = max(rest - 2, 0) if full > 8 else max(full - 2, 0)
                check(f"[{label}] target code P", code_at(r, ct) == "P")
                check(f"[{label}] target OT math", abs(num_at(r + 1, ct) - exp_ot) < 1e-6,
                      f"got {num_at(r + 1, ct)} want {exp_ot}")
                check(f"[{label}] target EOT math", abs(num_at(r + 2, ct) - exp_eot) < 1e-6,
                      f"got {num_at(r + 2, ct)} want {exp_eot}")
                check(f"[{label}] source keeps W", code_at(r, cf) == "W")
                check(f"[{label}] source zeroed",
                      abs(num_at(r + 1, cf)) < 1e-6 and abs(num_at(r + 2, cf)) < 1e-6)
            else:
                check(f"[{label}] target got duty P/OT/EOT",
                      code_at(r, ct) == "P"
                      and abs(num_at(r + 1, ct) - float(base[1])) < 1e-6
                      and abs(num_at(r + 2, ct) - float(base[2])) < 1e-6)
                check(f"[{label}] source becomes A/0/0",
                      code_at(r, cf) == "A"
                      and abs(num_at(r + 1, cf)) < 1e-6
                      and abs(num_at(r + 2, cf)) < 1e-6)

            excel.Calculate()
            p_after = ws.Cells(r, summary_p).Value
            a_after = ws.Cells(r, summary_a).Value
            if weekend_case:
                # W becomes a worked weekday: P +1, A -1 net
                p_ok = abs(float(p_after) - float(p_before) - 1) < 1e-6
                a_ok = abs(float(a_after) - float(a_before) + 1) < 1e-6
            else:
                # weekday-to-weekday: one P and one A swap places - net 0
                p_ok = abs(float(p_after) - float(p_before)) < 1e-6
                a_ok = abs(float(a_after) - float(a_before)) < 1e-6
            check(f"[{label}] monthly P/A totals correct",
                  p_ok and a_ok, f"P {p_before}->{p_after}, A {a_before}->{a_after}")
            check(f"[{label}] other workers untouched",
                  ws.Cells(*nbr_cell).Value == nbr_before)

            # undo must restore everything exactly
            res_u = excel.Run("UndoLastTransfer")
            print(f"    undo said: {res_u!r}")
            after = [ws.Cells(rr, cc).Value for rr, cc in
                     ((r, cf), (r + 1, cf), (r + 2, cf), (r, ct), (r + 1, ct), (r + 2, ct))]
            check(f"[{label}] undo restored all six cells",
                  all((str(a).strip().upper() == str(b).strip().upper()
                       if isinstance(a, str) else abs((a or 0) - (b or 0)) < 1e-6)
                      for a, b in zip(after, base)))
            has_undo = excel.Run("HasUndo")
            check(f"[{label}] HasUndo false after undo", not bool(has_undo))

        run_case(weekday, weekend_case=False)
        run_case(weekend, weekend_case=True)

        # ---------- form initialise (combos filled) via temp module ----------
        comp = wb.VBProject.VBComponents.Add(1)  # std module
        comp.Name = "TransferTestProbe"
        comp.CodeModule.AddFromString(
            "Public Function TestFormCheck() As String\n"
            "    Dim f As Object\n"
            "    Set f = New frmDutyTransfer\n"
            "    TestFormCheck = \"sections=\" & f.Controls(\"cboSection\").ListCount & _\n"
            "        \",dates=\" & f.Controls(\"cboFrom\").ListCount & _\n"
            "        \",undo=\" & f.Controls(\"btnUndo\").Enabled\n"
            "    Unload f\n"
            "End Function\n"
        )
        probe = str(excel.Run("TestFormCheck"))
        m = re.match(r"sections=(\d+),dates=(\d+),undo=(\w+)", probe)
        check("form initialised with combo data", bool(m), probe)
        if m:
            check(f"section combo filled ({m.group(1)} entries)", int(m.group(1)) >= 2)
            check(f"date combos filled ({m.group(2)} days)", int(m.group(2)) == ndates)
            check(f"undo button disabled on fresh file ({m.group(3)})", m.group(3) == "False")
        wb.VBProject.VBComponents.Remove(comp)

        # save untouched: confirm workbook has no persistent history sheet
        check("no _TransferLog in saved layout", "_TransferLog" not in wb.Worksheets or
              wb.Worksheets("_TransferLog").Visible == 2)
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()

    print("ALL TRANSFER CHECKS PASSED" if ok else "TRANSFER CHECKS FAILED")
    return ok


if __name__ == "__main__":
    main()
