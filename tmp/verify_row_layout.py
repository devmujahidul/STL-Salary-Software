"""Verify the new 3-row-per-worker layout against the raw jobcard files:

Row 1: P/A/W codes only in date columns (no OT numbers).
Row 2: daily OT hours, sum == monthly OT in the summary.
Row 3: daily EOT/totals, sum == monthly EOT in the summary.

Streams the sheet with iter_rows once (fast) and then indexes into memory.
"""
from __future__ import annotations

import sys
from pathlib import Path

from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from attendance_export.constants import DATA_START_ROW, IDENTITY_HEADERS, SUMMARY_KEYS


def check(workbook_path: str) -> bool:
    wb = load_workbook(workbook_path, data_only=True, read_only=True)
    problems: list[str] = []
    total_workers = 0
    total_ot_checked = 0
    total_eot_checked = 0
    date_count = 31

    for sheet_name in wb.sheetnames:
        if sheet_name in {"Lists", "Hidden"}:
            continue
        ws = wb[sheet_name]
        # stream all rows once; each row is a tuple of values
        grid = [list(r) for r in ws.iter_rows(values_only=True)]

        date_start = len(IDENTITY_HEADERS)
        summary_start = date_start + date_count
        ot_col = summary_start + SUMMARY_KEYS.index("OT")
        eot_col = summary_start + SUMMARY_KEYS.index("EOT")

        row = DATA_START_ROW + 1  # 1-indexed Excel rows
        sample_shown = False
        while row <= len(grid):
            name = grid[row - 1][1]
            if name in (None, ""):
                break
            total_workers += 1
            main = grid[row - 1]

            code_cells = main[date_start:summary_start]
            # 1) Main row: must NOT contain numeric OT values where a code exists
            bad_main = [v for v in code_cells if isinstance(v, (int, float))]
            if bad_main:
                problems.append(
                    f"{sheet_name} row {row} ({name}): numeric values in main date row: {bad_main[:5]}"
                )
            codes = {str(v).strip().upper() for v in code_cells if v not in (None, "")}
            unknown = codes - {"P", "W", "A", "CL", "SL", "EL", "ML", ""}
            if unknown:
                problems.append(f"{sheet_name} row {row} ({name}): odd codes {unknown}")

            # 2) OT row: sum of date cells == monthly OT
            ot_grid = grid[row]  # 0-indexed: row+1 in Excel
            ot_sum = sum(float(v or 0) for v in ot_grid[date_start:summary_start])
            monthly_ot = float(main[ot_col] or 0)
            if monthly_ot > 0:
                total_ot_checked += 1
                if abs(ot_sum - monthly_ot) > 0.01:
                    problems.append(
                        f"{sheet_name} row {row + 1} ({name}): daily OT sum {ot_sum} != monthly OT {monthly_ot}"
                    )

            # 3) EOT row: sum of date cells == monthly EOT
            eot_grid = grid[row + 1]
            eot_sum = sum(float(v or 0) for v in eot_grid[date_start:summary_start])
            monthly_eot = float(main[eot_col] or 0)
            if monthly_eot > 0:
                total_eot_checked += 1
                if abs(eot_sum - monthly_eot) > 0.01:
                    problems.append(
                        f"{sheet_name} row {row + 2} ({name}): daily totals sum {eot_sum} != monthly EOT {monthly_eot}"
                    )

            # all three rows must carry the subsection filter value
            filter_col = summary_start + 10
            for grid_row in (main, ot_grid, eot_grid):
                if grid_row[filter_col] in (None, ""):
                    problems.append(f"{sheet_name}: missing SubSection Filter value")

            if not sample_shown:
                print(f"--- {sheet_name}: first worker sample (row {row}) ---")
                print("codes  :", code_cells[:10])
                print("OT row :", ot_grid[date_start:date_start + 10])
                print("EOT row:", eot_grid[date_start:date_start + 10])
                sample_shown = True

            row += 3  # next worker block

    wb.close()
    print(f"\nWorkers checked: {total_workers}")
    print(f"Workers with monthly OT>0 checked: {total_ot_checked}")
    print(f"Workers with monthly EOT>0 checked: {total_eot_checked}")
    if problems:
        print(f"\nPROBLEMS ({len(problems)}):")
        for p in problems[:30]:
            print(" -", p)
        return False
    print("\nALL CHECKS PASSED")
    return True


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else str(
        Path(__file__).resolve().parent.parent / "output" / "Salary_Attendance_August_2026.xlsx"
    )
    ok = check(path)
    sys.exit(0 if ok else 1)
