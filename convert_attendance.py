from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path

from attendance_export.classify import group_workers, sheet_for_worker
from attendance_export.parser import parse_raw_files
from attendance_export.vba import attach_filter_macros
from attendance_export.writer import write_workbook


def main() -> None:
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="Convert Matrix jobcard raw files into the Spring Trade attendance workbook."
    )
    parser.add_argument("--raw-dir", default=str(root / "raw_files"))
    parser.add_argument("--output", default="")
    parser.add_argument("--skip-macros", action="store_true")
    args = parser.parse_args()

    print(f"Reading raw files from {args.raw_dir}")
    data = parse_raw_files(args.raw_dir)
    grouped = group_workers(data.workers)

    month_tag = data.end_date.strftime("%B_%Y")
    output = Path(args.output) if args.output else root / "output" / f"Salary_Attendance_{month_tag}.xlsx"
    print(f"Company: {data.company}")
    print(f"Period:  {data.period_label} ({data.start_date.isoformat()} to {data.end_date.isoformat()})")
    print(f"Days:    {len(data.date_headers)} -> {[d for d, _ in data.date_headers]}")
    print(f"Workers: {len(data.workers)}")
    for sheet in grouped:
        print(f"  {sheet}: {len(grouped[sheet])}")
    prefixes = Counter(sheet_for_worker(w) for w in data.workers)
    print(f"Sheet split: {dict(prefixes)}")

    path = write_workbook(data, output)
    print(f"Wrote {path}")

    if not args.skip_macros:
        ok, reason = attach_filter_macros(path)
        if ok:
            print(f"Added filter macros: {path.with_suffix('.xlsm')}")
        else:
            print("Macros were not added:")
            print(f"  {reason}")
            print(
                "The .xlsx still filters from the SubSection Filter column; "
                "A4 shows the current selection."
            )


if __name__ == "__main__":
    main()
