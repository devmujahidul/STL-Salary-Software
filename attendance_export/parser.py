from __future__ import annotations

import calendar
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from bs4 import BeautifulSoup

from .constants import (
    ADDRESS_FALLBACK,
    CLEANER_PREFIX,
    CLEANER_SUBSECTION,
    COMPANY_FALLBACK,
)

MONTH_OF_RE = re.compile(r"Month of\s*[-:]\s*([A-Za-z]+)\s*[-/]\s*(\d{4})", re.I)
DATE_HEADER_RE = re.compile(r"^(\d{1,2})(?:\s+([A-Za-z]{3}))?$")
ID_RE = re.compile(r"^[A-Za-z]{1,4}-\d+")


@dataclass
class Worker:
    name: str
    emp_id: str
    designation: str
    subsection: str
    attendance: list[str]
    summary: dict[str, float | int | str]
    daily_ot: list[float] = field(default_factory=list)
    daily_eot: list[float] = field(default_factory=list)


@dataclass
class AttendanceData:
    company: str
    address: str
    period_label: str
    start_date: date
    end_date: date
    date_headers: list[tuple[int, str]]  # (day_number, weekday)
    workers: list[Worker] = field(default_factory=list)


def id_prefix(emp_id: str) -> str:
    return emp_id.split("-", 1)[0].upper() if emp_id else ""


def subsection_for_worker(emp_id: str, raw_subsection: str) -> str:
    if id_prefix(emp_id) == CLEANER_PREFIX:
        return CLEANER_SUBSECTION
    return raw_subsection


def parse_raw_files(raw_dir: str | Path) -> AttendanceData:
    raw_dir = Path(raw_dir)
    files = sorted(
        p
        for p in raw_dir.iterdir()
        if p.is_file() and p.suffix.lower() in {".xls", ".xlsx", ".html", ".htm"}
    )
    if not files:
        raise FileNotFoundError(f"No raw files found in {raw_dir}")

    company = COMPANY_FALLBACK
    address = ADDRESS_FALLBACK
    period_label = ""
    start_date = end_date = None
    date_headers: list[tuple[int, str]] = []
    workers: list[Worker] = []
    seen_ids: set[str] = set()

    for path in files:
        text = path.read_text(encoding="utf-8", errors="replace")
        if "<html" not in text[:400].lower() and "<table" not in text[:800].lower():
            raise ValueError(f"{path.name} is not an HTML jobcard export")
        soup = BeautifulSoup(text, "lxml")
        tables = soup.find_all("table")
        if not tables:
            continue

        header_meta = _parse_company_table(tables[0])
        if header_meta.get("company"):
            company = header_meta["company"]
        if header_meta.get("address"):
            address = header_meta["address"]
        if header_meta.get("period_line"):
            period_label = header_meta["period_line"]
            start_date, end_date = period_from_month_line(header_meta["period_line"])

        for table in tables[1:]:
            table_workers, headers = _parse_section_table(table)
            if headers and not date_headers:
                date_headers = headers
            for worker in table_workers:
                if worker.emp_id in seen_ids:
                    continue
                seen_ids.add(worker.emp_id)
                workers.append(worker)

    if not workers:
        raise ValueError("No workers found in raw files")
    if not date_headers:
        raise ValueError("Could not read date columns from raw files")
    if start_date is None or end_date is None:
        start_date, end_date = _period_from_date_headers(date_headers)

    workers.sort(key=lambda w: (w.subsection.lower(), w.emp_id))
    return AttendanceData(
        company=company,
        address=address,
        period_label=_attendance_title(start_date, end_date),
        start_date=start_date,
        end_date=end_date,
        date_headers=date_headers,
        workers=workers,
    )


def period_from_month_line(line: str) -> tuple[date, date]:
    match = MONTH_OF_RE.search(line)
    if not match:
        raise ValueError(f"Cannot parse salary month from: {line}")
    month_name, year_s = match.group(1), match.group(2)
    month_num = _month_number(month_name)
    year = int(year_s)
    end = date(year, month_num, 25)
    if month_num == 1:
        start = date(year - 1, 12, 26)
    else:
        start = date(year, month_num - 1, 26)
    return start, end


def _attendance_title(start: date, end: date) -> str:
    # ATTENDANCE FOR <previous month> - <mentioned month> <yyyy>
    if start.year == end.year:
        return f"ATTENDANCE FOR {start:%B} - {end:%B %Y}"
    # year boundary (e.g. December 2026 - January 2027): show both years
    return f"ATTENDANCE FOR {start:%B %Y} - {end:%B %Y}"


def _month_number(name: str) -> int:
    name = name.strip().title()
    for i in range(1, 13):
        if calendar.month_name[i] == name or calendar.month_abbr[i] == name:
            return i
    raise ValueError(f"Unknown month name: {name}")


def _period_from_date_headers(headers: list[tuple[int, str]]) -> tuple[date, date]:
    days = [d for d, _ in headers]
    if not days:
        raise ValueError("Empty date headers")
    today = date.today()
    start_day = days[0]
    end_day = days[-1]
    if start_day >= 26:
        start_month = today.month - 1 if today.day < 26 else today.month
        year = today.year
        if start_month < 1:
            start_month = 12
            year -= 1
        start = date(year, start_month, start_day)
        if start_month == 12:
            end = date(year + 1, 1, end_day)
        else:
            end = date(year, start_month + 1, end_day)
        return start, end
    raise ValueError("Date headers do not start on the 26th")


def _parse_company_table(table) -> dict[str, str]:
    rows = table.find_all("tr")
    texts = [r.get_text(" ", strip=True) for r in rows if r.get_text(strip=True)]
    out: dict[str, str] = {}
    if texts:
        out["company"] = texts[0]
    if len(texts) > 1:
        out["address"] = texts[1]
    if len(texts) > 2:
        out["period_line"] = texts[2]
    return out


def _parse_section_table(table) -> tuple[list[Worker], list[tuple[int, str]]]:
    rows = table.find_all("tr")
    subsection = ""
    date_headers: list[tuple[int, str]] = []
    workers: list[Worker] = []
    i = 0
    while i < len(rows):
        cells = _row_cells(rows[i])
        if not cells:
            i += 1
            continue
        first = cells[0]
        if first.startswith("Sub Section"):
            subsection = first.split(":", 1)[-1].strip()
            i += 1
            continue
        if first == "SL":
            date_headers = _extract_date_headers(cells)
            i += 1
            continue
        if _is_worker_row(cells):
            # The rows immediately after a worker row are the per-day value
            # breakdown: daily OT hours (sum == monthly OT) followed by the
            # daily totals row (sum == monthly EOT).
            daily_ot_cells, daily_eot_cells = _collect_daily_rows(
                rows, i + 1, len(date_headers)
            )
            worker = _worker_from_row(
                cells,
                subsection,
                daily_ot_cells,
                len(date_headers),
                daily_eot_cells,
            )
            workers.append(worker)
            i += 1
            if daily_ot_cells is not None:
                i += 1  # skip the consumed OT row
            if daily_eot_cells is not None:
                i += 1  # skip the consumed EOT row
            continue
        i += 1
    return workers, date_headers


def _row_cells(row) -> list[str]:
    return [c.get_text(" ", strip=True) for c in row.find_all(["td", "th"])]


def _is_worker_row(cells: list[str]) -> bool:
    if len(cells) < 10:
        return False
    if not cells[0].isdigit():
        return False
    return bool(ID_RE.match(cells[2].replace(" ", "")))


def _extract_date_headers(header_cells: list[str]) -> list[tuple[int, str]]:
    dates: list[tuple[int, str]] = []
    for cell in header_cells:
        match = DATE_HEADER_RE.match(cell.strip())
        if not match:
            continue
        day = int(match.group(1))
        weekday = (match.group(2) or "").title()
        dates.append((day, weekday))
    return dates


def _collect_daily_rows(
    rows,
    start: int,
    date_count: int,
) -> tuple[list[str] | None, list[str] | None]:
    """Collect the numeric value rows that follow a worker row.

    The raw jobcard has two numeric rows under each worker row: the first is
    the daily OT hours (its sum equals the monthly OT) and the second is the
    daily totals row (its sum equals the monthly EOT). Return them as a pair
    (ot_cells, eot_cells); whichever rows are missing give None.
    """
    if not date_count:
        return None, None
    collected: list[list[str] | None] = []
    j = start
    while j < len(rows) and len(collected) < 2:
        nxt = _row_cells(rows[j])
        if len(nxt) != date_count or not _is_all_numeric(nxt):
            break
        if not any(c.strip() for c in nxt):
            break  # skip blank rows instead of treating them as values
        collected.append(nxt)
        j += 1
    while len(collected) < 2:
        collected.append(None)
    ot_cells, eot_cells = collected
    return ot_cells, eot_cells


def _is_all_numeric(cells: list[str]) -> bool:
    """True when every cell is a number (or empty), used to detect the per-day
    OT breakdown row that immediately follows each worker row."""
    if not cells:
        return False
    for c in cells:
        t = c.strip()
        if t == "":
            continue
        t2 = t.replace(".", "", 1)
        if not t2.isdigit():
            return False
    return True


def _worker_from_row(
    cells: list[str],
    subsection: str,
    daily_ot_cells: list[str] | None = None,
    date_count: int | None = None,
    daily_eot_cells: list[str] | None = None,
) -> Worker:
    emp_id = cells[2].replace(" ", "")
    name = cells[1].strip()
    designation = cells[3].strip() if len(cells) > 3 else ""
    attendance: list[str] = []
    # date values sit between designation and summary columns P,A,...
    summary_start = _find_summary_start(cells)
    raw_days = cells[4:summary_start]
    attendance = raw_days

    # Per-day value breakdown rows (31 numeric cells each under the worker
    # row): daily OT hours (sum == monthly OT total) and daily totals
    # (sum == monthly EOT total).
    target = date_count if date_count is not None else len(attendance)

    def _daily_from(cells: list[str] | None) -> list[float]:
        values: list[float] = []
        if cells:
            for x in cells:
                try:
                    values.append(float(x))
                except ValueError:
                    values.append(0.0)
            if len(values) < target:
                values.extend([0.0] * (target - len(values)))
            else:
                values = values[:target]
        return values

    daily_ot = _daily_from(daily_ot_cells)
    daily_eot = _daily_from(daily_eot_cells)

    summary_vals = cells[summary_start : summary_start + 9]
    keys = ["P", "A", "OT", "EOT", "TOT", "CL", "SL", "EL", "H"]
    summary: dict[str, float | int | str] = {}
    for key, val in zip(keys, summary_vals):
        summary[key] = _as_number(val)
    return Worker(
        name=name,
        emp_id=emp_id,
        designation=designation,
        subsection=subsection_for_worker(emp_id, subsection),
        attendance=attendance,
        summary=summary,
        daily_ot=daily_ot,
        daily_eot=daily_eot,
    )


def _find_summary_start(cells: list[str]) -> int:
    # Last 9 value columns before any trailing empties are P,A,OT,EOT,TOT,CL,SL,EL,H
    # Worker rows in raw files are 44 cells: 4 identity + N dates + 9 summary
    if len(cells) >= 13:
        return len(cells) - 9
    return 4


def _as_number(val: str) -> float | int | str:
    if val is None:
        return ""
    text = str(val).strip()
    if text == "":
        return 0
    try:
        number = float(text)
    except ValueError:
        return text
    if number.is_integer():
        return int(number)
    return number
