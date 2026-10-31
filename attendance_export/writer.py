from __future__ import annotations

import struct
from pathlib import Path

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from .classify import group_workers
from .constants import (
    ALL_LABEL,
    DATA_START_ROW,
    FOOTER_LINE,
    FOOTER_MARGIN,
    FOOTER_SIGNATURES,
    HEADER_ROWS,
    IDENTITY_HEADERS,
    LIST_SHEET,
    PRINT_TITLE_LAST_ROW,
    SHEET_ORDER,
    SHEET_PROTECT_PASSWORD,
    SUBSECTION_PREFIX,
    SUMMARY_HEADERS,
    SUMMARY_KEYS,
    LOGO_HEIGHT_PX,
    LOGO_PATH,
    LOGO_RIGHT_PAD_PX,
    LOGO_SECTION_ROW_HEIGHT_PT,
    LOGO_WIDTH_PX,
)
from .parser import AttendanceData, Worker

HELPER_NAME_COL = 46  # AU, hidden
HELPER_FLAG_COL = 47  # AV, hidden
LIST_COL = 48  # AW, hidden
LABEL_TEXT_COL = 49  # AX, hidden

# Widths (Excel character units) of the summary columns right of the date
# columns, in order: P A OT EOT TOT CL SL EL H Signature. _set_widths()
# applies them and the company-logo anchor walks them Signature-first to keep
# the logo inside the print area.
SUMMARY_COL_WIDTHS = (2.9, 2.9, 3.1, 4.6, 4.6, 2.9, 2.9, 2.9, 2.9, 8.0)


def write_workbook(data: AttendanceData, output_path: str | Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    grouped = group_workers(data.workers)
    workbook = xlsxwriter.Workbook(
        str(output_path), {"use_future_functions": True}
    )
    workbook.set_calc_mode("auto")
    formats = _formats(workbook)

    sheets = {name: workbook.add_worksheet(name) for name in SHEET_ORDER}

    for sheet_name in SHEET_ORDER:
        workers = grouped.get(sheet_name, [])
        _write_sheet(
            sheets[sheet_name],
            workbook,
            formats,
            sheet_name,
            data,
            workers,
        )

    # Signature options live on the Lists sheet; the printed footer is built
    # from them. In the macro (.xlsm) workbook the ApplyFooterSignatures macro
    # re-reads the same sheet, so editing the names there and reopening (or
    # running the macro) refreshes the footers without touching the code.
    _write_lists_sheet(workbook, formats, list(FOOTER_SIGNATURES))
    footer_left, footer_center, footer_right = _footer_zones(FOOTER_SIGNATURES)
    footer_options = {"margin": FOOTER_MARGIN}
    for ws_obj in sheets.values():
        ws_obj.set_footer(
            f"&L{footer_left}&C{footer_center}&R{footer_right}", footer_options
        )

    workbook.close()
    return output_path


def _footer_zones(names: list[str]) -> tuple[str, str, str]:
    """Map signature names onto the left / center / right footer zones.

    The first three names take the three zones; any extras are appended to
    the right zone. A name prints as 'Name: ____signature line____' at 9pt.
    """
    padded = list(names[:3]) + [""] * (3 - len(names[:3]))

    def part(name: str) -> str:
        return f"&9{name}: {FOOTER_LINE}" if name else ""

    left = part(padded[0])
    center = part(padded[1])
    right = part(padded[2])
    for extra in names[3:]:
        if extra and right:
            right += f"        {part(extra)}"
        elif extra:
            right = part(extra)
    return left, center, right


def _write_lists_sheet(
    workbook: xlsxwriter.Workbook,
    formats: dict,
    signatures: list[str],
):
    """Store the footer signature options on the Lists sheet.

    Edit the names in column A, save, and reopen the file (or run the
    ApplyFooterSignatures macro) to refresh every printed page footer.
    """
    ws = workbook.add_worksheet(LIST_SHEET)
    ws.write(0, 0, "Footer Signatures", formats["list_header"])
    for index, name in enumerate(signatures, start=1):
        ws.write(index, 0, name)
    ws.write(
        0,
        2,
        "These names fill the signature footer at the bottom of every printed "
        "page. Edit them, save, then reopen the file (or run the "
        "ApplyFooterSignatures macro) to refresh the footers.",
        formats["cell_left"],
    )
    ws.set_column(0, 0, 24)
    ws.set_column(1, 1, 3)
    ws.set_column(2, 2, 90)
    workbook.define_name(
        "FooterSignatures",
        f"={LIST_SHEET}!$A$2:$A${1 + max(len(signatures), 3)}",
    )
    return ws


def _formats(workbook: xlsxwriter.Workbook) -> dict:
    border = {"border": 1, "border_color": "#8EA9DB"}
    thin = {"border": 1, "border_color": "#BFBFBF"}
    return {
        "company": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 18,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "font_color": "#1F1F1F",
            }
        ),
        "address": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "font_color": "#1F1F1F",
            }
        ),
        "period": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12.5,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "font_color": "#1F1F1F",
            }
        ),
        "subsection": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12.5,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "font_color": "#1F3864",
                # Keep the subsection dropdown writable on protected sheets.
                "locked": False,
            }
        ),
        "header": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10.5,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "bg_color": "#D9E2F3",
                "font_color": "#1F1F1F",
                **border,
            }
        ),
        "header_weekend": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 10.5,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
                "bg_color": "#F8CBAD",
                "font_color": "#1F1F1F",
                **border,
            }
        ),
        "cell": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                **thin,
            }
        ),
        "cell_left": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "left",
                "valign": "vcenter",
                "shrink": True,
                **thin,
            }
        ),
        "cell_weekend": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#FCE4D6",
                **thin,
            }
        ),
        "cell_absent": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "font_color": "#C00000",
                **thin,
            }
        ),
        "number": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                **thin,
            }
        ),
        "signature": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "center",
                "valign": "vcenter",
                **thin,
            }
        ),
        "filter_col": workbook.add_format(
            {
                "font_name": "Calibri",
                "font_size": 11,
                "align": "left",
                "valign": "vcenter",
                "bg_color": "#F2F2F2",
                **thin,
            }
        ),
        "list_header": workbook.add_format(
            {"font_name": "Calibri", "font_size": 10.5, "bold": True}
        ),
    }


def _row_pixels(height_pt: float) -> int:
    """Row points -> pixels, matching xlsxwriter's own _size_row conversion."""
    return int(4.0 / 3.0 * height_pt)


def _col_pixels(width: float) -> int:
    """Excel character units -> pixels (Calibri 11), matching xlsxwriter's
    own _size_col conversion: pixels = int(width * 7 + 0.5) + 5."""
    return int(width * 7 + 0.5) + 5


def _logo_anchor(signature_col: int, logo_w_px: int) -> tuple[int, int]:
    """Return (anchor_col, x_offset) so the logo's right edge stays inside the
    print area, LOGO_RIGHT_PAD_PX before the Signature column's right edge.

    A large logo is wider than the Signature column alone, so the anchor walks
    left over the summary columns (SUMMARY_COL_WIDTHS, Signature first) until
    the span is wide enough. The resulting offset is always smaller than the
    anchor column's own width, so xlsxwriter keeps the anchor cell as chosen.
    """
    need = logo_w_px + LOGO_RIGHT_PAD_PX
    span_px = 0
    steps_left = 0
    while True:
        idx = len(SUMMARY_COL_WIDTHS) - 1 - steps_left
        if idx < 0:
            break  # wider than all summary columns; anchor at the first one
        span_px += _col_pixels(SUMMARY_COL_WIDTHS[idx])
        if span_px >= need:
            break
        steps_left += 1
    return signature_col - steps_left, max(span_px - need, 0)


def _logo_scale_factors(logo_path: Path, target_w_px: int, target_h_px: int):
    """Return (x_scale, y_scale) so xlsxwriter renders the logo at exactly the
    target pixel size.

    xlsxwriter sizes an image from its embedded pixel density:
        rendered_px = source_px * (96 / source_dpi) * scale
    (workbook._process_png falls back to 96 dpi when the PNG has none). We
    replicate that here so a PNG stored at any DPI still prints at the same
    standard size, while the full-resolution PNG stays embedded in the file.
    """
    data = logo_path.read_bytes()
    if len(data) < 24 or data[:8] != b"\x89PNG\r\n\x1a\n":
        return 1.0, 1.0

    src_w, src_h = struct.unpack(">II", data[16:24])

    # Locate the pHys chunk (physical pixel density, pixels per metre).
    x_density = y_density = 0
    pos = 8
    while pos + 12 <= len(data):
        length = struct.unpack(">I", data[pos : pos + 4])[0]
        chunk_type = data[pos + 4 : pos + 8]
        if chunk_type == b"pHYs" and pos + 20 <= len(data):
            x_density, y_density = struct.unpack(">II", data[pos + 8 : pos + 16])
            break
        pos += 12 + length  # length + type + data + CRC

    x_dpi = x_density * 0.0254 if x_density else 96.0
    y_dpi = y_density * 0.0254 if y_density else 96.0

    base_w = src_w * 96.0 / x_dpi
    base_h = src_h * 96.0 / y_dpi
    return target_w_px / base_w, target_h_px / base_h


def _write_sheet(
    ws,
    workbook,
    formats,
    sheet_name,
    data: AttendanceData,
    workers: list[Worker],
):
    date_count = len(data.date_headers)
    # Date headings are labels, not values. Writing them as text prevents Excel
    # from showing "###" when the narrow date columns use a larger font.
    headers = IDENTITY_HEADERS + [str(day) for day, _ in data.date_headers] + SUMMARY_HEADERS
    last_col = len(headers) - 1
    signature_col = last_col - 1
    filter_col = last_col
    date_start_col = len(IDENTITY_HEADERS)
    summary_start = date_start_col + date_count

    # Each worker occupies three rows, mirroring the raw jobcard files: the
    # main attendance row, the daily OT hours row, and the daily totals
    # (EOT) row. This keeps every daily OT value visible.
    ROWS_PER_WORKER = 3
    worker_count = max(len(workers), 1)
    last_data_row = DATA_START_ROW + ROWS_PER_WORKER * worker_count - 1

    weekend_flags = _weekend_flags(data.date_headers)
    _set_widths(ws, date_count)
    _set_page(ws, last_data_row, signature_col)

    # Rows 1-3 span the printable table. They are repeated on every printed page.
    ws.merge_range(0, 0, 0, signature_col, data.company, formats["company"])
    ws.merge_range(1, 0, 1, signature_col, data.address, formats["address"])
    ws.merge_range(2, 0, 2, signature_col, data.period_label, formats["period"])

    # Company logo on the top-right of the header band. It spans the company /
    # address / period rows plus the Sub Section row on the right side, where
    # nothing is written, so the info rows keep their original compact
    # heights. Its right edge stays LOGO_RIGHT_PAD_PX inside the print area
    # (Signature column) so it prints on every page - title rows 1-5 repeat
    # on every printed page. The image is embedded in the workbook package,
    # not linked, so it remains even if the source PNG is deleted later.
    if LOGO_PATH.exists():
        logo_x_scale, logo_y_scale = _logo_scale_factors(
            LOGO_PATH, LOGO_WIDTH_PX, LOGO_HEIGHT_PX
        )
        band_px = (
            _row_pixels(24)
            + _row_pixels(17)
            + _row_pixels(17)
            + _row_pixels(LOGO_SECTION_ROW_HEIGHT_PT)
        )
        if LOGO_HEIGHT_PX + 4 > band_px:
            print(
                f"Warning: logo ({LOGO_HEIGHT_PX}px) is taller than the "
                f"header band ({band_px}px) - raise "
                "LOGO_SECTION_ROW_HEIGHT_PT in constants.py."
            )
        logo_col, logo_x_offset = _logo_anchor(signature_col, LOGO_WIDTH_PX)
        y_offset = max((band_px - LOGO_HEIGHT_PX) // 2, 0)
        ws.insert_image(
            0,
            logo_col,
            str(LOGO_PATH),
            {
                "x_scale": logo_x_scale,
                "y_scale": logo_y_scale,
                "positioning": 2,  # don't move or size with cells
                "x_offset": logo_x_offset,
                "y_offset": y_offset,
                "description": "Company logo",
            },
        )
    else:
        print(
            f"Warning: company logo not found at {LOGO_PATH}; "
            "header logo was not embedded in this workbook."
        )

    # The subsection label is deliberately not merged: a normal formula in A4
    # recalculates when the SubSection Filter column is filtered.
    subsections = sorted({w.subsection for w in workers if w.subsection})
    formula = _label_formula(last_data_row, len(subsections))
    list_range = _write_subsection_list(ws, formats, subsections)
    ws.write_formula(3, 0, formula, formats["subsection"])
    ws.data_validation(
        3,
        0,
        3,
        0,
        {
            "validate": "list",
            "source": list_range,
            "dropdown": True,
            "input_title": "Sub Section",
            "input_message": (
                "Pick one subsection (automatic filtering uses the macro "
                "file). For several, use the check boxes in the SubSection "
                "Filter column; the label will show them joined with +."
            ),
        },
    )

    for col, title in enumerate(headers):
        is_date = date_start_col <= col < date_start_col + date_count
        weekend = bool(is_date and weekend_flags[col - date_start_col])
        fmt = formats["header_weekend"] if weekend else formats["header"]
        ws.write(4, col, title, fmt)

    for index, worker in enumerate(workers, start=1):
        base_row = DATA_START_ROW + (index - 1) * ROWS_PER_WORKER
        _write_worker_block(
            ws,
            formats,
            base_row,
            index,
            worker,
            date_count,
            weekend_flags,
        )

    if not workers:
        for r in range(DATA_START_ROW, DATA_START_ROW + ROWS_PER_WORKER):
            for col in range(last_col + 1):
                ws.write(r, col, "", formats["cell"])

    _write_filter_helpers(ws, filter_col, last_data_row, subsections)

    ws.autofilter(4, 0, last_data_row, last_col)
    ws.freeze_panes(HEADER_ROWS, 0)
    ws.repeat_rows(0, PRINT_TITLE_LAST_ROW)
    ws.print_area(0, 0, last_data_row, signature_col)
    # Rows 1-3 keep their original compact heights; the Sub Section row (4)
    # carries the extra height the logo needs on the right side. All five
    # repeat on every printed page.
    ws.set_row(0, 24)
    ws.set_row(1, 17)
    ws.set_row(2, 17)
    ws.set_row(3, LOGO_SECTION_ROW_HEIGHT_PT)
    ws.set_row(4, 32)
    for row in range(DATA_START_ROW, last_data_row + 1):
        # 15.1pt lets 43 attendance rows fit on an A4 landscape page after
        # Excel's fit-to-width scaling, avoiding a near-empty final page.
        ws.set_row(row, 15.1)

    # Protect the sheet so normal Excel editing cannot alter or delete the
    # embedded logo ("locked objects") or attendance data. The A4 subsection
    # picker stays editable (its format is unlocked) and the
    # AutoFilter dropdowns stay usable ("autofilter" allowed). The VBA side
    # re-applies the same protection at open with UserInterfaceOnly so the
    # macros keep full access - see vba.py / ProtectAttendanceSheets.
    ws.protect(
        SHEET_PROTECT_PASSWORD,
        {
            # NOTE on xlsxwriter's inverted semantics for these two keys:
            # "objects"/"scenarios" mean "editing is ALLOWED", so they must be
            # False to write objects="1" in the XML (logo cannot be edited
            # or deleted). All other keys mean "allowed" and must be True.
            "objects": False,  # lock objects -> the logo can't be deleted
            "scenarios": False,
            "format_cells": False,     # cell contents/formats stay locked
            "format_columns": True,    # drag column edges to resize + hide
            "format_rows": True,       # drag row edges to resize + hide
            "insert_columns": False,
            "insert_rows": False,
            "insert_hyperlinks": False,
            "delete_columns": False,
            "delete_rows": False,
            "select_locked_cells": True,  # view/copy stays possible
            "select_unlocked_cells": True,
            "sort": False,
            "autofilter": True,  # dropdown filtering keeps working
            "pivot_tables": False,
        },
    )

    workbook.define_name(
        f"{sheet_name}_SubSectionCell",
        f"='{sheet_name}'!$A$4",
    )
    workbook.define_name(
        f"{sheet_name}_SubSections",
        f"='{sheet_name}'!{list_range.lstrip('=')}",
    )

    # The macro reads this text and puts the live label formula back in A4.
    ws.write_string(1, LABEL_TEXT_COL, formula)
    workbook.define_name(
        f"{sheet_name}_LabelFormula",
        f"='{sheet_name}'!${xl_col_to_name(LABEL_TEXT_COL)}$2",
    )


def _write_worker_block(
    ws,
    formats,
    base_row,
    sl,
    worker: Worker,
    date_count: int,
    weekend_flags: list[bool],
):
    """Write three rows for one worker, mirroring the raw jobcard file:

    Row 1 - main attendance: SL, Name, Id No, Designation, daily P/W/A codes
          (codes only - no OT hours), monthly summary (P, A, OT, EOT, TOT,
          CL, SL, EL, H), Signature.
    Row 2 - daily OT hours under each date column.
    Row 3 - daily totals (EOT) under each date column.

    SL, Name, Id No, Designation, the monthly summary columns and Signature
    are merged vertically across all three rows, exactly like the raw
    jobcard; only the date columns stay as separate rows.

    All three rows carry the SubSection Filter value so AutoFilter hides/shows
    them together.
    """
    date_start_col = len(IDENTITY_HEADERS)
    summary_start = date_start_col + date_count
    signature_col = summary_start + 9
    filter_col = summary_start + 10

    days = list(worker.attendance[:date_count])
    while len(days) < date_count:
        days.append("")
    daily_ot = list(worker.daily_ot[:date_count])
    while len(daily_ot) < date_count:
        daily_ot.append(0.0)
    daily_eot = list(worker.daily_eot[:date_count])
    while len(daily_eot) < date_count:
        daily_eot.append(0.0)

    def _write_daily(r_main, col, value):
        # The main row always shows the attendance code (P/A/W/...);
        # OT hours are never written here - they live on the rows below.
        if weekend_flags[col - date_start_col]:
            ws.write(r_main, col, value, formats["cell_weekend"])
        elif str(value).upper() == "A":
            ws.write(r_main, col, value, formats["cell_absent"])
        else:
            ws.write(r_main, col, value, formats["cell"])

    # ---- Row 1: main attendance row ----
    r_main = base_row
    r_ot = base_row + 1
    r_eot = base_row + 2

    def _merged(col, value, fmt_key: str):
        """Merge a column across the worker's three rows, like the raw file."""
        ws.merge_range(r_main, col, r_eot, col, value, formats[fmt_key])

    _merged(0, sl, "cell")
    _merged(1, worker.name, "cell_left")
    _merged(2, worker.emp_id, "cell")
    _merged(3, worker.designation, "cell_left")
    for i, value in enumerate(days):
        _write_daily(r_main, date_start_col + i, value)

    # Monthly summary columns are live formulas so that editing any daily
    # code or value recalculates them automatically (like the raw jobcard
    # totals).
    date_first = xl_col_to_name(date_start_col)
    date_last = xl_col_to_name(date_start_col + date_count - 1)
    excel_r_main = r_main + 1
    excel_r_ot = r_ot + 1
    excel_r_eot = r_eot + 1
    code_range = f"${date_first}${excel_r_main}:${date_last}${excel_r_main}"
    ot_range = f"${date_first}${excel_r_ot}:${date_last}${excel_r_ot}"
    eot_range = f"${date_first}${excel_r_eot}:${date_last}${excel_r_eot}"
    ot_letter = xl_col_to_name(summary_start + SUMMARY_KEYS.index("OT"))
    eot_letter = xl_col_to_name(summary_start + SUMMARY_KEYS.index("EOT"))
    formulas = {
        "P": f'=COUNTIF({code_range},"P")',
        "A": f'=COUNTIF({code_range},"A")',
        "OT": f"=SUM({ot_range})",
        "EOT": f"=SUM({eot_range})",
        "TOT": f"={ot_letter}{excel_r_main}+{eot_letter}{excel_r_main}",
        "CL": f'=COUNTIF({code_range},"CL")',
        "SL": f'=COUNTIF({code_range},"SL")',
        "EL": f'=COUNTIF({code_range},"EL")',
        "H": f'=COUNTIF({code_range},"H")',
    }
    for offset, key in enumerate(SUMMARY_KEYS):
        col = summary_start + offset
        fmt = formats["number"]
        ws.merge_range(r_main, col, r_eot, col, "", fmt)
        # The parser value doubles as the cached formula result so the file
        # shows the correct numbers even before Excel recalculates.
        ws.write_formula(
            r_main, col, formulas[key], fmt, worker.summary.get(key, 0)
        )
    _merged(signature_col, "", "signature")
    ws.write_string(r_main, filter_col, worker.subsection, formats["filter_col"])

    # ---- Row 2: per-day OT hours ----
    # leave identity / summary / signature blank; daily OT hours are shown here
    # exactly like the raw jobcard: a 0 is printed where no OT was earned.
    for i, ot_val in enumerate(daily_ot):
        col = date_start_col + i
        ws.write_number(r_ot, col, float(ot_val), formats["number"])
    ws.write_string(r_ot, filter_col, worker.subsection, formats["filter_col"])

    # ---- Row 3: per-day totals (EOT) ----
    for i, eot_val in enumerate(daily_eot):
        col = date_start_col + i
        ws.write_number(r_eot, col, float(eot_val), formats["number"])
    ws.write_string(r_eot, filter_col, worker.subsection, formats["filter_col"])


def _weekend_flags(date_headers: list[tuple[int, str]]) -> list[bool]:
    return [weekday[:3].title() == "Fri" for _, weekday in date_headers]


def _set_widths(ws, date_count: int):
    date_start = len(IDENTITY_HEADERS)
    date_end = date_start + date_count - 1
    summary_start = date_start + date_count

    ws.set_column(0, 0, 3.2)
    ws.set_column(1, 1, 20)
    ws.set_column(2, 2, 8.2)
    # Designation: stored but hidden by default so it never appears in the
    # printout, even when macros are disabled (the Workbook_BeforePrint
    # event also hides it for print-only if someone un-hides it on screen).
    ws.set_column(3, 3, 14, None, {"hidden": True})
    ws.set_column(date_start, date_end, 1.9)

    # P A OT EOT TOT CL SL EL H Signature
    for offset, width in enumerate(SUMMARY_COL_WIDTHS):
        ws.set_column(summary_start + offset, summary_start + offset, width)
    # SubSection Filter - kept on screen for filtering, outside the print area
    ws.set_column(summary_start + 10, summary_start + 10, 15)
    # Hidden helper cells, dropdown list and label formula text
    ws.set_column(HELPER_NAME_COL, LABEL_TEXT_COL, 8.43, None, {"hidden": True})


def _set_page(ws, _last_row: int, _last_print_col: int):
    ws.set_paper(9)  # A4
    ws.set_landscape()
    # A generous bottom margin reserves a footer-only strip so the table
    # never overflows into the signature footer printed there. The footer
    # itself sits 0.35" from the paper edge, inside that reserved strip.
    ws.set_margins(0.25, 0.25, 0.25, 0.75)
    # NOTE: set_footer() below must carry matching margin options — it resets
    # xlsxwriter's internal footer margin to 0.3 whenever none is passed.
    ws.fit_to_pages(1, 0)  # one page wide, as many pages tall as needed
    ws.center_horizontally()
    ws.hide_gridlines(2)
    ws.set_footer("", {"margin": FOOTER_MARGIN})  # keep reserved strip
    ws.set_header("")
    ws.print_across()
    ws.set_zoom(80)


def _write_subsection_list(ws, formats, subsections: list[str]) -> str:
    ws.write(0, LIST_COL, "Sub sections", formats["list_header"])
    ws.write(1, LIST_COL, ALL_LABEL)
    for i, name in enumerate(subsections, start=2):
        ws.write(i, LIST_COL, name)
    excel_last = 2 + len(subsections)
    start = xl_col_to_name(LIST_COL)
    return f"=${start}$2:${start}${excel_last}"


def _write_filter_helpers(
    ws,
    filter_col: int,
    last_data_row: int,
    subsections: list[str],
):
    """Write hidden helper cells that list the currently visible subsections.

    The helpers deliberately avoid FILTER/SORT/UNIQUE so the label also works
    in Excel versions without dynamic arrays. SUBTOTAL(103, ...) ignores rows
    hidden by AutoFilter, so filtering recalculates the label automatically.
    """
    if not subsections:
        return

    first_row = DATA_START_ROW + 1
    last_row = last_data_row + 1
    filter_letter = xl_col_to_name(filter_col)
    name_letter = xl_col_to_name(HELPER_NAME_COL)
    sl_range = f"$A${first_row}:$A${last_row}"
    filter_range = f"${filter_letter}${first_row}:${filter_letter}${last_row}"

    for index, name in enumerate(subsections):
        row = DATA_START_ROW + index
        excel_row = row + 1
        name_ref = f"${name_letter}${excel_row}"
        ws.write_string(row, HELPER_NAME_COL, name)
        ws.write_formula(
            row,
            HELPER_FLAG_COL,
            (
                f"=IF(SUMPRODUCT(SUBTOTAL(103,OFFSET($A${first_row},"
                f"ROW({sl_range})-ROW($A${first_row}),0)),"
                f"--({filter_range}={name_ref}))>0,{name_ref},\"\")"
            ),
        )


def _label_formula(
    last_data_row: int,
    subsection_count: int,
) -> str:
    first_row = DATA_START_ROW + 1
    last_row = last_data_row + 1
    sl_range = f"$A${first_row}:$A${last_row}"
    helper_last = DATA_START_ROW + max(subsection_count, 1)
    flag_letter = xl_col_to_name(HELPER_FLAG_COL)
    flag_range = f"{flag_letter}{first_row}:{flag_letter}{helper_last}"
    join = f'TEXTJOIN(" + ",TRUE,{flag_range})'
    return (
        f'="{SUBSECTION_PREFIX}"&IF(AND(SUBTOTAL(103,{sl_range})'
        f'<COUNTA({sl_range}),{join}<>""),{join},"{ALL_LABEL}")'
    )
