from __future__ import annotations

COMPANY_FALLBACK = "Spring Trade Ltd."
ADDRESS_FALLBACK = "Plot# 42, Konapara,Tenguri, Zirani, BKSP, Savar, Dhaka."

SHEET_PRODUCTION = "PRODUCTION"
SHEET_LO_PO = "LO_PO"
SHEET_FIXED = "FIXED"
SHEET_ORDER = (SHEET_LO_PO, SHEET_FIXED, SHEET_PRODUCTION)

PRODUCTION_PREFIXES = {"LI", "GPL", "TR", "IR"}
LO_PO_PREFIXES = {"LO", "PO"}
CLEANER_PREFIX = "CL"
CLEANER_SUBSECTION = "Cleaner and Sweaper"

IDENTITY_HEADERS = ["SL", "Name", "Id No", "Designation"]
SUMMARY_HEADERS = [
    "P",
    "A",
    "OT",
    "EOT",
    "TOT",
    "CL",
    "SL",
    "EL",
    "H",
    "Signature",
    "SubSection Filter",
]
SUMMARY_KEYS = ["P", "A", "OT", "EOT", "TOT", "CL", "SL", "EL", "H"]

HEADER_ROWS = 5  # company, address, period, subsection, table header
PRINT_TITLE_LAST_ROW = 4  # 0-based, rows 1-5
DATA_START_ROW = 5  # 0-based, first worker row (Excel row 6)

ALL_LABEL = "All"
SUBSECTION_PREFIX = "Sub Section: "

# Sheet that stores user-editable options, e.g. the printed footer signatures.
LIST_SHEET = "Lists"
# Default footer signature names; also written into the Lists sheet so they
# can be changed from inside Excel without touching the code.
FOOTER_SIGNATURES = ("Prepared by", "Checked by", "Approved by")
# Underscore line drawn after each signature name in the page footer.
FOOTER_LINE = "_" * 22
# Distance of the footer text from the paper edge, in inches. Must stay small
# against the reserved bottom margin (0.75") so it prints inside the strip.
FOOTER_MARGIN = 0.35
