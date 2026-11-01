"""Build a repro workbook whose sub sections look like the user's data
("AO Barteck", "AO Finishing", "AO Hole Button", ...) and drive the A4 picker.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

RAW_DIR = ROOT / "tmp" / "raw_user"
OUT = ROOT / "tmp" / "user_sim.xlsx"

SRC = ROOT / "raw_files" / "matrix_jobcardxl.xls"
RAW_DIR.mkdir(parents=True, exist_ok=True)

text = SRC.read_text(encoding="utf-8", errors="replace")

# AO-Finishing -> "AO Finishing"
text = text.replace("AO-Finishing", "AO Finishing")
# AO-Knitting -> alternate "AO Barteck" / "AO Hole Button" (per occurrence)
counter = {"n": 0}


def _knitting(_match):
    counter["n"] += 1
    return "AO Barteck" if counter["n"] % 2 == 1 else "AO Hole Button"


text = re.sub(r"AO-Knitting", _knitting, text)
# AO-Jacquard -> "AO Jacquard"
text = text.replace("AO-Jacquard", "AO Jacquard")

(RAW_DIR / "matrix_jobcardxl.xls").write_text(text, encoding="utf-8")
print("raw patched:", RAW_DIR / "matrix_jobcardxl.xls")

from attendance_export.classify import group_workers
from attendance_export.parser import parse_raw_files
from attendance_export.vba import attach_filter_macros
from attendance_export.writer import write_workbook

data = parse_raw_files(RAW_DIR)
grouped = group_workers(data.workers)
for sheet, workers in grouped.items():
    subs = sorted({w.subsection for w in workers})
    print(f"{sheet}: {len(workers)} workers, subs={subs}")

path = write_workbook(data, OUT)
ok, reason = attach_filter_macros(path)
print("macros:", ok, reason)
print("built:", path.with_suffix(".xlsm"))
