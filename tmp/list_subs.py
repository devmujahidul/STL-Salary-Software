"""List the sub sections per sheet in the current raw data."""
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from attendance_export.classify import group_workers
from attendance_export.parser import parse_raw_files

data = parse_raw_files(Path(__file__).resolve().parent.parent / "raw_files")
grouped = group_workers(data.workers)
for sheet, workers in grouped.items():
    counts = Counter(w.subsection for w in workers)
    print(sheet, dict(counts))
