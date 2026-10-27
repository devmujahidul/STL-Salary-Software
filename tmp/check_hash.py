import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = pymupdf.open(pdf)
hits = []
for page in doc:
    text = page.get_text()
    if "#" in text:
        for line in text.splitlines():
            if "#" in line:
                hits.append((page.number + 1, line.strip()))
print("pages with #:", len({p for p, _ in hits}))
for item in hits[:30]:
    print(item)
