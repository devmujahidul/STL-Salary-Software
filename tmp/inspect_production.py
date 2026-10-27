import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = pymupdf.open(pdf)

for pno in (26, 27, 36):
    page = doc[pno]
    print(f"\n===== PAGE {pno + 1} =====")
    rows = {}
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text:
                continue
            y = round(line["bbox"][1], 1)
            rows.setdefault(y, []).append((line["bbox"][0], text))
    for y in sorted(rows)[:12] + sorted(rows)[-4:]:
        items = sorted(rows[y])
        print(f"y={y:6.1f} :: " + " | ".join(t for _, t in items)[:220])
    print("bottom-most text y:", max(rows) if rows else None)
    print("page height:", page.rect.height)
