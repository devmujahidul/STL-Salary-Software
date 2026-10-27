import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = pymupdf.open(pdf)

for pno in (0, 1, 2, 3, 20):
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
            y = round(line["bbox"][1] / 4) * 4
            rows.setdefault(y, []).append((line["bbox"][0], text, round(line["spans"][0]["size"], 1)))
    for y in sorted(rows):
        if y > 300 and pno in (0,):
            continue
        items = sorted(rows[y])
        joined = " | ".join(t for _, t, _ in items)
        print(f"y={y:4d} size={items[0][2]:4.1f} :: {joined[:300]}")
