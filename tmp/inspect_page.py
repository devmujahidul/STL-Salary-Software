import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = pymupdf.open(pdf)
page = doc[0]
print("PAGE", page.number + 1, page.rect)
data = page.get_text("dict")
for block in data["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        text = "".join(span["text"] for span in line["spans"]).strip()
        if not text:
            continue
        x0, y0, x1, y1 = line["bbox"]
        sizes = sorted({round(span["size"], 1) for span in line["spans"]})
        print(f"y={y0:6.1f}-{y1:6.1f} x={x0:6.1f}-{x1:6.1f} size={sizes} | {text[:130]}")
