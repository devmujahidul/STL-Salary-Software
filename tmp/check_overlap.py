import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = sys.argv[1] if len(sys.argv) > 1 else r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = pymupdf.open(pdf)

name_over = []
desig_over = []
for page in doc:
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if not text:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                # Name column data starts around x=40 and should end before the ID column at 133.7.
                if 38 < x0 < 45 and x1 > 132:
                    name_over.append((page.number + 1, round(y0, 1), round(x1, 1), text))
                # Designation column data starts around x=166 and should end before date cells at 227.9.
                if 164 < x0 < 170 and x1 > 226:
                    desig_over.append((page.number + 1, round(y0, 1), round(x1, 1), text))

print("name overflows:", len(name_over))
for item in name_over[:20]:
    print(item)
print("designation overflows:", len(desig_over))
for item in desig_over[:20]:
    print(item)
