import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

doc = pymupdf.open(sys.argv[1] if len(sys.argv) > 1 else r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf")
max_name = (0, None)
max_desig = (0, None)
max_id = (0, None)
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
                if 38 < x0 < 45:
                    if x1 > max_name[0]:
                        max_name = (x1, (page.number + 1, round(y0, 1), text))
                elif 164 < x0 < 170:
                    if x1 > max_desig[0]:
                        max_desig = (x1, (page.number + 1, round(y0, 1), text))
                elif 130 < x0 < 140:
                    if x1 > max_id[0]:
                        max_id = (x1, (page.number + 1, round(y0, 1), text))
print("max name x1:", max_name)
print("max designation x1:", max_desig)
print("max id x1:", max_id)
# boundaries from header
page = doc[0]
for block in page.get_text("dict")["blocks"]:
    if block["type"] != 0:
        continue
    for line in block["lines"]:
        for span in line["spans"]:
            if span["text"].strip() in ("Name", "Id No", "Designation"):
                print(span["text"].strip(), span["bbox"])
