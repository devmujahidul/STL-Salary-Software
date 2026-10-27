import sys
import types

# pymupdf 2.0 auto-loads optional layout/ocr helpers whose onnxruntime import
# crashes on this machine. Stub them out so the core PDF renderer can be used.
sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import fitz

pdf = r"D:\MERND\tmp\Salary_Attendance_August_2026.pdf"
doc = fitz.open(pdf)
print("pages:", doc.page_count)
for i in [0, 1, 5, 20]:
    page = doc[i]
    pix = page.get_pixmap(dpi=110)
    out = rf"D:\MERND\tmp\page_{i + 1}.png"
    pix.save(out)
    print(out, page.rect)
