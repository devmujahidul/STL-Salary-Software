import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = sys.argv[1]
doc = pymupdf.open(pdf)
print("pages:", doc.page_count)

for pno in range(doc.page_count):
    page = doc[pno]
    lines = []
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            text = "".join(s["text"] for s in line["spans"]).strip()
            if text:
                lines.append((line["bbox"][1], line["bbox"][0], text))
    data = [x for x in lines if x[0] > 90]
    data.sort()
    first = data[0][2] if data else ""
    last = data[-1][2] if data else ""
    print(f"p{pno+1:02d} rows~{len(data):2d} first: {first[:60]:60s} last: {last[:60]}")
