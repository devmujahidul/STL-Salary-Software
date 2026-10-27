import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = sys.argv[1]
doc = pymupdf.open(pdf)

for pno in (0, 2, 3, 24, 25, 34):
    page = doc[pno]
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
    print(f"\n=== PAGE {pno + 1} ===")
    ys = sorted(rows)
    for y in ys[:6]:
        print(f"y={y:6.1f} :: " + " | ".join(t for _, t in sorted(rows[y]))[:180])
    if len(ys) > 10:
        print("...")
        for y in ys[-3:]:
            print(f"y={y:6.1f} :: " + " | ".join(t for _, t in sorted(rows[y]))[:180])
    print("bottom text y:", max(ys), "page bottom margin:", page.rect.height - 18)
