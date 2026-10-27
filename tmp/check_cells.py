import sys
import types

sys.modules["pymupdf.layout"] = types.ModuleType("pymupdf.layout")
sys.modules["pymupdf4llm"] = types.ModuleType("pymupdf4llm")

import pymupdf

pdf = sys.argv[1]
doc = pymupdf.open(pdf)

# Cell boundaries from the vertical rules on page 1.
page = doc[0]
xs = set()
for dr in page.get_drawings():
    for item in dr["items"]:
        if item[0] == "l":
            p1, p2 = item[1], item[2]
            if abs(p1.x - p2.x) < 0.5:
                xs.add(round(p1.x, 1))
        elif item[0] == "re":
            r = item[1]
            if r.width < 1:
                xs.add(round(r.x0, 1))
xs = sorted(xs)
print("boundaries:", xs)

name_left, name_right = 37.7, 127.7
desig_left, desig_right = 166.1, 230.1

name_over = []
desig_over = []
name_sizes = {}
desig_sizes = {}
for page in doc:
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if not text or "Sub Section:" in text:
                    continue
                x0, y0, x1, y1 = span["bbox"]
                if name_left - 1 <= x0 < name_right:
                    if x1 > name_right + 0.5:
                        name_over.append((page.number + 1, round(y0, 1), round(x1, 1), text))
                    name_sizes[round(span["size"], 2)] = name_sizes.get(round(span["size"], 2), 0) + 1
                if desig_left - 1 <= x0 < desig_right:
                    if x1 > desig_right + 0.5:
                        desig_over.append((page.number + 1, round(y0, 1), round(x1, 1), text))
                    desig_sizes[round(span["size"], 2)] = desig_sizes.get(round(span["size"], 2), 0) + 1

print("pages:", len(doc), "size:", page.rect)
print("name sizes:", sorted(name_sizes.items(), reverse=True)[:12])
print("designation sizes:", sorted(desig_sizes.items(), reverse=True)[:12])
print("name overflows:", len(name_over))
for item in name_over[:20]:
    print(item)
print("designation overflows:", len(desig_over))
for item in desig_over[:20]:
    print(item)
