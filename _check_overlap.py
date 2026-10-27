from pathlib import Path
from collections import defaultdict
from bs4 import BeautifulSoup

ids = defaultdict(list)
raw_dir = Path(r"D:\MERND\raw_files")
for f in sorted(raw_dir.iterdir()):
    soup = BeautifulSoup(f.read_text(encoding="utf-8", errors="replace"), "lxml")
    for t in soup.find_all("table")[1:]:
        rows = t.find_all("tr")
        i = 0
        while i < len(rows):
            cells = [c.get_text(" ", strip=True) for c in rows[i].find_all(["td", "th"])]
            if len(cells) >= 3 and cells[0].isdigit() and "-" in cells[2]:
                ids[cells[2]].append(f.name)
                i += 3
                continue
            i += 1

dups = {k: v for k, v in ids.items() if len(set(v)) > 1 or len(v) > 1}
print("total unique", len(ids), "dup ids", len(dups))
for k, v in list(dups.items())[:15]:
    print(k, v)
