from pathlib import Path
from bs4 import BeautifulSoup
from collections import Counter

path = Path(r"D:\MERND\raw_files\matrix_jobcard (2).xls")
soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "lxml")
tables = soup.find_all("table")
print("tables", len(tables))
for i, t in enumerate(tables):
    rows = t.find_all("tr")
    kinds = Counter()
    sub = None
    for r in rows:
        cells = [c.get_text(" ", strip=True) for c in r.find_all(["td", "th"])]
        if not cells:
            kinds["empty"] += 1
            continue
        t0 = cells[0]
        if t0.startswith("Sub Section"):
            kinds["subsection"] += 1
            sub = t0
        elif t0 == "SL":
            kinds["header"] += 1
        elif t0.isdigit() and len(cells) >= 4 and "-" in cells[2]:
            kinds["worker"] += 1
        elif len(cells) == 31:
            kinds["extra31"] += 1
        else:
            kinds["other"] += 1
            if kinds["other"] <= 3:
                print(" other example", cells[:8], "n", len(cells))
    print(f"table {i} rows={len(rows)} kinds={dict(kinds)} firstsub={sub}")
