from pathlib import Path
from bs4 import BeautifulSoup

path = Path(r"D:\MERND\raw_files\matrix_jobcard (1).xls")
text = path.read_text(encoding="utf-8", errors="replace")
print("len", len(text), "tables_tag", text.lower().count("<table"))
soup = BeautifulSoup(text, "lxml")
tables = soup.find_all("table")
print("n tables", len(tables))
for i, t in enumerate(tables[:12]):
    rows = t.find_all("tr")
    print(f"table {i}: rows={len(rows)}")
    for r in rows[:8]:
        cells = [c.get_text(" ", strip=True) for c in r.find_all(["td", "th"])]
        if any(cells):
            print(" ", cells[:25], "ncols", len(cells))
