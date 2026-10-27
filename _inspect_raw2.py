from pathlib import Path
from bs4 import BeautifulSoup
from collections import Counter, defaultdict

path = Path(r"D:\MERND\raw_files\matrix_jobcard (1).xls")
soup = BeautifulSoup(path.read_text(encoding="utf-8", errors="replace"), "lxml")
tables = soup.find_all("table")

# full header
t = tables[1]
rows = t.find_all("tr")
header_cells = [c.get_text(" ", strip=True) for c in rows[1].find_all(["td", "th"])]
print("HEADER COLS", len(header_cells))
for i, h in enumerate(header_cells):
    print(i, repr(h))

print("\n=== first worker 3 rows full ===")
for ridx in range(2, 5):
    cells = [c.get_text(" ", strip=True) for c in rows[ridx].find_all(["td", "th"])]
    print("row", ridx, "ncols", len(cells))
    for i, v in enumerate(cells):
        print(f"  {i}:{v!r}")

print("\n=== second worker first row last cols ===")
cells = [c.get_text(" ", strip=True) for c in rows[5].find_all(["td", "th"])]
print("ncols", len(cells))
for i, v in enumerate(cells):
    print(f"  {i}:{v!r}")

# scan all files for subsections and ID prefixes
raw_dir = Path(r"D:\MERND\raw_files")
all_subs = Counter()
id_prefixes = Counter()
designations = Counter()
id_examples = defaultdict(list)
status_vals = Counter()
company_lines = set()
month_lines = set()

for f in sorted(raw_dir.iterdir()):
    print("\nFILE", f.name)
    soup = BeautifulSoup(f.read_text(encoding="utf-8", errors="replace"), "lxml")
    tables = soup.find_all("table")
    if tables:
        t0 = tables[0]
        t0rows = t0.find_all("tr")
        for r in t0rows:
            txt = r.get_text(" ", strip=True)
            if txt:
                company_lines.add(txt)
                print("  hdr:", txt)
    n_workers = 0
    for t in tables[1:]:
        rows = t.find_all("tr")
        if not rows:
            continue
        first = rows[0].get_text(" ", strip=True)
        if first.startswith("Sub Section"):
            sub = first.split(":", 1)[-1].strip()
            all_subs[sub] += 1
        # find header
        hdr = None
        start = 0
        for i, r in enumerate(rows[:4]):
            cells = [c.get_text(" ", strip=True) for c in r.find_all(["td", "th"])]
            if cells and cells[0] == "SL":
                hdr = cells
                start = i + 1
                break
        i = start
        while i < len(rows):
            cells = [c.get_text(" ", strip=True) for c in rows[i].find_all(["td", "th"])]
            if not cells:
                i += 1
                continue
            if cells[0] in ("SL",) or cells[0].startswith("Sub Section"):
                i += 1
                continue
            # worker row typically starts with a number SL and has name
            if len(cells) >= 4 and cells[2] and any(ch.isalpha() for ch in cells[2]):
                n_workers += 1
                emp_id = cells[2]
                prefix = emp_id.split("-")[0] if "-" in emp_id else emp_id[:2]
                id_prefixes[prefix] += 1
                if len(id_examples[prefix]) < 3:
                    id_examples[prefix].append((emp_id, cells[1], cells[3], first))
                designations[cells[3]] += 1
                for v in cells[4:]:
                    if v and not v.replace(".", "", 1).isdigit():
                        status_vals[v] += 1
                i += 3
                continue
            i += 1
    print("  workers", n_workers)

print("\n=== company/month ===")
for x in sorted(company_lines):
    print(x)
print("\n=== subsections ===")
for k, v in all_subs.most_common():
    print(f"{v:4} {k}")
print("\n=== id prefixes ===")
for k, v in id_prefixes.most_common():
    print(f"{v:4} {k} examples {id_examples[k]}")
print("\n=== designations ===")
for k, v in designations.most_common():
    print(f"{v:4} {k}")
print("\n=== status sample ===")
print(status_vals.most_common(40))
