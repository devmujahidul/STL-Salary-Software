import xlrd

path = r"D:\MERND\Format.xls"
book = xlrd.open_workbook(path, formatting_info=True)
sh = book.sheet_by_name("Format")
xf_list = book.xf_list
fonts = book.font_list

def describe_xf(xf_index):
    xf = xf_list[xf_index]
    font = fonts[xf.font_index]
    align = xf.alignment
    return {
        "font": font.name,
        "height": font.height / 20,
        "bold": font.bold,
        "italic": font.italic,
        "colour": font.colour_index,
        "horz": align.hor_align,
        "vert": align.vert_align,
        "wrap": getattr(align, "text_wrapped", None),
        "border_t": xf.border.top_line_style,
        "fill": xf.background.fill_pattern,
        "pat_fore": xf.background.pattern_colour_index,
    }

print("default style", describe_xf(15) if len(xf_list) > 15 else None)
for r in range(5):
    cell = sh.cell(r, 0)
    print("r", r, "xf", cell.xf_index, describe_xf(cell.xf_index), "val", cell.value)

print("header row 4 styles sample")
for c in [0, 1, 2, 3, 34, 43, 44]:
    cell = sh.cell(4, c)
    print("c", c, describe_xf(cell.xf_index), "val", cell.value)

print("col widths:")
for c in range(sh.ncols):
    info = sh.colinfo_map.get(c)
    if info:
        print(c, "width", info.width, "hidden", info.hidden)

print("row heights:")
for r in range(6):
    info = sh.rowinfo_map.get(r)
    if info:
        print(r, "height", info.height, "hidden", info.hidden)

print("sheet attrs related:")
for a in sorted(dir(sh)):
    if any(k in a.lower() for k in ("print", "page", "margin", "fit", "orient", "paper", "header", "footer", "show")):
        try:
            print(a, getattr(sh, a))
        except Exception as e:
            print(a, e)

print("book.user_name", book.user_name)
print("nsheets", book.nsheets)
