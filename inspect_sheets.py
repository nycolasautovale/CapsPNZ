import openpyxl

wb = openpyxl.load_workbook("data/producao2026.xlsx", data_only=True)
print("ABAS:", wb.sheetnames)
print("=" * 80)

for name in wb.sheetnames:
    ws = wb[name]
    print(f"\n### ABA: {name}  (dims={ws.dimensions}, max_row={ws.max_row}, max_col={ws.max_column})")
    # Print first 12 rows, all columns, to understand layout
    for r in range(1, min(ws.max_row, 14) + 1):
        cells = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is None:
                v = ""
            cells.append(str(v)[:18])
        print(f"R{r:>2}: " + " | ".join(cells))
