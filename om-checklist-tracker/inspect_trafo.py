from openpyxl import load_workbook

wb = load_workbook("Trafo_Checklist_Nov_2026_Populated.xlsx", read_only=True, data_only=True)
sheet = wb.active
print("Sheet title:", sheet.title)
for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=12, values_only=True), start=1):
    vals = [str(v) if v is not None else "" for v in row[:10]]
    print(f"Row {i}: {vals}")
