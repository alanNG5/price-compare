from openpyxl import load_workbook

xls = load_workbook('input.xlsx')

sheet = xls['sheet1']

data = []
for row in sheet[1:sheet.max_row]:
    row_data = []
    for cell in row:
        row_data.append(cell.value)
    data.append(row_data)

# for row in data:
#     print(row)
if __name__ == "__main__":
    print(data[2])
    print(data[6])
