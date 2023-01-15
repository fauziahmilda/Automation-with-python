import openpyxl as xl
from openpyxl.chart import BarChart, Reference

wb = xl.load_workbook('transactions.xlsx')
sheet = wb['Sheet1']
cell = sheet['a1']
cell2 = sheet.cell(1, 1)
print(cell.value)

print(sheet.max_row)

for row in range(2, sheet.max_row + 1):  # include 2, 3, 4
    cell3 = sheet.cell(row, 3)  # cell in price section
    print(cell3.value)
    corrected_price = cell3.value * 0.9
    # add to new collumn
    corrected_price_cell = sheet.cell(row, 4)  # get the collumn
    corrected_price_cell.value = corrected_price  # put the value into the column

values = Reference(sheet,
                   min_row=2,
                   max_row=sheet.max_row,
                   min_col=4,
                   max_col=4)

chart = BarChart()
chart.add_data(values)
sheet.add_chart(chart, 'e2')


wb.save('transactions.xlsx')

# then, add a chart--------------------
# ->> import from openpyxl.chart import BarChart, Refference
