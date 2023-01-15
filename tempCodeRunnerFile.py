r row in range(2, sheet.max_row + 1):
    cell4 = sheet.cell(row, 4)  # select the value
    print(cell4.value)
