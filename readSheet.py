import xlrd

workbook = xlrd.open_workbook('Financial Sample.xlsx')

# Load sheet by name.
#worksheet = workbook.sheet_by_name('Sheet1')
# Load sheet by index, in this case, first sheet.
worksheet = workbook.sheet_by_name(0)

# gives you a list of the names of the sheets present in the file, which helps you iterate over the sheets.
workbook.sheet_names()

# Value of 1st row and 1st column
#worksheet.cell(0, 0).value

if worksheet.cell(0, 0).value == xlrd.empty_cell.value:
    print("Cell is empty.")