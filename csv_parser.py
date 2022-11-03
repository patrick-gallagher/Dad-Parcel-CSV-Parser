import os.path
import openpyxl

csv_path = 'C:\\Temp\\'

wb = openpyxl.load_workbook(csv_path + 'Parcel_CSV_Example.xlsx')
raw_sheet = wb.worksheets[0]
parsed_sheet = wb.worksheets[1]

start_cell = 2
end_cell = 7

for row in raw_sheet.iter_rows(min_row=start_cell, max_col=1, max_row=end_cell):
    for cell in row:
        print(cell.value)