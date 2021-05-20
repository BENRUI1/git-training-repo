import openpyxl
import configparser
import os

print(os.getcwd())

# Read ini file
parser = configparser.ConfigParser()
parser.read('createScoreCustomers.ini')

customersFile=parser['files']['customersFile']

print(customersFile)

# Give the location of the file
loc = ("scoreCustomers.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
print(sheet.cell_value(0, 0))
print(sheet)
print(wb)
print(sheet.nrows)
print(sheet.ncols)

for i in range(sheet.ncols):
    print(sheet.cell_value(0, i))

for z in range(sheet.nrows):
    print(sheet.cell_value(z,0))

excel_data_df = pandas.read_excel('Names.xlsx', sheet_name='Sheet1')
json_str = excel_data_df.to_json()
print('Excel Sheet to JSON:\n', json_str)
