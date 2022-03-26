import xlrd
from openpyxl.workbook import Workbook
#from openpyxl import load_workbook # for load excel file

wb = xlrd.open_workbook("PREZENTA DECEMBRIE 2021.xls")

ws = wb.sheet_by_index(0)

print(ws[2][4])
