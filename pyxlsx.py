import openpyxl
from openpyxl import load_workbook

wb2 = load_workbook('PR_April15.xlsx', read_only=True)


print wb2.get_sheet_names()
# sheet1 = wb2['FirstSheet']
sheets = wb2.get_sheet_names()
print sheets
sheet1 = wb2[sheets[0]]
print sheet1['A1'].value
#print sheet1.cell('A','1')
