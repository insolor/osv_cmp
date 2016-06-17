import xlrd
rb = xlrd.open_workbook('d:/final.xls',formatting_info=True)
sheet = rb.sheet_by_index(0)
for rownum in range(sheet.nrows):
row = sheet.row_values(rownum)
for c_el in row:
print c_el
