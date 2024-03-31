import sys
import xlrd

excel = xlrd.open_workbook('./test.xlsx')

sheet = excel.sheet_by_name('源数据')
# sheet = excel.sheets()[0] #打开第一个表

sheet.nrows #一共有几行

for i in range(1, sheet.nrows):
    value_3 = sheet.cell_value(i, 2)
    value_4 = sheet.cell_value(i, 3)
    print(value_3, value_4, value_3 + value_4)


