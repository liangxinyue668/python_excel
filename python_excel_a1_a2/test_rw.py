import xlwt
import xlrd

excel_l = xlrd.open_workbook('./test.xlsx')

sheet_1 = excel_l.sheet_by_name('源数据')
# sheet = excel.sheets()[0] #打开第一个表

excel_2 = xlwt.Workbook(encoding='utf-8')
sheet_2 = excel_2.add_sheet('aa')
sheet_2.write(0,0,'相加结果')

for i in range(1, sheet_1.nrows):
    value_3 = sheet_1.cell_value(i, 2)
    value_4 = sheet_1.cell_value(i, 3)
    print(value_3, value_4, value_3 + value_4)
    sheet_2.write(i, 0, value_3 + value_4)

excel_2.save('./new.xlsx')

