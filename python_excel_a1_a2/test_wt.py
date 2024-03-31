import xlwt
import xlrd

excel_l = xlrd.open_workbook('./test.xlsx')

sheet_1 = excel_l.sheet_by_name('源数据')
# sheet = excel.sheets()[0] #打开第一个表

excel_2 = xlwt.Workbook(encoding='utf-8')
sheet_2 = excel_2.add_sheet('aa')

sheet_2.write(0, 0, sheet_1.cell_value(0, 0))
sheet_2.write(0, 1, sheet_1.cell_value(0, 1))
sheet_2.write(0, 2, sheet_1.cell_value(0, 2))
sheet_2.write(0, 3, sheet_1.cell_value(0, 4))
sheet_2.write(0, 4, sheet_1.cell_value(0, 5))


for i in range(1, sheet_1.nrows):
    sheet_2.write(i*2 - 1, 0, sheet_1.cell_value(i, 0))
    sheet_2.write(i*2 - 1, 1, sheet_1.cell_value(i, 1))
    sheet_2.write(i*2 - 1, 2, sheet_1.cell_value(i, 2))
    sheet_2.write(i*2 - 1, 3, sheet_1.cell_value(i, 4))
    sheet_2.write(i*2 - 1, 4, sheet_1.cell_value(i, 5))

    sheet_2.write(i*2, 0, sheet_1.cell_value(i, 0))
    sheet_2.write(i*2, 1, sheet_1.cell_value(i, 1))
    sheet_2.write(i*2, 2, sheet_1.cell_value(i, 3))


excel_2.save('./new.xlsx')

