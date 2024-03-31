import xlrd
import xlwt

excel_src_parh = './test.xlsx'
excel_new_path = './new.xlsx'

a = 5

excel_src = xlrd.open_workbook(excel_src_parh)
sheet_src = excel_src.sheet_by_name('源数据')
# sheet_src = excel_src.sheets()[0] #打开第一个表

new_data = []
new_data.append([
    sheet_src.cell_value(0, 0),
    sheet_src.cell_value(0, 1),
    sheet_src.cell_value(0, 2),
    sheet_src.cell_value(0, 4),
    sheet_src.cell_value(0, 5),
])
print(new_data)

for i in range(1, sheet_src.nrows):
    line1 = []
    line1.append(sheet_src.cell_value(i, 0))
    line1.append(sheet_src.cell_value(i, 1))
    line1.append(sheet_src.cell_value(i, 2))
    line1.append(sheet_src.cell_value(i, 4))
    line1.append(sheet_src.cell_value(i, 5))
    print(line1)
    new_data.append(line1)

    line2 = []
    line2.append(sheet_src.cell_value(i, 0))
    line2.append(sheet_src.cell_value(i, 1))
    line2.append(sheet_src.cell_value(i, 3))
    line2.append('')
    line2.append('')
    print(line2)
    new_data.append(line2)

excel_new = xlwt.Workbook(encoding='utf-8')
sheet_new = excel_new.add_sheet('new')

i = 0
for line in new_data:
    for j in range(5):
        sheet_new.write(i, j, line[j])
    i += 1

excel_new.save(excel_new_path) #如果文件已存在会被覆盖,执行此句python时,不能打开new.xlsx
