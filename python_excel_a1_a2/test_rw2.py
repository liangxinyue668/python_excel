import xlwt
import xlrd

excel_l = xlrd.open_workbook('./test.xlsx')

sheet_1 = excel_l.sheet_by_name('源数据')
# sheet = excel.sheets()[0] #打开第一个表

# dic = {
#     '拼接':       0,
#     'SPU':        1,
#     'A2':         2,
#     'A1':         3,
#     '库龄0~30天':  4,
#     '库龄30~60天': 5,
# }

# dic['拼接'] = 0
# dic['SPU'] = 1


dic = {}

for i in range(6):
    dic[sheet_1.cell_value(0, i)] = i

excel_2 = xlwt.Workbook(encoding='utf-8')
sheet_2 = excel_2.add_sheet('aa')

# sheet_2.write(0, 0, sheet_1.cell_value(0, dic['拼接']))
# sheet_2.write(0, 1, sheet_1.cell_value(0, dic['SPU']))
# sheet_2.write(0, 2, sheet_1.cell_value(0, dic['A2']))
# sheet_2.write(0, 3, sheet_1.cell_value(0, dic['库龄0~30天']))
# sheet_2.write(0, 4, sheet_1.cell_value(0, dic['库龄30~60天']))

titles = ['拼接', 'SPU', 'A2', '库龄0~30天', '库龄30~60天']

for i in range(len(titles)):
    sheet_2.write(0, i, titles[i])

old_data = []

for i in range(1, sheet_1.nrows):
    row = []
    for j in range(len(titles)):
        row.append(sheet_1.cell_value(i, dic[titles[j]]))
    old_data.append(row)

    row = []
    row.append('a2')
    row.append(sheet_1.cell_value(i, dic['SPU']))
    row.append(sheet_1.cell_value(i, dic['A1']))
    old_data.append(row)

print(old_data)

for i in range(len(old_data)):
    for j in range(len(old_data[i])):
        sheet_2.write(i + 1, j, old_data[i][j])



excel_2.save('./new.xlsx')

