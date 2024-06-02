import openpyxl
import excel_tool

excel_1 = openpyxl.load_workbook('./文件1.xlsx')
excel_2 = openpyxl.load_workbook('./文件2.xlsx')
sheet_1 = excel_1['Sheet1']
sheet_2 = excel_2['Sheet1']
sheet_3 = excel_2['Sheet2']

# 文件2第一个表格的数据需要移动到文件2的第二个表格
excel_tool.copy_data(sheet_2, sheet_3, start_row_src=1, start_col_src=1, start_row_tgt=8, start_col_tgt=1, num_rows=5, num_cols=7)

# 文件1数据复制到文件2
excel_tool.copy_data(sheet_1, sheet_2, start_row_src=1, start_col_src=1, start_row_tgt=1, start_col_tgt=1, num_rows=5, num_cols=7)

excel_2.save('./文件2.xlsx')

excel_2.close()
