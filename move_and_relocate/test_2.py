import openpyxl

excel_1 = openpyxl.load_workbook('./文件1.xlsx')
excel_2 = openpyxl.load_workbook('./文件2.xlsx')
sheet_1 = excel_1['Sheet1']
sheet_2 = excel_2['Sheet1']
sheet_3 = excel_2['Sheet2']

def copy_data(source_sheet, target_sheet, start_row_src, start_col_src, start_row_tgt, start_col_tgt, num_rows, num_cols):
    """
    从源工作表复制数据到目标工作表。
    
    参数：
    source_sheet: 源工作表
    target_sheet: 目标工作表
    start_row_src: 源工作表起始行
    start_col_src: 源工作表起始列
    start_row_tgt: 目标工作表起始行
    start_col_tgt: 目标工作表起始列
    num_rows: 需要复制的行数
    num_cols: 需要复制的列数
    """
    data = []

    # 从源工作表读取数据
    for i in range(start_row_src, start_row_src + num_rows):
        row = []
        for j in range(start_col_src, start_col_src + num_cols):
            row.append(source_sheet.cell(i, j).value)
        data.append(row)
    
    # 将数据写入目标工作表
    for i in range(len(data)):
        for j in range(len(data[i])):
            target_sheet.cell(start_row_tgt + i, start_col_tgt + j, data[i][j])

# 文件2第一个表格的数据需要移动到文件2的第二个表格
copy_data(sheet_2, sheet_3, start_row_src=1, start_col_src=1, start_row_tgt=8, start_col_tgt=1, num_rows=5, num_cols=7)

# 文件1数据复制到文件2
copy_data(sheet_1, sheet_2, start_row_src=1, start_col_src=1, start_row_tgt=1, start_col_tgt=1, num_rows=5, num_cols=7)


excel_2.save('./文件2.xlsx')

excel_2.close()


