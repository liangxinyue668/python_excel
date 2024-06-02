
# source_sheet: 源工作表
# target_sheet: 目标工作表
# start_row_src: 源工作表起始行
# start_col_src: 源工作表起始列
# start_row_tgt: 目标工作表起始行
# start_col_tgt: 目标工作表起始列
# num_rows: 需要复制的行数
# num_cols: 需要复制的列数

def copy_data(source_sheet, target_sheet, start_row_src, start_col_src, start_row_tgt, start_col_tgt, num_rows, num_cols):

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
