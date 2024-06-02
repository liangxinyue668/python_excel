
# source_sheet: 源工作表
# target_sheet: 目标工作表
# start_row_src: 源工作表起始行
# start_col_src: 源工作表起始列
# start_row_tgt: 目标工作表起始行
# start_col_tgt: 目标工作表起始列
# num_rows: 需要复制的行数
# num_cols: 需要复制的列数

def copy_data(source_sheet, target_sheet, start_row_src, start_col_src, start_row_tgt, start_col_tgt, num_rows, num_cols):

    # 从源工作表读取数据
    for i in range(num_rows):
        for j in range(num_cols):
            # 将数据写入目标工作表
            value =  source_sheet.cell(start_row_src + i, start_col_src + j).value
            target_sheet.cell(start_row_tgt + i, start_col_tgt + j, value)
