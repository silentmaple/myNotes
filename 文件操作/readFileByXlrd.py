"""
xlrd和xlwt的基本用法
"""
import os
import xlrd
import xlwt

"""
基本思路:
xlrd的使用方法与excel中的框架一致。
workbook：工作簿，worksheet：工作表，cell：单元格。
"""
# 1.打开Excel文件
filepath = os.getcwd() + "\\test.xls"
wb_1 = xlrd.open_workbook(filepath)

# 2.返回工作簿的所有sheet
all_sheets = wb_1.sheet_names()

# 3.打开指定sheet
# 通过sheet索引打开
sh_1 = wb_1.sheets()[0]
sh_2 = wb_1.sheet_by_index(0)

# 通过名称获取
sh_3 = wb_1.sheet_by_name("sheet1")

# 4.对sheet表的行操作
# 获取sheet表的行数
row_1 = sh_1.nrows()
# 返回由该行中所有的单元格对象组成的列表，列表内是键值对。
res_1 = sh_1.row(1)
res_2 = sh_1.row_slice(1)
# 返回指定行的所有单元格数值组成的列表。
sh_1.row_values(1)
# 返回指定行的有效长度。
sh_1.row_len(1)

# 5.对sheet表的列操作
# 返回指定sheet表的有效列数
col_num = sh_1.ncols()
# 返回由该列中所有的单元格对象组成的列表，列表内是键值对。
sh_1.col(1)
sh_1.col_slice(1)
# 返回指定列的所有单元格数值组成的列表，可以截取部分行。
sh_1.col_values(1, 0, 10)

# 6.对sheet表单元格的操作
# 返回单元格对象
sh_1.cell(1, 1)
# 返回对应位置单元格中的数据类型
sh_1.cell_type(1, 1)
# 返回对应位置单元格中的数据
sh_1.cell_value(1, 1)
