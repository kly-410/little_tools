
#!/usr/bin/env python3
# import common
from common.common import *
import common.common
import openpyxl
import xlrd


# file    = r'test1.xlsx'
file    = r'1、23年9月份管理费用预算执行表-安全环保部.xlsx'
sheet_name = "安环部预算执行表"
# sheet_name = "功率"
row_val_list = []

column_val_list =[]

tmp = 0





column_copy(file, sheet_name, 17, 4, 6, 118)


# file.save(file)

# excel = SumExcel(file, sheet_name, 1, 1, 5)
# total = excel.return_column_sum()


# val = read_from_excel_cell(file, sheet_name, 8, 1)
# print(val)
# # column_swag(file, sheet_name, 1, 2, 1, 5)

# result = []
# print(find_string_in_excel(file, sheet_name, "哈哈", result))

# print(result[0])
# print(result[1])

"""
主程序步骤
1. 由透视后表格创建各部门新的表格
2. 各部门新的表格，数据处理，产生底稿
3. 数据处理后，删除一部分，产生发送版


1. 数据校验
2. 格式美化

"""

























# 1. python修改单元格的值，公式会自动计算
# 2. 读取公式，会读回来公式的表达式，而不是公式计算的值
# 3. 获取字符串在单元格的坐标



# row_read_from_excel(file, sheet_name, 1, 1, 3, row_val_list)
# column_read_from_excel(file, sheet_name, 1, 1, 3, cachelist)







# for i in range(len(row_val_list)):
#     print(row_val_list[i])#打印单元格的值



# for i in range(len(cachelist)):
#     print(cachelist[i])#打印单元格的值





