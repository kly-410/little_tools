
#!/usr/bin/env python3
# import common
from common.common import *
import common.common
from config.config import *
import openpyxl
import xlrd
import os

# file = r"6、23年9月份管理费快乐部门.xlsx"
# sheetname="安环部预算执行表"
# wb = openpyxl.load_workbook(file,data_only=True)   #加载
# set_cell_color(wb, sheetname,5, 3, 1)
# set_row_color(wb, sheetname, 5, 3, 6, 1)

def process_my_files(fold_path):
    flag =0
    file_names = os.listdir(fold_path)
    for file_name in file_names:
        file_path = os.path.join(fold_path, file_name) #获取完整的文件路径

        if os.path.isfile(file_path):
            print("开始处理文件：", file_name)

            ret = process_single_excel(file_path)
            if ret == -2:
                print("已进行处理: 底稿>>>发送版")

            print("success")
            print("************************************************")

        else:
            print("忽略文件夹:", file_name)
        flag += 1
        # print("已经第搞定%d个"% flag)
        
    print("提示：所有文件处理成功，底稿>>>>发送版")


print_log()
process_my_files(FOLD_PATH)



# 在Python中，可以使用openpyxl库来操作Excel文件。要合并和拆分指定行列的单元格，可以使用openpyxl中的merge_cells()和unmerge_cells()方法。

# 合并单元格的方法是使用merge_cells()方法，该方法接受一个参数，即要合并的单元格范围。例如，要合并第1行到第3行，第1列到第2列的单元格，可以使用以下代码：

 
# from openpyxl import Workbook

# wb = Workbook()
# ws = wb.active

# # 合并单元格
# ws.merge_cells(start_row=1, end_row=3, start_column=1, end_column=2)

# wb.save("merged_cells.xlsx")
# 拆分单元格的方法是使用unmerge_cells()方法，该方法接受一个参数，即要拆分的单元格范围。例如，要拆分第1行到第3行，第1列到第2列的单元格，可以使用以下代码：

 
# from openpyxl import load_workbook

# wb = load_workbook("merged_cells.xlsx")
# ws = wb.active

# # 拆分单元格
# ws.unmerge_cells(start_row=1, end_row=3, start_column=1, end_column=2)

# wb.save("unmerged_cells.xlsx")






















# 1. python修改单元格的值，公式会自动计算
# 2. 读取公式，会读回来公式的表达式，而不是公式计算的值
# 3. 获取字符串在单元格的坐标



# row_read_from_excel(file, sheet_name, 1, 1, 3, row_val_list)
# column_read_from_excel(file, sheet_name, 1, 1, 3, cachelist)







# for i in range(len(row_val_list)):
#     print(row_val_list[i])#打印单元格的值



# for i in range(len(cachelist)):
#     print(cachelist[i])#打印单元格的值





