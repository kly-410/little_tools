
#!/usr/bin/env python3
# import common
from common.common import *
import common.common
from config.config import *
import openpyxl
import xlrd
import os

# sheet_name = "安环部预算执行表"

# file = r"1、23年9月份管理费用预算执行表-安全环保部.xlsx"
# # wb = openpyxl.load_workbook(file)   #加载

# process_single_excel(file)

def process_my_files(fold_path):
    flag =0
    file_names = os.listdir(fold_path)
    for file_name in file_names:
        file_path = os.path.join(fold_path, file_name) #获取完整的文件路径

        if os.path.isfile(file_path):
            print("开始处理文件：", file_path)
            process_single_excel(file_path)
            print("文件处理成功：", file_path)
            print("************************************************")

        else:
            print("忽略文件夹:", file_name)
        flag += 1
        print("已经第搞定%d个"% flag)
        
    # print("提示：所有文件处理成功，底稿 》》》》发送版")


print_log()
process_my_files(FOLD_PATH)


# def process_files(folder_path):
#     # 获取文件夹中的所有文件名
#     file_names = os.listdir(folder_path)
    
#     # 遍历每一个文件名
#     for file_name in file_names:
#         # 获取完整的文件路径
#         file_path = os.path.join(folder_path, file_name)
        
#         # 判断是否为文件（而不是文件夹）
#         if os.path.isfile(file_path):
#             print("Processing file:", file_path)
#             # 在这里添加处理文件的代码
#             # ...
#         else:
#             print("Ignoring directory:", file_path)

# # 调用函数，传入文件夹路径
# process_files("F:/处理9月份表格/2_各部门分表")













# # file    = r'test1.xlsx'
# file    = r'1、23年9月份管理费用预算执行表-安全环保部.xlsx'
# sheet_name = "安环部预算执行表"
# # sheet_name = "功率"
# # sheet_name = "Sheet1"

# wb = openpyxl.load_workbook(file,data_only=True)   #加载
# copy_data_from_src(wb, sheet_name)
# del_none_row_and_col(wb, sheet_name)
# write_sum_to_xiaoji_row(wb, sheet_name)
# wb.save(file)




# 如果已经处理，就不处理了
"""
主程序步骤
1. 由透视后表格创建各部门新的表格
2. 各部门新的表格，数据处理，产生底稿
3. 数据处理后，删除一部分，产生发送版

1. 删除一行
2. 增减一行
3. 删除一列
4. 增加一列
如果已经处理，就不处理了


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





