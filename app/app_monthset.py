
#!/usr/bin/env python3
# import common
import sys
sys.path.append('../config')
sys.path.append('../script')
# import script
# from script.mon_settle import *
import openpyxl
import xlrd
import os

file = r"NC-管理费用2309_bak.xlsx"
sheet_name= "管理透视"
wb = openpyxl.load_workbook(file,data_only=True)   #加载




def get_data_from_nc(wb,sheet_name):
    work_sheet = wb[sheet_name]

    heji_row_list = []
    heji_col_list = []

    find_string_in_excel(wb, work_sheet, "总计", heji_row_list, heji_col_list)  #以 "总计" 为标识符，确定最大行和最大列
    print("总计1：行_%d,列_%d" %(heji_row_list[0],heji_col_list[0]))
    print("总计2：行_%d,列_%d" %(heji_row_list[1],heji_col_list[1]))




get_data_from_nc(file,sheet_name)


"""
1. 从nc表格，创建多个部门的表格（不需要，只需要，复制某个月份的进原始表格就可以）
2. 识别属于哪个部门
3. 费用分类写入指定部门


"""






# def process_my_files(fold_path):
#     flag =0
#     file_names = os.listdir(fold_path)
#     for file_name in file_names:
#         file_path = os.path.join(fold_path, file_name) #获取完整的文件路径

#         if os.path.isfile(file_path):
#             print("开始处理文件：", file_name)

#             ret = process_single_excel(file_path)
#             if ret == -2:
#                 print("已进行处理: 底稿>>>发送版")

#             print("success")
#             print("************************************************")

#         else:
#             print("忽略文件夹:", file_name)
#         flag += 1
#         # print("已经第搞定%d个"% flag)
        
#     print("提示：所有文件处理成功，底稿>>>>发送版")


# print_log()
# process_my_files(FOLD_PATH)
































