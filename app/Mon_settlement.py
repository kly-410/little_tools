
#!/usr/bin/env python3
# import common
from common.common import *
from common.style import *
import common.common
from config.config import *
import openpyxl
import xlrd
import os

# file = r"1、23年9月份管理费用预算执行表-安全环保部.xlsx"
# sheetname="安环部预算执行表"
# wb = openpyxl.load_workbook(file,data_only=True)   #加载









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
































