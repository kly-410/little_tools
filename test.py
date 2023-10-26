
#!/usr/bin/env python3
# import common
from common.common import *
import common.common
from config.config import *
import openpyxl
import xlrd
import os


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





