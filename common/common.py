#!/usr/bin/env python3
import openpyxl
import xlrd
from openpyxl import load_workbook

# """
# 按sheet number 写
# """
# def write_to_excel_cell_num(path, sheetnumber, row, column, value):
#     file = openpyxl.load_workbook(path)   #加载
#     sheetnumber = file.active   #激活
#     sheetnumber.cell(row,column,value)  #修改
#     file.save(path) #保存


"""按sheet名字cell写"""
def write_to_excel_cell(path, sheetname, row, column, value):
    file = openpyxl.load_workbook(path)   #加载
    work_sheet = file[sheetname]
    work_sheet.cell(row,column,value)  #修改
    file.save(path) #保存



""" row write """
def row_write_to_excel(path, sheetname, row, column_start, rowlist):
    for i in range(len(rowlist)):
        write_to_excel_cell(path, sheetname, row, column_start + i, rowlist[i])


""" column write """
def column_write_to_excel(path, sheetname, column, row_start, columnlist):
    for i in range(len(columnlist)):
        write_to_excel_cell(path, sheetname, row_start + i, column, columnlist[i])


"""按sheetname cell读，如果读到公式，将返回公式表达式"""
def read_from_excel_cell(path, sheetname, row, column):
    file = openpyxl.load_workbook(path)   #加载
    # file = xlrd.open_workbook(path)   #加载
    work_sheet = file[sheetname]
    ret = work_sheet.cell(row, column).value
    file.save(path) #保存
    return ret


"""如果读到公式，则返回为公式的值, 而不是公式的表达式"""
def data_only_read_from_excel_cell_(path, sheetname, row, column):
    file = openpyxl.load_workbook(path, data_only=True)   #加载
    # file = xlrd.open_workbook(path)   #加载
    work_sheet = file[sheetname]
    ret = work_sheet.cell(row, column).value
    file.save(path) #保存
    return ret



""" row read """
def row_read_from_excel(path, sheetname, row, column_start, num, row_val_list):
    for i in range(num):
        row_val_list.append(read_from_excel_cell(path, sheetname, row, column_start + i))
        # print(row_val_list[i])


""" column read """
def column_read_from_excel(path, sheetname, column, row_start, num, column_val_list):
    for i in range(num):
        column_val_list.append(read_from_excel_cell(path, sheetname, row_start + i, column))
        # print(column_val_list[i])



""" row swag"""
def row_swag(path, sheetname, row0, row1,column_start, num):
    row_list_0 = []
    row_list_1 = []
    row_read_from_excel(path, sheetname, row0, column_start, num, row_list_0)
    row_read_from_excel(path, sheetname, row1, column_start, num, row_list_1)

    for i in range(num):
        tmp = row_list_0[i]
        row_list_0[i] = row_list_1[i]
        row_list_1[i] = tmp

    row_write_to_excel(path, sheetname, row0, column_start, row_list_0)
    row_write_to_excel(path, sheetname, row1, column_start, row_list_1)
    return



""" column swag"""
def column_swag(path, sheetname, column0, column1, row_start, num):
    column_list_0 = []
    column_list_1 = []
    column_read_from_excel(path, sheetname, column0, row_start, num, column_list_0)
    column_read_from_excel(path, sheetname, column1, row_start, num, column_list_1)

    for i in range(num):
        tmp = column_list_0[i]
        column_list_0[i] = column_list_1[i]
        column_list_1[i] = tmp

    column_write_to_excel(path, sheetname, column0, row_start, column_list_0)
    column_write_to_excel(path, sheetname, column1, row_start, column_list_1)
    return




""" row copy"""
def row_copy(path, sheetname, row_source, row_target,column_start, num):
    row_list_tmp = []
    row_read_from_excel(path, sheetname, row_source, column_start, num, row_list_tmp)
    row_write_to_excel(path, sheetname, row_target, column_start, row_list_tmp)
    return

""" column copy"""
def column_copy(path, sheetname, column_source, column_target, row_start, num):
    column_list_tmp = []
    column_read_from_excel(path, sheetname, column_source, row_start, num, column_list_tmp)
    column_write_to_excel(path, sheetname, column_target, row_start, column_list_tmp)
    return


"""字符串查找  查找该字符串的单元格位置和内容"""
def find_string_in_excel(file_path, sheetname, target_string, resultlist):
    file = load_workbook(filename=file_path)
    sheet = file[sheetname]
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == target_string:
                resultlist.append(cell.row)
                resultlist.append(cell.column)





class SumExcel:
    def __init__(self, path, sheetname, target_c_or_r, start_of_target, num):
        self.path = path
        self.sheetname = sheetname
        self.target = target_c_or_r
        self.start = start_of_target
        self.num =num
        self.cachelist = []
        self.tmp = 0

    def return_column_sum(self):
        column_read_from_excel(self.path, self.sheetname, self.target, self.start, self.num, self.cachelist)
        for i in range(len(self.cachelist)):
            if self.cachelist[i] == None:
                self.cachelist[i] = 0  
            self.tmp += self.cachelist[i]
        print("第%d列, %d-%d行的和为：%d" % (self.target, self.start, self.num + self.start, self.tmp))
        return self.tmp

    def return_row_sum(self):
        row_read_from_excel(self.path, self.sheetname, self.target, self.start, self.num, self.cachelist)
        for i in range(len(self.cachelist)):
            if self.cachelist[i] == None:
                self.cachelist[i] = 0  
            self.tmp += self.cachelist[i]
        print("第%d行, %d-%d列的和为：%d" % (self.target, self.start, self.num + self.start, self.tmp))
        return self.tmp












