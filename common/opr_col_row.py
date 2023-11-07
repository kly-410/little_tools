
from opr_cell import *


""" row write """
def row_write_to_excel(file, sheetname, row, column_start, rowlist):
    for i in range(len(rowlist)):
        write_to_excel_cell(file, sheetname, row, column_start + i, rowlist[i])


""" column write """
def column_write_to_excel(file, sheetname, column, row_start, columnlist):
    for i in range(len(columnlist)):
        # print(columnlist[i])
        write_to_excel_cell(file, sheetname, row_start + i, column, columnlist[i])

""" row read """
def row_read_from_excel(file, sheetname, row, column_start, num, row_val_list):
    for i in range(num):
        row_val_list.append(read_from_excel_cell(file, sheetname, row, column_start + i))
        # print(row_val_list[i])

""" column read """
def column_read_from_excel(file, sheetname, column, row_start, num, column_val_list):
    for i in range(num):
        column_val_list.append(read_from_excel_cell(file, sheetname, row_start + i, column))
        # print(column_val_list[i])

""" row swag"""
def row_swag(file, sheetname, row0, row1,column_start, num):
    row_list_0 = []
    row_list_1 = []
    row_read_from_excel(file, sheetname, row0, column_start, num, row_list_0)
    row_read_from_excel(file, sheetname, row1, column_start, num, row_list_1)

    for i in range(num):
        tmp = row_list_0[i]
        row_list_0[i] = row_list_1[i]
        row_list_1[i] = tmp

    row_write_to_excel(file, sheetname, row0, column_start, row_list_0)
    row_write_to_excel(file, sheetname, row1, column_start, row_list_1)


""" column swag"""
def column_swag(file, sheetname, column0, column1, row_start, num):
    column_list_0 = []
    column_list_1 = []
    column_read_from_excel(file, sheetname, column0, row_start, num, column_list_0)
    column_read_from_excel(file, sheetname, column1, row_start, num, column_list_1)

    for i in range(num):
        tmp = column_list_0[i]
        column_list_0[i] = column_list_1[i]
        column_list_1[i] = tmp

    column_write_to_excel(file, sheetname, column0, row_start, column_list_0)
    column_write_to_excel(file, sheetname, column1, row_start, column_list_1)

""" row copy"""
def row_copy(file, sheetname, row_source, row_target,column_start, num):
    row_list_tmp = []
    row_read_from_excel(file, sheetname, row_source, column_start, num, row_list_tmp)
    row_write_to_excel(file, sheetname, row_target, column_start, row_list_tmp)

""" column copy"""
def column_copy(file, sheetname, column_source, column_target, row_start, num):
    column_list_tmp = []
    column_read_from_excel(file, sheetname, column_source, row_start, num, column_list_tmp)
    column_write_to_excel(file, sheetname, column_target, row_start, column_list_tmp)

"""删除的行/列放进一个list,然后删除这个list里边的行/列"""
def del_row_in_excel(file, sheetname, to_del_row_list):
    work_sheet = file[sheetname]
    for i in range(len(to_del_row_list)):
        work_sheet.delete_rows((to_del_row_list[i]-i)) #删除后下边的单元格上移动

def del_column_in_excel(file, sheetname, to_del_column_list):
    work_sheet = file[sheetname]
    for i in range(len(to_del_column_list)):
        work_sheet.delete_cols((to_del_column_list[i]-i)) #删除后下边的单元格上移动

