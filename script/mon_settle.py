from common.opr_advance import *



def del_none_row_and_col(wb, sheet_name):
    work_sheet = wb[sheet_name]
    to_del_row_list = []
    to_del_col_list = []

    row_val_list = []
    column_val_list = []
    maxrow = GetSearchInExcel(wb, sheet_name, "合计")
    maxrow_num = maxrow.get_max_row()
    maxcol = GetSearchInExcel(wb, sheet_name, "1-12月合计")
    if None == maxcol.get_max_column():
        return -1
    else:
        maxcol_num = maxcol.get_max_column() +3  #由表格获得,+3是多删除一点空白列
    # print("最大列：", maxcol_num)


    find_string_in_excel(wb, sheet_name, "本月预算", row_val_list, column_val_list)
    # print("基地址(本月预算)：行_%d,列_%d" %(row_val_list[0],column_val_list[0]))

    find_string_in_excel(wb, sheet_name, "累计费用", row_val_list, column_val_list)
    # print("基地址(累计费用)：行_%d,列_%d" %(row_val_list[1],column_val_list[1]))


    #获取坐标
    base_row_start = row_val_list[0]
    base_column_start = column_val_list[0]
    base_row_end = row_val_list[1]
    base_column_end = column_val_list[1]

    if row_val_list[0] != row_val_list[1]:
        print("基地址开始行，获取错误！！！！")
        return
    
    column_size = base_column_end - base_column_start + 1 
    # print("开始行：%d, 最大行：%d" %(base_row_start, maxrow_num ))


    # print("************************************************")
    # print("****************开始删除表格********************")

    for row_num in range(base_row_start, maxrow_num):

        flag_all_none = 0
        for i in range(base_column_start, (base_column_end +1)):
            # cell = work_sheet.cell(row=row_num, column=i)
            tmp = work_sheet.cell(row_num, i).value
            # print(tmp)
            if tmp == None or tmp == 0:
                flag_all_none += 1
                
        if flag_all_none == column_size:
            to_del_row_list.append(row_num)
        # del_row_list.append(row_num)
            # tmp1 = work_sheet.cell(row, base_column_start - 1).value
            # if tmp1 == "小计":
            #     tmp2 = work_sheet.cell(row, base_column_start - 2).value
            #     if tmp2 != None:
            #         work_sheet.delete_rows(row - flag_has_been_del)
            #         flag_has_been_del += 1
            #     else:
            #         work_sheet.delete_rows(row - flag_has_been_del)
            #         flag_has_been_del += 1

            # elif tmp1 != "小计":
            #     tmp3 = work_sheet.cell(row, base_column_start - 2).value
            #     if tmp3 != None:
            #         row_copy(wb, sheet_name, row, row+1, (base_column_start - 2), 1)
            #         work_sheet.delete_rows(row - flag_has_been_del)
            #         flag_has_been_del += 1
            #     else:
            #         row_copy(wb, sheet_name, row, row+1, (base_column_start - 2), 1)
            #         work_sheet.delete_rows(row - flag_has_been_del)
            #         flag_has_been_del += 1                
    # print(to_del_row_list)
    del_row_in_excel(wb, sheet_name, to_del_row_list)



    del_col_start = base_column_end + 2
    del_col_end = maxcol_num 
    for i in range(del_col_start, del_col_end + 1):
        to_del_col_list.append(i)

    del_column_in_excel(wb, sheet_name, to_del_col_list) 
    # print("删除空白行，列成功！") 



def write_sum_to_xiaoji_row(wb, sheet_name):
    list_ABC= [ {'1': 'A', 
                 '2': 'B', 
                 '3': 'C',
                 '4': 'D',
                 '5': 'E',
                 '6': 'F',
                 '7': 'G',
                 '8': 'H',
                 '9': 'I'}]

    work_sheet = wb[sheet_name]

    row_base_addr_list = []
    column_base_addr_list = []
    xioaji_and_tushu_row_list = []

    maxrow = GetSearchInExcel(wb, sheet_name, "合计")
    maxrow_num = maxrow.get_max_row()


    find_string_in_excel(wb, sheet_name, "本月预算", row_base_addr_list, column_base_addr_list)
    # print("基地址(本月预算)：行_%d,列_%d" %(row_base_addr_list[0],column_base_addr_list[0]))
    col_base_start = column_base_addr_list[0]

    for column in range(col_base_start,col_base_start + 5):
    # column = 3
        row_start = row_base_addr_list[0] +1
        row_end = None
        for i in range(row_base_addr_list[0] +1, maxrow_num):
            tmp1 = work_sheet.cell(i, column_base_addr_list[0] - 1).value
            if tmp1 == "小计":
                # print(i)
                row_end = i - 1
                # print("[%d - %d]"%(row_start,row_end))
                value = "=sum(" + str(list_ABC[0][str(column)]) + str(row_start) + ":" + str(list_ABC[0][str(column)]) + str(row_end) +")"
                # print(value)
                work_sheet.cell(i,column,value)
                row_start = i + 1
            if tmp1 == "图书杂志费":
                row_start += 1

    #获取小记所在行,
    for i in range(row_base_addr_list[0] +1, maxrow_num):
        tmp1 = work_sheet.cell(i, column_base_addr_list[0] - 1).value
        if tmp1 == "小计":
            xioaji_and_tushu_row_list.append(i)
        if tmp1 == "图书杂志费":
            xioaji_and_tushu_row_list.append(i)   
    # print(xioaji_and_tushu_row_list)

    # #写入合计
    for column in range(col_base_start,col_base_start + 5 ):
        # print("column:",column)
    # column = 3
        str_to_write = "="
        for i in range(len(xioaji_and_tushu_row_list)):
            str_to_write += "+" + str(list_ABC[0][str(column)]) + str(xioaji_and_tushu_row_list[i])
        # print(str_to_write)
        work_sheet.cell(maxrow_num,column,str_to_write)
    print("合计成功！")


def copy_data_from_src(wb, sheet_name):
    work_sheet = wb[sheet_name]
    systime = SysTime()
    month = systime.month
    now = systime.now
    # print("现在是%d月，进行%d月表格处理" %(month, (int(month) - 1)))

    maxrow = GetSearchInExcel(wb, sheet_name, "合计")
    maxrow_num = maxrow.get_max_row()

    row_base_addr_list = []
    column_base_addr_list = []
    month_row_list = []
    month_column_list = []

    row_of_month_real_money = 0
    col_of_month_real_money = 0
    row_of_month_plan_money = 0
    col_of_month_plan_money = 0

    month_str = (str(month - 1) + "月")

    find_string_in_excel(wb, sheet_name, "本月预算", row_base_addr_list, column_base_addr_list)
    find_string_in_excel(wb, sheet_name, month_str, month_row_list, month_column_list)
    n = 4
    for i in range(len(month_row_list)):
        tmp1 = work_sheet.cell(month_row_list[i], month_column_list[i]).value


        if tmp1 == (str(month - 1) + "月") and month_row_list[i] ==row_base_addr_list[0]:
            # print("%d月 ：行_%d,列_%d" %(month, month_row_list[i], month_column_list[i]))
            row_of_month_real_money = month_row_list[i]
            col_of_month_real_money = month_column_list[i]
            row_of_month_plan_money = month_row_list[i]
            col_of_month_plan_money = month_column_list[i]
            # print(month_row_list[i]+1)

            # print(maxrow_num - month_row_list[i]-1)
           
            column_copy(wb, sheet_name, col_of_month_real_money, n, month_row_list[i]+1, maxrow_num - month_row_list[i]-1)
            n = n-1




def have_processed_check1(file, sheet_name):
    wb = openpyxl.load_workbook(file,data_only=True)   #加载

    target_str_row = []
    target_str_row.append(None)
    target_str_col = [] 

    find_string_in_excel(wb, sheet_name, "1-12月合计", target_str_row, target_str_col)
    # print("处理标记行：",target_str_row[0])
    if target_str_row[0] == None:
        print("#已为：发送版，无需再次处理")
        # print("%s已完成：透视表 》》》》》底稿，无需再次处理",file)
        return -1
    wb.save(file)







def unmerge_my_cell(wb, sheet_name):

    work_sheet = wb[sheet_name]
    row_base_addr_list = []
    column_base_addr_list = []
    xioaji_and_tushu_row_list = []

    maxrow = GetSearchInExcel(wb, sheet_name, "合计")
    maxrow_num = maxrow.get_max_row()

    find_string_in_excel(wb, sheet_name, "本月预算", row_base_addr_list, column_base_addr_list)
    # print("基地址(本月预算)：行_%d,列_%d" %(row_base_addr_list[0],column_base_addr_list[0]))
    col_base_start = column_base_addr_list[0]


    row_start = row_base_addr_list[0] +1
    row_end = None
    for i in range(row_base_addr_list[0] +1, maxrow_num):
        # print(row_start)
        tmp1 = work_sheet.cell(i, column_base_addr_list[0] - 1).value
        if tmp1 == None:
            break
        if tmp1 == "小计":
            if i == maxrow_num:
                break
            if row_start == i  and row_start != 6:
                pass

            else:
                # print("[%d-%d]"%(row_start, i))
                unmerge_cells_value(wb, sheetname=sheet_name, range_string=None, start_row=row_start, start_column=1, end_row=i, end_column = 1)

                value = read_from_excel_cell(wb, sheet_name, row_start, 1) #获取大类费用名称
                # print("第%d行的值:%s" %(row_start, value))
                for ummerge_i in range(row_start + 1,i+1):
                    # print("unmerge_i=",ummerge_i)
                    work_sheet.cell(ummerge_i,1,value)
                    
            
            # print("[%d-%d]"%(row_start, i))
            row_start = i + 1
        if tmp1 == "图书杂志费":
            row_start = i + 1



        # if tmp1 == "图书杂志费":
        #     row_start += 1




def merge_my_cell(wb, sheet_name):

    work_sheet = wb[sheet_name]
    row_base_addr_list = []
    column_base_addr_list = []
    xioaji_and_tushu_row_list = []

    maxrow = GetSearchInExcel(wb, sheet_name, "合计")
    maxrow_num = maxrow.get_max_row()

    find_string_in_excel(wb, sheet_name, "本月预算", row_base_addr_list, column_base_addr_list)
    # print("基地址(本月预算)：行_%d,列_%d" %(row_base_addr_list[0],column_base_addr_list[0]))
    col_base_start = column_base_addr_list[0]
    row_start = row_base_addr_list[0] +1



    # merge_cells_value(wb, sheetname=sheet_name, range_string=None, start_row=6, start_column=1, end_row=11, end_column=1)
    # merge_cells_value(wb, sheetname=sheet_name, range_string=None, start_row=12, start_column=1, end_row=14, end_column=1)
    for i in range(row_base_addr_list[0]+1, maxrow_num):
        # print("开始：%d结束：%d"%(row_base_addr_list[0]+1, maxrow_num))
        tmp1 = work_sheet.cell(i, column_base_addr_list[0] - 1).value
        if tmp1 == "小计":
            row_end = i
            merge_cells_value(wb, sheetname=sheet_name, range_string=None, start_row=row_start, start_column=1, end_row=row_end, end_column=1)
            row_start = i + 1
        if tmp1 == "图书杂志费":
            row_start = i + 1







def process_single_excel(file):
    # sheet_name = "安环部预算执行表"
    # tmp = have_processed_check1(file, sheet_name)
    # if tmp == -1:
    #     return
    sheet_list = []
    
    wb = openpyxl.load_workbook(file,data_only=True)   #加载
    sheet_list= wb.sheetnames
    sheet_name = sheet_list[0]

    unmerge_my_cell(wb, sheet_name)

    copy_data_from_src(wb, sheet_name)



    ret = del_none_row_and_col(wb, sheet_name)
    if ret == -1:
        return -2
    

    write_sum_to_xiaoji_row(wb, sheet_name)
    merge_my_cell(wb, sheet_name) #TODO

    unfreeze_and_unfilt_cell(wb, sheet_name)

    # # # set_color_of_sheet(wb, sheet_name)#TODO

    wb.save(file)



"""从原始表格里获取数据存放在list中"""





def get_data_from_nc(wb,sheet_name):
    work_sheet = wb[sheet_name]

    heji_row_list = []
    heji_col_list = []

    find_string_in_excel(wb, work_sheet, "总计", heji_row_list, heji_col_list)  #以 "总计" 为标识符，确定最大行和最大列
    print("总计1：行_%d,列_%d" %(heji_row_list[0],heji_col_list[0]))
    print("总计2：行_%d,列_%d" %(heji_row_list[1],heji_col_list[1]))






"""
处理文件
1.



"""

















