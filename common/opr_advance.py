import datetime
from opr_cell import *
from opr_col_row import *




class SysTime:
    def __init__(self):
        self.now = datetime.datetime.now()
        self.year = self.now.year
        self.month = self.now.month
        self.day = self.now.day
        self.hour = self.now.hour
        self.minute = self.now.minute


class SumExcel:
    def __init__(self, file, sheetname, target_c_or_r, start_of_target, num):
        self.file = file
        self.sheetname = sheetname
        self.target = target_c_or_r
        self.start = start_of_target
        self.num =num
        self.cachelist = []
        self.tmp = 0

    def return_column_sum(self):
        column_read_from_excel(self.file, self.sheetname, self.target, self.start, self.num, self.cachelist)
        for i in range(len(self.cachelist)):
            if self.cachelist[i] == None:
                self.cachelist[i] = 0  
            self.tmp += self.cachelist[i]
        # print("第%d列, %d-%d行的和为：%d" % (self.target, self.start, self.num + self.start, self.tmp))
        return self.tmp

    def return_row_sum(self):
        row_read_from_excel(self.file, self.sheetname, self.target, self.start, self.num, self.cachelist)
        for i in range(len(self.cachelist)):
            if self.cachelist[i] == None:
                self.cachelist[i] = 0  
            self.tmp += self.cachelist[i]
        # print("第%d行, %d-%d列的和为：%d" % (self.target, self.start, self.num + self.start, self.tmp))
        return self.tmp



class GetSearchInExcel:
    def __init__(self, wb, sheet_name, str ):
        self.wb = wb
        self.sheet_name = sheet_name
        self.str = str
        self.row_val_list = []
        self.column_val_list = []

    #找一个最大行，列的字符串,用于搜索    
    def get_max_row(self):
        find_string_in_excel(self.wb, self.sheet_name, self.str , self.row_val_list, self.column_val_list)
        return self.row_val_list[-1]

    def get_max_column(self):
        find_string_in_excel(self.wb, self.sheet_name, self.str , self.row_val_list, self.column_val_list)
        if len(self.column_val_list) != 0:
            return self.column_val_list[-1]
        else:
            print("提示：已经过p2处理")#已经没有最大行flag, 








