#encoding=utf-8
import xlrd,xlwt
from xlutils.copy import copy

class use_xlrd():
    def __init__(self,filename):
        self.filename = filename
        self.sheets = []
        self.get_sheets()

    def get_sheets(self):
        try:
            self.data = xlrd.open_workbook(self.filename,formatting_info=True)#格式
            self.sheets = self.data.sheets()
        except IOError,e:
            self.sheets = None
            print '打开文件夹失败:',e

    def get_sheets_num(self):
        if self.sheets != None:
            return len(self.sheets)
        else:
            return 0

    def get_sheet(self,sheet_index):
        if self.sheets == None:
            return None
        try:
            sheet = self.sheets[sheet_index]
            return sheet
        except IndexError,e:
            print 'sheet_index error:',e
            return None
    
    def get_sheet_rows(self,sheet_index):
        sheet = self.get_sheet(sheet_index)
        if sheet != None:
            return sheet.nrows
        else:
            return 0

    def get_sheet_cols(self,sheet_index):
        sheet = self.get_sheet(sheet_index)
        if sheet != None:
            return sheet.ncols
        else:
            return 0

class use_xlwt():
    def __init__(self,filename):
        self.filename = filename
        self.read_xl_obj = use_xlrd(filename)
        if self.read_xl_obj.sheets == None:
            self.workbook = xlwt.Workbook(encoding = 'utf-8')
        else:
            self.workbook = copy(self.read_xl_obj.data)
        self.dict_sheet_row = {}
        self.max_sheet_rows = 200
        #encoding
        self.workbook.encoding = 'utf-8'
        self.workbook.__dict__['_Workbook__sst'].encoding = 'utf-8'

    def add_default_sheet(self):
        sheet_num = self.read_xl_obj.get_sheets_num() 
        if sheet_num == 0:
            sheetname = '第' + str(sheet_num) + '页'
            self.workbook.add_sheet(sheetname)
        else:
            nrows = self.read_xl_obj.get_sheet_rows(sheet_num - 1)
            if nrows == self.max_sheet_rows:
                sheet_num += 1
                sheetname = '第' + str(sheet_num) + '页'
                self.workbook.add_sheet(sheetname)

    #获取当前xlsx中可用的sheet的rows(程序中做处理,每个sheet最多200行)
    def get_min_rows(self):
        sheet_num = self.read_xl_obj.get_sheets_num()
        for i in range(sheet_num):
            nrows = self.read_xl_obj.get_sheet_rows(i)
            if nrows < self.max_sheet_rows:
                self.dict_sheet_row[i] = nrows
                break
        return self.dict_sheet_row


    #写入excel中的内容 
    def write_row_content(self,dict_info):
        print 'write_row_content start'
        min_row = 0
        active_sheet_index = 0
        col = 0

        print self.dict_sheet_row
        for key in self.dict_sheet_row:
            min_row = self.dict_sheet_row[key]
            active_sheet_index = key

        print 'min_row:',min_row
        print 'befor:',self.dict_sheet_row
        sheet = self.workbook.get_sheet(active_sheet_index)

        
        #写入内容
        for eachKW in dict_info:
            '''
            if eachKW == 'img_f' or eachKW == 'img_v': 
                style = easyxf('font: underline single')
                sheet.write(min_row,col,xlwt.Formula('HYPERLINK("%s";"%s")' % dict_info[eachKW],eachKW),style)
                #sheet.insert_bitmap(dict_info[eachKW],min_row,col)
                col += 1
            else:
                sheet.write(min_row,col,dict_info[eachKW])
                col += 1
            '''
            if eachKW == 'label':
                sheet.write(min_row,0,dict_info[eachKW])
            elif eachKW == 'score':
                sheet.write(min_row,1,dict_info[eachKW])
            elif eachKW == 'time':
                sheet.write(min_row,2,dict_info[eachKW])
            elif eachKW == 'img_f':
                style = xlwt.easyxf('font: underline single')
                sheet.write(min_row,3,xlwt.Formula('HYPERLINK("%s";"%s")' % (dict_info[eachKW],eachKW)),style)
            elif eachKW == 'img_v':
                style = xlwt.easyxf('font: underline single')
                sheet.write(min_row,4,xlwt.Formula('HYPERLINK("%s";"%s")' % (dict_info[eachKW],eachKW)),style)
        #行数++
        min_row += 1

        min_sheet_num = 0#sheet number
        if len(self.dict_sheet_row) == 0:
            self.dict_sheet_row[min_sheet_num] = min_row

        for key in self.dict_sheet_row:
            self.dict_sheet_row[key] = min_row
            min_sheet_num = key

        print 'after:',self.dict_sheet_row

        if min_row >= self.max_sheet_rows:#一个sheet最多200行,超过了新建一个sheet
            min_sheet_num += 1
            sheet_name = '第' + str(min_sheet_num) + '页'
            self.workbook.add_sheet(sheet_name)
            min_row = 0
            self.dict_sheet_row[min_sheet_num] = min_row

        print 'wirte_row_content finsh'

    def save_file(self):
        self.workbook.save(self.filename)
