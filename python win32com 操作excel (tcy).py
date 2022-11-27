python win32com 操作excel (tcy)

tcy23456

于 2021-08-21 08:00:53 发布

2941
 收藏 39
分类专栏： python
版权

python
专栏收录该内容
168 篇文章14 订阅
订阅专栏
本篇主要讲述win32com操作excel的读写的基本语法及用途实例。
并在easyExcel类的基础上封装了一个简单的excel VBA python操作。（90%变更）

特点：
1）能够多个工作薄多个工作表操作（不必显示打开），这种方式功能多但有点麻烦
2）你也可以用激活的工作薄工作表操作，这种方式较简单
   当前工作薄工作表range,cell属性已经封装，你可以直接采用比较直观 如
   current_wb ='file.xlsx'  == active_wb('file.xlsx')
   current_ws = 'sheet1'    == active_ws('sheet1')
3）代码主要供数据操作，没有涉及图标；如需要你可以运行宏（有实例）或完善相关代码
4）提供一个A1 R1C1字符串数字相互转换函数

备注：
1）所有处理经过异常检测
2）使用前将所有要操作的excel 放到当前工作目录下；
   如不想放置需要指明当前操作excel文件的默认路径self.wb_path='...'
   注意：所有操作的excel基于该路径

缺点：
运行比较慢（对比其他excel库如openpyxl是在内存中操作）但功能强大

其他：类经过测试（附测试代码）有建议或发现错误请指明。谢谢

使用注意事项：
尽量不要用self.get_cell遍历数据，效率太低，很耗时间用 self.get_range
第一部分：基本函数
1.appliation启动 Excel

    import win32com.client as win32
    from win32com.client import Dispatch

    xls = Dispatch('Excel.Application')#打开excel操作环境
    xls = win32com.client.DispatchEx('Excel.Application')

    #后台运行, 不显示, 不警告
    xls.Visible = False         # true打开excel程序界面
    xls.DisplayAlerts =False    # 禁止弹窗-不显示警告信息
    # xls.EnableEvents = False  # 禁用事件
2.workbooks

2.1.说明：
1）工作簿索引从1开始 Workbooks or WorkBooks都可
2）Microsoft对象模型共同特征—对集合依赖。集合看作是列表和字典之间的交叉；
    通过数字索引(括号或方括号)或命名字符串键访问(必须使用括号)
    ws = xlBook.Sheets(1).Name  基于集合索引从1开始 'Sheet1'
    ws = xlBook.Sheets[1].Name  基于真实位置        'Sheet2'
3）关键字参数
    Python和Excel都支持关键字参数。调用只需提供所需的参数；关键字大小写必须完全正确。
    Microsoft通常对除Filename之外所有内容都使用大小写混合

    WorkBook.SaveAs(Filename, FileFormat, Password, WriteResPassword, ReadOnlyRecommended,
     CreateBackup, AddToMru, TextCodePage, TextVisualLayout)
    调用：xlBook.SaveAs(Filename='C:\\temp\\mysheet.xls')

2.2.访问：
    wb = xlApp.ActiveWorkbook
    wb = xlApp.Workbooks(1)
    wb = xlApp.Workbooks("Book1")    新打开的工作簿（实质上取用工作簿中Name属性的值）
    wb = xlApp.Workbooks('xxx.xlsx') 文件必须已经打开（不能是全路径）

2.3.属性：
    wb_n = xls.Workbooks.Count
    wb_path = xls.Workbooks(i).Path
    wb_name = xls.Workbooks(i).Name
    wb.Checkcompatibility = False

2.4.方法：
    xls.Workbooks.Add().Name='book1'      # 创建新的工作簿
    wb = xls.Workbooks.Open('./xxx.xls')  # 打开excel文件
    wb = xls.Workbooks.Open(fullPath, ReadOnly = False)
    wb = xls.Workbooks.Open(filepath,UpdateLinks=3,ReadOnly=False,Format = None, Password=passWords)

    wb.save()                             # 保存当前工作簿
    wb.SaveAs('xxx.xls')                  # 将工作簿另存为
    xls.Workbooks(name_idx).Close(SaveChanges=0) # 关闭当前打开文件,不保存文件

    xls.Workbooks(i).Activate()

    wb.Quit()
    xls.Application.Quit()
    xls.Quit()                            # 关闭excel操作环境
    del xls

    wb.Application.Run(VBA)               #宏
    wb.RunAutoMacros(2)                   #1:打开宏，2:禁用宏
3.sheets

3.1.说明：

3.2.访问：
    ws = xlApp.ActiveSheet
    ws = xlApp.ActiveWorkbook.ActiveSheet
    ws = xlApp.Workbooks(1).Sheets(1) == wb.Sheets(1)
    ws = xlApp.Workbooks("Book1").Sheets("Sheet1") == wb.Sheets("Sheet1")

3.3.属性：
    wb.Worksheets.Count              # 获取工作表个数
    wb.Worksheets(2).Name = 'Details'# 工作簿更名
    sht_names = [ws.Name for ws in xlbook.Worksheets]
    sht_data =[sheetObj.UsedRange.Value for sheetObj in wb.Worksheets]

3.4.方法：
    wb.Worksheets.Add().Name = 'sheet1'           # 新建工作表sheet1
    ws = xls.Sheets(1)=== wb.Worksheets('Sheet1') # 选择工作表-默认Sheet1
    ws.Activate()                                 # 激活当前工作表
    ws.Shapes.AddPicture(picturename, 1, 1, Left, Top, Width, Height)#添加图片

    ws = wb.Worksheets(1).Copy(None, sheets(1))   #copy 工作簿
    shts = self.wb.Worksheets
    shts(1).Copy(None,shts(1))
4.ranges/cells

4.1.说明：

4.2.访问：
    #一个单元格
    wr = ws.Cells(1,1) == xlApp.ActiveSheet.Cells(1,1)
    wr = ws.Cells(1,1) == xlApp.ActiveWorkbook.ActiveSheet.Cells(1,1)
    wr = ws.Cells(1,1) == xlApp.Workbooks("Book1").Sheets("Sheet1").Cells(1,1)
    wr = wb.Sheets(1).Cells(1,1) == xlApp.Workbooks(1).Sheets(1).Cells(1,1).Value
    wr = wb.Sheets(1).Cells(1,1) == xlApp.Workbooks(1).Sheets(1).Cells(1,1).Value

    #多个单元格
    wr = ws.Range("B5:C10")
    wr = ws.Range(ws.Cells(2,2), ws.Cells(3,8))

4.3.属性：
    # 获取Excel Data的范围
    n_row = ws.UsedRange.Rows.Count                      # 获取使用区域的行数
    n_col = ws.UsedRange.Columns.Count
    list(ws.Range(ws.Cells(1,2),ws.Cells(row,col)).Value)# 获取所有数据

    # 向Excel单元格中写入数据
    wr.Value = 1
    wr.Value = list

    ws.Cells(2,1).Value = 3
    ws.Cells(1,1).Value = None                           # Cells(row,col)
    ws.Cells(11, 5).offset(3,2).Value =1

    ws.Range('D' + str(10)).value = 3.14                 # 第十行
    ws.Range('A1').Value = 'Win32com On Excel'

    ws.Range('A2:F2').Value = list('abcDeF')
    ws.Range(ws.Cells(3,1),ws.Cells(3,6)).Value = tuple(range(6))

4.4.方法：
    ws.UsedRange.Copy()              # 复制
    ws.UsedRange.Clear()             # 清除内容
    ws.UsedRange.ClearContents()     # 对当前使用区域清除内容
第二部分：pyExcel类封装
#!/usr/bin/env python
# -*- coding: utf-8 -*-

# *****************************************************************
# *****************************************************************
# @Project:ha_image
# @Module:aa2
# @Use:

# @Author:       tcy
# @Emial:        3615693665@qq.com
# @Date:         2021/8/16  8:14
# @Version:     1.0
# @Last Modified time:
# @City:           China Shanghai Songjiang Xiaokunshan
# @Company:  TCY Machinery Manufacturing Priority
# *****************************************************************
# *****************************************************************
from win32com.client import Dispatch
import win32com
import xlsxwriter

import os,sys,shutil,typing
from typing import List, Tuple, Dict
from datetime import datetime
import pandas as pd

from xls_constants import constants


str_int = typing.TypeVar('str_int', int, str)
str_int_obj = typing.TypeVar('str_int_obj', int, str,win32com.client.CDispatch)

class PyExcel:
    """
    用途：更易使用Excel 实用程序。可在多个工作簿上操作
    说明：类主要面对数据操作，未涉及图形；有简单的格式操作
               其他复杂操作可运行宏（本类支持）
    操作：
        方法1：使用激活工作簿工作表工作（简单但不够灵活）
        方法2：指定工作簿工作表操作（麻烦但灵活）
    备注：数据保存不直接保存关闭保存副本
    """

    def __init__(self,wb_path=''):
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = False                                             # true打开excel程序界面
        self.xlApp.DisplayAlerts = False                                  # 禁止弹窗-不显示警告信息
        self.wb_path = wb_path if wb_path else os.getcwd() # 设置当前操作工作簿路径

        self.__init__xlsformat__()

    def __del__(self):
        self.xlApp.Quit()
        del self.xlApp

    def __init__xlsformat__(self):
        self.xlsformat = pd.Series(dtype=object)
        self.xlsformat.Font_Size = 15              # 字体大小
        self.xlsformat.Font_Bold = True           # 是否黑体
        self.xlsformat.Interior_ColorIndex = 3   # 表格背景
        self.xlsformat.Borders_LineStyle = constants.xlDouble
        self.xlsformat.RowHeight = 30              # 行高
        self.xlsformat.HorizontalAlignment = constants.xlLeft  # -4131 水平居中xlCenter=-4108
        self.xlsformat.VerticalAlignment = constants.xlTop      # -4160

    @property
    def current_wb(self) :
        if (not self.wb_count): self.xlApp.WorkBooks.Add()
        return self.xlApp.ActiveWorkBook

    @current_wb.setter
    def current_wb(self, wb_name: str_int_obj):  # 设置当前工作簿
        '''工作簿打开可输入文件名或全名或索引
          工作簿未打开必须输入全名或在设置的当前路径self.wb_path下的文件名'''
        wb = self.__get_wb__(wb_name)
        wb.Activate()

    @property
    def wb_count(self) -> int:
        """返回当期工作簿数量 """
        try:
            return self.xlApp.Workbooks.Count
        except Exception:
            return 0

    @property
    def wb_name(self) -> str:
        """返回当期工作簿名称 """
        try:
            return self.xlApp.ActiveWorkbook.Name
        except Exception:
            return ''

    @property
    def current_ws(self) :
        if (not self.wb_count): self.xlApp.WorkBooks.Add()
        return self.xlApp.ActiveWorkbook.ActiveSheet

    @current_ws.setter
    def current_ws(self,ws_name:str_int):
        wb = self.xlApp.ActiveWorkBook
        ws = self.__get_ws__(ws_name,wb)
        ws.Activate()

    @property
    def current_ws_count(self) -> int:
        if not self.wb_count: return 0
        return self.xlApp.ActiveWorkbook.WorkSheets.Count

    @property
    def current_ws_name(self)->str:
        if not self.wb_count:return ''
        return self.xlApp.ActiveWorkbook.ActiveSheet.Name

    @property
    def current_ws_names(self):
        if (not self.wb_count): return {}
        wb = self.xlApp.ActiveWorkbook
        ws = wb.Sheets
        d = {ws(j + 1).Name: (j + 1) for j in range(ws.Count)}
        return {wb.Name:d}

    @property
    def current_usedrange(self):
        if (not self.wb_count): raise Exception('not workbook open.')
        return self.xlApp.ActiveWorkbook.ActiveSheet.UsedRange

    @property
    def current_sheet_size(self) -> Tuple[int,int]:
        '''获取使用区域的行列数'''
        if not self.wb_count:return 0,0
        used_range = self.xlApp.ActiveWorkbook.ActiveSheet.UsedRange
        return used_range.Rows.Count,used_range.Columns.Count

    def get_sheet_size(self,ws_name:str_int=None,wb_name:str_int_obj=None) -> Tuple[int,int]:
        '''获取使用区域的行列数'''
        wb = self.__get_wb__(wb_name)
        ws = self.__get_ws__(ws_name)
        used_range = ws.UsedRange
        return used_range.Rows.Count,used_range.Columns.Count

    @property
    def wb_names(self)->dict:
        '''返回当期工作簿名'''
        n = self.wb_count
        if (not n): return {}
        return {self.xlApp.WorkBooks(i + 1).Name:(i+1) for i in range(n)}

    @property
    def ws_names(self)->dict:
        '''  返回当期工作表名 '''
        return self.__ws_names_base__()
    def __ws_names_base__(self) -> dict:  # 返回当期工作表名
        n = self.wb_count
        if (not n):return {}
        wb_names, ws_names = {}, {}

        for i in range(n):
            wb = self.xlApp.WorkBooks(i + 1)
            ws = self.xlApp.WorkBooks(i + 1).Worksheets
            tmp ={ws(j+1).Name:(j+1)   for j in range(ws.Count )}# 获取工作表个数
            ws_names[wb.Name]=tmp

        return ws_names

    def __get_dirname__(self, wb_name: str) -> str:
        if (not isinstance(wb_name, str)):return ''
        basename = os.path.basename(wb_name)
        if not basename:return ''
        return self.wb_path + '\\' + basename

    def __get_basename__(self, wb_name: str) -> str:
        if (not isinstance(wb_name,str)):return ''
        return os.path.basename(wb_name)

    def __get_name__(self,wb_name:str):
        dirname  = self.__get_dirname__(wb_name)
        basename = self.__get_basename__(wb_name)
        return basename,dirname

    def __get_wb_index_name__(self,index:str_int_obj)->str:
        if isinstance(index,str):
            basename = self.__get_basename__(index)
            if basename:return basename
        elif isinstance(index,int):
            if index  in self.wb_names.values():
                return self.xlApp.WorkBooks(index).Name
        elif isinstance(index,win32com.client.CDispatch):
            basename = index.Name
            if basename and basename in self.wb_names.keys():
                return basename
        raise Exception('param err.')

    def __get_wb_base__(self,wb_name:str_int):
        if isinstance(wb_name,str):
            basename, dirname = self.__get_name__(wb_name)
            if basename in self.wb_names.keys():
                return self.xlApp.WorkBooks(basename)
            if  self.exists_file(dirname) :
                self.xlApp.WorkBooks.Open(dirname)
                return self.xlApp.WorkBooks(basename)

        elif isinstance(wb_name,int):
            if wb_name in self.wb_names.values():
                return self.xlApp.WorkBooks(wb_name)

        raise Exception('param err.')

    def __get_wb__(self,wb_name:str_int_obj=None):
        if not wb_name:
            return self.xlApp.ActiveWorkBook
        elif isinstance(wb_name,(str,int)):
            return self.__get_wb_base__(wb_name)
        elif  isinstance(wb_name, win32com.client.CDispatch):
            return wb_name
        else:
            raise Exception('param err.')

    def __get_ws_base__(self, ws_name: str_int,wb) :
        ws_names = self.ws_names[wb.Name]
        b1 = isinstance(ws_name, str) and (ws_name in ws_names.keys())
        b2 = isinstance(ws_name, int) and (ws_name in ws_names.values())
        if b1 or b2: return wb.Sheets(ws_name)
        raise Exception('not exists work sheet name.')

    def __get_ws__(self, sht_name:str_int_obj=None,wb_name:str_int_obj = None):
        wb = self.__get_wb__(wb_name)
        if not sht_name:
            return wb.ActiveSheet
        elif isinstance(sht_name,(str,int)):
            return self.__get_ws_base__(sht_name, wb)
        elif  isinstance(sht_name, win32com.client.CDispatch):
            return sht_name
        else:
            raise Exception('not exists work sheet name.')
     def is_open_wb(self,wb_name:str_int_obj)->bool:
        if (not self.wb_count): return False
        elif isinstance(wb_name,str):
            basename = self.__get_basename__(wb_name)
            return basename in self.wb_names.keys()
        elif isinstance(wb_name,int):
            return wb_name in self.wb_names.values()
        elif isinstance(wb_name,win32com.client.CDispatch):
            try:
                return wb_name.Name in self.wb_names.keys()
            except Exception:
                return False
        else:
            return False

    def is_empty_wb(self)->bool:
        return self.wb_count ==0

    def open(self,wb_name:str='')->None:#可打开不存在的文件
        '''存在工作簿名打开，为空打开空工作簿，不存在则打开并重命名'''
        basename, dirname = self.__get_name__(wb_name)
        if not wb_name:
            self.xlApp.WorkBooks.Add()
        elif basename in self.wb_names.keys():
            self.xlApp.WorkBooks(basename).Activate()
        elif self.exists_file(dirname):
            self.xlApp.Workbooks.Open(dirname)
        else:
            self.xlApp.WorkBooks.Add()
            self.xlApp.ActiveWorkbook.SaveAs(dirname)

    def __save_all_wb__(self):
        for i in range(self.wb_count):
            self.xlApp.WorkBooks(i+1).Save()

    def save(self, wbname: str_int = None,save_all=False)->None:
        ''' 存在工作簿则保存该工作薄忽略save_all参数
            不存在工作簿save_all=True保存所有工作簿，false保存当前工作簿
        '''
        if self.is_empty_wb():
            return
        elif save_all:
            self.__save_all_wb__()
        elif wbname:
            is_open = self.is_open_wb(wbname)
            wb = self.__get_wb__(wbname)
            wb.Save()
            if not is_open:wb.Close()
        else:
            self.xlApp.ActiveWorkBook.Save()

    def save_as(self,new_wbname:str, old_wbname: str_int_obj =None,file_cover= True)->None:
        '''
        用途：
            old_wbname为空保存当前工作簿；
            不存在工作簿退出；
            存在工作簿则保存该工作簿
        警告：
            new_wbname存在则覆盖该工作簿并不警告
        '''

        if not old_wbname and self.is_empty_wb():return
        is_open = self.is_open_wb(old_wbname)
        wb = self.__get_wb__(old_wbname)

        if wb.Name == self.__get_basename__(new_wbname):
            wb.Save()
        else:
            new_basename = self.__get_basename__(new_wbname)
            new_dirname = self.__get_dirname__(new_wbname)

            if file_cover and new_basename in self.wb_names.keys():
                self.xlApp.WorkBooks(new_basename).Close()
            elif not file_cover and self.exists_file(new_dirname):
                raise Exception(' save as file exists!')

            wb.SaveAs(new_dirname)
            if not is_open: wb.Close()


    def exist_dir_file(self,fileorpath: str) -> bool:
        return os.path.exists(fileorpath)  # 能够判断文件和文件夹是否存在

    def exists_file(self,file: str) -> bool:#文件存在判断
        return os.path.isfile(file)

    def del_file(self, file: str) -> bool:  # 删除单文件不删除目录
        os.remove(file)

    def del_all_dirfile(self, file: str) :  # 删除文件夹及内容-空目录、有内容的目录都可以删
        shutil.rmtree(file)

    def __get_backup_filename__(self,file:str = '',no = 0):
        if not file:
            dirname = self.wb_path+'\\backup.xlsx'
        else:
            dirname = self.__get_dirname__(file)

        if not dirname:raise Exception('file name err.')
        time = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
        lst = os.path.splitext(file)
        file = '%s_backup_%s_%s%s'%(lst[0],no,time,lst[1])
        return file.replace('backup_backup_','backup_')

    def __save_backup_base__(self,wb,wb_name:str,no = 0):
        backup_dirname = self.__get_backup_filename__(wb_name, no)
        backup_basename = self.__get_basename__(backup_dirname)

        if backup_basename in self.wb_names.keys():
             self.xlApp.WorkBooks(backup_basename).Close()
        wb.SaveAs(backup_dirname)

    def __save_backup__(self,wb,wb_name:str_int_obj,no = 0):
        basename = self.__get_wb_index_name__(wb_name)
        self.__save_backup_base__(wb,basename,no)

    def __save_backup_all_(self):
        for i in range(self.wb_count):
            wb = self.xlApp.WorkBooks(i + 1)
            dirname = self.__get_backup_filename__(no = i)
            basename = os.path.basename(dirname)

            if basename in self.wb_names.keys():
                self.xlApp.WorkBooks(basename).Close()
            wb.SaveAs(dirname)

    def close(self, wb_name: str_int_obj = None,save=False,save_backup= False):
        '''
        save 关闭工作簿是否保存
        save_backup 关闭文件不保存时True保存备份文件，False不备份'''
        if self.is_empty_wb() :return
        if wb_name and not self.is_open_wb(wb_name):return
        wb = self.__get_wb__(wb_name)

        if wb_name and not save and save_backup:
            self.__save_backup__(wb, wb_name, 0)
        elif not wb_name and save:
            self.__save_all_wb__()
        elif not wb_name and not save and save_backup:
            self.__save_backup_all_()

        if wb_name:
            wb.Close(SaveChanges=save)
        else:
            self.xlApp.WorkBooks.Close()


    def activate_wb(self, wb_name:str_int_obj): #名字(带后缀wb.add不加后缀)或索引
        self.__get_wb__(wb_name).Activate()

    def exists_sheet(self,sht_name:str_int,wb_name:str_int_obj=None):
        wb = self.__get_wb__(wb_name)
        if isinstance(sht_name,str):
            return sht_name in self.ws_names[wb.Name].keys()
        elif isinstance(sht_name, int):
            return sht_name in self.ws_names[wb.Name].values()
        else:
            return False

    def add_sheet(self, sht_name:str=None,wb_name: str_int_obj = None):
        '''    新建工作表sheet1,存在表退出 '''
        wb = self.__get_wb__(wb_name)
        if (sht_name):
            if sht_name in self.ws_names[wb.Name]:return
            wb.Worksheets.Add().Name = sht_name
        else:
            wb.Worksheets.Add()

    def del_sheet(self, sht_name: str_int_obj, wb_name: str_int_obj = None):  # 删除工作表
        if not self.exists_sheet(sht_name, wb_name): return
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Delete()

    def __rename_copy_sheet_name__(self, wb,new_sht_name:str=''):
        old_sht_name = wb.ActiveSheet.Name
        if old_sht_name.find('(') == -1 : return
        if new_sht_name:
            wb.ActiveSheet.Name = new_sht_name
        else:
            wb.ActiveSheet.Name = old_sht_name.replace('(','_').replace(')','')

    def copy_sheet(self, old_sht_name:str_int,new_sht_name:str=None,wb_name:str_int_obj=None):
        '''  工作表拷贝，旧名不存在退出，重名报错 '''
        wb = self.__get_wb__(wb_name)
        if not self.exists_sheet(old_sht_name, wb_name): return
        ws = self.__get_ws__(old_sht_name, wb)
        ws.Copy(None, ws)
        self.__rename_copy_sheet_name__(wb,  new_sht_name)

    def rename_sheet(self, sht_oldname:str_int, sht_newname:str,wb_name: str_int_obj = None):
        '''  工作表更名，旧名不存在退出，重名报错 '''
        wb = self.__get_wb__(wb_name)
        if not self.exists_sheet(sht_oldname,wb_name) : return
        ws = self.__get_ws__(sht_oldname,wb)
        ws.Name =  sht_newname

    def activate_sheet(self, sht_name:str_int, wb_name: str_int_obj = None):
        '''  工作表激活，不存在退出 '''
        if not self.exists_sheet(sht_name, wb_name): return
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Activate()
  def a1_to_r1c1(self, a1_style_cell_str:str,is_one=True):
        '''
         xlsxwriter.utility.xl_cell_to_rowcol(cell_str)
        用途：将A1表示法中的单元格引用转换为零索引行和列
        参数：cell_str(string) – A1
        样式字符串, 绝对或相对
        返回：(row, col)的整数元组
        实例：
        (row, col) = xl_cell_to_rowcol('A1')  # (0, 0)
        (row, col) = xl_cell_to_rowcol('B1')  # (0, 1)
        (row, col) = xl_cell_to_rowcol('C2')  # (1, 2)
        (row, col) = xl_cell_to_rowcol('$C2')  # (1, 2)
        (row, col) = xl_cell_to_rowcol('C$2')  # (1, 2)
        (row, col) = xl_cell_to_rowcol('$C$2')  # (1, 2)
        '''
        rst =xlsxwriter.utility.xl_cell_to_rowcol(a1_style_cell_str.upper())
        if is_one:
            return tuple([v+1 for v in rst])
        else:
            return rst

    def a1_to_int(self, a1_style_cell_str:str,is_one=True):
        '''
         将a1样式的列名转换为对应第几列
         is_one =True,‘A'为第1列;False 为第0 列
        '''
        a1_style_cell_str = a1_style_cell_str.upper()+'1'
        rst =xlsxwriter.utility.xl_cell_to_rowcol(a1_style_cell_str.upper())
        return rst[1]+1 if is_one else rst[1]

    def r1c1_to_a1(self, col:int, col_abs:bool=False,is_one=True):
        '''
         xlsxwriter.utility.xl_col_to_name(col[, col_abs] )
        用途：将零索引列单元格引用转换为字符串
        参数：
        col(int) – 单元格列
        col_abs(bool) – 使列成为绝对值的可选标志.
        返回：列样式字符串
        实例：
        column = xl_col_to_name(0)  # A
        column = xl_col_to_name(1)  # B
        column = xl_col_to_name(702)  # AAA
        column = xl_col_to_name(0, False)  # A
        column = xl_col_to_name(0, True)  # $A
        column = xl_col_to_name(1, True)  # $B
        '''
        col = col-1 if is_one else col
        return xlsxwriter.utility.xl_col_to_name(col, col_abs )


# 设置和获取单元格：可指定工作表名或索引、行和列
    def get_cell(self, row, col,sht_name:str_int=None, wb_name: str_int = None):  # 获取一个单元格的值
        ws = self.__get_ws__(sht_name, wb_name)
        return ws.Cells(row, col).Value

    def set_cell(self, row, col, value,sht_name:str_int=None,wb_name: str_int = None):  # 设置一个单元格的值
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Cells(row, col).Value = value

    def get_range(self, row1:str_int, col1:int=None, row2:int=None, col2:int=None,ws_name:str_int=None,wb_name: str_int = None):  # 返回一个二维数组（即元组元组）
        ws = self.__get_ws__(ws_name, wb_name)
        if isinstance(row1,str):
            return ws.Range(row1).Value
        else:
            if(not row2 or not col2):
                return self.__get_contiguous_range__( row1, col1,ws_name,wb_name)
            else:
                return ws.Range(ws.Cells(row1, col1), ws.Cells(row2, col2)).Value

     # 插入一个数据块-只需指定第一个单元格；无需计算行数
    def set_range(self, data,left_col:str_int, top_row:int=None, ws_name:str_int=None,wb_name: str_int = None):  # 从给定位置开始插入一个二维数组
        ws = self.__get_ws__(ws_name, wb_name)
        if isinstance(left_col,str): top_row,left_col = self.a1_to_r1c1(left_col)
        bottomRow = top_row + len(data) - 1
        rightCol = left_col + len(data[0]) - 1
        ws.Range( ws.Cells(top_row, left_col),ws.Cells(bottomRow, rightCol) ).Value = data

    # 获取数据：不知有多少行列时，从起点开始向下和向右扫描直到遇到空白
    def __get_contiguous_range__(self, row, col,sht_name:str_int=None,wb_name: str_int = None):
        ws = self.__get_ws__(sht_name, wb_name)

        # 找到底行
        bottom = row
        while ws.Cells(bottom + 1, col).Value not in [None,'']:
            bottom = bottom + 1

        # right column
        right = col
        while ws.Cells(row, right + 1).Value not in [None, '']:
            right = right + 1

        return ws.Range(ws.Cells(row, col), ws.Cells(bottom, right)).Value

    def set_cell_format(self,  row, col,sht_name:str_int=None,wb_name: str_int = None):  # 设置单元格的数据
        "set value of one cell"
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Cells(row, col).Font.Size = self.xlsformat.Font_Size    # 字体大小
        ws.Cells(row, col).Font.Bold = self.xlsformat.Font_Bold    # 是否黑体
        ws.Cells(row, col).Name = "Arial"                                        # 字体类型
        ws.Cells(row, col).Interior.ColorIndex = self.xlsformat.Interior_ColorIndex      # 表格背景
        ws.Cells(row, col).Borders.LineStyle = self.xlsformat.Borders_LineStyle
        ws.Cells(row, col).BorderAround(1, 4)                                                              # 表格边框
        ws.Rows(3).RowHeight = self.xlsformat.RowHeight                                        # 行高
        ws.Cells(row, col).HorizontalAlignment = self.xlsformat.HorizontalAlignment  # 水平居中xlCenter
        ws.Cells(row, col).VerticalAlignment = self.xlsformat.VerticalAlignment

    def insert_row(self,  pos_row,number =1,sht_name:str_int=None,wb_name: str_int = None):
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Rows(pos_row).Insert(number)

    def insert_col(self, pos_row, number=1,sht_name: str_int = None, wb_name: str_int = None):
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Columns(pos_row).Insert(number)

    def del_row(self, row,sht_name:str_int=None,wb_name: str_int = None):
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Rows(row).Delete()  # 删除行

    def del_col(self, row, sht_name: str_int = None, wb_name: str_int = None):
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Columns(row).Delete()  # 删除列

    def add_picture(self, pictureName, Left, Top, Width, Height,sht_name:str_int=None,wb_name: str_int = None):
        "Insert a picture in sheet"
        ws = self.__get_ws__(sht_name, wb_name)
        ws.Shapes.AddPicture(pictureName, 1, 1, Left, Top, Width, Height)
# 返回数组常包含Unicode字符串或COM日期。根据需要在每列基础上转换它们（有时不需要）
    def convert_str_date(self, aMatrix):
        # 转换所有 unicode 字符串和时间
        newmatrix = []
        for row in aMatrix:
            newrow = []
            for cell in row:
                if type(cell) is self.xlApp.UnicodeType:
                    newrow.append(str(cell))
                elif type(cell) is self.xlApp.TimeType:
                    newrow.append(int(cell))
                else:
                    newrow.append(cell)
            newmatrix.append(tuple(newrow))
        return newmatrix
# 运行宏
    def run_macro(self,macro_name,xlsm_file):
        xlApp = win32com.client.DispatchEx("Excel.Application")
        xlApp.Visible = True
        xlApp.DisplayAlerts = 0
        wb = xlApp.Workbooks.Open(xlsm_file)
        wb.Application.Run(macro_name) #宏
        wb.Save()
        wb.Close()
        xlApp.quit()

    #py中实现和excel宏相同的功能
    def __test_run_macro_py__(self, xlsm_file, results, sheet_name):
        xlApp = win32com.client.DispatchEx("Excel.Application")
        xlApp.Visible = True  # 进程可见-测试用
        xlApp.DisplayAlerts = 0
        wb = xlApp.Workbooks.Open(xlsm_file, False)

        ws = wb.Worksheets(sheet_name)  # 找到要操作的sheet
        for i in range(len(results)):                  # sheet数据,日期列格式为date
            for j in range(len(results[0])):
                ws.Cells(i + 2, j + 1).Value = results[i][j]  # 从第二行第二列开始填入数据
        print("write finish.")

        wb.Save()
        wb.Close(SaveChanges=0)
        xlApp.quit()  # 关闭excel操作环境

    # 操作
    def test_run_macro(self):
        results=((5,2,1),(7,2,7),(4,0,3))    #这是要填入的数据，一个3*3元组
        xlsm_file=r"C:\Users\Administrator\Desktop\工作簿1.xlsm" #文件
        sheet_name= "sheet1"
        macro_name = "macro_1"
        self.__test_run_macro_py__(xlsm_file,results,sheet_name)
        self.run_macro(macro_name,xlsm_file)


if __name__ == '__main__':
    a = PyExcel()
    print(a.__dict__)
第三部分：测试代码：
#!/usr/bin/env python
# -*- coding: utf-8 -*-

# *****************************************************************
# *****************************************************************
# @Project:ha_image
# @Module:run_aa2
# @Use:

# @Author:       tcy
# @Date:         2021/8/18  7:07
# @Version:     1.0
# @Last Modified time:
# @City:           China Shanghai Songjiang Xiaokunshan
# @Company:  TCY Machinery Manufacturing Priority
# *****************************************************************
# *****************************************************************
from aa2 import PyExcel
import win32com


def test_attribute(self):
    file1 = 'file1.xlsx'
    file2 = 'file2.xlsx'
    file_notexists = self.wb_path + '\\file_notexists.xlsx'
    self.close_wb_backup_file = False

    # test wb_count:
    self.close()
    assert self.wb_count == 0
    self.open(file1)
    assert self.wb_count == 1
    self.open(file2)
    assert self.wb_count == 2

    if self.exists_file(file_notexists):
        self.del_file(file_notexists)
    self.open(file_notexists)
    assert self.wb_count == 3

    self.close(file1)
    assert self.wb_count == 2
    self.close()

    print(self.wb_names)
    assert self.wb_count == 0

    # test self.current_wb
    self.close()
    # assert self.current_wb.Name== '工作簿1'
    print('1.1 current_wb = ', self.current_wb.Name)
    self.open(file1)
    assert self.current_wb.Name == 'file1.xlsx'

    self.current_wb = file2
    assert self.current_wb.Name == 'file2.xlsx'
    self.close(file1)
    assert self.__get_wb__(file2).Name == 'file2.xlsx'

    if self.exists_file(file_notexists):
        self.del_file(file_notexists)
    try:
        self.__get_wb__(file_notexists).Name
    except Exception as e:
        print('1.2 err:not exists file.create file.')
        self.open(file_notexists)
        assert self.__get_wb__(file_notexists).Name == 'file_notexists.xlsx'

    # test sefl.current_ws
    self.close()

    self.open(file1)
    self.add_sheet('ws1')
    self.add_sheet('ws2')
    self.add_sheet('ws3')
    print('2.1 current ws,wb =', self.current_ws.Name, self.current_wb.Name)

    self.current_ws = 'ws1'
    assert self.current_ws.Name == 'ws1'

    self.current_ws = 'ws2'
    assert self.current_ws.Name == 'ws2'

    print('2.2 current wb count =', self.current_ws_count)
    print('2.3 current wb name =', self.current_ws_names)
    print('2.4 wb names =', self.wb_names)
    print('2.5 ws names =', self.ws_names)

    # test range
    self.current_ws.Cells(1, 3).Value = 100
    self.current_ws.Cells(3, 1).Value = 200
    print('3.1 usedrange_count =', self.current_usedrange_count)
    print('3.2 usedrange.Value =', self.current_usedrange.Value)

    self.current_ws.Cells(2, 4).Value = 300
    self.current_ws.Cells(4, 2).Value = 400
    print('3.3 usedrange_count =', self.current_usedrange_count)
    print('3.4 usedrange.Value =', self.current_usedrange.Value)

    self.close(file1)
def test_open(self, file1, file2, file3, file4):
    self.close()
    if self.exists_file(file4):
        self.del_file(file4)

    print('1.1wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.open()
    print('1.2wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.open(file1)
    print('1.3 wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.open(file2)
    self.open(file3)
    print('1.4 wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.open(file4)
    print('1.5 wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.close()
    print('1.6 wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

def set_range_value(self,ws,row,col,value):
    if not isinstance(ws.Cells(row,col).Value, float):
        ws.Cells(row,col).Value = value
    else:
        ws.Cells(row,col).Value = ws.Cells(row,col).Value + value

def open_wb(self,file1,file2,file3=None):
    self.open(file1)
    self.open(file2)
    if file3:self.open(file3)
def test_close_write_wb(self,file,save,backup,value):
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    x0 = ws.Cells(1, 1).Value
    set_range_value(self, ws, 1, 1, value)
    x1 = ws.Cells(1, 1).Value

    self.close(file, save, backup)
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    x2 = ws.Cells(1, 1).Value
    return x0,x1,x2

def __test_close_write_wb1__(self,file1,file2,save,backup,v1,v2):
    x11,x12,x13 = test_close_write_wb(self, file1, save, backup, v1)
    x21, x22, x23 = test_close_write_wb(self, file2, save, backup, v2)
    return x11,x12,x13,x21,x22,x23

def __test_close_write_wb2__(self,save,backup,v1):
    lst = []
    i=0
    for k,v in self.ws_names.items():
        i=i+1
        x1,x2,x3= test_close_write_wb(self, k, save, backup, v1+i*100)
        lst.append((x1,x2,x3))
    return lst

def __test_close_view__(self,file1,file2,save,backup,v1,v2,no):
    self.close()
    x11, x12, x13, x21, x22, x23 = __test_close_write_wb1__(self, file1, file2, save=save, backup=backup, v1=v1, v2=v2)
    print('%s.1.x11= %s; x12= %s; x13= %s'%(no,x11,x12,x13))
    print('%s.2.x21= %s; x22= %s; x23= %s' % (no,x21, x22, x23))
    print()
    self.close()

def __test_close_view2__(self,save,backup,v1,no):
    lst = __test_close_write_wb2__(self, save=save, backup=backup, v1=v1)
    for i,v in enumerate(lst):
        print('%s.%s.value=  %s'%(no,i,v))
    print()
    self.close()

def test_close(self, file1, file2, file3):
    '''
    close(self, wb_name: str_int = None,save=False,save_backup= False)
    '''
    self.close()
    open_wb(self, file1, file2,file3)
    print('0.1.close=', self.wb_names, self.ws_names, self.wb_count)
    print()
    self.close()

    __test_close_view__(self, file1, file2, save=False, backup=False, v1=1, v2=10, no='1')
    __test_close_view__(self, file1, file2, save=False, backup=True, v1=1, v2=10, no='2')
    __test_close_view__(self, file1, file2, save=True, backup=False, v1=1, v2=10, no='3')
    __test_close_view__(self, file1, file2, save=True, backup=True, v1=1, v2=10, no='4')

    open_wb(self, file1, file2)
    __test_close_view2__(self, save=False, backup=False, v1=10, no='1')

    open_wb(self, file1, file2)
    __test_close_view2__(self, save=True, backup=False, v1=10, no='2')

    open_wb(self, file1, file2)
    __test_close_view2__(self, save=False, backup=True, v1=10, no='3')

    open_wb(self, file1, file2)
    __test_close_view2__(self, save=True, backup=True, v1=10, no='4')
def write_cells(self,file,value):
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    x0 = ws.Cells(1, 1).Value
    set_range_value(self, ws, 1, 1, value)
    x1 = ws.Cells(1, 1).Value
    return x1

def get_cells(self,file):
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    return ws.Cells(1, 1).Value

def test_save_write_wb(self,file,save_all,value):
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    x0 = ws.Cells(1, 1).Value
    set_range_value(self, ws, 1, 1, value)
    x1 = ws.Cells(1, 1).Value

    self.save(file,save_all)
    self.close(file)
    self.open(file)
    self.current_wb = file
    ws = self.current_wb.Sheets(1)
    x2 = ws.Cells(1, 1).Value
    return x0,x1,x2

def test_save_write_wb2(self,save_all,v1):
    i,lst = 0,[]
    for k in self.ws_names.keys():
        i=i+1
        x1,x2,x3= test_save_write_wb(self, k,  save_all,v1+i*100)
        lst.append((x1,x2,x3))
    return lst

def test_save1(self,file1,file2,save_all=True,value=100):
    ''' def save(self, wbname: str_int = None,save_all=False) '''
    # test 1
    self.close()

    open_wb(self,file1,file2)
    x0=get_cells(self, file2)
    write_cells(self,file2, -1000)
    x1=get_cells(self, file2)
    self.close(file2)
    lst = test_save_write_wb(self,file1,save_all=save_all,value=value)
    self.open(file2)
    x2= get_cells(self, file2)
    print('1.1 file2 cells=',(x0,x1,x2))
    print('1.2.lst=',lst)
    '''
    1.1 file2 cells= (-4030.0, -5030.0, -4030.0)
    1.2.lst= (2880.0, 2980.0, 2980.0)
    '''
def test_save2(self,file1,file2,save_all=False,v1=100,v2=200):
    ''' def save(self, wbname: str_int = None,save_all=False) '''
    # test 3
    open_wb(self, file1, file2)
    print('1.1 file cells=', (get_cells(self, file1), get_cells(self, file2)))
    write_cells(self, file1, v1)
    write_cells(self, file2, v2)
    print('1.2 file cells=', (get_cells(self, file1), get_cells(self, file2)))
    self.save(save_all=save_all)
    self.close()
    open_wb(self, file1, file2)
    print('1.3 file cells=', (get_cells(self, file1), get_cells(self, file2)))
    self.close()
    ''' 1.1 file cells= (3680.0, -3030.0)
        1.2 file cells= (3780.0, -2830.0)
        1.3 file cells= (3680.0, -2830.0)
    '''
def test_save(self,file1,file2):
    ''' def save(self, wbname: str_int = None,save_all=False) '''
    # test 1
    test_save1(self,file1,file2,save_all=True,value=100)
    self.close()

    # test 2
    test_save1(self, file1, file2, save_all=False, value=100)
    self.close()

    # test 3
    test_save2(self,file1,file2,save_all=False,v1=100,v2=200)
    # test 4
    test_save2(self, file1, file2, save_all=True, v1=100, v2=200)
def test_save_as(self):
    '''
    def save_as_wb(self, new_wbname:str,old_wbname: str_int = None)->None:
    '''
    file1 = 'file1.xlsx'
    file2 = 'file2.xlsx'

    #test1:
    file_notexists = 'file_notexists.xlsx'
    if(self.exists_file(file_notexists)):self.del_file(file_notexists)

    self.open(file1)
    self.open(file2)
    print('1.1.wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.current_wb = file1
    print('1.2.current wb=',self.current_wb.Name)

    file1_save = 'file1_save.xlsx'
    if (self.exists_file(file1_save)): self.del_file(file1_save)
    self.save_as(file1_save)
    print('1.3.file 1 save file1_save: \n wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.save_as(file1_save)
    print('1.4.file1 save file1_save 覆盖:\n wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)#覆盖
    print()

    #test2:
    self.close()
    self.open(file1)
    self.open(file2)

    file1_save = 'file1_save.xlsx'
    if (self.exists_file(file1_save)): self.del_file(file1_save)
    file2_save = 'file2_save.xlsx'
    if (self.exists_file(file2_save)): self.del_file(file2_save)
    print('2.1.wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.save_as(old_wbname=file1, new_wbname=file1_save)
    print('2.2.wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)
    self.save_as(old_wbname=file2, new_wbname=file2_save)
    print('2.3.wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)
    print()
def test_active_wb(self):
    '''
    def activate_wb(self, wb_name:str_int_obj):
    '''
    file1 = 'file1.xlsx'
    file2 = self.wb_path+'\\file2.xlsx'

    #test1:
    file_notexists = 'file_notexists.xlsx'
    if(self.exists_file(file_notexists)):self.del_file(file_notexists)

    self.open(file1)
    self.open(file2)
    print('1.1.wb_names=', self.wb_names, ' ; ws_names=', self.ws_names)

    self.current_wb = file1
    print('1.2.current wb=',self.current_wb.Name)
    self.current_wb = file2
    print('1.3.current wb=', self.current_wb.Name)

    self.activate_wb(file1)
    print('1.4.current wb=', self.current_wb.Name)
    self.activate_wb(file2)
    print('1.5.current wb=', self.current_wb.Name)

def test_workbook(self):
    file1 = 'file1.xlsx'
    file2 = self.wb_path + '\\file2.xlsx'
    file3 = 'file3.xlsx'
    file_notexists = 'file_notexists.xlsx'

    test_open(self, file1, file2, file3, file_notexists)
    test_close(self, file1, file2, file3)
    test_save(self, file1, file2)
    test_save_as(self)
    test_active_wb(self)
def test_ws(self):
    self.close()
    file3 = 'file3.xlsx'
    self.open(file3)

    print('1.1.ws=',self.current_ws_names)

    self.add_sheet('ws1')
    print('1.2.ws=', self.current_ws_names)

    self.del_sheet('ws1')
    print('1.3.exist ws=', self.exists_sheet('ws1'))

    self.add_sheet('ws1')
    self.add_sheet('ws2')
    self.add_sheet('ws3')
    self.add_sheet('ws3')

    print('1.4.ws=', self.current_ws_names,' ;exist ws=', self.exists_sheet('ws1'))

    self.del_sheet('ws11')
    self.del_sheet('ws12')
    print('2.1.ws=', self.current_ws_names)

    self.rename_sheet('ws1','ws11')
    print('2.2.ws=', self.current_ws_names)
    self.rename_sheet('notexist', 'ws12')
    print('2.3.ws=', self.current_ws_names)

    self.del_sheet('ws11')
    self.del_sheet('ws12')
    print('2.4.ws=', self.current_ws_names)

    # test copy sheet
    self.copy_sheet('ws3')
    print('3.1.ws=', self.current_ws_names)
    self.copy_sheet('ws3','new_ws3')
    print('3.2.ws=', self.current_ws_names)

    # test active sheet:
    print('4.1 active sheet name=',self.current_ws.Name)
    self.activate_sheet('ws2')
    print('4.2 active sheet name=', self.current_ws.Name)
    self.activate_sheet('notexists')
    print('4.3 active sheet name=', self.current_ws.Name)
    self.activate_sheet('ws3')
    print('4.4 active sheet name=', self.current_ws.Name)


def test_worksheet(self):
    test_ws(self)
def test_range(self):
    file1 = 'file1.xlsx'
    self.open(file1)
    print('==>',self.xlApp.ActiveCell.FormulaR1C1)
    r1= self.current_ws.Cells(1,2)
    r1.Value=200
    print('==>', self.xlApp.ActiveCell.FormulaR1C1)
    self.save()

    print('==>', r1.FormulaR1C1,self.current_ws_name,r1.Value)

def test_r1c1_to_a1(self):
    assert self.r1c1_to_a1(0+1)  == 'A'
    assert self.r1c1_to_a1(1+1)  == 'B'
    assert self.r1c1_to_a1(702+1)  == 'AAA'

    assert self.r1c1_to_a1(0+1, False)  == 'A'
    assert self.r1c1_to_a1(0+1, True)  == '$A'
    assert self.r1c1_to_a1(1+1, True)  == '$B'

    assert self.r1c1_to_a1(0,is_one=False) == 'A'
    assert self.r1c1_to_a1(1,is_one=False) == 'B'
    assert self.r1c1_to_a1(702,is_one=False) == 'AAA'

    assert self.r1c1_to_a1(0, False,is_one=False) == 'A'
    assert self.r1c1_to_a1(0, True,is_one=False) == '$A'
    assert self.r1c1_to_a1(1, True,is_one=False) == '$B'

def test_range1(self):
    self.open('工作21-8-21.xlsx')
    ws_name1 = 'ha_func'
    ws_name2 = 'Sheet1'

    row1, _ = self.get_sheet_size(ws_name1)
    row2, _ = self.get_sheet_size(ws_name2)

    col1 = self.a1_to_int('G')
    col2 = self.a1_to_int('A')
    s1 = 'G1:%s' % row1
    s2 = 'A1:%s' % row2

    t1 = self.get_range(1, col1, row1, col1, ws_name1)
    t2 = self.get_range(1, col2, row2, col2, ws_name2)
    lst1 = [v[0] for v in t1]
    lst2 = [v[0] for v in t2]
    print(lst1)
    print(lst2)

    print('===>', self.get_range('A1:G3', ws_name=ws_name1))
    data = [['Tom', 11], ['Bob', 22]]
    self.set_range(data, 2, 2, ws_name='Sheet2')
    print('===>', self.get_range(2, 2, 3, 3, ws_name='Sheet2'))  # (('Tom', 11.0), ('Bob', 22.0))
    print('===>', self.get_range('b2:c3', ws_name='Sheet2'))
    print('===>', self.get_range(2, 2, ws_name='Sheet2'))
def run(self):
    test_attribute(self)
    test_workbook(self)
    test_worksheet(self)
    test_range(self)
    test_a1_to_r1c1(self)
    test_r1c1_to_a1(self)
    test_range1(self)


if __name__ == '__main__':
    self = PyExcel()
    run(self)
备注：excel常量


xls_constants.zip
————————————————
版权声明：本文为CSDN博主「tcy23456」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/tcy23456/article/details/119834642