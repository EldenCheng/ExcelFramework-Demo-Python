from Common.CONST import CONST
from pathlib import Path
import shutil
import random

import time

from xlutils.copy import copy
import xlrd, xlwt

class Excel:
    def __init__(self,path=CONST.EXCELPATH, rw="r"):

        try:
            if rw == "r":
                self.excelrdwb = xlrd.open_workbook(path,formatting_info = True)
                self.excelrdst = ''
            elif rw == "w":
                self.excelrdwb = xlrd.open_workbook(path,formatting_info = True)
                self.excelrdst = ''
                self.excelwtwb = copy(self.excelrdwb)
                self.excelwtst = ''
                self.filepath = path
        except Exception as msg:
            print(msg)

        self.rw = rw

    def Get_All_Sheets_Name(self):
        return self.excelrdwb.sheet_names()

    def Get_Sheets_Numbers(self):
        #return len(self.excelrdwb.sheet_names())
        return self.excelrdwb.nsheets

    def Select_Sheet_By_Name(self,name):
        if self.rw == "r":
            self.excelrdst = self.excelrdwb.sheet_by_name(name)
            return self.excelrdst
        elif self.rw == "w":
            self.excelrdst = self.excelrdwb.sheet_by_name(name)
            self.excelwtst = self.excelwtwb.get_sheet(self.excelwtwb.sheet_index(name))
            return self.excelwtst

    def Get_Row_Numbers(self):

        rows = self.excelrdst.nrows

        if self.rw == "r":
            for i in range (0, rows):
                if self.excelrdst.cell_value(i,0) != '':
                    if i != rows - 1 :
                        continue
                    else:
                        return i + 1
                else:
                    return i
        elif self. rw == "w":
            return len(self.excelwtst.rows)


    def Get_Col_Numbers(self, sheetpassed=''):
        if sheetpassed !='':
            sheet = sheetpassed
        else:
            sheet = self.excelrdst


        for i in range (sheet.ncols-1,0,-1):
            if sheet.cell_value(0,i) == '':
                continue
            else:
                break
        return i

    def Get_All_Values_By_ColName(self,colname,sheetpassed=''):
        if sheetpassed !='':
            sheet = sheetpassed
        else:
            sheet = self.excelrdst

        values = []
        for c in range(sheet.ncols-1,-1,-1):
            if sheet.cell_value(0,c)  == colname:
                for r in range(1,sheet.nrows):
                    if sheet.cell_value(r,c)  != '':
                        values.append(sheet.cell_value(r,c))
        return values

    def Get_Value_By_ColName(self,colname,row,path = ''):

        for c in range(self.excelrdst.ncols - 1, -1, -1):
            if self.excelrdst.cell_value(0, c) == colname:
                if row < self.excelrdst.nrows:
                    value =  self.excelrdst.cell_value(row,c)

                    if str(value).find("random") != -1 and path !="":
                        excelTemp = Excel(CONST.EXCELPATH)
                        excelTemp.Select_Sheet_By_Name("data")
                        randomlist = excelTemp.Get_All_Values_By_ColName(value)
                        randomvalue = random.choice(randomlist)
                        file = Path(path) / Path("%d_%s_%s.random" % (row,colname,str(randomvalue)))
                        file.touch()
                        #print(str(file))
                        return randomvalue
                    else:
                        return value
        else:
            print("Cannot find any value")
            return None

    def Set_Value_By_ColName(self,content,colname,row,sheetpassed=''):
        if sheetpassed !='':
            sheet = sheetpassed
        else:
            sheet = self.excelwtst

        #if colname != "Screen capture":
            for c in range(self.excelrdst.ncols - 1, -1, -1):
                if self.excelrdst.cell_value(0, c) == colname:
                    if str(content).find("HYPERLINK") == -1:
                        sheet.write(row,c,content)
                    else:
                        #print(colname)
                        #print(row)
                        #print(content)
                        #link = '=' + content
                        sheet.write(row, c, xlwt.Formula('A2+B2'))

    def Set_Sheet_Name(self,old_name,new_name):

        self.excelwtwb.get_sheet(self.excelwtwb.sheet_index(old_name)).set_name(new_name)

    def Create_New_Report(self,folderpath= r".\TestReport" ):
        otime = time.strftime("%Y-%m-%d_%H%M%S", time.localtime())
        reprotfolderpath = Path(folderpath) / Path(otime)
        if reprotfolderpath.is_dir():
            pass
        else:
            reprotfolderpath.mkdir()
            reportfilepath = folderpath + "\\" + reprotfolderpath.name +  r"\Report_" + otime + ".xls"

        shutil.copy(CONST.EXCELPATH,reportfilepath)

        if Path(reportfilepath).is_file():
            return str(reportfilepath)
        else:
            print("Cannot copy excel file")

    def Get_Excution_DataSet(self,colname,dataset_name="Data set",keyword="Not Yet"):
        DaList = []
        for c in range(self.excelrdst.ncols-1,0,-1):
            if self.excelrdst.cell_value(0,c)  == colname:
                for r in range(1,self.excelrdst.nrows):
                    if self.excelrdst.cell_value(r,c)  == keyword:
                        if self.Get_Value_By_ColName(dataset_name,r) != '':
                            #print("The value of cell(%d,%d) is %r" % (r,c, self.Get_Value_By_ColName(dataset_name,r)))
                            DaList.append(int(self.Get_Value_By_ColName(dataset_name,r)))
        return DaList

    def Save_Excel(self,path=''):
        if path == '':
            savepath = self.filepath
        else:
            savepath = path

        self.excelwtwb.save(savepath)








