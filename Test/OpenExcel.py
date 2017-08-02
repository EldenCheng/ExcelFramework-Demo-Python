from Common.CONST import CONST
from Common.Excel import Excel
from pathlib import Path
import shutil

import time

from xlutils.copy import copy
import xlrd, xlwt

exr = Excel(r"..\TestCaseData\Kerry_data_2c.xls")

exw = Excel(r"..\TestReport\Kerry_data_2w.xls","w")

strd = exr.Select_Sheet_By_Name("1")

stwt = exw.Select_Sheet_By_Name("result-timestamp")

#path = r"file:///D:\百度云同步盘\项目\自动化测试\Kerry\KerryDemo\TestReport\2017-08-02_142534\TC4\Pass\TC4_Dataset_1_Step_Pass.png"

#sceencapturepath = Path(path[8:])
#sceencapturename = list(sceencapturepath.parts)[-1]
#print((sceencapturepath))
#print(sceencapturename)
#link = 'HYPERLINK("%s";"%s")' % (str(sceencapturepath), str(sceencapturename))

link = r'HYPERLINK("D:\百度云同步盘\项目\自动化测试\Kerry\KerryDemo\TestReport\2017-08-02_145238\TC4\Pass\TC4_Dataset_1_Step_Pass.png";"TC4_Dataset_1_Step_Pass.png")'

print(link)

exw.Set_Value_By_ColName(xlwt.Formula(link),"Screen capture",2)

exw.Save_Excel()

#print(stwt.cols)