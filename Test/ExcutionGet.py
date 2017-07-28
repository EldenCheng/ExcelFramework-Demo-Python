from Common.CONST import CONST
from Common.Excel import Excel
from pathlib import Path
import shutil

import time

from xlutils.copy import copy
import xlrd, xlwt

exr = Excel(r"..\TestCaseData\Kerry_data_2c.xls")

exw = Excel(r"..\TestReport\Kerry_data_2w.xls","w")

exr.Select_Sheet_By_Name("1")

exw.Select_Sheet_By_Name("result-timestamp")

#print(exw.excelwtst.rows)

print(exr.Get_Row_Numbers())

#exw.excelwtst.write(1,0,1)

#print(exw.Get_Row_Numbers())

#DaList = exw.Get_Excution_DataSet("execute","Case No","yes")

#print(DaList)

