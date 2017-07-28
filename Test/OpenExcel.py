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

stwt = exw.Select_Sheet_By_Name("1")

exw.Set_Value_By_ColName(exr.Get_Value_By_ColName("ID",1),"ID",3)

exw.Save_Excel()

#print(stwt.cols)