# coding: utf-8

from Common.CONST_j import CONST
from Common.Excel_x import Excel
from Common.Report import *

import sys
import unittest

if __name__ == '__main__':

    suite = unittest.TestSuite()

    excel = Excel(CONST.EXCELPATH)
    reportfilepath = Create_New_Report()

    test_loader = unittest.TestLoader()
    excel.Select_Sheet_By_Name("summary")
    for caseno in excel.Get_Excution_DataSet("execute","Case No","yes"):
        if excel.Get_Value_By_ColName("Script", caseno).lower() != "to do":
            exec("from TestCase_Scripts.TC_%d import TC%d" % (caseno, caseno))

            suite.addTest(eval("TC"+str(caseno))("test_Excute",reportfilepath, CONST.BROWSER))

    result = unittest.TextTestRunner().run(suite)

    excel = Excel(reportfilepath)
    #excel.Set_Sheet_Name("result-timestamp","result-%s" % (time.strftime("%Y-%m-%d_%H%M%S", time.localtime())))
    excel.Set_Sheet_Name("result-timestamp", "result")
    excel.Save_Excel()
    sys.exit(not result.wasSuccessful())

