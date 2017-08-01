from TestCase_Scripts.TC_1 import TC1
from TestCase_Scripts.TC_2 import TC2
from TestCase_Scripts.TC_3 import TC3
from TestCase_Scripts.TC_4 import TC4
from Common.CONST import CONST
from Common.Excel import Excel

import pathlib,time

import unittest

if __name__ == '__main__':

    suite = unittest.TestSuite()

    excel = Excel(CONST.EXCELPATH)
    reportfilepath = excel.Create_New_Report()

    test_loader = unittest.TestLoader()
    excel.Select_Sheet_By_Name("summary")
    for caseno in excel.Get_Excution_DataSet("execute","Case No","yes"):
        if excel.Get_Value_By_ColName("Script", caseno) !="to do":

            suite.addTest(eval("TC"+str(caseno))("test_Excute",reportfilepath, "Chrome"))

    result = unittest.TextTestRunner().run(suite)

    excel = Excel(reportfilepath,"w")
    #excel.Set_Sheet_Name("result-timestamp","result-%s" % (time.strftime("%Y-%m-%d_%H%M%S", time.localtime())))
    excel.Set_Sheet_Name("result-timestamp", "result")
    excel.Save_Excel()
    #sys.exit(not result.wasSuccessful())

