from Common.CONST import CONST
from Common.Excel_x import Excel
from Common.WebPage import WebPage
from Common.Report import *
from Common.Alias import *

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pathlib import Path

import time
import unittest


class TC1(unittest.TestCase):
    def __init__(self, testname, reportpath, browser="Firefox", Datasetrange="All"):

        super(TC1, self).__init__(testname)

        self.reportfilepath = reportpath
        self.Rowrange = Datasetrange
        self.browser = browser

    def setUp(self):

        self.excel = Excel()
        self.excel.Select_Sheet_By_Name("1")
        self.report = Excel(self.reportfilepath)
        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC1")
        self.stepsdirpath = self.casedirpath / Path("Steps")

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

    def test_Excute(self):

        for i in self.excel.Get_Excution_DataSet("executed"):

            try:
                self.page.Log_in(self.page, self.excel, 1, i, self.casedirpath)

                if self.page.Verify_Text(self.excel.Get_Value_By_ColName("Assertion", i),
                                         StartPageAlias_CSS['Login_UserName'],
                                         StartPageAlias_CSS['Login_UserName_expression']):
                    print("success")
                else:
                    self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
                    raise AssertionError("The element not contains the Assertion text")

            except Exception as msg:
                print(msg)

                Generate_Report(self.driver, self.excel, self.report, "fail", 1, self.casedirpath, i)

                self.driver.refresh()

            else:

                Generate_Report(self.driver, self.excel, self.report, "pass", 1, self.casedirpath, i)

                time.sleep(2)
                self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name("1")
        self.driver.quit()
        Generate_Final_Report(self.excel, self.report, self.reportfilepath, 1)
        self.report.Save_Excel()









