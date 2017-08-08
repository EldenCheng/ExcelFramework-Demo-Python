from Common.CONST_j import CONST
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
    def __init__(self, testname, reportpath, browser="Chrome", Datasetrange="All"):

        self.caseno = 1
        exec("super(TC%d, self).__init__(testname)" % self.caseno)

        self.reportfilepath = reportpath
        self.Rowrange = Datasetrange
        self.browser = browser

    def setUp(self):

        self.excel = Excel()
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        self.report = Excel(self.reportfilepath)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC%d" % self.caseno)
        self.stepsdirpath = self.casedirpath / Path("Steps")

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

    def test_Excute(self):

        for i in self.excel.Get_Excution_DataSet("executed"):

            try:
                bro = self.excel.Get_Value_By_ColName("Browser", i)
                if bro is not None:
                    self.browser = bro

                self.page = WebPage()
                self.driver = self.page.Start_Up(CONST.URL, self.browser)
                if i != '':
                    self.page.Log_in(self.excel, self.caseno, i, self.casedirpath)

                    if self.page.Verify_Text(self.excel.Get_Value_By_ColName("Assertion", i),
                                             StartPageAlias_CSS['Login_UserName']):
                        print("success")
                    else:
                        raise AssertionError("The element not contains the Assertion text")

            except Exception as msg:
                print(msg)

                Generate_Report(self.driver, self.excel, self.report, "fail", self.caseno, self.casedirpath, i)
                #if self.driver.title == "Start":
                #    self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                #    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
                #elif self.driver.title == "KV Login Page":
                #    self.driver.refresh()
                self.driver.quit()
                #time.sleep(10)

            else:

                Generate_Report(self.driver, self.excel, self.report, "pass", self.caseno, self.casedirpath, i)

                #time.sleep(2)
                #self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                #WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
                self.driver.quit()
                #time.sleep(10)

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        Generate_Final_Report(self.excel, self.report, self.caseno)
        self.report.Save_Excel()









