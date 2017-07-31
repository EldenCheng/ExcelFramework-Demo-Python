from Common.CONST import CONST
from Common.Excel import Excel
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
        self.report = Excel(self.reportfilepath,"w")
        self.report.Select_Sheet_By_Name("result-timestamp")
        self.wtrowindex =self.report.Get_Row_Numbers()
        self.report.Select_Sheet_By_Name("1")
        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC1")
        self.stepsdirpath = casedirpath / Path("Steps")

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

    def test_Excute(self):

        passdirpath = casedirpath / Path("Pass")
        faildirpath = casedirpath / Path("Fail")

        for i in self.excel.Get_Excution_DataSet("executed"):

            try:
                self.page.Input(self.excel.Get_Value_By_ColName("ID", i, self.casedirpath), LoginPageAlias_CSS['ID_Field'])
                self.driver.save_screenshot(str(self.stepsdirpath) + "/TC1_DataSet_%s_Step1.png" % str(i))
                self.page.Input(self.excel.Get_Value_By_ColName("PW", i, self.casedirpath), LoginPageAlias_CSS['PW_Field'])
                self.driver.save_screenshot(str(self.stepsdirpath) + "/TC1_DataSet_%s_Step2.png" % str(i))
                self.page.ButtonClick(LoginPageAlias_CSS['Login_Btn'])
                time.sleep(1)

                self.page.Verification_Code(self.excel.Get_Value_By_ColName("PW",i))

                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Start"))

                if self.page.Verify_Text(self.excel.Get_Value_By_ColName("Assertion", i),
                                         StartPageAlias_CSS['Login_UserName'], StartPageAlias_CSS['Login_UserName_expression']):
                    print("success")
                else:
                    self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
                    raise AssertionError("The element not contains the Assertion text")

            except Exception as msg:
                print(msg)

                if not faildirpath.is_dir():
                    faildirpath.mkdir()

                time.sleep(2)
                snapshot = str(faildirpath.absolute()) + r"\TC1_fail_on_dataset_%s.png" % str(
                    int(self.excel.Get_Value_By_ColName("Data set", i)))
                Generate_Report("fail", snapshot, i)

                self.driver.refresh()

            else:
                if not passdirpath.is_dir():
                    passdirpath.mkdir()

                snapshot = str(passdirpath.absolute()) + r"\TC1_pass_on_dataset_%s.png" % str(int(self.excel.Get_Value_By_ColName("Data set",i)))
                Generate_Report("pass", snapshot, i)

                time.sleep(2)
                self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name("1")
        self.driver.quit()
        Generate_Final_Report()
        self.report.Save_Excel()









