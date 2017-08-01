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
        self.report.Select_Sheet_By_Name("1")
        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC1")
        self.stepsdirpath = self.casedirpath / Path("Steps")

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

    def test_Excute(self):

        passdirpath = self.casedirpath / Path("Pass")
        faildirpath = self.casedirpath / Path("Fail")

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

                if not faildirpath.is_dir():
                    faildirpath.mkdir()

                time.sleep(2)
                snapshot = str(faildirpath.absolute()) + r"\TC1_Dataset_%s_Step_Fail.png" \
                                                         % str(int(self.excel.Get_Value_By_ColName("Data set", i)))
                Generate_Report(self.driver, self.report, "fail", 1, snapshot, i)
                # TO DO: Cut Steps captures of fail case to Fail folder

                # TO DO: Generate a Log file to place error msg

                self.driver.refresh()

            else:
                if not passdirpath.is_dir():
                    passdirpath.mkdir()

                snapshot = str(passdirpath.absolute()) + r"\TC1_pass_on_dataset_%s.png" \
                                                         % str(int(self.excel.Get_Value_By_ColName("Data set", i)))
                Generate_Report(self.driver, self.report, "fail", 1, snapshot, i)

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









