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


class TC7(unittest.TestCase):
    def __init__(self, testname, reportpath, browser="Firefox", Datasetrange="All"):
        self.caseno = 7
        exec("super(TC%d, self).__init__(%s)" % self.caseno, testname)

        self.reportfilepath = reportpath
        self.Rowrange = Datasetrange
        self.browser = browser

    def setUp(self):

        self.excel = Excel()
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        self.report = Excel(self.reportfilepath)
        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC%d" % self.caseno)
        self.stepsdirpath = self.casedirpath / Path("Steps")

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

    def test_Excute(self):

        for i in self.excel.Get_Excution_DataSet("executed"):

            try:
                current_step = self.page.Log_in(self.page, self.excel, self.caseno, i, self.casedirpath)

                self.page.LabelClick(StartPageAlias_CSS['Menu-PO Maintenance'])

                self.driver.save_screenshot(
                    str(self.stepsdirpath) + "/TC%d_DataSet_%d_Step%d.png" % (self.caseno, i, current_step + 1))

                self.page.VideoButtonClick(self.excel.Get_Value_By_ColName("Order Type", i),
                                           POMaintenancePageAlias_CSS['Order Type'], self.casedirpath, i, "Order Type")

                self.driver.save_screenshot(
                    str(self.stepsdirpath) + "/TC%d_DataSet_%d_Step%d.png" % (self.caseno, i, current_step + 2))

            except Exception as msg:
                print(msg)

                Generate_Report(self.driver, self.excel, self.report, "fail", self.caseno, self.casedirpath, i)

                self.driver.refresh()

            else:

                Generate_Report(self.driver, self.excel, self.report, "pass", self.caseno, self.casedirpath, i)

                time.sleep(2)
                self.page.ButtonClick(StartPageAlias_CSS['Logout_Btn'])
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        self.driver.quit()
        Generate_Final_Report(self.excel, self.report, self.caseno)
        self.report.Save_Excel()