from Common.CONST_j import CONST
from Common.Excel_x import Excel
from Common.WebPage import WebPage
from Common.Report import *
from Common.Alias import *

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from pathlib import Path

import time
import unittest


class TC4(unittest.TestCase):
    def __init__(self, testname, reportpath, browser="Firefox", Datasetrange="All"):
        self.caseno = 4
        exec("super(TC%d, self).__init__(testname)" % self.caseno)

        self.reportfilepath = reportpath
        self.Rowrange = Datasetrange
        self.browser = browser

    def setUp(self):
        self.excel = Excel(CONST.EXCELPATH)
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

                bro = self.excel.Get_Value_By_ColName("Browser", i)
                if bro is not None:
                    self.browser = bro

                self.page = WebPage()
                self.driver = self.page.Start_Up(CONST.URL, self.browser)

                current_step = self.page.Log_in(self.excel, self.caseno, i, self.casedirpath)

                self.page.LabelClick(StartPageAlias_CSS['Menu_Invitation'])

                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Invitation"))

                self.driver.save_screenshot(str(self.stepsdirpath) + "/TC%d_DataSet_%d_Step_%d.png"
                                            % (self.caseno, i, current_step + 1))

                self.page.Input(self.excel.Get_Value_By_ColName("Contact Person", i, self.casedirpath),
                                InvitationPageAlias_CSS['Contact Person'])
                self.driver.save_screenshot(str(self.stepsdirpath) + "/TC%d_DataSet_%d_Step_%d.png"
                                            % (self.caseno, i, current_step + 2))
                self.page.Input(self.excel.Get_Value_By_ColName("Phone Number", i, self.casedirpath),
                                InvitationPageAlias_CSS['Phone Number'])
                self.driver.save_screenshot(str(self.stepsdirpath) + "/TC%d_DataSet_%d_Step_%d.png"
                                            % (self.caseno, i, current_step + 3))

                self.page.ButtonClick(InvitationPageAlias_CSS['Submit_Btn'])

                time.sleep(1)

                #ActionChains(self.driver).move_to_element(self.driver.find_element(By.CSS_SELECTOR,"td#txfld_companyEmail-sideErrorCell"))

                #self.driver.find_element(By.CSS_SELECTOR, "td#txfld_companyEmail-sideErrorCell").click()

                #time.sleep(1)

                if self.driver.find_element(By.CSS_SELECTOR,
                                            InvitationPageAlias_CSS['Email_Error_icon']).is_displayed():
                    print("success")
                else:
                    raise AssertionError("The element was not shown!")

            except Exception as msg:
                print(msg)

                Generate_Report(self.driver, self.excel, self.report, "fail", self.caseno, self.casedirpath, i)

                self.driver.quit()

            else:

                Generate_Report(self.driver, self.excel, self.report, "pass", self.caseno, self.casedirpath, i)

                self.driver.quit()

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        # print(self.excel)
        Generate_Final_Report(self.excel, self.report,  self.caseno)
        self.report.Save_Excel()
