from Common.CONST import CONST
from Common.Excel import Excel
from Common.WebPage import WebPage

from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By

from pathlib import Path

import time
import unittest


class TC4(unittest.TestCase):
    def __init__(self, testname, reportpath, browser="Firefox", Datasetrange="All"):

        super(TC4, self).__init__(testname)

        self.reportfilepath = reportpath
        self.Rowrange = Datasetrange
        self.browser = browser

    def setUp(self):

        self.excel = Excel(CONST.EXCELPATH)
        self.excel.Select_Sheet_By_Name("4")
        self.report = Excel(self.reportfilepath,"w")
        self.report.Select_Sheet_By_Name("result-timestamp")
        self.wtrowindex =self.report.Get_Row_Numbers()
        self.report.Select_Sheet_By_Name("4")
        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)


    def test_Excute(self):

        casedirpath = Path(self.reportfilepath).parent / Path("TC4")
        passdirpath = casedirpath / Path("Pass")
        faildirpath = casedirpath / Path("Fail")
        stepsdirpath = casedirpath / Path("Steps")
        if not casedirpath.is_dir():
            casedirpath.mkdir()

        if not stepsdirpath.is_dir():
            stepsdirpath.mkdir()

        for i in self.excel.Get_Excution_DataSet("executed"):
            try:

                if self.driver.title == "KV Login Page":
                    self.page.Input(self.excel.Get_Value_By_ColName("ID", i), "userId")
                    self.driver.save_screenshot(str(stepsdirpath) + "/TC4_DataSet_%s_Step1.png" % str(i))
                    self.page.Input(self.excel.Get_Value_By_ColName("PW", i), "password")
                    self.driver.save_screenshot(str(stepsdirpath) + "/TC4_DataSet_%s_Step2.png" % str(i))
                    self.page.ButtonClick(r"btn_login-btnEl")
                    time.sleep(1)

                    self.page.Verification_Code(self.excel.Get_Value_By_ColName("PW", i))

                    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Start"))

                    time.sleep(2)

                    self.page.LabelClick(r"label-1030")

                    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Invitation"))

                    self.driver.save_screenshot(str(stepsdirpath) + "/TC4_DataSet_%s_Step3.png" % str(i))

                self.page.Input(self.excel.Get_Value_By_ColName("Contact Person", i, casedirpath), "contactPerson")
                self.driver.save_screenshot(str(stepsdirpath) + "/TC4_DataSet_%s_Step4.png" % str(i))
                self.page.Input(self.excel.Get_Value_By_ColName("Phone Number", i, casedirpath), "phoneNo")
                self.driver.save_screenshot(str(stepsdirpath) + "/TC4_DataSet_%s_Step5.png" % str(i))

                self.page.ButtonClick(r"btn_submit-btnEl")

                time.sleep(1)



                #ActionChains(self.driver).move_to_element(self.driver.find_element(By.CSS_SELECTOR,"td#txfld_companyEmail-sideErrorCell"))

                #self.driver.find_element(By.CSS_SELECTOR, "td#txfld_companyEmail-sideErrorCell").click()

                #time.sleep(1)

                if self.driver.find_element(By.CSS_SELECTOR,"td#txfld_companyEmail-sideErrorCell").is_displayed():
                    print("success")
                else:
                    raise AssertionError("The element was not shown!")


            except Exception as msg:
                print(msg)

                if not faildirpath.is_dir():
                    faildirpath.mkdir()

                time.sleep(2)
                snapshot = str(faildirpath.absolute()) + r"\TC4_fail_on_dataset_%s.png" % str(
                    int(self.excel.Get_Value_By_ColName("Data set", i)))
                self.Generate_Report("fail", snapshot, i)

                self.driver.refresh()

            else:
                if not passdirpath.is_dir():
                    passdirpath.mkdir()

                snapshot = str(passdirpath.absolute()) + r"\TC4_pass_on_dataset_%s.png" % str(
                    int(self.excel.Get_Value_By_ColName("Data set", i)))
                self.Generate_Report("pass", snapshot, i)

                time.sleep(2)
                self.page.ButtonClick(r"btn_logout-btnEl")
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))

    def Generate_Report(self, pass_fail, snapshotpath, rowindex):
        self.report.Select_Sheet_By_Name("4")
        self.driver.save_screenshot(snapshotpath)
        self.report.Set_Value_By_ColName(pass_fail, "Result", rowindex)
        self.report.Set_Value_By_ColName(r"file:///" + snapshotpath, "Screen capture", rowindex)
        self.report.Set_Value_By_ColName("done", "executed", rowindex)

    def Generate_Final_Report(self):

        randomfilepath = Path(self.reportfilepath).parent / Path("TC4")
        randomfiles = list(randomfilepath.rglob('*.random'))
        #print(randomfiles)

        randomrecords = []

        for l in randomfiles:
            p = str(list(l.parts)[-1:])
            #l = str(l)
            p = p[:p.find(".random")]
            #print("p is %s" % p)
            randomrecords.append(p.split("_"))

        randomrecords = list(randomrecords)

        for rowindex in range(1, self.excel.Get_Row_Numbers()):
            self.report.Select_Sheet_By_Name("result-timestamp")
            # print(self.wtrowindex)
            self.report.Set_Value_By_ColName(4, "Case No", self.wtrowindex)
            self.report.Set_Value_By_ColName(rowindex, "Data set", self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("Expected result", rowindex),
                                             "Expected result",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("Result", rowindex), "Result",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("Description", rowindex),
                                             "Description",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("Screen capture", rowindex),
                                             "Screen capture",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("executed", rowindex), "executed",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("Assertion", rowindex), "Assertion",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("ID", rowindex), "ID",
                                             self.wtrowindex)
            self.report.Set_Value_By_ColName(self.excel.Get_Value_By_ColName("PW", rowindex), "PW",
                                             self.wtrowindex)

            if randomrecords != '':
                for i in range(len(randomrecords)):
                    if int(randomrecords[i][0][-1:]) == rowindex:

                        self.report.Set_Value_By_ColName("%s(%s)" % (self.excel.Get_Value_By_ColName(randomrecords[i][1], rowindex), randomrecords[i][2]), randomrecords[i][1],
                                                         self.wtrowindex)

            self.wtrowindex = self.wtrowindex + 1

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name("4")
        self.driver.quit()
        # print(self.excel)
        self.Generate_Final_Report()
        self.report.Save_Excel()