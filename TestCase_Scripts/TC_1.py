from Common.CONST import CONST
from Common.Excel import Excel
from Common.WebPage import WebPage

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


    def test_Excute(self):

        casedirpath = Path(self.reportfilepath).parent / Path("TC1")
        passdirpath = casedirpath / Path("Pass")
        faildirpath = casedirpath / Path("Fail")
        stepsdirpath = casedirpath / Path("Steps")
        if not casedirpath.is_dir():
            casedirpath.mkdir()

        if not stepsdirpath.is_dir():
            stepsdirpath.mkdir()

        for i in self.excel.Get_Excution_DataSet("executed"):

            try:
                self.page.Input(self.excel.Get_Value_By_ColName("ID",i,casedirpath),"userId")
                self.driver.save_screenshot(str(stepsdirpath) + "/TC1_DataSet_%s_Step1.png" % str(i))
                self.page.Input(self.excel.Get_Value_By_ColName("PW",i,casedirpath),"password")
                self.driver.save_screenshot(str(stepsdirpath) + "/TC1_DataSet_%s_Step2.png" % str(i))
                self.page.ButtonClick(r"btn_login-btnEl")
                time.sleep(1)

                self.page.Verification_Code(self.excel.Get_Value_By_ColName("PW",i))

                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Start"))

                if self.page.Verify_Text(self.excel.Get_Value_By_ColName("Assertion",i),"LogonUserName-body",r"div#%s div"):
                    print("success")
                else:
                    self.page.ButtonClick(r"btn_logout-btnEl")
                    WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
                    raise AssertionError("The element not contains the Assertion text")


            except Exception as msg:
                print(msg)

                if not faildirpath.is_dir():
                    faildirpath.mkdir()

                time.sleep(2)
                snapshot = str(faildirpath.absolute()) + r"\TC1_fail_on_dataset_%s.png" % str(
                    int(self.excel.Get_Value_By_ColName("Data set", i)))
                self.Generate_Report("fail",snapshot,i)

                self.driver.refresh()

            else:
                if not passdirpath.is_dir():
                    passdirpath.mkdir()

                snapshot = str(passdirpath.absolute()) + r"\TC1_pass_on_dataset_%s.png" % str(int(self.excel.Get_Value_By_ColName("Data set",i)))
                self.Generate_Report("pass", snapshot, i)

                time.sleep(2)
                self.page.ButtonClick(r"btn_logout-btnEl")
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))

    def Generate_Report(self,pass_fail,snapshotpath,rowindex):
        #print(pass_fail)
        #print(str(rowindex))
        self.report.Select_Sheet_By_Name("1")
        self.driver.save_screenshot(snapshotpath)
        self.report.Set_Value_By_ColName(pass_fail, "Result", rowindex)
        self.report.Set_Value_By_ColName(r"file:///" + snapshotpath, "Screen capture", rowindex)
        self.report.Set_Value_By_ColName("done", "executed", rowindex)

    def Generate_Final_Report(self):

        randomfilepath = Path(self.reportfilepath).parent / Path("TC1")
        randomfiles = list(randomfilepath.rglob('*.random'))
        #print(randomfiles)

        randomrecords = []

        for l in randomfiles:
            p = str(list(l.parts)[-1:])
            p = p[:p.find(".random")]
            randomrecords.append(p.split("_"))

        randomrecords = list(randomrecords)

        for rowindex in range(1, self.excel.Get_Row_Numbers()):
            self.report.Select_Sheet_By_Name("result-timestamp")
            # print(self.wtrowindex)
            self.report.Set_Value_By_ColName(1, "Case No", self.wtrowindex)
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
        self.excel.Select_Sheet_By_Name("1")
        self.driver.quit()
        self.Generate_Final_Report()
        self.report.Save_Excel()









