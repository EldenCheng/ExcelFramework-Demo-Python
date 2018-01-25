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


class TC_Runner(unittest.TestCase):
    def __init__(self, testname, TC_NO, reportpath):

        self.caseno = TC_NO
        exec("super(TC_Runner, self).__init__(testname)")
        self.reportfilepath = reportpath

    def setUp(self):

        self.excel = Excel()
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        self.report = Excel(self.reportfilepath)
        self.casedirpath = Path(self.reportfilepath).parent / Path("TC%d" % self.caseno)
        self.stepsdirpath = self.casedirpath / Path("Steps")
        self.TC_Content = ''

        isscript = False

        if not self.casedirpath.is_dir():
            self.casedirpath.mkdir()

        if not self.stepsdirpath.is_dir():
            self.stepsdirpath.mkdir()

        try:
            with open(r'.\TestCase_Scripts\TC_%d.py' % self.caseno) as file:
                for line in file:
                    verifyTxt = ''.join(line.split())
                    if verifyTxt == '#@BEGIN':
                        isscript = True
                    if isscript == True:
                        self.TC_Content = self.TC_Content + line[8:]
                    if verifyTxt == "#@END":
                        isscript = False
                        break

            self.TC_Content = self.TC_Content.replace("\"", "\'")
            #print(self.TC_Content)

        except Exception as msg:
            print(msg)

    def test_Excute(self):
        print("Now testing TC_%d" % self.caseno)
        exec(self.TC_Content)

    def tearDown(self):
        self.report.Save_Excel()
        self.excel = Excel(self.reportfilepath)
        self.excel.Select_Sheet_By_Name(str(self.caseno))
        Generate_Final_Report(self.excel, self.report, self.caseno)
        self.report.Save_Excel()
