from Common.CONST import CONST
from Common.Excel import Excel
from Common.WebPage import WebPage
from Common.Alias import *


def Generate_Report(self, pass_fail, snapshotpath, rowindex):
    # print(pass_fail)
    # print(str(rowindex))
    self.report.Select_Sheet_By_Name("1")
    self.driver.save_screenshot(snapshotpath)
    self.report.Set_Value_By_ColName(pass_fail, "Result", rowindex)
    self.report.Set_Value_By_ColName(r"file:///" + snapshotpath, "Screen capture", rowindex)
    self.report.Set_Value_By_ColName("done", "executed", rowindex)


def Generate_Final_Report(self):
    randomfilepath = Path(self.reportfilepath).parent / Path("TC1")
    randomfiles = list(randomfilepath.rglob('*.random'))
    # print(randomfiles)

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
                    self.report.Set_Value_By_ColName("%s(%s)" % (
                    self.excel.Get_Value_By_ColName(randomrecords[i][1], rowindex), randomrecords[i][2]),
                                                     randomrecords[i][1],
                                                     self.wtrowindex)

        self.wtrowindex = self.wtrowindex + 1