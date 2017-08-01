from Common.CONST import CONST
from Common.Excel import Excel
from Common.WebPage import WebPage
from Common.Alias import *

from pathlib import Path


def Generate_Report(driver, report, pass_fail, test_case_no, snapshotpath, rowindex):
    # print(pass_fail)
    # print(str(rowindex))
    report.Select_Sheet_By_Name(str(test_case_no))
    driver.save_screenshot(snapshotpath)
    report.Set_Value_By_ColName(pass_fail, "Result", rowindex)
    report.Set_Value_By_ColName(r"file:///" + snapshotpath, "Screen capture", rowindex)
    report.Set_Value_By_ColName("done", "executed", rowindex)


def Generate_Final_Report(excel, report, reportfilepath, test_case_no):
    randomfilepath = Path(reportfilepath).parent / Path("TC%s" % str(test_case_no))
    randomfiles = list(randomfilepath.rglob('*.random'))
    # print(randomfiles)

    randomrecords = []
    report.Select_Sheet_By_Name("result-timestamp")
    wtrowindex = report.Get_Row_Numbers()

    for l in randomfiles:
        p = str(list(l.parts)[-1:])
        p = p[:p.find(".random")]
        randomrecords.append(p.split("_"))

    randomrecords = list(randomrecords)

    for rowindex in range(1, excel.Get_Row_Numbers()):
        # print(wtrowindex)
        report.Set_Value_By_ColName(test_case_no, "Case No", wtrowindex)
        report.Set_Value_By_ColName(rowindex, "Data set", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Expected result", rowindex), "Expected result", wtrowindex)
        res = excel.Get_Value_By_ColName("Result", rowindex)
        exe = excel.Get_Value_By_ColName("executed", rowindex)
        if exe != 'Skip':
            report.Set_Value_By_ColName(res, "Result", wtrowindex)
        elif exe == "Skip":
            report.Set_Value_By_ColName("Skipped", "Result", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Description", rowindex),"Description", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Screen capture", rowindex),"Screen capture", wtrowindex)
        report.Set_Value_By_ColName(exe, "executed", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Assertion", rowindex), "Assertion", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("ID", rowindex), "ID", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("PW", rowindex), "PW", wtrowindex)

        if randomrecords != '':
            for i in range(len(randomrecords)):
                if int(randomrecords[i][0][-1:]) == rowindex:
                    report.Set_Value_By_ColName("%s(%s)" % (
                    excel.Get_Value_By_ColName(randomrecords[i][1], rowindex), randomrecords[i][2]),
                                                     randomrecords[i][1], wtrowindex)

        wtrowindex = wtrowindex + 1