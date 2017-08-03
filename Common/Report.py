from Common.CONST import CONST
from pathlib import Path

import time,shutil


def Generate_Report(driver, excel, report, pass_fail, test_case_no, casedirpath, rowindex):
    # print(pass_fail)
    # print(str(rowindex))

    try:
        if pass_fail == "pass":
            passdirpath = casedirpath / Path("Pass")
            if not passdirpath.is_dir():
                passdirpath.mkdir()

            capturename = r"TC%s_Dataset_%s_Step_Pass.png" \
                          % (str(test_case_no), str(int(excel.Get_Value_By_ColName("Data set", rowindex))))
            snapshotpath = str(passdirpath.absolute()) + "\\" + capturename
            link = r'=HYPERLINK("%s","%s")' % (snapshotpath, capturename)

        elif pass_fail == "fail":
            faildirpath = casedirpath / Path("Fail")
            if not faildirpath.is_dir():
                faildirpath.mkdir()

            capturename = r"TC%s_Dataset_%s_Step_Fail.png" \
                          % (str(test_case_no), str(int(excel.Get_Value_By_ColName("Data set", rowindex))))
            snapshotpath = str(faildirpath.absolute()) + "\\" + capturename
            link = r'=HYPERLINK("%s","%s")' % (snapshotpath, capturename)

        report.Select_Sheet_By_Name(str(test_case_no))
        driver.save_screenshot(snapshotpath)
        report.Set_Value_By_ColName(pass_fail, "Result", rowindex)
        report.Set_Value_By_ColName(link, "Screen capture", rowindex)
        report.Set_Value_By_ColName("done", "executed", rowindex)

        randomfilepath = Path(casedirpath)
        randomfiles = list(randomfilepath.rglob('*.random'))
        if len(randomfiles) != 0:
            randomrecords = []
            for l in randomfiles:
                p = str(list(l.parts)[-1:])
                p = p[:p.find(".random")]
                randomrecords.append(p.split("_"))

            randomrecords = list(randomrecords)

            for i in range(len(randomrecords)):
                if int(randomrecords[i][0][-1:]) == rowindex:
                    report.Set_Value_By_ColName("%s(%s)" % (
                        excel.Get_Value_By_ColName(randomrecords[i][1], rowindex), randomrecords[i][2]),
                                                randomrecords[i][1], rowindex)

        # TO DO: Cut Steps captures of fail case to Fail folder

        # TO DO: Generate a Log file to place error msg

    except Exception as msg:
        print(msg)

def Generate_Final_Report(excel, report, reportfilepath, test_case_no):

    report.Select_Sheet_By_Name("result-timestamp")
    wtrowindex = report.Get_Row_Numbers()

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
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Screen capture", rowindex), "Screen capture",
                                    wtrowindex)
        report.Set_Value_By_ColName(exe, "executed", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("Assertion", rowindex), "Assertion", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("ID", rowindex), "ID", wtrowindex)
        report.Set_Value_By_ColName(excel.Get_Value_By_ColName("PW", rowindex), "PW", wtrowindex)

        wtrowindex = wtrowindex + 1

def Create_New_Report(folderpath= r".\TestReport" ):
    otime = time.strftime("%Y-%m-%d_%H%M%S", time.localtime())
    reprotfolderpath = Path(folderpath) / Path(otime)
    if reprotfolderpath.is_dir():
        pass
    else:
        reprotfolderpath.mkdir()
        reportfilepath = folderpath + "\\" + reprotfolderpath.name +  r"\Report_" + otime + ".xlsx"

    shutil.copy(CONST.EXCELPATH,reportfilepath)

    if Path(reportfilepath).is_file():
        return str(reportfilepath)
    else:
        raise Exception("Cannot copy excel file")