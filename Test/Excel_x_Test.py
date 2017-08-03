from Common.Excel_x import Excel

#excel = Excel(r"..\TestCaseData\Kerry_data_2c.xlsx")

excel = Excel(r"D:\百度云同步盘\项目\自动化测试\Kerry\KerryDemo\TestReport\2017-08-02_182545\Report_2017-08-02_182545.xlsx")

excel.Select_Sheet_By_Name("result")

print(excel.Get_Value_By_ColName("Screen capture", 1))

#print(excel.Get_Row_Numbers())

#print(excel.Get_All_Values_By_ColName("ID"))

#print(excel.Get_Value_By_ColName("ID", 1))

#print(excel.Get_Value_By_ColName("PW", 1, r"D:\VBA"))

#report= Excel(r"..\TestReport\Kerry_data_2w.xlsx")

#report.Select_Sheet_By_Name("1")

#report.Set_Sheet_Name("1","5")

#report.Set_Value_By_ColName(r'=HYPERLINK("D:\百度云同步盘\项目\自动化测试\Kerry\KerryDemo\TestReport\2017-08-02_145238\TC4\Pass\TC4_Dataset_1_Step_Pass.png","TC4_Dataset_1_Step_Pass.png")', "Screen capture", 1)

#report.Save_Excel()

