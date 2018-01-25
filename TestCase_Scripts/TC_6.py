#@BEGIN
for i in self.excel.Get_Excution_DataSet("executed"):

    try:
        bro = self.excel.Get_Value_By_ColName("Browser", i)
        if bro is not None:
            self.browser = bro

        self.page = WebPage()
        self.driver = self.page.Start_Up(CONST.URL, self.browser)
        if i != '':
            self.page.Log_in(self.excel, self.caseno, i, self.casedirpath)

            if self.page.Verify_Text(self.excel.Get_Value_By_ColName("Assertion", i),
                                     StartPageAlias_CSS['Login_UserName']):
                Log("success", self.caseno, i, self.casedirpath)
            else:
                raise AssertionError("The element not contains the Assertion text")

    except Exception as msg:
        Log(str(msg), self.caseno, i, self.casedirpath)
        Generate_Report(self.driver, self.excel, self.report, "fail", self.caseno, self.casedirpath, i)

        self.driver.quit()
    else:
        Generate_Report(self.driver, self.excel, self.report, "pass", self.caseno, self.casedirpath, i)
        self.driver.quit()

    #@END