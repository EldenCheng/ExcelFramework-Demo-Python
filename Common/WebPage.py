from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from Common.CONST import CONST
from Common.Alias import *

class WebPage:
    def __init__(self,driver=''):

        self.driver = driver
        self.browser =''

    def Start_Up(self,url,browser = "Firefox"):

        if browser == "Firefox":
            self.driver = webdriver.Firefox()
        elif browser == "Chrome":
            self.driver = webdriver.Chrome(CONST.CHROMEDRIVERPATH)
        elif browser == "IE":
            self.driver = webdriver.Ie(CONST.IEDRIVERPATH)
        else:
            print("Not support this kind of driver")

        self.browser = browser

        self.driver.get(url)
        self.driver.implicitly_wait(5)

        if self.browser != "Firefox":
            self.driver.maximize_window()
        else:
            time.sleep(5)

        if self.driver.title != "KV Login Page":
            self.By_Pass_External_Page()
            WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
        else:
            time.sleep(5)

        return self.driver

    def By_Pass_External_Page(self):
        WebDriverWait(self.driver,5,0.5).until(lambda x: x.find_element_by_css_selector("a[href='login-flow']")).click()
        #if self.driver.find_element(By.CSS_SELECTOR, "a[href='login-flow']"):
        #    self.driver.find_element(By.CSS_SELECTOR, "a[href='login-flow']").click()

    def Log_in(self, page, excel, test_case_no, data_set, case_dir_path, start_step=1):

        try:
            page.Input(excel.Get_Value_By_ColName("ID", data_set, case_dir_path), LoginPageAlias_CSS['ID_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                        % (str(test_case_no), str(data_set), start_step))
            page.Input(excel.Get_Value_By_ColName("PW", data_set, case_dir_path), LoginPageAlias_CSS['PW_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                        % (str(test_case_no), str(data_set), start_step + 1))
            page.ButtonClick(LoginPageAlias_CSS['Login_Btn'])
            time.sleep(1)

            page.Verification_Code(excel.Get_Value_By_ColName("PW", data_set))

            WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("Start"))

        except Exception as msg:
            print(msg)

    def Verification_Code(self,pwd):
        if self.driver.title == "KV Login Page":
            time.sleep(2)
            if self.driver.find_element(By.CSS_SELECTOR, LoginPageAlias_CSS['Verfidation_Code_Text']).text == "Please input verification code.":

                try:
                    Jscript = "var code =prompt('Please input the Verification Code Manually');"  \
                               "var inbox =document.querySelector(%s);"   \
                               "inbox.value = code;" % LoginPageAlias_CSS['Verfidation_Code_Field']
                    self.driver.execute_script(Jscript)
                    WebDriverWait(self.driver, 20, 1).until_not(EC.alert_is_present())

                except Exception as msg:
                    print(msg)

                if self.browser != "Firefox":

                    self.Input(pwd, LoginPageAlias_CSS['PW_Field'])

                elif self.browser == "Firefox":
                    try:
                        Jscript = "var inbox =document.querySelector('input[name=password]');"  \
                                   "inbox.value = %s;" % pwd
                        self.driver.execute_script(Jscript)

                    except Exception as msg:
                        print(msg)

                time.sleep(2)
                self.ButtonClick(LoginPageAlias_CSS['Login_Btn'])

        else:
            pass

    def Input(self, text, Element, Exception="input[name=%s]", driverT = ''):
        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Exception % Element).clear()
            driver.find_element(By.CSS_SELECTOR, Exception % Element).send_keys(str(text))
        except Exception as msg:
            print(msg)

    def ButtonClick(self,Element,Exception="button[id=%s]", driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Exception % Element).click()
        except Exception as msg:
            print(msg)

    def LabelClick(self,Element,Exception="label[id=%s]", driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Exception % Element).click()
        except Exception as msg:
            print(msg)

    def Verify_Text(self,text,Element,Exception,driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        element_text = driver.find_element(By.CSS_SELECTOR, Exception % Element).text
        element_text = ''.join(element_text.split())
        text = ''.join(text.split())

        #print("The element text is %s" % element_text[-len(text):])
        #print("The text to compare is %s" % text)

        try:
            if element_text[-len(text):] == text:
                print("The text of the element is equals %s" % text)
                return True

            #elif driver.find_element(By.CSS_SELECTOR, Exception % Element).text.find(text):
            #    print("The text of the element is contents %s" % text)
            #    return True
            else:
                print("The text of the element is not contents %s" % text)
                return False
        except Exception as msg:
            print(msg)
            return False