from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from pathlib import Path
import random
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

    def Log_in(self, page, excel, test_case_no, data_set, case_dir_path, start_step=1):

        try:
            page.Input(excel.Get_Value_By_ColName("ID", data_set, case_dir_path), LoginPageAlias_CSS['ID_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step%d.png"
                                        % (str(test_case_no), str(data_set), start_step))
            page.Input(excel.Get_Value_By_ColName("PW", data_set, case_dir_path), LoginPageAlias_CSS['PW_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step%d.png"
                                        % (str(test_case_no), str(data_set), start_step + 1))
            page.ButtonClick(LoginPageAlias_CSS['Login_Btn'])
            time.sleep(1)

            steps = page.Verification_Code(excel.Get_Value_By_ColName("PW", data_set), test_case_no, data_set,
                                           case_dir_path, start_step + 1)

            WebDriverWait(self.driver, 2, 0.5).until(EC.title_is("Start"))

            return steps

        except Exception as msg:
            print(msg)

    def Verification_Code(self,pwd, test_case_no, data_set, case_dir_path, step):
        if self.driver.title == "KV Login Page":
            time.sleep(2)
            if self.driver.find_element(By.CSS_SELECTOR, LoginPageAlias_CSS['Login_Error_Prompt']).text \
                    == "Please input verification code.":

                try:
                    Jscript = "var code =prompt('Please input the Verification Code Manually');"  \
                               "var inbox =document.querySelector(%s);"   \
                               "inbox.value = code;" % LoginPageAlias_CSS['Verfidation_Code_Field']
                    self.driver.execute_script(Jscript)
                    WebDriverWait(self.driver, 20, 1).until_not(EC.alert_is_present())

                    self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step%d.png"
                                                % (str(test_case_no), str(data_set), step + 1))

                except Exception as msg:
                    print(msg)

                if self.browser != "Firefox":

                    self.Input(pwd, LoginPageAlias_CSS['PW_Field'])

                    self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step%d.png"
                                                % (str(test_case_no), str(data_set), step + 2))

                elif self.browser == "Firefox":
                    try:
                        Jscript = "var inbox =document.querySelector('input[name=password]');"  \
                                   "inbox.value = %s;" % pwd
                        self.driver.execute_script(Jscript)

                        self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step%d.png"
                                                    % (str(test_case_no), str(data_set), step + 2))

                    except Exception as msg:
                        print(msg)

                time.sleep(2)
                self.ButtonClick(LoginPageAlias_CSS['Login_Btn'])

                return step + 2
        else:
            return step

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

    def VideoButtonClick(self, Element, Exception, path, row, colname):

        elements = self.driver.find_elements(By.CSS_SELECTOR, Exception)

        #for e in elements:
        #    print(e.get_attribute("id"))

        #print(POMaintenancePageAlias_CSS[Element])

        #print(random.choice(elements).get_attribute("id"))

        if Element.lower() != "randomindex":
            elements[POMaintenancePageAlias_CSS[Element]].click()
        else:
            randomvalue = random.choice(elements)
            index = elements.index(randomvalue)
            labels = self.driver.find_elements(By.CSS_SELECTOR, Exception.replace("input", "label"))
            file = Path(path) / Path("%d_%s_%s.random" % (row, colname, labels[index * 2 + 1].text))
            file.touch()
            randomvalue.click()


    def Verify_Text(self,text,Exception,driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        element_text = driver.find_element(By.CSS_SELECTOR, Exception).text
        element_text = ''.join(element_text.split())
        text = ''.join(text.split())
        try:
            if element_text[-len(text):] == text:
                print("The text of the element is equals %s" % text)
                return True
            else:
                print("The text of the element is not contents %s" % text)
                return False
        except Exception as msg:
            print(msg)
            return False
