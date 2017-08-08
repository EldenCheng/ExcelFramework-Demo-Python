from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from pathlib import Path
import random
import time
from Common.CONST_j import CONST
from Common.Alias import *

class WebPage:
    def __init__(self,driver=''):

        self.driver = driver
        self.browser =''

    def Start_Up(self,url,browser = "Chrome"):
        try:
            if browser == "Firefox":
                self.driver = webdriver.Firefox()
            elif browser == "Chrome":
                self.driver = webdriver.Chrome(CONST.CHROMEDRIVERPATH)
            elif browser == "IE":
                self.driver = webdriver.Ie(CONST.IEDRIVERPATH)
            else:
                raise Exception("Not support this kind of driver")
        except Exception as msg:
            print(msg)

        try:
            self.browser = browser
            self.driver.get(url)
            self.driver.implicitly_wait(5)

            if self.browser != "Firefox":
                self.driver.maximize_window()
            else:
                try:
                    self.driver.maximize_window()
                except Exception as msg:
                    print(msg)
        except Exception as msg:
            print(msg)

        if self.driver.title != "KV Login Page":
            try:
                self.By_Pass_External_Page()
                WebDriverWait(self.driver, 5, 0.5).until(EC.title_is("KV Login Page"))
            except Exception as msg:
                raise Exception("The Exception page cannot be skipped or cannot access the website")
        else:
            time.sleep(5)

        return self.driver

    def By_Pass_External_Page(self):
        try:
            WebDriverWait(self.driver, 5, 0.5).until(lambda x:
                                                     x.find_element_by_css_selector("a[href='login-flow']")).click()
        except Exception as msg:
            print(msg)

    def Log_in(self, excel, test_case_no, data_set, case_dir_path, start_step=1):

        try:
            self.Input(excel.Get_Value_By_ColName("ID", data_set, case_dir_path), LoginPageAlias_CSS['ID_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                        % (str(test_case_no), str(data_set), start_step))
            pw = excel.Get_Value_By_ColName("PW", data_set, case_dir_path)
            self.Input(pw, LoginPageAlias_CSS['PW_Field'])
            self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                        % (str(test_case_no), str(data_set), start_step + 1))
            time.sleep(2)
            if self.browser != "Firefox":
                self.ButtonClick(LoginPageAlias_CSS['Login_Btn'])
            else:
                self.Send_key(Keys.ENTER, LoginPageAlias_CSS['PW_Field'])
                #time.sleep(2)
                #if self.driver.title == "KV Login Page":
                #self.Send_key(Keys.ENTER, LoginPageAlias_CSS['PW_Field'])

            time.sleep(1)

            steps = self.Verification_Code(pw, test_case_no, data_set, case_dir_path, start_step + 1)

            WebDriverWait(self.driver, 2, 0.5).until(EC.title_is("Start"))

            return steps

        except Exception as msg:
            print(msg)

    def Verification_Code(self,pwd, test_case_no, data_set, case_dir_path, step):
        if self.driver.title == "KV Login Page":
            time.sleep(2)

            if self.driver.find_element(By.CSS_SELECTOR, LoginPageAlias_CSS['Login_Error_Prompt']).text \
                    == "Please input verification code.":
                self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                            % (str(test_case_no), str(data_set), step + 1))
                try:
                    Jscript = "var code =prompt('Please input the Verification Code Manually');"  \
                               "var inbox =document.querySelector(%s);"   \
                               "inbox.value = code;" % LoginPageAlias_CSS['Verfidation_Code_Field']
                    try:
                        self.driver.execute_script(Jscript)
                    except Exception as msg:
                        print(msg)
                    WebDriverWait(self.driver, 20, 1).until_not(EC.alert_is_present())
                    self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                                % (str(test_case_no), str(data_set), step + 2))
                except Exception as msg:
                    print(msg)

                if self.browser != "Firefox":

                    self.Input(pwd, LoginPageAlias_CSS['PW_Field'])
                    self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                                % (str(test_case_no), str(data_set), step + 3))

                elif self.browser == "Firefox":

                        Jscript = "var inbox =document.querySelector('input[name=password]');"  \
                                   "inbox.value = '%s';" % pwd
                        try:
                            self.driver.execute_script(Jscript)
                        except Exception as msg:
                            print(msg)
                        self.driver.save_screenshot(str(case_dir_path) + r"\Steps" + r"\TC%s_DataSet_%s_Step_%d.png"
                                                % (str(test_case_no), str(data_set), step + 3))
                time.sleep(2)

                if self.browser != "Firefox":
                    self.ButtonClick(LoginPageAlias_CSS['Login_Btn'])
                else:
                    self.Send_key(Keys.ENTER, LoginPageAlias_CSS['PW_Field'])

                return step + 3
        else:
            return step

    def Input(self, text, Element, Expression="input[name=%s]", driverT = ''):
        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Expression % Element).clear()
            driver.find_element(By.CSS_SELECTOR, Expression % Element).send_keys(str(text))
        except Exception as msg:
            print(msg)

    def ButtonClick(self,Element,Expression="button[id=%s]", driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Expression % Element).click()
        except Exception as msg:
            print(msg)

    def LabelClick(self,Element,Expression="label[id=%s]", driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Expression % Element).click()
        except Exception as msg:
            print(msg)

    def RadioButtonClick(self, Element, Expression, path, row, colname):

        try:
            elements = self.driver.find_elements(By.CSS_SELECTOR, Expression)

            if Element.lower() != "randomindex":
                elements[POMaintenancePageAlias_CSS[Element]].click()
            else:
                randomvalue = random.choice(elements)
                index = elements.index(randomvalue)
                labels = self.driver.find_elements(By.CSS_SELECTOR, Expression.replace("input", "label"))
                file = Path(path) / Path("%d_%s_%s.random" % (row, colname, labels[index * 2 + 1].text))
                file.touch()
                randomvalue.click()

        except Exception as msg:
            print(msg)

    def Send_key(self, key, Element, Expression="input[name=%s]", driverT = ''):
        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            driver.find_element(By.CSS_SELECTOR, Expression % Element).send_keys(key)
        except Exception as msg:
            print(msg)

    def Move_To_Click(self, Element, Expression="input#%s", driverT = ''):
        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
           ActionChains(self.driver).move_to_element(self.driver.find_element(By.CSS_SELECTOR,
                                                                              Expression % Element)).click().perform()

           #self.driver.find_element(By.CSS_SELECTOR, Expression % Element).click()
        except Exception as msg:
            print(msg)


    def Verify_Text(self,text,Expression,driverT = ''):

        if driverT != '':
            driver = driverT
        else:
            driver = self.driver

        try:
            element_text = driver.find_element(By.CSS_SELECTOR, Expression).text
            element_text = ''.join(element_text.split())
            text = ''.join(text.split())

            if element_text[-len(text):] == text:
                print("The text of the element is equals %s" % text)
                return True
            else:
                print("The text of the element is not contents %s" % text)
                return False
        except Exception as msg:
            print(msg)

