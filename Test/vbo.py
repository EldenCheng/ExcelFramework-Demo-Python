import win32gui
import win32api
import win32con

from selenium import webdriver
import unittest, time
import xlrd
from xlutils.copy import copy
import os
import shutil
import xlwt
import random
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


driver=webdriver.Chrome(r"..\WebDrivers\chromedriver.exe")
driver.implicitly_wait(30)
driver.maximize_window()


driver.get("http://www-uat.kerrylogistics.com/kerriervbo-demo/dispatcher/index-flow?execution=e4s1")

driver.find_element_by_xpath("//*[@id='ext-gen1018']/a").click()

driver.find_element_by_id("txtfld_id-inputEl").send_keys("jack.li")
driver.find_element_by_id("txtfld_pwd-inputEl").clear()

driver.find_element_by_id("txtfld_pwd-inputEl").send_keys("abc@12345")
time.sleep(1)
driver.find_element_by_id("btn_login-btnIconEl").click()

driver.find_element_by_id("label-1030").click()

driver.find_element_by_id("txfld_contactPerson-inputEl").send_keys("automation demo 0")
driver.find_element_by_id("txfld_position-inputEl").send_keys("automation demo 1")
driver.find_element_by_id("txfld_phoneNo-inputEl").send_keys("12345678")
driver.find_element_by_id("txfld_faxNo-inputEl").send_keys("87654321")
driver.find_element_by_id("txfld_companyEmail-inputEl").send_keys("automation@demo.com")

driver.find_element_by_id("txfld_companyName-inputEl").send_keys("automation demo 2")
driver.find_element_by_id("txfld_supplierEmail-inputEl").send_keys("automation@demo.com")
driver.find_element_by_id("txfld_subject-inputEl").send_keys("automation demo 3")

driver.find_element_by_id("button-1025-btnInnerEl").click()

time.sleep(1)

driver.find_element_by_xpath("//*[@id='container-1015']/a[1]").click()
ActionChains(driver).move_to_element(driver.find_element_by_id("panel_SPM-innerCt")).perform()
