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


driver = webdriver.Firefox()


url = r"http://www-uat.kerrylogistics.com/kerriervbo-demo/dispatcher/index-flow?execution=e4s1"
driver.get(url)
driver.implicitly_wait(5)

driver.delete_all_cookies()

WebDriverWait(driver, 40, 2).until(lambda x: x.find_element_by_css_selector("a[href='login-flow']")).click()

WebDriverWait(driver, 40, 2).until(lambda x: x.find_element_by_css_selector("input[name=userId]")).clear()

driver.find_element_by_css_selector("input[name=userId]").send_keys("1234")

driver.find_element_by_css_selector("input[name=password]").clear

driver.find_element_by_css_selector("input[name=password]").send_keys("1234")

time.sleep(2)

driver.find_element_by_css_selector("input[name=password]").send_keys(Keys.ENTER)

WebDriverWait(driver, 40, 2).until(lambda x: x.find_element_by_css_selector("img#img_randomCode")).screenshot\
    (r".\Webelementsnapshot.jpg")

#driver.find_element_by_css_selector("img#img_randomCode").screenshot(r".\Webelementsnapshot.jpg")

time.sleep(10)

driver.quit()
