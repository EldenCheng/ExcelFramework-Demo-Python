from Common.WebPage import WebPage
from Common.Alias import *
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

page = WebPage()

driver = page.Start_Up(r"http://www-uat.kerrylogistics.com/kerriervbo-demo/dispatcher/index-flow?execution=e4s1","Chrome")

page.Input('jack.li', LoginPageAlias_CSS['ID_Field'])

page.Input('abc@12345', LoginPageAlias_CSS['PW_Field'])

page.ButtonClick(LoginPageAlias_CSS['Login_Btn'])

WebDriverWait(driver, 2, 0.5).until(EC.title_is("Start"))

page.LabelClick(r"label-1054")

time.sleep(2)

elements = driver.find_elements(By.CSS_SELECTOR, "tr#ext-gen1256 td input")

for e in elements:
    #print(e.get_attribute("id"))
    e.click()
    time.sleep(3)
