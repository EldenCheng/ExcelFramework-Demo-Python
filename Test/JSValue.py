from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

driver = webdriver.Chrome(r"..\Webdrivers\chromedriver.exe")

driver.get("http://www.baidu.com")

Jscript = "var code =prompt('Please input the Verification Code Manually');"  \
           "var inbox =document.querySelector('input#kw');"  \
           "inbox.value = code;"


driver.execute_script(Jscript)


WebDriverWait(driver,20,1).until_not(EC.alert_is_present())

#time.sleep(5)

#print(driver.find_element(By.CSS_SELECTOR,"input#kw").text)
#print(vcode)

driver.quit()

