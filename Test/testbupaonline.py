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

rb = xlrd.open_workbook("c:/test_project/data/Bupa_Caesar_data_table_v1_170427.xls")
r_sheet = rb.sheet_by_name("Bupa online for web clinic")
num_rows = r_sheet.nrows
num_cols = r_sheet.ncols
caseno = "Case No"
casenos = []
username = "username"
usernames = []
password = "password"
passwords = []
providercode = "provider code"
providercodes = []
membership1 = "membership1"
membership1s = []
membership2 = "membership2"
membership2s = []
receiptno = "Receipt No"
receiptnos = []
outputs = []
otime=time.strftime("%Y-%m-%d %H%M%S", time.localtime())
hyperlink=[]
judgementrandom = "judgement random"
judgementrandoms=[]
yzmm="ID/PW"
yzmms=[]
ccache="clear cache"
ccaches=[]
execute="executed"
executes=[]
select1="Select1"
select1s=[]
select2="Select2"
select2s=[]
judgementrandom2="judgement random2"
judgementrandom2s=[]
pcode="physician code"
pcodes=[]
lcode="location code"
lcodes=[]
select3="Select3"
select3s=[]
other1="other1"
other2="other2"
other3="other3"
other1v="other1 money"
other2v="other2 money"
other3v="other3 money"
other1s=[]
other2s=[]
other3s=[]
other1vs=[]
other2vs=[]
other3vs=[]
fugai=[]
for i in range(num_rows):
    for j in range(num_cols):
        val = r_sheet.cell(i, j).value
        if val == caseno:
            for k in range(i+4, num_rows-7):
                casenos.append(int(r_sheet.cell(k,j).value))
        elif val == username:
            for k in range(i+1, num_rows-7):
                usernames.append(r_sheet.cell(k,j).value)
        elif val == password:
            for k in range(i+1, num_rows-7):
                passwords.append(r_sheet.cell(k,j).value)
        elif val == providercode:
            for k in range(i+1, num_rows-7):
                providercodes.append(r_sheet.cell(k, j).value)
        elif val == membership1:
            for k in range(i+1, num_rows-7):
                membership1s.append(r_sheet.cell(k,j).value)
        elif val == membership2:
            for k in range(i+1, num_rows-7):
                membership2s.append(r_sheet.cell(k, j).value)
        elif val == receiptno:
            for k in range(i+1, num_rows-7):
                receiptnos.append(r_sheet.cell(k, j).value)
        elif val == judgementrandom:
            for k in range(i+1, num_rows-7):
                judgementrandoms.append(r_sheet.cell(k,j).value)
        elif val == yzmm:
            for k in range(i+1, num_rows-7):
                yzmms.append(r_sheet.cell(k,j).value)
        elif val == ccache:
            for k in range(i+4, num_rows-7):
                ccaches.append(r_sheet.cell(k,j).value)
        elif val == execute:
            for k in range(i+4, num_rows-7):
                executes.append(r_sheet.cell(k,j).value)
        elif val ==select1:
            for k in range(i+1, num_rows-7):
                select1s.append(r_sheet.cell(k,j).value)
        elif val == select2:
            for k in range(i+1, num_rows-7):
                select2s.append(r_sheet.cell(k,j).value)
        elif val == judgementrandom2:
            for k in range(i+1, num_rows-7):
                judgementrandom2s.append(r_sheet.cell(k,j).value)
        elif val == pcode:
            for k in range(i+1, num_rows-7):
                pcodes.append(r_sheet.cell(k,j).value)
        elif val == lcode:
            for k in range(i+1, num_rows-7):
                lcodes.append(r_sheet.cell(k,j).value)
        elif val == select3:
            for k in range(i+1, num_rows-7):
                select3s.append(r_sheet.cell(k,j).value)
        elif val == other1:
            for k in range(i+1, num_rows-7):
                other1s.append(r_sheet.cell(k,j).value)
        elif val == other2:
            for k in range(i+1, num_rows-7):
                other2s.append(r_sheet.cell(k,j).value)
        elif val == other3:
            for k in range(i+1, num_rows-7):
                other3s.append(r_sheet.cell(k,j).value)
        elif val == other1v:
            for k in range(i+1, num_rows-7):
                other1vs.append(r_sheet.cell(k,j).value)
        elif val == other2v:
            for k in range(i+1, num_rows-7):
                other2vs.append(r_sheet.cell(k,j).value)
        elif val == other3v:
            for k in range(i+1, num_rows-7):
                other3vs.append(r_sheet.cell(k,j).value)

class Bupa01(unittest.TestCase):




    def test_bupa01(self):
     os.makedirs(r'c:/test_project/clinic/report&screenshots/{0}/passcase/'.format(otime))
     os.makedirs(r'c:/test_project/clinic/report&screenshots/{0}/failcase/'.format(otime))
     for caseno,username,password,providercode,membership1,membership2,receiptno,judgementrandom,yzmm,ccache,execute,select1,select2,judgementrandom2,pcode,lcode,select3,other1,other2,other3,other1v,other2v,other3v in zip(casenos,usernames,passwords,providercodes,membership1s,membership2s,receiptnos,judgementrandoms,yzmms,ccaches,executes,select1s,select2s,judgementrandom2s,pcodes,lcodes,select3s,other1s,other2s,other3s,other1vs,other2vs,other3vs):
        try:
         if execute=="Done":
           outputs.append("Done")
           hyperlink.append("Done")
           fugai.append("Done")
           continue
         if execute=="Skip":
            outputs.append("Skip")
            hyperlink.append("Skip")
            fugai.append("Skip")
            continue



         if execute=="Not Yet":


          self.driver = webdriver.Ie()
          self.driver.implicitly_wait(30)
          self.base_url = "https://203.83.252.135:61443/ecard/Jsp/security/login.jsp"
          driver= self.driver
          driver.get(self.base_url)
          if ccache=="Yes":
           driver.delete_all_cookies()
          else:
              pass
          driver.maximize_window()
          driver.get("javascript:document.getElementById('overridelink').click()")

          os.makedirs(r'c:/test_project/debug/{0}/'.format(caseno))
          WebDriverWait(driver,60).until(lambda x:x.find_element_by_id("htmlPageTopContainer_ecardForm_txtUserId")).clear()

          WebDriverWait(driver,60).until(lambda x:x.find_element_by_id("htmlPageTopContainer_ecardForm_txtUserId")).send_keys(username)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP1.jpg".format(caseno))
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_pwPassword").clear()
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_pwPassword").send_keys(password)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP2.jpg".format(caseno))
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_loginLink_imgLogin").click()
          time.sleep(3)
          try:
             yzmmaT=driver.switch_to.alert.text
          except Exception:
             yzmmaT="nalert"
             pass
          if yzmm=="Incorrect" and yzmmaT=="Please refer to the message in red at the top of the page.":
               driver.switch_to.alert.accept()
               print("success")
               outputs.append("success")
               fugai.append("Done")
               driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEPyzmm.jpg".format(caseno))
               hyperlink.append(
                   "c:/test_project/clinic/report&screenshots/{0}/passcase/{1}/TC{1}_STEPyzmm.jpg".format(otime, caseno))

               shutil.move("c:/test_project/debug/{0}/".format(caseno),
                           "c:/test_project/clinic/report&screenshots/{0}/passcase/".format(otime))
               driver.quit()
               continue
          if yzmm=="Incorrect" and yzmmaT =="nalert":
            
               print("fail")
               outputs.append("fail")
               fugai.append("Done")
               self.driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEPfail.jpg".format(caseno))
               hyperlink.append(
                   "c:/test_project/clinic/report&screenshots/{0}/failcase/{1}/TC{1}_STEPfail.jpg".format(otime, caseno))
               shutil.move("c:/test_project/debug/{0}/".format(caseno),
                           "c:/test_project/clinic/report&screenshots/{0}/failcase/".format(otime))
               driver.quit()
               continue
            
         
         
          time.sleep(1)
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtProviderCode").click()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtProviderCode").clear()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtProviderCode").send_keys(providercode)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP3.jpg".format(caseno))
          #driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSearchLocationCode").click()
          #driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSearch").click()
          #time.sleep(3)
          #try:
             #driver.switch_to.alert.accept()

          #except Exception:

             #pass

          #driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[2]/form/table[5]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr/td/table/tbody/tr/td[1]/input").click()
          #driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[2]/form/table[5]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/font/input").click()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtPhysicianCode").clear()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtPhysicianCode").send_keys(pcode)
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtLocationCode").clear()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtLocationCode").send_keys(lcode)


          #driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtMembershipNo1").click()

          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtMembershipNo1").clear()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtMembershipNo1").send_keys(membership1)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP4.jpg".format(caseno))
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtMembershipNo2").clear()
          driver.find_element_by_id("htmlPageTopContainer_ecardForm_txtMembershipNo2").send_keys(membership2)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP5.jpg".format(caseno))
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSave").click()
          time.sleep(5)
          alerttext=driver.switch_to.alert.text
          if alerttext=="Please refer to the message in red at the top of the page.":
           driver.switch_to.alert.accept()
           driver.find_element_by_xpath('//*[@id="htmlPageTopContainer_ecardForm_divWithReferralLetter"]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[3]/input[1]').click()
           driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSave").click()
           time.sleep(5)
           try:
            driver.switch_to.alert.accept()
           except Exception:
            pass
           try:
            time.sleep(5)
            driver.switch_to.alert.accept()
           except Exception:
               pass
          else:
           driver.switch_to.alert.accept()
           time.sleep(3)
           driver.switch_to.alert.accept()
          time.sleep(3)
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxMedicalServiceForm_btnClaimEntry1").click()
          time.sleep(3)
          s1=  "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[1]/input"
          s2 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input"
          s3 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[1]/input"
          s4 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[4]/td[1]/input"
          s5 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[5]/td[1]/input"
          s6 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[6]/td[1]/input"
          s7 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[1]/input"
          s8 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[8]/td[1]/input"
          s9 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[9]/td[1]/input"
          s10 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[10]/td[1]/input"
          s11 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[11]/td[1]/input"
          s12 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[1]/table/tbody/tr/td/table/tbody/tr[12]/td[1]/input"
          if judgementrandom == "Yes":
           xpaths=[s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12]
           random.shuffle(xpaths)
           driver.find_element_by_xpath(xpaths[0]).click()
          if judgementrandom=="s1":
           driver.find_element_by_xpath(s1).click()
          if judgementrandom=="s2":
           driver.find_element_by_xpath(s2).click()
          if judgementrandom=="s3":
           driver.find_element_by_xpath(s3).click()
          if judgementrandom=="s4":
           driver.find_element_by_xpath(s4).click()
          if judgementrandom=="s5":
           driver.find_element_by_xpath(s5).click()
          if judgementrandom=="s6":
           driver.find_element_by_xpath(s6).click()
          if judgementrandom=="s7":
           driver.find_element_by_xpath(s7).click()
          if judgementrandom=="s8":
           driver.find_element_by_xpath(s8).click()
          if judgementrandom=="s9":
           driver.find_element_by_xpath(s9).click()
          if judgementrandom=="s10":
           driver.find_element_by_xpath(s10).click()
          if judgementrandom=="s11":
           driver.find_element_by_xpath(s11).click()
          if judgementrandom=="s12":
           driver.find_element_by_xpath(s12).click()
          ss1 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[1]/td[1]/input"
          ss2 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[2]/td[1]/input"
          ss3 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[3]/td[1]/input"
          ss4 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[4]/td[1]/input"
          ss5 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[5]/td[1]/input"
          ss6 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[6]/td[1]/input"
          ss7 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[7]/td[1]/input"
          ss8 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[8]/td[1]/input"
          ss9 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[9]/td[1]/input"
          ss10 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[10]/td[1]/input"
          ss11 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[11]/td[1]/input"
          ss12 = "/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/table[7]/tbody/tr[4]/td[2]/table/tbody/tr/td/table/tbody/tr[12]/td[1]/input"
          if judgementrandom2 == "Yes":
              xpaths2 = [ss1, ss2, ss3, ss4, ss5, ss6, ss7, ss8, ss9, ss10, ss11, ss12]
              random.shuffle(xpaths2)
              driver.find_element_by_xpath(xpaths2[0]).click()
          if judgementrandom2 == "s1":
              driver.find_element_by_xpath(ss1).click()
          if judgementrandom2 == "s2":
              driver.find_element_by_xpath(ss2).click()
          if judgementrandom2 == "s3":
              driver.find_element_by_xpath(ss3).click()
          if judgementrandom2 == "s4":
              driver.find_element_by_xpath(ss4).click()
          if judgementrandom2 == "s5":
              driver.find_element_by_xpath(ss5).click()
          if judgementrandom2 == "s6":
              driver.find_element_by_xpath(ss6).click()
          if judgementrandom2 == "s7":
              driver.find_element_by_xpath(ss7).click()
          if judgementrandom2 == "s8":
              driver.find_element_by_xpath(ss8).click()
          if judgementrandom2 == "s9":
              driver.find_element_by_xpath(ss9).click()
          if judgementrandom2 == "s10":
              driver.find_element_by_xpath(ss10).click()
          if judgementrandom2 == "s11":
              driver.find_element_by_xpath(ss11).click()
          if judgementrandom2 == "s12":
              driver.find_element_by_xpath(ss12).click()

          driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxPreReceiptNo_txtPreReceiptNo").clear()
          name =driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxPreReceiptNo_txtPreReceiptNo")
          name.send_keys(receiptno)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP6.jpg".format(caseno))
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnShowOtherService1").click()
          time.sleep(3)
          try:
           if select1=="random":
            sel=WebDriverWait(driver,10).until(lambda x:x.find_element("name","htmlPageTopContainer_ecardForm_boxLabXray_dtLabXray_dtLabXrayTRRow0_dtLabXrayTDRow1_selLabXray_0"))
            alloptions1=sel.find_elements_by_tag_name("option")
            extraform1=[]
            for option1 in alloptions1:
             extraform1.append(option1.get_attribute("value"))
            del extraform1[0]
            random.shuffle(extraform1)
            Select(sel).select_by_value("{0}".format(extraform1[0]))
           else:
            sel = WebDriverWait(driver, 10).until(lambda x: x.find_element("name",
                                                                             "htmlPageTopContainer_ecardForm_boxLabXray_dtLabXray_dtLabXrayTRRow0_dtLabXrayTDRow1_selLabXray_0"))
            Select(sel).select_by_value("{0}".format(select1))
          except Exception:
              pass
          try:
           if select2=="random":
            sel2=driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxMinorOper_dtMinorOper_dtMinorOperTRRow0_dtMinorOperTDRow1_selMinorOperation_0")
            extraform2=[]
            alloptions=sel2.find_elements_by_tag_name("option")
            for option2 in alloptions:
              extraform2.append(option2.get_attribute("value"))
            del extraform2[0]
            random.shuffle(extraform2)
            Select(sel2).select_by_value("{0}".format(extraform2[0]))
           else:
            sel2 = driver.find_element_by_name(
                  "htmlPageTopContainer_ecardForm_boxMinorOper_dtMinorOper_dtMinorOperTRRow0_dtMinorOperTDRow1_selMinorOperation_0")
            Select(sel2).select_by_value("{0}".format(int(select2)))
          except Exception:
              pass
          try:
           if select3=="random":
            sel3=driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxOperThea_dtOperThea_dtOperTheaTRRow0_dtOperTheaTDRow1_selOperThea_0")
            extraform3=[]
            alloptions=sel3.find_elements_by_tag_name("option")
            for option3 in alloptions:
              extraform3.append(option3.get_attribute("value"))
            del extraform3[0]
            random.shuffle(extraform3)
            Select(sel3).select_by_value("{0}".format(extraform3[0]))
           else:
            sel3 = driver.find_element_by_name(
                  "htmlPageTopContainer_ecardForm_boxOperThea_dtOperThea_dtOperTheaTRRow0_dtOperTheaTDRow1_selOperThea_0")
            Select(sel3).select_by_value("{0}".format(int(select3)))
          except Exception:
              pass
          try:
           if other1=="blank":
              pass
           else:
              selo1=driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_selMedication_0")
              Select(selo1).select_by_value("M9999")
              driver.find_element_by_id(
                  "htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_0").click()
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_0").send_keys(str(other1))
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow2_txtMedicationCharge_0").send_keys(str(other1v))
           if other2=="blank":
              pass
           else:
              selo2=driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_selMedication_1")
              Select(selo2).select_by_value("M9999")
              driver.find_element_by_id(
                  "htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_1").click()
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_1").send_keys(str(other2))
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow2_txtMedicationCharge_1").send_keys(str(other2v))
           if other3=="blank":
              pass
           else:
              selo3=driver.find_element_by_name("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_selMedication_2")
              Select(selo3).select_by_value("M9999")
              driver.find_element_by_id(
                  "htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_2").click()
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow1_txtMedicationOther_2").send_keys(str(other3))
              driver.find_element_by_id("htmlPageTopContainer_ecardForm_boxExtraMedication_dtExtraMedication_dtExtraMedicationTRRow0_dtExtraMedicationTDRow2_txtMedicationCharge_2").send_keys(str(other3v))
          except Exception:
              pass
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP7.jpg".format(caseno))
          driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSave1").click()
          time.sleep(3)
          alerttext1=driver.switch_to.alert.text
          if alerttext1=="Please refer to the message in red at the top of the page.":
           driver.switch_to.alert.accept()
           time.sleep(3)
           jinggao=driver.find_element_by_xpath("//*[@id='divhtmlPageTopContainer_ecardForm_errorMessage']").text
          #print (jinggao)
          #alerttext1=driver.switch_to.alert.text
          #if alerttext1=="Please refer to the message in red at the top of the page.":
           if jinggao=="Please specify $ charge to selected service":

             print("write")
             #time.sleep(3)
             #driver.switch_to.alert.accept()
             time.sleep(3)
             target=driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[5]/input")
             driver.execute_script("arguments[0].scrollIntoView();",target)
             driver.find_element_by_xpath("//*[@id='htmlPageTopContainer_ecardForm_boxLabXray_dtLabXray_dtLabXrayTRRow0_dtLabXrayTDRow2_txtLabXrayCodeInput_0']").send_keys("20")
             driver.find_element_by_xpath("//*[@id='htmlPageTopContainer_ecardForm_boxLabXray_dtLabXray_dtLabXrayTRRow0_dtLabXrayTDRow2_txtLabXrayCodeInput_0']").clear()
             time.sleep(5)
             sel=WebDriverWait(driver,10).until(lambda x:x.find_element("xpath","/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr/td/div/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]/select"))
             if select1=="random":
              Select(sel).select_by_value("{0}".format(extraform1[0]))
             else:
              Select(sel).select_by_value("{0}".format(select1))
             time.sleep(5)
             driver.find_element_by_xpath("//*[@id='htmlPageTopContainer_ecardForm_boxLabXray_dtLabXray_dtLabXrayTRRow0_dtLabXrayTDRow4_txtLabXrayCharge_0']").send_keys("20")
             driver.find_element_by_name("htmlPageTopContainer_ecardForm_btnSave1").click()
             time.sleep(3)
             driver.switch_to.alert.accept()
           if jinggao=="Receipt No. already exists.":
              driver.switch_to.alert.accept()



          else:
             print("no jinggao")
             driver.switch_to.alert.accept()
          time.sleep(3)
          #driver.switch_to.alert.accept()
          num1=driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td[2]").text
          print (num1)
          num2=driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[2]/form/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr/td/table[3]/tbody/tr/td/table/tbody/tr/td/table[2]/tbody/tr/td[2]/b/font").text
          print (num2)
          time.sleep(1)
          driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEP8.jpg".format(caseno))
          driver.find_element_by_xpath("/html/body/table/tbody/tr[3]/td[1]/table/tbody/tr[6]/td/a/img").click()
         else:
             print("error execute input")
             break

        except BaseException:




            print("fail")
            outputs.append("fail")
            fugai.append("Done")
            self.driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEPfail.jpg".format(caseno))
            hyperlink.append("c:/test_project/clinic/report&screenshots/{0}/failcase/{1}/TC{1}_STEPfail.jpg".format(otime,caseno))
            shutil.move("c:/test_project/debug/{0}/".format(caseno),"c:/test_project/clinic/report&screenshots/{0}/failcase/".format(otime))
            self.driver.quit()
        else:
           print("success")
           fugai.append("Done")
           outputs.append("success"+num2)
           #self.driver.get_screenshot_as_file("c:/test_project/debug/{0}/TC{0}_STEPpass.jpg".format(caseno))
           hyperlink.append("c:/test_project/clinic/report&screenshots/{0}/passcase/{1}/TC{1}_STEP8.jpg".format(otime,caseno))
           os.remove("c:/test_project/debug/{0}/TC{0}_STEP1.jpg".format(caseno))
           os.remove("c:/test_project/debug/{0}/TC{0}_STEP2.jpg".format(caseno))
           os.remove("c:/test_project/debug/{0}/TC{0}_STEP3.jpg".format(caseno))
           os.remove("c:/test_project/debug/{0}/TC{0}_STEP4.jpg".format(caseno))
           os.remove("c:/test_project/debug/{0}/TC{0}_STEP5.jpg".format(caseno))
           shutil.move("c:/test_project/debug/{0}/".format(caseno),
                       "c:/test_project/clinic/report&screenshots/{0}/passcase/".format(otime))
           driver.quit()
     cd=len(hyperlink)
     wb = copy(rb)

     w_sheet = wb.get_sheet(3)
     for tmp0 in range(0,cd):
      w_sheet.write(4+tmp0,7,xlwt.Formula('HYPERLINK("{0}";"{0}")'.format(hyperlink[tmp0])))
     for tmp1 in range(0,cd):
       w_sheet.write(4+tmp1,8,hyperlink[tmp1])  
     for tmp in range(0, cd):
      w_sheet.write(4+tmp,6,outputs[tmp])
     for tmp2 in range(0,cd):
         w_sheet.write(4 + tmp2, 12, fugai[tmp2])
     wb.save('c:/test_project/debug/report.xls')
     shutil.copy("c:/test_project/debug/report.xls", "c:/test_project/clinic/report&screenshots/{0}/".format(otime))

if __name__ == "__main__":
    unittest.main(warnings='ignore')


