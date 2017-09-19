from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
import unittest, time, re
import os

class Abn(unittest.TestCase):
    def setUp(self):
        ffprofile = webdriver.FirefoxProfile()
        ffprofile.set_preference("browser.download.dir", os.getcwd()+'\\tmp')
        ffprofile.set_preference("browser.download.folderList",2);
        ffprofile.set_preference("browser.helperApps.neverAsk.saveToDisk", 
            ",application/octet-stream" + 
            ",application/vnd.ms-excel" + 
            ",application/vnd.msexcel" + 
            ",application/x-excel" + 
            ",application/x-msexcel" + 
            ",application/xls" + 
            ",application/vnd.ms-excel" +
            ",application/vnd.ms-excel.addin.macroenabled.12" +
            ",application/vnd.ms-excel.sheet.macroenabled.12" +
            ",application/vnd.ms-excel.template.macroenabled.12" +
            ",application/vnd.ms-excelsheet.binary.macroenabled.12" +
            ",application/vnd.ms-fontobject" +
            ",application/vnd.ms-htmlhelp" +
            ",application/vnd.ms-ims" +
            ",application/vnd.ms-lrm" +
            ",application/vnd.ms-officetheme" +
            ",application/vnd.ms-pki.seccat" +
            ",application/vnd.ms-pki.stl" +
            ",application/vnd.ms-word.document.macroenabled.12" +
            ",application/vnd.ms-word.template.macroenabed.12" +
            ",application/vnd.ms-works" +
            ",application/vnd.ms-wpl" +
            ",application/vnd.ms-xpsdocument" +
            ",application/vnd.openofficeorg.extension" +
            ",application/vnd.openxmformats-officedocument.wordprocessingml.document" +
            ",application/vnd.openxmlformats-officedocument.presentationml.presentation" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slide" +
            ",application/vnd.openxmlformats-officedocument.presentationml.slideshw" +
            ",application/vnd.openxmlformats-officedocument.presentationml.template" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" +
            ",application/vnd.openxmlformats-officedocument.spreadsheetml.template" +
            ",application/vnd.openxmlformats-officedocument.wordprocessingml.template" +
            ",application/x-ms-application" +
            ",application/x-ms-wmd" +
            ",application/x-ms-wmz" +
            ",application/x-ms-xbap" +
            ",application/x-msaccess" +
            ",application/x-msbinder" +
            ",application/x-mscardfile" +
            ",application/x-msclip" +
            ",application/x-msdownload" +
            ",application/x-msmediaview" +
            ",application/x-msmetafile" +
            ",application/x-mspublisher" +
            ",application/x-msschedule" +
            ",application/x-msterminal" +
            ",application/x-mswrite" +
            ",application/xml" +
            ",application/xml-dtd" +
            ",application/xop+xml" +
            ",application/xslt+xml" +
            ",application/xspf+xml" +
            ",application/xv+xml" +
            ",application/excel")

        self.driver = webdriver.Firefox(ffprofile)
        self.driver.implicitly_wait(30)
        self.base_url = "https://b2b.abn.ru/"
        self.verificationErrors = []
        self.accept_next_alert = True
    
    def test_abn(self):
        driver = self.driver
        driver.get(self.base_url + "/index.php")
        driver.find_element_by_id("login").clear()
        driver.find_element_by_id("login").send_keys("146348_1509_130006")
        driver.find_element_by_id("pass").clear()
        driver.find_element_by_id("pass").send_keys("146348_12345")
        driver.find_element_by_name("Submit").click()
#       driver.find_element_by_css_selector("div.image").click()
#       driver.find_element_by_css_selector("a.download_ico.xls").click()
#       driver.get(self.base_url + "get_price_list.php?action=proccess&login=72392&reqstr=13_6575,6582,6584,6608,6638,6696,6665,6616_1_1_1_0_0&fileMode=xls")  --
#       driver.get(self.base_url + "get_price_list.php?action=proccess&login=72392&reqstr=-1_-1_1_1_1_0_0&fileMode=xls")
        driver.get(self.base_url + "/get_price_list.php?action=proccess&login=72392&reqstr=-1_-1_6701,6575,6582" + 
            ",6630,6710,6584,7044,6608,6657,7047,6638,6618,6619,6681,6639,7365,6621,6641,6622,7081,6589,6660" + 
            ",6610,6624,6591,7369,6625,6684,6578,6696,6627,7413,6665,6667,6646,7410,6596,7199,6671,6658,6672" + 
            ",6598,6704,6698,7366,6649,6632,6687,6651,7364,6690,6614,6615,6602,6616,6635,7352,6708,6712,7378" + 
            ",6599,7348,6692,7362_1_1_1_0_0&fileMode=xls")
        time.sleep(200)     


    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
        except NoSuchElementException: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
