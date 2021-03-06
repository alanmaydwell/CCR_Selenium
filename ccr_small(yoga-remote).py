#!/usr/bin/evn python

#Standard Selenium imports
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, NoSuchElementException
from selenium.webdriver.support.ui import Select, WebDriverWait
#Firefox profile support
from selenium.webdriver.firefox.webdriver import FirefoxProfile
import sys

print "Starting"

##firefox_profile = r"E:\TEST\Firefox_Profiles\james_conlon_ccr_aws"
firefox_profile = "/home/user/Selenium/james_conlon_ccr_aws(with_firebug)"
url = "http://ccr.hosting.legalaid.technology/ccr/AutoLogin"
selenium_grid_url = "http://localhost:32768/wd/hub"

#Access CCR using Firefox profile
ffp_object  = FirefoxProfile(firefox_profile)

#Direct driver
##driver = webdriver.Firefox(ffp_object)

#Remote driver - Selenium grid access
capabilities = webdriver.DesiredCapabilities.FIREFOX.copy()
driver = webdriver.Remote(desired_capabilities=capabilities,
                          browser_profile=ffp_object,
                          command_executor=selenium_grid_url)

#Open the application under test
driver.get(url)
#Wait for page
WebDriverWait(driver, 20).until(lambda driver:
                        "Search For Claims" in driver.page_source
                        or "<h2>Login Error</h2>" in driver.page_source)
#Search for a case
driver.find_element_by_id("caseNumber").clear()
driver.find_element_by_id("caseNumber").send_keys("T20132011")
driver.find_element_by_xpath("//input[@value='Search']").click()
#Wait for sarch results
WebDriverWait(driver,10).until(lambda driver:
                "Search Results" in driver.page_source
                or "No claims found" in driver.page_source,driver)
#Crude count of number of results
print "Crude count of search results:",
count = len(driver.find_elements_by_class_name("dataRowo"))
print count

#Simple pass/fail (1-fail, 0 pass)
result = 1
if count==2:
    result = 0
#Issue result as exit code
try:
    sys.exit(result)
except SystemExit as e:
    print "Finished with exit code:", result