# -*- coding: utf-8 -*-
import csv
from time import sleep
from selenium import webdriver
from selenium.webdriver.support.ui import Select

chromedriver = 'chromedriver'
options = webdriver.ChromeOptions()
prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': 'Data path'}
options.add_experimental_option('prefs', prefs)
driver = webdriver.Chrome(chromedriver, chrome_options=options)

driver.get('https://erdb.epa.gov.tw/DataRepository/EnvMonitor/AirQualityMonitorMonData.aspx?topic1=%u5927%u6c23&topic2=%u74b0%u5883%u53ca%u751f%u614b%u76e3%u6e2c&subject=%u7a7a%u6c23%u54c1%u8cea')
Area = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_ucSearchCondition_ddlAirArea'))
Area.select_by_index(6)
for i in range(10) :
    Year = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_ucSearchCondition_ddlYearE'))
    Year.select_by_index(i)
    sleep(0.5)
    for k in range(12) :
        Month = Select(driver.find_element_by_id('ctl00_ContentPlaceHolder1_ucSearchCondition_ddlMonthE'))
        Month.select_by_index(k)
        sleep(0.5)
        Search = driver.find_element_by_id('ctl00_ContentPlaceHolder1_imgSearch').click()
        sleep(0.5)
        Download = driver.find_element_by_id('ctl00_ContentPlaceHolder1_ucShareAndExport_ibtnExcel').click()
        sleep(0.5)

