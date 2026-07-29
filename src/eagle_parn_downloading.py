#!/usr/bin/env python
# coding: utf-8

# # Pull Holdings and Derivative Data from Eagle

# How to wait until Element is Visible in Selenium Python
#
# https://pythonexamples.org/python-selenium-wait-until-element-is-visible/

# In[1]:

print("\n\n#####################################################")
print("#                                                   #")
print("#      START 1/4 derv_checker_downloading.py   X    #")
print("#                                                   #")
print("#####################################################\n\n")

# create a time difference function
import time

start_time = time.time()
start_time_overlord = start_time
print("Importing libraries ...")

# load libraries
from utilities import timediff
import pandas as pd
import numpy as np
import os
import csv
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# import networkdays as networkdays

print(f"Importing libraries completed: {timediff(start_time, time.time())}", "\n")


# In[3]:


start_time = time.time()
print(f"Setting up paths ...")

# set paths to the driver, urls, and report parameters
os.environ["PATH"] = r"C:/SeleniumDrivers"  # + os.pathsep + os.getenv("PATH")
# https://stackoverflow.com/questions/61213005/modify-beginning-of-path-variable-with-os-environ-in-python
eagle_portal = r"https://eagleportal.prescient.co.za"
url_default = eagle_portal + r"/Default.aspx"
url_derv = eagle_portal + r"/Queries/Query.aspx?rpt=DerivativeExposure"
url_parn = eagle_portal + r"/Queries/Query.aspx?rpt=PortfolioAnalytics"
pth = r"P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm"  # variables stored here

print(f"Setting up paths completed: {timediff(start_time, time.time())}", "\n")


# In[4]:


start_time = time.time()
print("Collecting input data ...")

# get user details, report date, and fund names from the py_report.xlsm sheet
import xlwings as xw

wb = xw.Book(
    r"P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm"
)  # open calc workbook as an object

# get variables from source sheet and then close the sheet
when = wb.sheets["parn"].range("B1").value
when_fmt = wb.sheets["parn"].range("B1").value.strftime("%#m/%#d/%Y")
aladdin = wb.sheets["creds"].range("A1").value
sesame = wb.sheets["creds"].range("A2").value
list_funds = (
    wb.sheets["parn"].range("A:A")[1:].value
)  # https://stackoverflow.com/questions/48513093/xlwings-how-to-select-an-entire-column-without-headers
df_fnds_ = [s for s in list_funds if s != None]
fnds_ = (",").join(df_fnds_)
wb.close()

# set reporting date
day_ = f"{when:%#d}"  # f'{when:%d}' report date with leading zeroes
month_year_ = f"{when:%B}, {when:%Y}"
month_ = f"{when:%b}"
year_ = f"{when:%Y}"

# check inputs
print(
    f" Report date:          {when.strftime('%A %d %B %Y')}",
    "\n",
    f"Portfolios ({len(df_fnds_)}):      {fnds_}",
)
print(f"Collecting input data completed: {timediff(start_time, time.time())}", "\n")


# In[5]:


start_time = time.time()
print("Assigning webdriver ...")

# assign the browser driver
from selenium import webdriver

driver = webdriver.Firefox()

# open the browser on the Eagle web page
driver.get(url_default)  # default page
wait = WebDriverWait(
    driver, 10
)  # https://selenium-python.readthedocs.io/waits.html, max wait for elements to appear

# login with credentials
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_UserName").send_keys(
    aladdin
)
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_Password").send_keys(
    sesame
)
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_LoginButton").click()

print(f"Assigning webdriver completed: {timediff(start_time, time.time())}", "\n")


# In[6]:


start_time = time.time()
print("Downloading and then saving holdings data ...")

##get the PAR-N holdings

# having logged in, open the reporting page
driver.get(url_parn)  # fund holdings page

# switch to the query page
driver.find_element(By.CSS_SELECTOR, "#ModifyLinkLabel").click()

edit_criteria_window = driver.window_handles[0]  # save curent window handle
# https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python

# update the calendar
driver.find_element(
    By.CSS_SELECTOR,
    'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]',
).click()
driver.find_element(
    By.CSS_SELECTOR,
    'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_NMC"]',
).click()
lmonth_selector = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_PMC"]',
        )
    )
)
while (
    driver.find_element(
        By.XPATH,
        '//td[@id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_TC"]',
    ).text
    != month_year_
):
    lmonth_selector.click()
day_selector = driver.find_element(
    By.XPATH,
    f'//td[@class="dxeCalendarDay"][text()={day_}] | //td[@class="dxeCalendarDay dxeCalendarWeekend"][text()={day_}]',
)
day_selector.click()

# get the web element for the FUND LIST and assign values to it
fund_selector = driver.find_element(
    By.CSS_SELECTOR,
    'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]',
)
driver.execute_script(f'arguments[0].value = "{fnds_}";', fund_selector)

# get the web element of the 'Submit' button and then click it
submit_button = driver.find_element(
    By.CSS_SELECTOR, 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]'
)
submit_button.click()

# Wait for and then click the export button and then the xls download button
# https://stackoverflow.com/questions/56085152/selenium-python-error-element-could-not-be-scrolled-into-view
WebDriverWait(driver, 1000).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]'))
).click()
WebDriverWait(driver, 1000).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'td[id="ExportMnu_DXI4_T"]'))
).click()

time.sleep(30)  # wait for 30 seconds after the holdings data downloads

print(
    f"Downloading and then saving holdings data completed: {timediff(start_time, time.time())}",
    "\n",
)


# In[7]:


start_time = time.time()
print("Downloading and then saving derivative data ...")

# get the derivative metrics

# having logged in, open the reporting page
driver.get(url_derv)

# switch to the query page
driver.find_element(By.CSS_SELECTOR, "#ModifyLinkLabel").click()

edit_criteria_window = driver.window_handles[0]  # save curent window handle
# https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python

# update the 'from' calendar
driver.find_element(
    By.CSS_SELECTOR,
    'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]',
).click()
driver.find_element(
    By.CSS_SELECTOR,
    'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_NMC"]',
).click()
lmonth_selector = wait.until(
    EC.element_to_be_clickable(
        (
            By.CSS_SELECTOR,
            'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_PMC"]',
        )
    )
)
while (
    driver.find_element(
        By.XPATH,
        '//td[@id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_TC"]',
    ).text
    != month_year_
):
    lmonth_selector.click()
day_selector = driver.find_element(
    By.XPATH,
    f'//td[@class="dxeCalendarDay"][text()={day_}] | //td[@class="dxeCalendarDay dxeCalendarWeekend"][text()={day_}]',
)
day_selector.click()

# get the web element for the fund list and assign values to it
fund_selector = driver.find_element(
    By.CSS_SELECTOR,
    'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]',
)
driver.execute_script(f'arguments[0].value = "{fnds_}";', fund_selector)

# get the web element of the 'Submit' button and then click it
submit_button = driver.find_element(
    By.CSS_SELECTOR, 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]'
)
submit_button.click()

# Wait for and then click the export button and then the xls download button
# https://stackoverflow.com/questions/56085152/selenium-python-error-element-could-not-be-scrolled-into-view
WebDriverWait(driver, 1000).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]'))
).click()
WebDriverWait(driver, 1000).until(
    EC.element_to_be_clickable((By.CSS_SELECTOR, 'td[id="ExportMnu_DXI4_T"]'))
).click()

print(
    f"Downloading and then saving derivative data completed: {timediff(start_time, time.time())}",
    "\n",
)
print(
    f"Roundtrip time for getting holdings and derivative data: {timediff(start_time_overlord, time.time())}",
    "\n",
)

driver.quit()  # close the web driver

print("\n\n#####################################################")
print("#                                                   #")
print("#       END 1/4 derv_checker_downloading.py   X     #")
print("#                                                   #")
print("#####################################################\n\n")
