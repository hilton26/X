#!/usr/bin/env python
# coding: utf-8

# # Eagle Downloader

# In[15]:


# timekeeper function
import datetime, time

def timediff(start, end, decimals = 1):
    if   int((end - start) / 3600) > 0: # non-zero hours
        return str(int(  (end - start) / 3600))           + 'hr '  + \
               str(int(  (end - start) /   60))           + 'min ' + \
               str(round((end - start) %   60, decimals)) + 'sec'
    elif int((end - start) /   60) > 0: # non-zero hours and minutes
        return str(int(  (end - start) /   60))           + 'min ' + \
               str(round((end - start) %   60, decimals)) + 'sec'
    else:
        return str(round((end - start) %   60, decimals)) + 'sec'


# In[16]:


# libraries, libraries!
start_time          = time.time()
start_time_overlord = start_time
print('Importing libraries ...')

import pandas as pd
import numpy as np
from pathlib import Path
import os
import csv
import win32com.client # to insert date in file downloaded Excel file
from datetime import date, datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException # to trigger when data is not yet available in Eagle

from tqdm import tqdm, notebook # notebook version of tqdm
#" ...any time you see a loop somewhere in your code in you can simply wrap it in either tdqm() or tqdm_notebook() in Jupyter" 

print(f'Importing libraries completed: {timediff(start_time, time.time())}', '\n')


# In[17]:


# set up paths as strings
start_time = time.time()
print(f'Setting up paths ...')

# set paths to the driver, urls, and report parameters
os.environ["PATH"] = r'C:/SeleniumDrivers' # + os.pathsep + os.getenv("PATH")
# https://stackoverflow.com/questions/61213005/modify-beginning-of-path-variable-with-os-environ-in-python
eagle_portal = r'https://eagleportal.prescient.co.za'
url_default  = eagle_portal + r'/Default.aspx'
url_r28      = eagle_portal + r'/Queries/Query.aspx?rpt=Reg28withExposure'
pth          = r'P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm' # variables stored here
pth_dl       = str(Path.home() / 'Downloads')
r28N         = 'Reg 28 Report - Incl Effective Exposure'  # prefix of the file from Eagle

print(f'Setting up paths completed: {timediff(start_time, time.time())}', '\n')


# In[18]:


# utility function to save downloaded file with reporting date
def dater(folder_path, fund_name, dte):
    start_time = time.time()

    fls     = os.listdir(folder_path)
    a       = max([os.path.abspath(os.path.join(folder_path, fl)) for fl in fls if r28N in fl], key = os.path.getmtime)
    z       = 'xlsx' if a[len(a) - 3:] == 'lsx' else 'xls'
    wb      = excel.Workbooks.Open(a)
    wb.Worksheets('Reg 28 Report - Incl Effective ').Range("J1").Value = dte.strftime("%d%b%Y")
    wb.SaveAs(os.path.abspath(os.path.join(folder_path, f'{fund_name} lookthrough {dte.strftime("%d%b%Y")}.{z}')))
    wb.Close()
    
    print(f'Downloaded file found and saved in {timediff(start_time, time.time())}') # time to get file name


# In[19]:


# get data to pass to Eagle
start_time = time.time()
print('Collecting input data ...')

#get user details, report date, and fund names from the py_report.xlsm sheet
import xlwings as xw
wb             = xw.Book(r'P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm') # open calc workbook as an object

# get variables from source sheet and then close the sheet
when           = wb.sheets['downloader'].range('d1').value
when_fmt       = wb.sheets['downloader'].range('d1').value.strftime('%#m/%#d/%Y')  
aladdin        = wb.sheets['creds'].range('A1').value
sesame         = wb.sheets['creds'].range('A2').value
fund_codes     = [s for s in wb.sheets['downloader'].range('A:A')[1:].value  if s != None] # list type
fnds_          = (',').join(fund_codes).upper()                                         # str type
wb.close()

# set reporting date
day_           = f'{when:%#d}' # f'{when:%d}' report date with leading zeroes
month_year_    = f'{when:%B}, {when:%Y}'
month_         = f'{when:%b}'
year_          = f'{when:%Y}'

# check inputs
print(f' Report date   : {when.strftime("%A %d %B %Y")}', '\n', \
      f'Portfolios ({len(fund_codes)}): {fnds_}', \
      )
print(f'Collecting input data completed: {timediff(start_time, time.time())}', '\n')


# In[20]:


# assign web driver
start_time = time.time()
print('Assigning the web driver ...')

# assign the browser driver
from selenium import webdriver
driver = webdriver.Firefox()
    
print(f'Assigning the web driver completed: {timediff(start_time, time.time())}', '\n')


# In[21]:


# iterate through the list of reports to be downloaded from Eagle
start_time = time.time()
print(f'Downloading lookthrough holdings as at {when.strftime("%A %d %B %Y")} ...', '\n')
    
excel = win32com.client.Dispatch("Excel.Application") # start Excel; to insert date in file downloaded from Eagle
excel.DisplayAlerts = False # suppress alerts

funds_not_downloaded = []
sets = 1 # number of sets in which to download the reports from Eagle
for a in notebook.tqdm(range(0, len(fund_codes), sets)):
    fnds_ = (',').join(fund_codes[a:a + sets]).upper()  # Eagle only uses upper case fund codes 

    # open the default page
    driver.get(url_default)          # default page
    wait = WebDriverWait(driver, 10) # https://selenium-python.readthedocs.io/waits.html, max wait for elements to appear

    # login with credentials
    driver.find_element(By.CSS_SELECTOR, '#LoginCtrl_MainLoginControl_UserName'   ).send_keys(aladdin)
    driver.find_element(By.CSS_SELECTOR, '#LoginCtrl_MainLoginControl_Password'   ).send_keys(sesame)
    driver.find_element(By.CSS_SELECTOR, '#LoginCtrl_MainLoginControl_LoginButton').click()

    # having logged in, open the reporting page
    driver.get(url_r28)

    # switch to the query page
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#ModifyLinkLabel'))).click()

    edit_criteria_window = driver.window_handles[0] # save curent window handle
    # https://stackoverflow.com/questions/10629815/how-to-switch-to-new-window-in-selenium-for-python        

    # pass in the reporting date
    driver.find_element(By.CSS_SELECTOR, 'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]').click()
    driver.find_element(By.CSS_SELECTOR, 'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_NMC"]').click()
    lmonth_selector = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_PMC"]')))
    while driver.find_element(By.XPATH,'//td[@id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_DDD_C_TC"]').text != month_year_:
        lmonth_selector.click()
    day_selector = driver.find_element(By.XPATH,f'//td[@class="dxeCalendarDay"][text()={day_}] | //td[@class="dxeCalendarDay dxeCalendarWeekend"][text()={day_}]')
    day_selector.click()

    # pass in the fund code to be looked up
    fund_selector  = driver.find_element(By.CSS_SELECTOR, 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]')
    driver.execute_script(f'arguments[0].value = "{fnds_}";', fund_selector)

    # print current fund and number of funds left
    print(f'{a + 1} {fnds_} {when.strftime("%A %d %B %Y")} loading, {len(fund_codes) - a - 1} ({round(100 - (a + sets) * 100 / len(fund_codes), 2)}%) remaining.')

    # wait for the 'Submit' button to be clickable and then click it
    WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]'))).click()

    #elem = WebDriverWait(driver, 50).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, 'a[id="DataMessageText"]'), "No data returned for the input criteria."))
    
    try:
        #Wait for and then click the export button and then the xls download button
        #https://stackoverflow.com/questions/56085152/selenium-python-error-element-could-not-be-scrolled-into-view
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR,  'a[id="DistrBtn"]'        ))).click()
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'td[id="ExportMnu_DXI4_T"]'))).click()

        # time.sleep(5)             # wait for 5 seconds for the download to complete
        dater(pth_dl, fnds_, when)  # save the report date in the downloaded file
        # print(f'dater() when = {when}')
        
    except TimeoutException:
    #except elem:
        funds_not_downloaded.append(fnds_)
        print(f'{fnds_} lookthrough dated {when.strftime("%d %b %Y")} is not available in Eagle') 
        continue # continue the for loop
        
print('\n', f'Funds not downloaded ({len(funds_not_downloaded)} out of {len(fund_codes)}): {", ".join(funds_not_downloaded)}', '\n')
driver.quit()
excel.DisplayAlerts = True # unspress Windows warning dialogue
#excel.Quit()
os.startfile(os.path.realpath(Path.home() / 'Downloads')) # open the Downloads folder

print(f'Downloading lookthrough holdings for {len(fund_codes) - len(funds_not_downloaded)} funds as \
at {when.strftime("%d %B %Y")} completed: {timediff(start_time, time.time())} \
({(time.time() - start_time) / (len(fund_codes) - len(funds_not_downloaded)): ,.1f}sec/fund)')


# In[22]:


# C:\Users\hilton.netta\Downloads


# In[ ]:




