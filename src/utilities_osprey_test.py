#!/usr/bin/env python
# coding: utf-8

# In[109]:


# precursors to osprey()
import time
start_time_osprey_test = time.time()

from datetime import datetime
import pandas as pd

df1 = pd.read_excel('P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm', sheet_name = 'creds', usecols = 'A', header = None, nrows = 2)
rpt_type = 'parn'
fund     = 'GEMSMEDC' # funds get passed to the application in step (11)
funds    = '3BBCIINC,ABMMBND,ADRRC,ADVMB,AMMARF,AMPBQP,ASBTOS,ASHFLX,BCIFIF,BPROV,\
CCNPF,CMPFFLEX,CMPFINC,CSIRBQP,ECICBALC,ELCIPF,ENGENIP,ENGENMBF,FEMPBF,GAEMBF,\
GEMSMEDC,GMRETF,GMRETF2,GRFINV,GTCWP2,HOLADC,HOLDINC,HOLIDC,HOLYPF,HOSMED,IJGCOR,IJGIPF,\
IMPALA,IMPBAL_C,IMPREF_C,IPIPF,ISPFP,LEZAFFI,LEZALDI,LIBTAA,LPIIFTAA,MASAINC,MASASI,MEDINC,\
MOMBBF,MOMFLX,MOMIPF,MOMPRET,MOMTAAHI,MOMTAALI,MOMTAAMI,MULTICH,MWPFEQU,MWPFILB,MYQIP,\
NESEQU,NFMWAGG,NGKINC,OMMAIF,PABS,PBNDQ,PCAEF,PCBF,PGCBF,PCCEF,PCEQTF,PCGEARF,PCGEF,PCSHQ,\
PEQ,PEQF,PETFIP,PEYF,PFFIF,PGCEF,PGEMZAR,PGIPFA,PGPCEM_C,PGPCGE_C,PGPCZAR,PGPGARF,PGPGBF_C,\
PGPGIF_C,PGPRF_C,PICPROV,PIF,PIMBAL,PIPF,PIPFP,PLMED,PLPRNA,POIF_C,PORFIP,PPEF,PPOS,PPSBAL_C,\
PPSFLEX,PPSNAM,PPSRBNQ,PPWEQU,PRPABF,PRPDBF,PSIF,PSILB,PSPAM,PSSPFEQU,PSTIF,QIFFGIF,RETFGD,\
RETFPE,RETFPP,RETPIE,SAAMCAU,SAAMINC,SAAMMOD,SABCCSHF,SDINCPM,SILAIF,SILEQF,SISLP,SMMAIF,\
SMMIBF,SMMRRF,STBFIP,TRFINC,TRFWLTH,UCT3YQ,UCTINC,UCTRFBAL,UCTRFINC,UNISABAL,USAIPF,UWRFCON,UWRFGRO,VMPTAA_C'
d_from   = datetime(2025, 9, 16).date()
d_to     = datetime(2025, 9, 16).date()
name     = 'Utilities_Test'
sfx      = 'csv'
al       = df1.iloc[0,0]
xe       = df1.iloc[1,0]


# In[110]:


# # an Eagle report lookup function, given eight parameters
# # (0) function definition
# def osprey(rpt_type, funds, d_from, d_to, name, sfx, al, xe):
# # def osprey(rpt_type = 'r28i', funds = 'PABS, SMMAIF' as string, d_from as datetime, d_to as datetime, name, sfx = 'csv' as string, 
         # al = 'username' as string, xe = 'psw' as string):
    # """Downloads a report from the online fund accounting system for a specified report type, funds, format, and dates

    # Args:
    #   rpt_type: A report name under Queries of the Eagle online application, including r28i, parn, derv, trad, scty, dflw, utps, fnav, tcrf, and cact
    #   funds: A comma-separated string of fund codes, including the ones appended with "_C", .e.g., 'PABS, SMMAIF'
    #   d_from: A start date for the report in datetime format, e.g., datetime(2025,5,1)
    #   d_to: An end date for the report in datetime format, e.g., datetime(2025,5,30)
    #   name: A descriptive name to be added to the downloaded report to make it more identifiable
    #   sfx: A file name extension specifying the file format, i.e., 'xls' or 'csv'
    #   al: The user name for the online application
    #   xe: The user credential for the online application

    # Returns:
    #   A downloaded report in the local Downloads folder renamed to identify it.
    # """

# (1) eagle report types, their short codes, and their URLs
eagle_root = r"https://eagleportal.prescient.co.za/Queries/Query.aspx?rpt="
report_types_dict = {
    "r28i": [
        "Reg 28 Report - Incl Effective Exposure",
        eagle_root + "Reg28withExposure",
    ],
    "parn": ["Portfolio Analytics Report - New", eagle_root + "PortfolioAnalytics"],
    "derv": ["Derivative Exposure", eagle_root + "DerivativeExposure"],
    "trad": ["Trades Report", eagle_root + "TRANSACTION"],
    "scty": ["Security Cross Reference", eagle_root + "SecurityCrossRef"],
    "dflw": ["Daily Flows", eagle_root + "FLOWS"],
    "utps": ["Unit Trust Prices", eagle_root + "UTPRICES"],
    "fnav": ["Fund Net Asset Value", eagle_root + "NetAsset"],
    "tcrf": ["Trades Cross Reference", eagle_root + "TRADES%20REFERENCE"],
    "cact": ["Cash Activity Details", eagle_root + "CSHACTIVITY"],
}


# In[111]:


# (2) load libraries
import time
start_time_osprey = time.time()
from datetime import datetime, timedelta
from utilities import timediff, latest_file
import os
from pathlib import Path
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException # to handle the eagleportal.prescient.co.za alerts
from selenium.common.exceptions import TimeoutException
# https://stackoverflow.com/questions/38022658/selenium-python-handling-no-such-element-exception
# https://www.selenium.dev/selenium/docs/api/py/common/selenium.common.exceptions.html

# (2a) for fnav rpt_type, first remove "_C" from the list of funds else the FNAV report will return "No data returned for the input criteria."
if rpt_type =='fnav':
    sfx   = "csv"
    lkup  = pd.DataFrame(pd.Series(funds.split(',')), columns = ['funds_ante']) # create a dataframe to look up before and after fund codes
    lkup['funds_post'] = lkup['funds_ante'].apply(lambda x: x.replace("_C", "") if x.endswith("_C") else x)
    funds = ','.join(lkup['funds_post'].astype(str))

print(f' rpt_type: {rpt_type} \n sfx: {sfx} \n fund: {fund} \n funds: {funds}')


# In[112]:


# (3) set the path to the web driver, urls, and to the report parameters
import os
# os.environ["PATH"] = r"C:/SeleniumDrivers"  # + os.pathsep + os.getenv("PATH")
t = "0" if sfx == "csv" else "4"            # report format: DXI4[0] for .xls[.csv]

# (4) assign the browser driver
from selenium import webdriver
driver = webdriver.Firefox()

# (5) open the browser on the default web page
url_default = r"https://eagleportal.prescient.co.za/Default.aspx"
driver.get(url_default)  # default page
wait = WebDriverWait(driver, 10)  # https://selenium-python.readthedocs.io/waits.html, max wait for elements to appear

# (6) log in
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_UserName").send_keys(al)
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_Password").send_keys(xe)
driver.find_element(By.CSS_SELECTOR, "#LoginCtrl_MainLoginControl_LoginButton").click()


# In[113]:


# (7) having logged in, switch to the selected report page
report_link = report_types_dict[rpt_type][1]
driver.get(report_link)  # a hyperlink for the report page selected in the function osprey()

# (7(a)) test for the presence of an alert
# this solution from Gemini prompt 17 Sep 2025: "python selenium test for the presence of alert text"
try:
    WebDriverWait(driver, 3).until(EC.alert_is_present())
    alert = driver.switch_to.alert
    alert.accept() # Or alert.dismiss()
except TimeoutException:
    print("No TimeoutException alert")
    pass
except NoAlertPresentException:
    print(" TimeoutException or NoAlertPresentException alert.")
    pass

###### ALTERNATIVE TEST FOR ALERT
# alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
# alert_text = alert.text
# print(f"Alert text: {alert_text}")
# alert.accept()
###### ALTERNATIVE TEST FOR ALERT END


# In[114]:


# (8) switch to the query fields of that report page
driver.find_element(By.CSS_SELECTOR, "#ModifyLinkLabel").click()

######## PREVIOUS ERROR AT (8) switch to the query page
# submit_button = driver.find_element(By.CSS_SELECTOR, 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]')
# UnexpectedAlertPresentException: Alert Text: SELECT method of object CAnalsReport returned FALSE indicating an error.
# Please check the  log file for more information.
# Portal message from eagleportal.prescient.co.za:
#  "SELECT method of object CAnalsReport returned FALSE indicating an error. Please check the log file for more information."
#  Message: Unexpected alert dialog detected. Performed handler "dismiss"
######## PREVIOUS ERROR END


# In[115]:


# (9) update the FROM calendar
clue_from_calendar        = 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_I"]'
clue_from_calendar_button = 'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_From_B-1"]'
date_selector_fr = driver.find_element(By.CSS_SELECTOR, clue_from_calendar)  # FROM date element
driver.execute_script(f'arguments[0].value = "{d_from.strftime("%#m/%#d/%Y")}";', date_selector_fr)  # FROM date without leading zeroes
driver.find_element(By.CSS_SELECTOR, clue_from_calendar).click()  # click inside FROM calendar
driver.find_element(By.CSS_SELECTOR, clue_from_calendar_button).click()  # update the FROM calendar

# (10) if it exists, update the TO calendar
try:  # https://stackoverflow.com/questions/38022658/selenium-python-handling-no-such-element-exception
    clue_to_calendar        = 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_I"]'
    clue_to_calendar_button = 'td[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_DATE1_DateCtrl_To_B-1"]'
    date_selector_to = driver.find_element(By.CSS_SELECTOR, clue_to_calendar)  # calendar
    driver.execute_script(f'arguments[0].value = "{d_to.strftime("%m/%d/%Y")}";', date_selector_to)
    driver.find_element(By.CSS_SELECTOR, clue_to_calendar).click() # click inside TO calendar
    driver.find_element(By.CSS_SELECTOR, clue_to_calendar_button).click()  # update the TO calendar
except NoSuchElementException:  # in the event that the selected report does not have a "to" calendar
    pass


# In[116]:


# (11) get the web element for the FUND LIST and assign values to it
clue_fund_selector = 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedIds"]'
fund_selector = driver.find_element(By.CSS_SELECTOR, clue_fund_selector)
driver.execute_script(f'arguments[0].value = "{funds}";', fund_selector) # send fund codes to fund selector

# (12) click the table header where "Entity ID" resides
clue_entity_ID = 'table[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedItemsGrid_DXHeaderTable"]'
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, clue_entity_ID))).click()  # fund code banner

######## ERROR AT (12): 
# ElementClickInterceptedException: Message: 
# Element <table id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedItemsGrid_DXHeaderTable" 
# class="dxgvTable_Eagle"> is not clickable at point (637,145) because another element 
# <div id="ctl00_c_qc_QueryInputs_QueryInputsPopup_FUND0_SelectedItemsGrid_LD" class="dxgvLoadingDiv_Eagle"> obscures it
########

time.sleep(5)  # arbitrary 5 second wait


# In[117]:


# (13) get the web element of the 'Submit' button and then click it
clue_submit_button = 'input[id="ctl00_c_qc_QueryInputs_QueryInputsPopup_RunBtn"]'
submit_button = driver.find_element(By.CSS_SELECTOR, clue_submit_button)
submit_button.click()

# (13(a)) alert pop-up management
try:
    WebDriverWait(driver, 3).until(EC.alert_is_present())
    alert = driver.switch_to.alert
    alert.accept() # or alert.dismiss()
except TimeoutException:
    print("No TimeoutException alert")
    pass
except NoAlertPresentException:
    print("No NoAlertPresentException alert.")
    pass

# ###################
# ERROR MANAGEMENT

    # Error types:

    # when a report is not available after clicking the 'Submit' button, variation 1:
    # table id="c_InboxGrid_DXMainTable", td class="dxgv"
    # <div>No data to display.</div>

    # when a report is not available after clicking the 'Submit' button, variation 2:
    # <span id="DataMessageText">No data returned for the input criteria.</span>

    # when an unknown fund code was submitted
    # <span id="DataMessageText">All required criteria have not been selected. Select criteria above to view data.</span>

# ###################


# In[118]:


# # (14)-(17) click the export button, close the webdriver, and save the download wit a distinguishable name
# try:
# (14) Wait for and then click the export button and then the report download button
# https://stackoverflow.com/questions/56085152/selenium-python-error-element-could-not-be-scrolled-into-view
WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[id="DistrBtn"]'))).click()
WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.CSS_SELECTOR, f'td[id="ExportMnu_DXI{t}_T"]'))).click()

# (14(a)) alert pop-up management
try:
    WebDriverWait(driver, 3).until(EC.alert_is_present())
    alert = driver.switch_to.alert
    alert.accept() # or alert.dismiss()
except TimeoutException:
    print("No TimeoutException alert")
    pass
except NoAlertPresentException:
    print("No NoAlertPresentException alert.")
    pass

time.sleep(5)  # wait for 5 seconds after the data downloads

# Pop-up: "Cancel All Downloads?"
# "If you cancel now, 1 download will be canceled. Are you sure you want to exit?"
# Two options: "Cancel 1 download" and "Don't exit"


# In[ ]:


# (15) having downloaded the requested report, close the web driver
driver.quit()
# print(f'Roundtrip time for getting holdings and derivative data: {timediff(start_time_osprey, time.time())}', '\n')

# (16) find the latest downloaded file and rename it and set the input variables for the latest file
folder_path   = str(Path.home() / "Downloads")
file_type     = sfx
to_date       = " to " + d_to.strftime("%d%b%Y") if d_to != d_from else ""
new_file_name = f'{rpt_type.upper()} {name}({len(funds.split(","))}) {d_from.strftime("%d%b%Y")}{to_date}'

# (17) run the file name change function
latest_file(folder_path, file_type, new_file_name)  # gets the latest file of that type in the given folder and renames it to new_file_name

# (17(a)) for fnav report, convert fund codes back to include "_C" suffix
if rpt_type == 'fnav':
    filen = os.path.join(folder_path,new_file_name + f'.{sfx}') # full path name of the nav file
    fnav = pd.read_csv(filen) # dataframe the fnav file
    fnav = fnav.merge(lkup, how = 'left', left_on = 'NAV Entity ID', right_on = 'funds_post') # merge the fnav and lookup dataframes
    fnav['NAV Entity ID'] = fnav['funds_ante'] # recover the original fund names
    fnav.drop(columns = ['funds_ante','funds_post'], axis = 1, inplace = True) # drop the merged lookup columns
    fnav.to_csv(filen, index = False) # resave the NAV file

# open the file
# os.system(f'start EXCEL.EXE "{os.path.join(folder_path, new_file_name)}"')
# https://stackoverflow.com/questions/35940748/use-python-to-launch-excel-file

# print(f"  {timediff(start_time_osprey, time.time())} to download the {rpt_type.upper()} report","\n",)
        
# except Exception as e:
#     print(e)
#     print('\n', f"  Report not completed. {timediff(start_time_osprey, time.time())} to download the {rpt_type.upper()} report","\n",)


# In[ ]:


print(f"  {timediff(start_time_osprey_test, time.time())} roundrip for {len(funds)} {rpt_type} reports \n",)


# In[ ]:


# !jupyter nbconvert --to script utilities_osprey_test.ipynb # convert from .ipynb to .py

