#!/usr/bin/env python
# coding: utf-8

# In[1]:


# libraries, libraries!
import time

start_time_portal = time.time()

import pandas as pd
from datetime import datetime, timedelta
import os
from tqdm import tqdm
from utilities import timediff, latest_file
from constants import pthPtl, ptl_login, ptl_b_rpt, pth_dl, p_al, p_xe

# some constants
breach_types = [
    "new",
    "open",
    "unclassified",
]  # in small letters; here, exclude "Resolved"
file_path = os.path.join(
    pthPtl, f"{datetime.today().strftime('%Y%m%d_%HH%M')} PortalBreachReg.xlsx"
)
delay = 10


# loop to obtain each breach report
start_time = time.time()

# selenium suite of tools
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import (
    NoAlertPresentException,
)  # to handle the eagleportal.prescient.co.za alerts
from selenium.common.exceptions import TimeoutException

# (1) assign the browser driver
from selenium import webdriver

driver = webdriver.Firefox()

# (2) open the login page and log in
driver.get(ptl_login)  # lands on the "Compliance System Status" page
driver.find_element(
    By.CSS_SELECTOR, 'img[src="/img/icons/icon-login.png"]'
).click()  # find user login icon and click it
# wait.until(EC.presence_of_element_located((By.ID, 'img[src="/img/icons/icon-login.png"]'))).click()
driver.find_element(By.CSS_SELECTOR, "#username").send_keys(p_al)
driver.find_element(By.CSS_SELECTOR, "#password").send_keys(p_xe)
driver.find_element(By.CSS_SELECTOR, "#submit").click()

print(
    f"Dataframing the {len(breach_types)} reports and writing each as a sheet in a workbook here: {file_path} ..."
)
with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
    # breach_type = breach_types[2]
    for breach_type in tqdm(breach_types):
        # (3) open the browser on the default web page
        driver.get(ptl_b_rpt)  # open the browser on the default web page

        # (4) open the "Breach Age Analysis" tab using the <a class="f-nav-click" ...> anchor tag
        baa = 'a[class="f-nav-link"]'
        driver.find_element(By.CSS_SELECTOR, baa).click()
        # try:
        #     # wait up to 10 seconds for the element with ID 'myDynamicElement' to be present
        #     element = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, baa)))
        #     element.click()
        # except TimeoutException:
        #     print(f"Loading took more than {delay} seconds or element not found.")
        # finally:
        #     element = WebDriverWait(driver, delay).until(EC.presence_of_element_located((By.CSS_SELECTOR, baa)))
        #     element.click()

        # (5) click on the "Teams" drop-down
        taa = "#select2-team_selected-container"
        driver.find_element(By.CSS_SELECTOR, taa).click()

        # (6) select 'ALL' from the 'Teams' dropdown list
        li = "//li[contains(@class,'select2-results__option') and text()='ALL']"
        driver.find_element(By.XPATH, li).click()

        # (7) select from the four items in the "Breach type" drop-down: 'New', 'Open', 'Unclassified', 'Resolved'
        dropdown = Select(driver.find_element(By.ID, "breach_type"))
        dropdown.select_by_value(breach_type)

        # (8) click on the 'Submit' button
        driver.find_element(By.CSS_SELECTOR, "#submit_button").click()

        # (9) wait for confirmation of the breach type search
        start_time_bt = time.time()
        wait = WebDriverWait(driver, 10)
        txt = "th[aria-label^='portfolio_code']"
        element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, txt)))
        # print(f"{timediff(start_time_bt, time.time())} for '{breach_type.capitalize()}' to appear\n")

        # (10) click the 'CSV' button
        csv_it = 'button[class="btn btn-secondary buttons-csv buttons-html5"]'  # to save as a csv else '-copy' to copy
        driver.find_element(By.CSS_SELECTOR, csv_it).click()

        # (11) get the latest file from the downloads folder
        sfx = "csv"
        fln = os.path.join(pth_dl, rf"{breach_type.capitalize()}")
        fname = f"{fln}.{sfx}"
        latest_file(pth_dl, sfx, fln)

        # (12) dataframe the downloaded file
        df = pd.read_csv(fname)
        df["datestamp"] = pd.to_datetime(df["datestamp"])
        df["datestamp"] = df["datestamp"].dt.date

        # (13) save the dataframe as a sheet in an Excel workbook for that day
        df.to_excel(writer, sheet_name=breach_type.capitalize(), index=False)

# (14) close the webdriver instance
driver.quit()

print(f"{file_path}\n")
print(f"{timediff(start_time_portal, time.time())} total run time")
