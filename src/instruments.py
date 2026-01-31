#!/usr/bin/env python
# coding: utf-8

# ### Extract instrument columns from instrument data dump out of the PIM Portal

# libraries, libraries!
import time

start_time_portal = time.time()
start_time = time.time()

import pandas as pd
from datetime import datetime
import os
import json
from tqdm import tqdm
from pathlib import Path
from constants import ptl_login, pth_dl, p_al, p_xe, ptl_instr, pth_instr
from utilities import timediff

# (1) import the selenium suite of tools
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

# (2) assign the browser driver
from selenium import webdriver

from dotenv import load_dotenv  # to access environment variables from .env file

load_dotenv()  # take environment variables from .env file

driver = webdriver.Firefox()

# (2a) assign an arbitrary wait time for element sto be clicked
wait = WebDriverWait(driver, 10)

# (3) open the portal login page and log in
driver.get(ptl_login)  # lands on the "Compliance System Status" page
driver.find_element(
    By.CSS_SELECTOR, 'img[src="/img/icons/icon-login.png"]'
).click()  # find user login icon and click it
# wait.until(EC.presence_of_element_located((By.ID, 'img[src="/img/icons/icon-login.png"]'))).click()
driver.find_element(By.CSS_SELECTOR, "#username").send_keys(os.getenv("PORTAL_UN"))
driver.find_element(By.CSS_SELECTOR, "#password").send_keys(os.getenv("PORTAL_PW"))
driver.find_element(By.CSS_SELECTOR, "#submit").click()

# (4) access the portal instrument maintenance page
driver.get(ptl_instr)

# (6) wait a few seconds for the data to load onto the page
# driver.wait
time.sleep(2)

# # Alternative way to wait for the data to load: wait until the info div is present
# <div class="dataTables_info" id="instruments_table_info" role="status" aria-live="polite">Showing 1 to 20 of 7,524 entries</div>
# ind_data = "instruments_table_info"
# info_div = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ind_data)))
# full_text = info_div.text
# print(full_text)

# (7) click the "CSV" button to download the instruments data
csv_btn = "button.buttons-csv.buttons-html5"
driver.find_element(By.CSS_SELECTOR, csv_btn).click()

# (8) close the webdriver instance
driver.quit()

# (9) get the latest CSV starting with data from the downloads folder, source: ChatGPT 27Jan2026
csv_files = [
    f
    for f in Path(
        pth_dl
    ).iterdir()  # first convert pth_dl from type str to type pathlib.WindowsPath
    if f.is_file() and f.name.lower().startswith("data") and f.suffix.lower() == ".csv"
]
if not csv_files:
    raise FileNotFoundError("No CSV files starting with 'data' found")
latest_csv = max(
    csv_files, key=lambda f: f.stat().st_mtime
)  # uses last modified time to find the newest file
# print("Latest CSV file:", latest_csv)

# (10) read and then dataframe the downloaded file
df1 = pd.read_csv(latest_csv)
df1["Datestamp"] = pd.to_datetime(df1["Datestamp"])
df1["Datestamp"] = df1["Datestamp"].dt.date

# (10) save the dataframe as a sheet in an Excel workbook for that day
sht_nm = f"instr_{datetime.today().strftime('%d%b%Y')}"
with pd.ExcelWriter(pth_instr, engine="openpyxl") as writer:
    df1.to_excel(writer, sheet_name=sht_nm, index=False)


# (11) define a function to prefix 'R' to SA Gov security codes and apply it
def zar(k):
    return (
        "R" + str(k)
        if (((len(str(k)) == 3) | (len(str(k)) == 4)) & isinstance(k, int))
        else k
    )


df1["Instrument Code"] = df1["Instrument Code"].apply(
    zar
)  # prefix 'R' to SA Gov security codes

# (12) split out the instrument tags into separate columns
df2 = pd.DataFrame()
print(f"Joining {len(df1):,.0f} instruments over 44 columns")
for tag in tqdm(list(df1["Tags"])):
    df2 = pd.concat([df2, pd.DataFrame([json.loads(tag)])], ignore_index=True)

df = pd.concat([df1["Instrument Code"], df2], axis=1)

# (13) writing to a sheet
with pd.ExcelWriter(
    pth_instr, engine="xlsxwriter"
) as writer:  # instantiate an Excel worksheet writer
    df.to_excel(
        writer, sheet_name=f"split_{datetime.today().strftime('%d%b%Y')}", index=False
    )
    df1.to_excel(
        writer, sheet_name=f"instr_{datetime.today().strftime('%d%b%Y')}", index=False
    )
writer.close

print(f"{pth_instr}\n")
print(f"{timediff(start_time_portal, time.time())} total run time")
