import time

start_time_zip = time.time()

from datetime import datetime
import pandas as pd
import os
from tqdm import tqdm
from constants import pth_dl, pthPy, pthPortal
from utilities import (
    timediff,
    last_working_day,
    prior_month_end,
    latest_file
)

# import selenium library
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException


# access and then load variables from .env file
from dotenv import load_dotenv

load_dotenv()

# get list of funds
funds = pd.read_excel(
    pthPy, sheet_name="arc", usecols="N", header=None, skiprows=1
).dropna()

# get reporting date
df = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
rptDate = (
    df.iloc[1, 0].date()
    if isinstance(df.iloc[1, 0], datetime)
    else prior_month_end(datetime.today().date())
)

# Initiate the WebDriver (e.g., Chrome)
driver = webdriver.Firefox()
# driver.minimize_window()

# sign in to the Prime Portal
sign_in = pthPortal + r"/signin"
driver.get(sign_in)

driver.find_element(By.CSS_SELECTOR, "#username").send_keys(os.getenv("PORTAL_UN"))
driver.find_element(By.CSS_SELECTOR, "#password").send_keys(os.getenv("PORTAL_PW"))
driver.find_element(By.CSS_SELECTOR, "#submit").click()

# go to the Positions page on the Prime Portal
positions = pthPortal + r"/positions"
driver.get(positions)

# enter date
date_field = driver.find_element(By.CSS_SELECTOR, "#date_selected")
driver.execute_script(
    "arguments[0].value = arguments[1];"
    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
    date_field,
    rptDate.strftime("%Y-%m-%d"),
)

# enter postion option, direct or lookthrough
position_options = Select(driver.find_element(By.CSS_SELECTOR, "#position_options"))
position_options.select_by_visible_text("LOOKTHROUGH")


# enter fund code
fund = "PIPF"
driver.execute_script(
    "$('#funds').val(arguments[0]).trigger('change');", fund
)

# click the submit button
driver.find_element(By.CSS_SELECTOR, ".btn.prime-btn-primary").click()

latest_file(pth_dl, "csv", "(13)LT_Portal")



submit_button = driver.find_element(By.CSS_SELECTOR, "#submit")



# # submit entries
# driver.find_element(By.CSS_SELECTOR, "#date_selected").send_keys(rptDate.strftime("%Y/%m/%d"))

# driver.find_element(By.CSS_SELECTOR, "#password").send_keys(os.getenv("PORTAL_PW"))
# driver.find_element(By.CSS_SELECTOR, "#submit").click()


# for fund in tqdm(funds):
#     driver.get(positions)


# print(
#     f"{timediff(start_time_zip, time.time())} total \
# time to download and save fund loo-through holdings \
# for {len(funds)} funds from Prime Portal"
# )
