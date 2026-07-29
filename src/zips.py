#!/usr/bin/env python
# coding: utf-8

# # Obtaining JSE and MSCI Indices and Bond Data a Month End
#
# ##### Python's zipfile: Manipulate Your ZIP Files Efficiently
# https://realpython.com/python-zipfile/#:~:text=Python%27s%20zipfile%20module%20lets%20you,ZIP%20archives%20using%20different%20algorithms.

# #### Dependencies:
#
# ##### folder P:\Investment Operations\GRC\Compliance\Reporting Requirements\
# PIM PPSBAL\Zips\PPS_MSCI_indexes_yyyymmdd.zip
# ##### format PPS_MSCI_indexes_yyyymmdd.zip

# libraries, libraries!

import time

start_time_zip = time.time()

import os, shutil
from pathlib import Path
from datetime import datetime, timedelta
import re  # regex for identifying property companies
import pandas as pd
import numpy as np
import zipfile
from tqdm import tqdm
from constants import (
    msci_zips,
    msci_dict,
    pth_dl,
    pth_struct,
    pth_BX,
    jse_data,
    credit_meta,
    pth_m_reports,
    pth_EC,
    pth_PPSBAL,
)
from utilities import (
    timediff,
    last_working_day,
    prior_month_end,
    latest_file_in_folder,
    property,
)

# import selenium for picking up the Jxxx indices and bond indices from Prime Portal
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# function to return matching MSCI index codes
def msci(txt):
    for key, pattern in msci_dict.items():
        if re.search(pattern, str(txt).upper()):
            return f"{key.upper()}"
    return "---xxx---"


# (1) create latest month reporting folders

# take reporting date as last working day of prior month end from today's date
rptDate = last_working_day(prior_month_end(datetime.today()))
print(rptDate)

# create the month end reporting folders if they don't yet exist - https://flexiple.com/python/python-make-directory

start_time = time.time()
print(f"Creating the month-end reporting folders, if they don't yet exist\n")


# create the month-end_17 reporting folder if it doesnt yet exist
pth = os.path.join(pth_m_reports, rptDate.strftime("%Y"), rptDate.strftime("%Y %m"))
if not os.path.exists(pth):
    os.makedirs(pth)

# create the month reporting folder for ECICBALC
fldr_ECICBALC = os.path.join(pth_EC, f"{rptDate.strftime('%Y%m')}")
if not os.path.exists(fldr_ECICBALC):
    os.makedirs(fldr_ECICBALC)

# create the month-end reporting folder for PPSBAL:
fldr_PPSBAL = os.path.join(pth_PPSBAL, f"{rptDate.strftime('%Y%m')}")
if not os.path.exists(fldr_PPSBAL):
    os.makedirs(fldr_PPSBAL)

print(
    f" {timediff(start_time, time.time())} creating the month-end \
reporting folders, if they didn't yet exist\n"
)

# (2) get MSCI data

# construct name of latest zipped archive in the folder
start_time = time.time()
print(f"Getting zipped MSCI data\n")

m = os.listdir(msci_zips)  # list of files in the Zips directory
# k = max([x[-8:-4] + x[-12:-10] + x[-10:-8] for x in m if x.endswith('.zip')]) # sub-list of files ending in ".zip"
# zip_file     = f'PPS_MSCI_indexes_{k[4:6] + k[-2:] + k[0:4]}.zip'
# TUE 4 NOV 2025: AMEND TO LOOK FOR \d{8}.zip
k = max(
    [
        x[-12:-4]
        for x in m
        if x.endswith(".zip") and os.path.getsize(os.path.join(msci_zips, x)) > 5 * 1024
    ]
)  # latest date suffix among ".zip" files larger than 5kB

# sub-list of files ending in ".zip"
# zip_file     = f'PPS_MSCI_indexes_{k}.zip' # latest zip file in the folder
zip_file = f"PPS_MSCI_indexes_{k}.zip"  # latest zip file in the folder
zipped_file = os.path.join(msci_zips, zip_file)

print(f" Unzipping {zipped_file} \n")

# TEST
# print('',zip_file,'\n',zipped_file)

# find the .cn1 file in the zipped archive
with zipfile.ZipFile(zipped_file, mode="r") as archive:
    for filename in archive.namelist():
        if filename.endswith(".cn1"):
            cn1 = filename
            break

# print('', f'{zip_file} is the latest zipped archive','\n')
# print('', f'{cn1} is the file in the archive to be read','\n')

# extract bytes data from the zipped .cn1 file
try:
    x = []
    with zipfile.ZipFile(zipped_file, mode="r") as archive:
        with archive.open(cn1, mode="r") as cn1:
            for line in cn1:
                # print(line.split(b','))
                x.append(line.split(b","))
except zipfile.BadZipFile as error:
    print(error)

# dataframe the retrieved bytes
df = pd.DataFrame(x)


# convert bytes to strings
def de_byte(byt):
    """function to decode bytes to str"""
    return byt.decode("utf-8")


df = df.map(de_byte)

# make the first row the header row
# https://stackoverflow.com/questions/31328861/replacing-header-with-top-row
new_header = df.iloc[0]  # grab the first row for the header
df = df[1:]  # take the data less the header row
df.columns = new_header  # set the header row as the df header

# drop the source sub-header row
df.drop(1, inplace=True)

# reset the index
df.reset_index(level=None, drop=True, inplace=True)

# make a unique, separate copy, df_MSCI, of the just 
# unzipped df so that modifications to the copy will not affect the original df
df_MSCI = df.copy(
    deep=True
)  # deep=True is the default and can be omitted, changes to the 
# copy do not affect the original

# rename the last column which contains "\r\n"
df_MSCI.rename(columns={"price currency\r\n": "price currency"}, inplace=True)
# df_MSCI.info()

# remove "\r\n" from the last column
df_MSCI["price currency"] = df_MSCI["price currency"].str.replace(
    "\r\n", "", regex=False
)

# convert some columns from type string to type float
headings = ["market cap 1", "market cap 2", "weight 1", "price"]
df_MSCI[headings] = df_MSCI[headings].astype(float)

# get MSCI report date
date_MSCI = datetime.strptime(df_MSCI["end date"].iloc[0], "%Y%m%d")

# use specific headings for the dataframe of indices
msci_heads = [
    "ticker",
    "end date",
    "series long name",
    "Statpro ISIN",
    "security name",
    "price currency",
    "market cap 1",
    "weight 1",
    "Statpro Sedol",
    "provider",
]
df_MSCI = df_MSCI[msci_heads]

# get the unique sets of indices
series = df_MSCI["series long name"].unique()

# derive reporting date from date of MSCI index data upload
# rptDate = last_working_day(date_MSCI)

zar = df_MSCI.drop_duplicates(subset=["Statpro ISIN"])

# apply msci() to the 'series long name' column
df_MSCI.iloc[:, 9] = df_MSCI.iloc[:, 2].apply(msci)

print(
    f" {len(series)} indices for {date_MSCI.strftime('%a %d %b %Y')} \
over {len(df_MSCI)} rows:\n {(', ').join(series)}"
)
print(f" all of which are covered by the dictionary \n")
print(
    f" {len(df_MSCI.iloc[:, 3].unique())} unique ISINs \
among the {len(series)} sets of MSCI indices \
including {len(zar[zar['price currency'] == 'ZAR'])} JSE tickers \n"
)
print(
    f"{timediff(start_time, time.time())} to get zipped MSCI \
data for {date_MSCI.strftime('%a %d %b %Y')}"
)

# (3) Get JSE data and combine with MSCI data and prior month-end BX data, then save as a file

# define MSCI and JSE index heading names
bx_heads = [
    "PIM Ticker",
    "MC (ZAR)",
    "SISS",
    "Sedol",
    "Name",
    "Bloomberg Ticker",
    "GICS Code",
    "Status",
    "Exchange",
    "Domiciled",
    "Currency",
    "ISIN",
    "RE",
    "Date of Upload",
    "Index",
]
msci_heads_r = {
    "ticker": "PIM Ticker",
    "Statpro ISIN": "ISIN",
    "market cap 1": "MC (ZAR)",
    "security name": "Name",
    "Statpro Sedol": "Sedol",
    "provider": "Index",
    "end date": "Date of Upload",
    "price currency": "Domiciled",
}
msci_new_cols = ["SISS", "Bloomberg Ticker", "GICS Code", "Status", "Exchange", "RE"]
JSE_heads = [
    "trade date",
    "index",
    "equity_alpha_code",
    "last_mod_datetime",
    "constituent_name",
    "iso_code",
    "price",
    "number_of_shares_in_issue",
    "investibility_weighting_factor",
    "adjusted_market_cap_net",
    "weight",
]
JSE_heads_r = {
    "trade_date": "Date of Upload",
    "index": "Index",
    "equity_alpha_code": "PIM Ticker",
    "constituent_name": "Name",
    "number_of_shares_in_issue": "SISS",
    "adjusted_market_cap_net": "MC (ZAR)",
}
JSE_new_cols = [
    "Sedol",
    "Bloomberg Ticker",
    "GICS Code",
    "Status",
    "Exchange",
    "Domiciled",
    "Currency",
    "ISIN",
    "RE",
]

# dataframe the unique, by ISIN, MSCI constituent data
start_time = time.time()
print("Dataframing the unique MSCI constituents")

# https://stackoverflow.com/questions/43184491/df-unique-on-whole-dataframe-based-on-a-column
# df_MSCI      = df_MSCI.drop_duplicates(subset = ['isin'])
# date_dfz = datetime.strptime(dfz['end date'].iloc[0], "%Y%m%d").date()

# rename MSCI columns and add new columns to match BX

df_MSCI = df_MSCI.rename(columns=msci_heads_r)  # rename MSCI columns

for col in msci_new_cols:  # add new, empty columns
    df_MSCI[col] = np.nan

df_MSCI["Currency"] = df_MSCI["Domiciled"]

df_MSCI = df_MSCI[bx_heads]  # reorder MSCI index column headings
df_MSCI["Date of Upload"] = pd.to_datetime(df_MSCI["Date of Upload"])

print(f"{timediff(start_time, time.time())} dataframing the unique MSCI constituents")

# configure the MSCI dataframe
start_time = time.time()
print("Populating the empty MSCI dataframe columns")

# convert float64 columns to string, else the .loc assignment below will fail
cols_to_str = ["Bloomberg Ticker", "GICS Code", "Status", "Exchange"]
for col in cols_to_str:
    df_MSCI[col] = df_MSCI[col].astype(str)

# populate the MSCI columns
df_MSCI["PIM Ticker"] = df_MSCI["ISIN"]
df_MSCI["Bloomberg Ticker"] = df_MSCI["ISIN"]
df_MSCI["Domiciled"] = df_MSCI["Domiciled"].str[:2]
df_MSCI["Status"] = "ACTV"
df_MSCI["RE"] = df_MSCI["Name"].apply(property)
df_MSCI.loc[df_MSCI["RE"] == "P", "GICS Code"] = str("60102030")
df_MSCI.loc[df_MSCI["RE"] != "P", "GICS Code"] = str("40102030")
df_MSCI.loc[df_MSCI["Domiciled"] == "ZA", "Exchange"] = "XJSE"
df_MSCI.loc[df_MSCI["Domiciled"] != "ZA", "Exchange"] = "XNSE"

print(
    f"{timediff(start_time, time.time())} populating \
the empty MSCI dataframe columns"
)

# get the JSE indices
start_time = time.time()

print(
    f"\n\nDownloading and then dataframing the JSE index \
constituents for {rptDate.date()}"
)

d = rptDate.strftime("%Y-%m-%d")

# read in the list of indices to be obtained from Prime Portal
df = pd.read_excel(pth_struct, sheet_name="dervs", usecols=[1]).dropna()
indices = list(df.iloc[:, 0])  # all the rows, zeroth column

# list only the JSE indices
pattern = "J\d{3}"  # e.g., 'J123'
indices = [s for s in indices if re.search(pattern, str(s).upper())]

print(f"  {(', ').join(indices)}")

links = []  # empty list to bunch each link into a list
df_J = pd.DataFrame()  # empty dataframe to hold Jxxx indices
for index in tqdm(indices):
    # pull the index into the Downloads folder
    link = jse_data + rf"/{index}/{d}/{d}/True/False"
    links.append(link)
    driver = webdriver.Chrome()  # click the link using Chrome, or Edge (crashes), or Firefox (takes long to close), etc.
    driver.get(link)
    time.sleep(10)
    driver.quit()

    # dataframe the latest file in the Downloads folder
    fl = latest_file_in_folder(pth_dl)
    df = pd.read_csv(fl)
    df_J = pd.concat([df_J, df])

    # print result
    # print('' , f'{index} downloaded as {fl}')

print(
    f"\n{timediff(start_time, time.time())} downloading and \
dataframing the {len(indices)} JSE index constituents \
{(', ').join(indices)} for {rptDate.date()}\n"
)

# configure the JSE constituent dataframe
start_time = time.time()
print("Configuring the JSE constituent dataframe \n")

# rename JSE index columns and add new columns to match BX
df_J = df_J.rename(columns=JSE_heads_r)  # rename JSE columns
for col in JSE_new_cols:  # add new, empty columns
    df_J[col] = np.nan
df_J = df_J[bx_heads]  # reorder MSCI index column headings

# get date of the index
df_J["Date of Upload"] = pd.to_datetime(
    df_J["Date of Upload"]
)  # convert date column to type datetime
date_J = df_J["Date of Upload"].iloc[0].date()
print(f" JSE indices date: {date_J}", "\n")

# convert df_J column "GICS Code" to string, else the .loc assignment below will fail
df_J["GICS Code"] = df_J["GICS Code"].astype(str)

# populate the JSE dataframe columns
df_J["PIM Ticker"] = df_J["PIM Ticker"] + " SJ"
df_J["Bloomberg Ticker"] = df_J["PIM Ticker"] + " Equity SJ"
df_J["Status"] = "ACTV"
df_J["Domiciled"] = "ZA"
df_J["Currency"] = "ZAR"
df_J["Exchange"] = "XJSE"
df_J["RE"] = df_J["Name"].apply(property)
df_J.loc[df_J["RE"] == "P", "GICS Code"] = str("60102030")
df_J.loc[df_J["RE"] != "P", "GICS Code"] = str("40102030")

print(
    f"{timediff(start_time, time.time())} configuring \
the JSE constituent dataframe"
)

# dataframe the previous month end BX sheet
start_time = time.time()
p = prior_month_end(min(date_MSCI.date(), date_J))
print("", f"Getting prior month-end market caps for {p.strftime('%a %d %b %Y')}")

prior_BX_name = rf"{pth_BX}\BX {p.strftime('%d%b%Y')}.xlsx"
df_prior_BX = pd.read_excel(prior_BX_name, usecols="A:O")
df_prior_BX = df_prior_BX[
    df_prior_BX["Index"].isnull()
]  # exclude constituents not linked to indices

# df_prior_BX
print(
    f"\n{timediff(start_time, time.time())} getting prior \
month-end market caps for {p.strftime('%a %d %b %Y')}"
)

# dataframe the BSKxxx basket constituents
start_time = time.time()
print("", f"Getting BSK basket constituents from prior month")

# dataframe the treasuries and BSKxxx basket indices
df_tsy = pd.read_excel(prior_BX_name, sheet_name="Treasuries", usecols="A:O").dropna(
    subset=["PIM Ticker"]
)
df_bsk = pd.read_excel(prior_BX_name, sheet_name="BSK", usecols="A:O").dropna(
    subset=["PIM Ticker"]
)

print(
    f"\n{timediff(start_time, time.time())} getting BSK basket \
constituents from prior month:\n   {df_bsk['Index'].unique()[0]}"
)

# combine the unique constituents of the MSCI, JSE, and prior BX dataframes
start_time = time.time()
print("Combining the MSCI, JSE, and prior BX dataframes")
MSCI_unique = df_MSCI.drop_duplicates(subset=["ISIN"])
JSE_unique = df_J.drop_duplicates(subset=["PIM Ticker"])
df_BX = pd.concat(
    [MSCI_unique, JSE_unique, df_prior_BX], axis=0
)  # axis=0 is default for rows

# Convert the datetime column to a date column
df_BX["Date of Upload"] = df_BX["Date of Upload"].dt.date

print(f" {len(df_BX):,.0f} unique securities")

print(
    f"{timediff(start_time, time.time())} combining the \
MSCI, JSE, and prior BX dataframes"
)

# # TEST
# # confirm an old ticker is still included in the current BX file
# df_BX[df_BX["PIM Ticker"] == "SEA SJ"]

# (4) Get bond data from Prime Portal, then save the indices and bond data to a file

# get bond data
start_time = time.time()
print("Downloading and dataframing bond data from Prime Portal")

try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import NoSuchElementException, TimeoutException

    # Set up the WebDriver (e.g., Chrome)
    driver = webdriver.Chrome()
    driver.get(credit_meta)

    # open the Bond Mapping page on Prime Portal
    # <a class="f-nav-link active" href="#bond-mapping" data-bs-toggle="tab" data-bs-target="#nav-3">Bond Mapping</a>
    link_element_css_clue = 'a[class="f-nav-link"][href="#bond-mapping"]'
    link_element = (
        WebDriverWait(driver, 10)
        .until(EC.element_to_be_clickable((By.CSS_SELECTOR, link_element_css_clue)))
        .click()
    )

    # click the CSV link on the Bond Mapping page
    # <button class="btn btn-secondary buttons-csv buttons-html5" tabindex="0" aria-controls="bond_mapping" type="button"><span>CSV</span></button>
    time.sleep(10)
    csv_element_css_clue = 'button[class="btn btn-secondary buttons-csv buttons-html5"][aria-controls="bond_mapping"]'
    csv_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, csv_element_css_clue))
    )
    csv_element.click()

except NoSuchElementException:
    print("Element not found. Handling NoSuchElementException.")
    # You can add actions here, like logging the error, taking a screenshot, or retrying
except TimeoutException:
    print("Operation timed out. Handling TimeoutException.")
    # Specific handling for timeout issues
except EmptyDataError as d:
    print(f"Empty data error occurred: {d}")
    # Catch for empty data error
except ParserError as p:
    print(f"Error tokenising data: {p}")
    # ParserError: Error tokenizing data. C error: EOF inside string starting at row 4324
except Exception as e:
    print(f"An unexpected error occurred: {e}")
    # Catch-all for other potential exceptions

finally:
    # This block will always execute, regardless of whether an exception occurred
    driver.quit()  # close the browser

# dataframe the latest file in the Downloads folder
time.sleep(5)  # arbitrary 5-second timeout
try:
    df_B = pd.read_csv(latest_file_in_folder(pth_dl))
    print(f" {len(df_B):,.0f} debt securities")
except:
    # use last month's bond data
    df_B = pd.read_excel(prior_BX_name, sheet_name="Bonds", usecols="A:H").dropna(
        subset=["Bond Code"]
    )
    print(f" {len(df_B):,.0f} debt securities from {p.strftime('%d %b %Y')} BX file")

print(
    f"\n{timediff(start_time, time.time())} downloading and dataframing bond data from Prime Portal"
)

# (5) Write the indices and bond dataframes to a workbook for review, then copy as the newest 'BX.xlsx' file

# write the indices and bond dataframes to a workbook
start_time = time.time()
print(f"Writing the new BX dataframe to a file for {rptDate.date()}")

# writing BX.xlsx for review
fln_BX = os.path.join(
    pth_BX, f"BX {prior_month_end(datetime.today()).strftime('%d%b%Y')}.xlsx"
)
with pd.ExcelWriter(
    fln_BX, engine="xlsxwriter"
) as writer:  # instantiate an Excel worksheet writer
    df_BX.to_excel(
        writer, sheet_name="BBG", index=False
    )  # write the BX indices to a sheet
    df_MSCI.to_excel(
        writer, sheet_name="MSCI", index=False
    )  # write the MSCI indices to a sheet
    df_J.to_excel(
        writer, sheet_name="JSE_indices", index=False
    )  # write the JSE indices to a sheet
    df_B.to_excel(
        writer, sheet_name="Bonds", index=False
    )  # write the bond data to a sheet
    df_bsk.to_excel(
        writer, sheet_name="BSK", index=False
    )  # write the BSK basket data to a sheet
    df_tsy.to_excel(
        writer, sheet_name="Treasuries", index=False
    )  # write the treasury bond data to a sheet
writer.close

# (6) Copy the new workbook as the newest 'BX.xlsx' file

# copy as the newest 'BX.xlsx' file
shutil.copyfile(fln_BX, os.path.join(pth_BX, "BX.xlsx"))

print(
    f"{timediff(start_time, time.time())} writing the new BX dataframe to a file for {rptDate.date()}\n"
)
print(" ", fln_BX)
print(
    f"\n{timediff(start_time_zip, time.time())} roundtrip time to collect index and bond data for {rptDate.strftime('%a %d %b %Y')}"
)

# (6) Create the latest month reporting folders
