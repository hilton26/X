#!/usr/bin/env python
# coding: utf-8

# In[1]:


# libraries, libraries!
import time

start_time = time.time()
start_time_r28_download = start_time
print("Importing libraries and setting up paths ...")

from datetime import datetime
import pandas as pd
import os
from pathlib import Path
from tqdm import tqdm
from utilities import timediff, osprey

# set paths to the driver, urls, and report parameters
eagle_portal = r"https://eagleportal.prescient.co.za"
url_default = eagle_portal + r"/Default.aspx"
url_r28 = eagle_portal + r"/Queries/Query.aspx?rpt=Reg28withExposure"
pth_py = r"P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm"  # variables stored here
pth_dl = str(Path.home() / "Downloads")
r28N = "Reg 28 Report - Incl Effective Exposure"  # prefix of the file from Eagle

print(
    f"{timediff(start_time, time.time())} importing libraries and setting up paths",
    "\n",
)

# get inputs to pass to Eagle
start_time = time.time()
print("Collecting input data ...")

# get user details, report date, and fund names from the py_report.xlsm sheet
df_downloader = pd.read_excel(
    pth_py, sheet_name="downloader", header=None, usecols="D", nrows=1
)
rptDate = df_downloader.iloc[0, 0]  # datetime type
df_funds = pd.read_excel(pth_py, sheet_name="downloader", usecols="A").dropna()
funds = df_funds["Entity Name"].apply(str.upper)
fund_codes = df_funds.iloc[:, 0].tolist()  # list type
fnds_ = (",").join(fund_codes)  # str type

# get credentials
df = pd.read_excel(pth_py, sheet_name="creds", header=None, usecols="A", nrows=2)
aladdin = df.iloc[0, 0]
sesame = df.iloc[1, 0]

# check inputs
print(
    f" Report date   : {rptDate.strftime('%A %d %B %Y')}",
    "\n",
    f"Portfolios ({len(funds)}): {(', ').join(funds)}",
)
print(f"{timediff(start_time, time.time())} collecting input data", "\n")

# determine which files still to be downloaded

# list and then pick out the .csv files from the Downloads folder and then ...
start_time = time.time()
print(
    f"Identifying the Reg 28 csv files dated {rptDate.strftime('%d %b %Y')} already in the local Downloads folder ..."
)

# pick out the csv files in the Downloads folder
import re

pattern = r"^R28I.*\(1\) " + f"{datetime.strftime(rptDate, '%d%b%Y')}" + r"\.csv$"

r28_csvs = [s for s in os.listdir(pth_dl) if re.match(pattern, s)]

# fund R28I reports still to be downloaded
done = pd.Series(r28_csvs).apply(
    lambda x: x.replace("R28I ", "").replace(
        f"(1) {rptDate.strftime('%d%b%Y')}.csv", ""
    )
)
undone = sorted(list(set(funds) ^ set(done)))

print(
    f" {len(r28_csvs)} ({len(funds)} in py_reports.xlsm) {rptDate.strftime('%d %b %Y')} R28I csv files in the Downloads folder, \
{len(funds) - len(r28_csvs)} still to be done:",
    "\n",
    "",
    (",").join(undone),
)

print(
    "\n",
    f"{timediff(start_time, time.time())} identifying the {len(undone)} Reg 28 csv files dated {rptDate.strftime('%d %b %Y')} already in \
the local Downloads folder",
    "\n",
)

# iterate through the list of reports to be downloaded from Eagle

start_time = time.time()
print(
    f"Downloading the remaining {len(undone)} fund R28I holdings for Reg 28 reporting as at {rptDate.strftime('%A %d %B %Y')} ...",
    "\n",
)

r28_not_downloaded = []
for fund in tqdm(undone):
    osprey("r28i", fund, rptDate, rptDate, fund, "csv", aladdin, sesame)

# open the Downloads folder
os.startfile(os.path.realpath(Path.home() / "Downloads"))

print(
    "\n\n",
    f"{timediff(start_time_r28_download, time.time())} downloading the \
        fund{'s' if len(undone) != 1 else ''} remaining {len(undone)} \
        fund R28I holdings for Reg 28 reporting as at \
        {rptDate.strftime('%d %B %Y')} ({(time.time() - start_time_r28_download) / (len(funds)):,.1f}sec/fund)",
)
