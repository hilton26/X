print("\n\n##########################")
print("#   START lt_classify.py   #")
print("##########################\n\n")

import time

start_time = time.time()
start_time_classify_dl = time.time()

# libraries, libraries!
print("Importing libraries ...")
from datetime import datetime
import pandas as pd
import os, re
from pathlib import Path
from tqdm import tqdm
from constants import pthPy, pth_dl, pthTest
from utilities import (
    timediff,
    last_working_day,
    prior_month_end,
    osprey,
    batch_list,
    r_classifier,
)
import subprocess

print(f" {timediff(start_time, time.time())} importing libraries\n")

# get inputs to pass to Eagle
start_time = time.time()
print("Collecting input data ...")

# get fund codes from 'arc' tab of the py_report.xlsm sheet
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df1.iloc[:, 0].str.upper()
df1

# get report parameters date
df = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
k = df.iloc[1, 0]
rptDate = (
    k if isinstance(k, datetime) else prior_month_end(datetime.today().date())
)  # prior month end or report date override; type is datetime()

s = "" if len(funds) == 1 else "s"
lt_fname = os.path.join(
    pthTest, f"LT holdings ({len(funds)}) {rptDate.strftime('%d%b%Y')}.xlsx"
)

print(f"{rptDate}\n{lt_fname}")

# run the Reg 28 reporting script
start_time = time.time()
print(f"Classifying the lookthrough holdings\n")
r_classifier("lt", lt_fname)
print(f" \n{timediff(start_time, time.time())} classifying the lookthrough holdings\n")

print(
    "\n",
    f"{timediff(start_time_classify_dl, time.time())} roundtrip to \
download and merge lookthroughs\n",
)

print("\n\n#####################")
print("#    END lt_classify.py   #")
print("#####################\n\n")
