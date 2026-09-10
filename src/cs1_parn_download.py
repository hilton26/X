#!/usr/bin/env python
# coding: utf-8

print("\n\n################################################")
print("#                                              #")
print("#      START cs1_PARN_download_csv.py   X      #")
print("#                                              #")
print("################################################\n\n")

import time

start_time = time.time()
start_time_cs1_PARN_download = start_time

# libraries, libraries!
print("Importing libraries ...")
from datetime import datetime
import pandas as pd
import os
from tqdm import tqdm
from constants import pthPy, pth_dl
from utilities import timediff, prior_month_end, osprey, batch_list

print(f" {timediff(start_time, time.time())} importing libraries\n")

# get inputs to pass to Eagle
start_time = time.time()
print("Collecting input data ...")

# get fund codes from r28_cs1 tab of the py_report.xlsm sheet
df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df.iloc[:, 0].apply(str.upper)

# get report date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
k = df1.iloc[1, 0]
rptDate = (
    k.date() if k == k else prior_month_end(datetime.today()).date()
)  # prior month end or report date override; type is datetime()

# print inputs
print(
    "\n",
    f"PARN CS1 fund report{'' if len(funds) == 1 else 's'} to be downloaded as at {rptDate.strftime('%A %d %b %Y')}",
    "\n",
    f"{len(funds)} fund{'' if len(funds) == 1 else 's'}: ",
    "\n",
    f"{(', ').join(funds)}",
)
print(f"\n{timediff(start_time, time.time())} collecting input data\n")


# download the PARN reports in batches

num_batches = 2 if len(funds) > 1 else 1
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start=1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    print(f"{filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}\n")
    # print(f" {len(batch)} files:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)
    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath} ...\n")
        osprey("parn", (",").join(batch), rptDate, rptDate, fln, "csv")

for batch_filepath in batch_filepaths:
    print(batch_filepath)

print("\n\n################################################")
print("#                                              #")
print("#       END cs1_PARN_download_csv.py   X       #")
print("#                                              #")
print("################################################\n\n")
