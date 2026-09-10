#!/usr/bin/env python
# coding: utf-8

# # Download the lookthroughs for Reg28 and Reg30 classification

print("\n\n#####################################")
print("#                                   #")
print("#           START lt_dl.py          #")
print("#                                   #")
print("#####################################\n\n")

import time

start_time = time.time()
start_time_lt_dl = time.time()

# libraries, libraries!
print("Importing libraries ...")
from datetime import datetime
import pandas as pd
import os, re
from pathlib import Path
from tqdm import tqdm
from constants import pthPy, pth_dl
from utilities import (
    timediff,
    prior_month_end,
    osprey,
    batch_list,
)

# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException

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
num_batches = df.iloc[2, 0]

# check inputs
s = "" if len(funds) == 1 else "s"
print(
    f"{len(funds)} fund lookthrough{s} as at {rptDate.strftime('%A %d %b %Y')} \
to be downloaded:\n  {(', ').join(funds)}"
)
print(f"\n{timediff(start_time, time.time())} collecting input data\n")

# download the lookthrough reports in batches
start_time = time.time()
suffix = "csv"
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), max(batch_size, 1)))
print(
    f"Downloading the {len(funds)} lookthroughs for \
{rptDate.strftime('%a %d %b %Y')} in {num_batches} batches ..."
)

batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start=1)):
    # give status of downloads
    dy = Path(pth_dl)
    pattern = re.compile(
        rf"^R28I.*of_{len(batches)}.*{rptDate.strftime('%d%b%Y')}\.{suffix}$"
    )
    matches = [f.name for f in dy.iterdir() if f.is_file() and pattern.match(f.name)]
    done = [
        re.search(r"R28I\s+(\d+)_of", f).group(1)
        for f in matches
        if re.search(r"R28I\s+(\d+)_of", f)
    ]  # type(done) = class 'list'
    done = sorted(done, key=int)  # sort the items in the done list
    es = "es" if len(batches) - len(done) != 1 else ""
    print(
        f"\nBatches {(', ').join(list(done))}, i.e., \
{len(done)} (\
{len(done) / len(batches) * 100:.1f}%) done \
out of {len(batches)} batches,"
    )
    print(
        f"{len(batches) - len(done)} \
batch{es} ({(1 - len(done) / len(batches)) * 100:.1f}\
%) to go\n"
    )

    fln = f"{index}_of_{len(batches)}"
    filename = f"R28I {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.{suffix}"
    s = "" if len(batch) == 1 else "s"
    print(f"Get {filename}, a batch of {len(batch)} file{s}:\n   {(', ').join(batch)}")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        start_time_1 = time.time()
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath} ...")
        osprey("r28i", (",").join(batch), rptDate, rptDate, fln, suffix)
        print(
            f" {timediff(start_time_1, time.time())} attempting \
to download batch {index} of {len(batches)}\n"
        )

if len(done) == len(batches):
    print(
        f"\n{timediff(start_time, time.time())} downloading \
the {len(funds)} lookthroughs for {rptDate.strftime('%a %d %b %Y')} \
in {num_batches} batches\n"
    )
    print("\n\n#######################################")
    print("#                                     #")
    print("#             END lt_dl.py            #")
    print("#                                     #")
    print("#######################################\n\n")
else:
    raise Exception(
        f"{len(batches) - len(done)} batch{es} \
not downloaded"  #: {(',').join(list(not_done))}"
    )
