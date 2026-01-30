#!/usr/bin/env python
# coding: utf-8

# # Pull Holdings and Derivative Data from Eagle

# How to wait until Element is Visible in Selenium Python
#
# https://pythonexamples.org/python-selenium-wait-until-element-is-visible/

print("######################################")
print("# START 1/4 DERV_CHECKER_DOWNLOADING #")
print("######################################")

import time

start_time = time.time()
start_time_derivative_downloading = time.time()
print("Importing libraries and setting paths for derv_checker_downloading.ipynb ...")

# load libraries
import pandas as pd
from datetime import datetime
import numpy as np
import os
from pathlib import Path
from constants import pthPy, pthEXPORTS, pth_dl, pthLOCAL
from utilities import timediff, osprey, prior_working_day

# get report date and selected summary sheet option
df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
    subset=[0]
)
k = df.iloc[2, 1]
rptDate = (
    k if isinstance(k, datetime) else prior_working_day(datetime.today())
)  # prior working day or report date override; has type datetime()
summ_yn = df.iloc[3, 1]
full = df[0].iloc[1:]
funds = (",").join(full.tolist())

print(
    f"{timediff(start_time, time.time())} importing libraries \
        and setting paths for derv_checker_downloading.py",
    "\n",
)

# derive file names
start_time = time.time()
print("Deriving file names ...")

filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
fPARN = os.path.join(
    Path.home(),
    "Downloads",
    f"PARN ({len(full)}) {rptDate.strftime('%d%b%Y')}.csv",
)
fDE = os.path.join(
    Path.home(),
    "Downloads",
    f"DERV ({len(full)}) {rptDate.strftime('%d%b%Y')}.csv",
)

hlf = int(len(full) / 2)
half1 = (",").join(full[:hlf])
half1_name = f"PARN half1({len(full[:hlf])}) {rptDate.strftime('%d%b%Y')}.csv"
half2 = (",").join(full[hlf:])
half2_name = f"PARN half2({len(full[hlf:])}) {rptDate.strftime('%d%b%Y')}.csv"
full_name = f"PARN ({len(full)}) {rptDate.strftime('%d%b%Y')}.csv"
derv_name = f"DERV ({len(full)}) {rptDate.strftime('%d%b%Y')}.csv"

print(
    "\n",
    f" {rptDate.strftime('%A %d %b %Y')} for {len(full)} funds:",
    "\n",
    f" {funds}",
    "\n",
)
print(f" {'No' if summ_yn == 'No' else 'A'} summary sheet is required", "\n")

print(f"{timediff(start_time, time.time())} deriving file names", "\n")

# get the fund holdings in "portfolio analytics review - new" format
start_time = time.time()
print("Downloading and then saving holdings data ...")

if len(full) > 100:  # if more than 100 funds are in the list ...
    # ... get holdings for the first half of funds in the list
    start_time_1 = time.time()
    print(f"  ... downloading first of two subsets of holdings: {half1_name}")
    # check if the file was already downloaded before running osprey()
    # if os.path.isfile(pth_dl + r'\\' + half1_name):
    if (
        os.path.exists(os.path.join(pth_dl, half1_name))
        and os.path.getsize(os.path.join(pth_dl, half1_name)) > 0
    ):
        print(f"  {half1_name} already exists\n")
        pass
    else:
        osprey("parn", half1, rptDate, rptDate, "half1", "csv")
        print(
            f"  {timediff(start_time_1, time.time())} downloading first of two subsets of holdings"
        )

    # ... get holdings for the second half of funds in the list
    start_time_2 = time.time()
    print(f"  ... downloading second of two subsets of holdings: {half2_name}")
    # check if the file was already downloaded before running osprey()
    # if os.path.isfile(pth_dl + r'\\' + half2_name):
    if (
        os.path.exists(os.path.join(pth_dl, half2_name))
        and os.path.getsize(os.path.join(pth_dl, half2_name)) > 0
    ):
        print(f"  {half2_name} already exists\n")
        pass
    else:
        osprey("parn", half2, rptDate, rptDate, "half2", "csv")
        print(
            f"  {timediff(start_time_2, time.time())} downloading second of two subsets of holdings"
        )

    # combine dataframes of the two half sets of holdings https://pandas.pydata.org/docs/user_guide/merging.html
    df1 = pd.read_csv(os.path.join(pth_dl, half1_name))
    df2 = pd.read_csv(os.path.join(pth_dl, half2_name))
    df3 = pd.concat([df1, df2])
    # print(len(df1), len(df2), len(df1) + len(df2), len(df3))

    # write the combined dataframe to a csv file in the Downloads folder
    df3.to_csv(
        os.path.join(
            pth_dl,
            f"PARN ({len(full[hlf:]) + len(full[:hlf])}) {rptDate.strftime('%d%b%Y')}.csv",
        ),
        index=False,
    )

else:  # else get all the holdings in one go
    # check if the file was already downloaded before running osprey()
    # if os.path.isfile(pth_dl + r'\\' + full_name):
    if (
        os.path.exists(os.path.join(pth_dl, full_name))
        and os.path.getsize(os.path.join(pth_dl, full_name)) > 0
    ):
        print(f"  {full_name} already exists\n")
        pass
    else:
        start_time_h = time.time()
        osprey("parn", funds, rptDate, rptDate, "", "csv")
        print(
            f"  {timediff(start_time_h, time.time())} downloading \
all holdings"
        )

    print("", full_name, "\n")

print(
    f"{timediff(start_time, time.time())} downloading and then \
saving holdings data\n"
)

# download derivative metrics
start_time = time.time()
print("Downloading and then saving derivative data ...")

# check if the file was already downloaded before running osprey()
if os.path.isfile(os.path.join(pth_dl, derv_name)):
    print(f"  {derv_name} already exists")
    pass
else:
    osprey("derv", funds, rptDate, rptDate, "", "csv")

print(
    f"{timediff(start_time, time.time())} downloading and then \
saving derivative data\n"
)

print("\n Expected downloads:")
print(f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}")
print(f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}")
print(f" A summary sheet is{' not' if summ_yn == 'No' else ''} required", "\n")
print(
    f"{timediff(start_time, time.time())} getting the reporting date and \
        latest downloaded holdings and derivatives files",
    "\n",
)

print(
    f"{timediff(start_time_derivative_downloading, time.time())} downloading \
        holdings and derivative metrics. Next step is compiling.",
    "\n",
)

print("######################################")
print("#  END 1/4 DERV_CHECKER_DOWNLOADING  #")
print("######################################")
