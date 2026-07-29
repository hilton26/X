#!/usr/bin/env python
# coding: utf-8

# # Downloading PGF UCITS Hedge Share Class NAVs and Holdings

print("\n\n#######################################")
print("#                                     #")
print("#    START 1/2 pgf_downloading.py X   #")
print("#                                     #")
print("#######################################\n\n")

# get input data off source sheet

import time

start_time_pgf_dl = time.time()
start_time = time.time()

# libraries, libraries!
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
from constants import pthPy, pth_dl
from utilities import osprey, timediff, prior_working_day

print("Getting input variables ...")
# get report inputs
df = pd.read_excel(pthPy, sheet_name="arc", usecols="G:I")
k = df.iloc[1, 2]
rptDate = (
    k
    if isinstance(k, datetime) and not pd.isna(k)
    else prior_working_day(datetime.today())
)  # prior working day or report date override; has type datetime()

# get portfolio lists as strings
clss_navs = (",").join(df["pgf: UT prices"].dropna())  # for unit trust prices
clss_hldgs = (",").join(df["pgf: PAR-N"].dropna())  # for class holdings

# check inputs
print(
    f" Report date:          {rptDate.date()}",
    "\n",
    f"Class NAV codes:      {clss_navs}",
    "\n",
    f"Class holdings codes: {clss_hldgs}",
)
print(f"{timediff(start_time, time.time())} getting input variables", "\n")

# get class NAVs in csv format
start_time = time.time()
print("Getting the unit trust prices report ...")

# print the names of the downloaded unit trust prices and the holdings files
parN_nm = os.path.join(
    pth_dl,
    f"UTPS PGF_UT_prices({len(df['pgf: UT prices'].dropna())}) {rptDate.strftime('%d%b%Y')}.csv",
)
utP_nm = os.path.join(
    pth_dl,
    f"PARN PGF_Holdings({len(df['pgf: PAR-N'].dropna())}) {rptDate.strftime('%d%b%Y')}.csv",
)

try:
    start_time_prices = time.time()
    print("... downloading PGF unit trust prices")
    # check if the file was already downloaded before running osprey()
    if os.path.isfile(parN_nm):
        print(f"  {parN_nm} already exists")
        pass
    else:
        osprey("utps", clss_navs, rptDate, rptDate, "PGF_UT_prices", "csv")
        print(
            f"  {timediff(start_time_prices, time.time())} getting PGF unit trust prices"
        )
except Exception as e:
    print(e)

print(f"{timediff(start_time, time.time())} getting the unit trust prices report", "\n")

# get class holdings in csv format
start_time = time.time()
print("Getting the class holdings report ...")

try:
    start_time_navs = time.time()
    print("... downloading PGF unit trust prices")
    # check if the file was already downloaded before running osprey()
    if os.path.isfile(utP_nm):
        print(f"  {utP_nm} already exists")
        pass
    else:
        osprey("parn", clss_hldgs, rptDate, rptDate, "PGF_Holdings", "csv")
        print(
            f"  {timediff(start_time_navs, time.time())} getting PGF unit trust prices"
        )
except Exception as e:
    print(e)

print(f"{timediff(start_time, time.time())} getting the class holdings report", "\n")
print(
    f"{timediff(start_time_pgf_dl, time.time())} roundtrip time to get unit trust prices and holdings reports",
    "\n",
)

print(f"{rptDate.strftime('%a %d %B %Y')} downloads")
print(f"  {parN_nm} which {'exists' if os.path.exists(parN_nm) else 'does not exist'}")
print(f"  {utP_nm} which {'exists' if os.path.exists(utP_nm) else 'does not exist'}")

print("\n\n#######################################")
print("#                                     #")
print("#     END 1/2 pgf_downloading.py X    #")
print("#                                     #")
print("#######################################\n\n")
