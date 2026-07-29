#!/usr/bin/env python
# coding: utf-8

# # Pull Holdings and Derivative Data from Eagle

# How to wait until Element is Visible in Selenium Python
#
# https://pythonexamples.org/python-selenium-wait-until-element-is-visible/

print("\n\n###############################################")
print("#                                             #")
print("#   START 1/4 derv_checker_downloading.py X   #")
print("#                                             #")
print("###############################################\n\n")

import time

start_time = time.time()
start_time_derivative_downloading = time.time()
print("\n\nImporting libraries ...\n")

# load libraries
import pandas as pd
import os, sys, subprocess
from send2trash import send2trash
from constants import pthEXPORTS, pth_dl, pthTest, pthOverdrafts
from utilities import timediff, osprey, parn_de

fund_load = 200  # else 2 separate "half1" and "half2" osprey() calls
cact_sets = 6  # number of batches of CACT download files

# get report date and selected summary sheet option
fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()


print(f"{timediff(start_time, time.time())} importing libraries\n")

# derive file names
start_time = time.time()
print("Deriving file names ...")

filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
fPARN = os.path.join(
    pth_dl,
    f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)
fDE = os.path.join(
    pth_dl,
    f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)

full_name = f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
derv_name = f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
bank_file_sa = pthOverdrafts + rf"\{rptDate.strftime('%Y%m%d')}_overdrafts_sa.xlsx"

print(
    f"\n {rptDate.strftime('%A %d %b %Y')} \
for {len(funds)} funds:\n {(',').join(funds)}\n",
)

# download derivative metrics
start_time = time.time()
print("Downloading and then saving derivative data ...")

# check if the file was already downloaded before running osprey()
if os.path.isfile(os.path.join(pth_dl, derv_name)):
    print(f"  {derv_name} already exists")
    pass
else:
    osprey("derv", (",").join(funds), rptDate, rptDate, "", "csv")

print(
    f" {timediff(start_time, time.time())} downloading and then \
saving derivative data\n"
)

# get the fund holdings in "portfolio analytics review - new" format
start_time = time.time()
print("Downloading and then saving holdings data ...")

if len(funds) > fund_load:  # if more than 100 funds are in the list ...
    # ... get holdings for the first half of funds in the list
    start_time_1 = time.time()
    print(f"  ... downloading first of two subsets of holdings: {half1_name}")

    # derive holdings full and half downl;oad file names expected
    hlf = int(len(funds) / 2)
    half1 = (",").join(funds[:hlf])
    half1_name = f"PARN half1({len(funds[:hlf])}) {rptDate.strftime('%d%b%Y')}.csv"
    half2 = (",").join(funds[hlf:])
    half2_name = f"PARN half2({len(funds[hlf:])}) {rptDate.strftime('%d%b%Y')}.csv"

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
            f"  {timediff(start_time_1, time.time())} downloading \
first of two subsets of holdings"
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
            f"  {timediff(start_time_2, time.time())} downloading \
second of two subsets of holdings"
        )

    df1 = pd.read_csv(os.path.join(pth_dl, half1_name))
    df2 = pd.read_csv(os.path.join(pth_dl, half2_name))
    df_parn = pd.concat([df1, df2])
    # print(len(df1), len(df2), len(df1) + len(df2), len(df_parn))

    # write the combined dataframe to a csv file in the Downloads folder
    df_parn.to_csv(
        os.path.join(
            pth_dl,
            f"PARN ({len(funds[hlf:]) + len(funds[:hlf])}) {rptDate.strftime('%d%b%Y')}.csv",
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
        osprey("parn", (",").join(funds), rptDate, rptDate, "", "csv")
        print(
            f"  {timediff(start_time_h, time.time())} downloading \
all holdings"
        )

    print("", full_name, "\n")

print(
    f"{timediff(start_time, time.time())} downloading and then \
saving holdings data\n"
)

print(
    f"\n {timediff(start_time_derivative_downloading, time.time())} \
downloading holdings, derivative metrics, and cash activities \
metrics. Next step is compiling.\n",
)

# save daily bank reconciliation files
start_time = time.time()
print("\n\nSaving daily bank recons from Outlook")
bank = os.path.join(os.path.dirname(__file__), "overdrafts.py")
subprocess.run([sys.executable, bank])
print(
    f" {timediff(start_time, time.time())} \
saving daily bank recons from Outlook\n\n"
)

# test that holdings (fPARN) downloaded
if os.path.exists(fPARN):
    df_fPARN = pd.read_csv(fPARN)
    test_fPARN = (
        os.path.exists(fPARN) and df_fPARN.columns[1] == "Valuation First Level"
    )
    if not test_fPARN:
        send2trash.sendtotrash(fPARN)

# test that derivatives (fDE) downloaded
if os.path.exists(fDE):
    df_fDE = pd.read_csv(fDE)
    test_fDE = os.path.exists(fDE) and df_fDE.columns[6] == "Effective Exposure"
    if not test_fDE:
        send2trash.sendtotrash(fDE)

print(f"\n Expected downloads for {rptDate.strftime('%A %d %B %Y')}:")
print(
    f"  {os.path.basename(fPARN)} which {'exists' if test_fPARN else 'does not exist'}"
)
print(f"  {os.path.basename(fDE)} which {'exists' if test_fDE else 'does not exist'}")
print(
    f"  {os.path.basename(bank_file_sa)} which {'exists' if os.path.exists(bank_file_sa) else 'does not exist'}"
)

print("\n\n###############################################")
print("#                                             #")
print("#    END 1/4 derv_checker_downloading.py X    #")
print("#                                             #")
print("###############################################\n\n")
