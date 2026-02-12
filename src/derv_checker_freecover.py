#!/usr/bin/env python
# coding: utf-8

# # Prepare the Derivative Free Cover Sheet

print("\n\n##############################################")
print("#     START 4/4 derv_checker_freecover.py    #")
print("##############################################\n\n")

# Import libraries

import time

start_time_free_cover = time.time()
start_time = time.time()

print("Importing libraries and setting up paths for the free cover report ...")

from datetime import datetime
import pandas as pd
import numpy as np
import xlwings as xw
import os
from constants import pthPy, frcv_file, pthEXPORTS, pthSttlmnt
from utilities import timediff, prior_working_day

# create a lookup table for fund UT status and investment team
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A:B")

print(
    f"{timediff(start_time, time.time())} importing libraries\
and setting up paths for the free cover report\n"
)

# Get report date and selected summary sheet option

start_time = time.time()
print(
    "Getting the reporting date and the comparative \
prior reporting date ..."
)

# extract override report date from cell "E1" in dervs sheet of py_reports.xlsm
df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="E", nrows=5)
k = df.iloc[2, 0]
rptDate = (
    k.date() if isinstance(k, datetime) else prior_working_day(datetime.today())
)  # prior working day or report date override; of type datetime.datetime()
prrDate = prior_working_day(rptDate)
dervthreshold = df.iloc[4, 0] * 100

print(
    f"\n {rptDate.strftime('%a %d %b %Y')} current \
report date\n {prrDate.strftime('%a %d %b %Y')} prior \
report date \n {dervthreshold:.1f}% threshold\n"
)

print(
    f"{timediff(start_time, time.time())} getting the \
reporting date and the comparative prior reporting date"
)

# dataframe the current and prior working day derivative summary files
start_time = time.time()
print(
    "Dataframing current and prior working day \
derivative summary files ..."
)

pr_fln = os.path.join(
    pthEXPORTS, f"Derv {prior_working_day(rptDate).strftime('%d%b%Y')}.xlsx"
)
cu_fln = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")

print(f"\n  {pr_fln}\n  {cu_fln}\n {dervthreshold}\n")

prior_day = pd.read_excel(
    pr_fln, sheet_name="Summary", usecols=[0, 1, 2, 3]
)  # prior day fund code, UT, #, and % columns
currn_day = pd.read_excel(
    cu_fln, sheet_name="Summary", usecols=[0, 1, 2, 3]
)  # current day day fund code, UT, #, and % columns

print(
    f"{timediff(start_time, time.time())} dataframing \
        current and prior working day derivative summary files\n"
)

# configure current day and prior day derv calc reports into a single datafrme
start_time = time.time()
print("Dataframing current and prior day derivative summaries ...")

fc = currn_day.merge(
    twoA, left_on=currn_day.columns[0], right_on="Fund Code"
)  # 6 columns
fc = fc.merge(
    prior_day, left_on="Fund Code", right_on=prior_day.columns[0]
)  # 10 columns

# drop superfluous columns
fc.drop(
    axis=1,
    columns=[fc.columns[4], fc.columns[6], fc.columns[7], fc.columns[8]],
    inplace=True,
)  # 6 columns

# rename column headings
fc = fc.set_axis(
    [
        "Fund Code",
        "UT",
        "#",
        f"Free Cover {rptDate.strftime('%a %d %b %Y')} (% NAV)",
        "Fund Name",
        f"Free Cover {prrDate.strftime('%a %d %b %Y')} (% NAV)",
    ],
    axis=1,
)

# reorder columns - https://www.geeksforgeeks.org/change-the-order-of-a-pandas-dataframe-columns-in-python/
fc = fc.iloc[:, [0, 4, 3, 5, 1, 2]]

# convert numbers to float type
fc[fc.columns[2:4]] = fc[fc.columns[2:4]].astype(float)

# identify funds below the dervthreshold
fc = fc[(fc[fc.columns[2]] < dervthreshold) & (fc["UT"] == "UT") & (fc["#"] != 0)]

# drop 'UT' and '#' columns
fc.drop(axis=1, columns=["UT", "#"], inplace=True)  # 4 columns

# add an empty 'Comment' column https://www.geeksforgeeks.org/how-to-add-empty-column-to-dataframe-in-pandas/
fc["Comment"] = ""

# fc

print(
    f"{timediff(start_time, time.time())} dataframing \
    current and prior day derivative summaries\n"
)

# paste the comparative dataframe onto the 'Summary' sheet in Free Cover.xlsm
start_time = time.time()
print("Saving the newest Free Cover.xlsm file ...")

# Open the Excel application
# !pip install pywin32
import win32com.client as win32

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

# open the workbook and select the worksheet
wb = excel.Workbooks.Open(frcv_file)
ws = wb.Worksheets("Summary")

# clear the contents of the MSCI_ALL sheet
ws.Cells.ClearContents()

# write the datframe headings to the first row of the Summary sheet
ws.Range(ws.Cells(1, 1), ws.Cells(1, len(fc.columns))).Value = list(fc)

# write the dataframe Summary sheet
ws.Range(
    ws.Cells(2, 1), ws.Cells(len(fc.index) + 1, len(fc.columns))
).Value = fc.values.tolist()

# save and close the workbook and then quit the Excel application
wb.Save()
wb.Close()
# excel.Visible = True

print(
    f"{timediff(start_time, time.time())} saving the newest\
        P:...\Free Cover.xlsm file\n"
)


print(
    f"{timediff(start_time_free_cover, time.time())} roundtrip \
        time for free cover file"
)

print("\n\n##############################################")
print("#      END 4/4 derv_checker_freecover.py     #")
print("##############################################\n\n")
