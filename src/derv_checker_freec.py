#!/usr/bin/env python
# coding: utf-8

# # Prepare the Derivative Free Cover Sheet

print("\n\n################################################")
print("#                                              #")
print("#     START 3/3 derv_checker_freecover.py X    #")
print("#                                              #")
print("################################################\n\n")

# import libraries
print(
    "Importing libraries and setting \
up paths for the free cover report ..."
)

import time
import pandas as pd
import os
from constants import frcv_file, pthEXPORTS, pthSttlmnt
from utilities import timediff, prior_working_day, parn_de

start_time_free_cover = time.time()
start_time = time.time()

# get report variables from parn_de() function in utilities.py
(fPARN, fDE, funds, rptDate, summ_yn, dervthreshold, batches) = parn_de()

# print funds and reporting date
print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds:\n",
    f" {(', ').join(funds.tolist())}\n",
)

# derive prior report date
prrDate = prior_working_day(rptDate)
print(f" A summary sheet is{' not' if summ_yn == 'No' else ''} required")
print(f" {dervthreshold:.1f}% is the cover threshold")
print(f" {prrDate.strftime('%a %d %b %Y')} is the prior report date\n")

# derive current file name
filename = os.path.join(pthEXPORTS, rptDate.strftime("%Y%m%d") + "_derv_calc.xlsx")
# print(filename)

# create a lookup table for fund UT status and investment team
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A:B")

print(
    f"{timediff(start_time, time.time())} getting the \
reporting date and the comparative prior reporting date"
)
####################

# dataframe the current and prior working day derivative summary files
start_time = time.time()
print(
    "\n\nDataframing current and prior working day \
derivative summary files ..."
)

# get fund code, current & prior day, and fund name columns
total_cols = len(pd.read_excel(filename, nrows=0).columns)
df = pd.read_excel(
    filename, sheet_name="Summary", usecols=[0, 1, 2, 3, 4, total_cols - 1]
)  # current day day fund code, UT, #, and % columns
# print(df)
summary = df[(df["Cash Cover"] < 10) & df["UT?"].isin(["UT"]) & (df["#"] != 0)].copy()
cols_to_drop = ["UT?", "#"]
summary.drop(columns=cols_to_drop, inplace=True)

# rearrange the order of the columns
summary = summary.iloc[:, [0, 3, 1, 2]]

# rename 'Cash Cover' column
renamed_cols = {
    "Cash Cover": f"Cash Cover \
{rptDate.strftime(r'%a %d %b %Y')}"
}
summary = summary.rename(columns=renamed_cols)
summary["Comment"] = None

# paste summary onto the 'Summary' sheet in Free Cover.xlsm
import win32com.client as win32  # !pip install pywin32

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(frcv_file)
ws = wb.Worksheets("Summary")

# (1) clear the contents of the Summary sheet
ws.Cells.ClearContents()

# (2) paste the dataframe headings
ws.Range(ws.Cells(1, 1), ws.Cells(1, len(summary.columns))).Value = list(
    summary.columns
)

# (3) paste the dataframe values
ws.Range(
    ws.Cells(2, 1), ws.Cells(len(summary.index) + 1, len(summary.columns))
).Value = summary.values.tolist()

wb.Save()
wb.Close()

print(f"\n{timediff(start_time_free_cover, time.time())} roundtrip\n\n")

print("\n\n################################################")
print("#                                              #")
print("#      END 3/3 derv_checker_freecover.py X     #")
print("#                                              #")
print("################################################\n\n")
