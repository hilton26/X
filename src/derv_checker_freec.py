#!/usr/bin/env python
# coding: utf-8

# # Prepare the Derivative Free Cover Sheet

print("\n\n################################################")
print("#                                              #")
print("#     START 3/3 derv_checker_freecover.py X    #")
print("#                                              #")
print("################################################\n\n")

# import libraries
print("Importing libraries and setting up paths for the free cover report ...")

import time

start_time_free_cover = time.time()
start_time = time.time()

from datetime import datetime
import pandas as pd
import os, re
from constants import frcv_file, pthEXPORTS, pthSttlmnt
from utilities import timediff, prior_working_day, parn_de

# (1) get report date, selected summary sheet option, and derivative cover calc sheets
# df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
#     subset=[0]
# )
# k = df.iloc[2, 1]
# rptDate = (
#     k if isinstance(k, datetime) else prior_working_day(datetime.today())
# )  # prior working day or report date override; has type datetime()
# summ_yn = df.iloc[3, 1]
# funds = df[0].iloc[1:]
# funds = (",").join(funds.tolist())

# # derive holdings and derivatives file paths
# fPARN = os.path.join(
#     pth_dl,
#     f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )

# fDE = os.path.join(
#     pth_dl,
#     f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )

# # check if the required files have been downloaded, else
# if not os.path.exists(fPARN) or not os.path.exists(fDE):
#     sys.exit(
#         f"Stopping: missing expected download(s):\n"
#         f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
#         f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}"
#     )

fPARN, fDE, fCACT, funds, rptDate, summ_yn, dervthreshold = parn_de()

print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds:",
    "\n",
    f"{(', ').join(funds.tolist())}",
)

prrDate = prior_working_day(rptDate)

print(f" A summary sheet is{' not' if summ_yn == 'No' else ''} required", "\n")
print(
    f" {timediff(start_time, time.time())} getting the reporting \
date and latest downloaded holdings and derivatives files\n",
)

# derive current file name
filepath = os.path.join(pthEXPORTS, rptDate.strftime("%Y%m%d") + "_derv_calc.xlsx")
# print(filepath)
# get and then dataframe the current and most recent derv checker tables

# derive prior file name
rpt_str = rptDate.strftime("%Y%m%d")
pattern = re.compile(r"^(\d{8})_derv_calc\.xlsx$")

file_dates = [
    match.group(1)
    for fname in os.listdir(pthEXPORTS)
    if (match := pattern.match(fname))
]
prior_dates = sorted(d for d in file_dates if d < rpt_str)

if not prior_dates:
    raise FileNotFoundError(
        f"No '_derv_calc.xlsx' file found in {pthEXPORTS} before {rpt_str}"
    )

prior_filepath = os.path.join(pthEXPORTS, prior_dates[-1] + "_derv_calc.xlsx")
prior_file_date = datetime.strptime(prior_dates[-1], "%Y%m%d")
# print("\n\n", filepath, "\n", prior_filepath, "\n\n")

# create a lookup table for fund UT status and investment team
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A:B")

print(
    f"\n {rptDate.strftime('%a %d %b %Y')} current \
report date\n {prior_file_date.strftime('%a %d %b %Y')} prior \
report date \n {dervthreshold:.1f}% threshold\n"
)

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

prior_day = pd.read_excel(
    prior_filepath, sheet_name="Summary", usecols=[0, 1, 2, 3]
)  # prior day fund code, UT, #, and % columns

currn_day = pd.read_excel(
    filepath, sheet_name="Summary", usecols=[0, 1, 2, 3]
)  # current day day fund code, UT, #, and % columns

# print(prior_day)
# print(currn_day)

print(
    f" {timediff(start_time, time.time())} dataframing \
current and prior working day derivative summary files\n"
)

# (2) configure current day and prior day
# derv calc reports into a single dataframe
start_time = time.time()

fc = currn_day.merge(
    twoA, left_on=currn_day.columns[0], right_on="Fund Code"
)  # 6 columns

fc = fc.merge(
    prior_day, left_on="Fund Code", right_on=prior_day.columns[0]
)  # 10 columns

smmry = fc.iloc[:, [0, 3, 4, 7]]

# 1. Convert columns to a modifiable list
columns_list = list(smmry.columns)

# 2. Change the names using index positions
columns_list[1] = f"Free Cover {rptDate.strftime('%a %d %b %Y')} (% NAV)"
columns_list[3] = f"Free Cover {prior_file_date.strftime('%a %d %b %Y')} (% NAV)"

# 3. Reassign the list back to the DataFrame
smmry.columns = columns_list

# 4. Rearrange columns to be C, A, B (indices 2, 0, 1)
smmry = smmry.iloc[:, [0, 2, 1, 3]]

# 5. Add a 'Comment' column
smmry["Comment"] = None

# 6. Keep only the rows where the third column is below 10
smmry = smmry[smmry.iloc[:, 2] < 10]

# 7.Sort the second column in ascending order
smmry = smmry.iloc[smmry.iloc[:, 2].argsort()]

# 8. Paste smmry onto the 'Summary' sheet in Free Cover.xlsm
import win32com.client as win32  # !pip install pywin32

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(frcv_file)
ws = wb.Worksheets("Summary")

# (1) clear the contents of the Summary sheet
ws.Cells.ClearContents()

# (2) paste the dataframe at the top left corner of the sheet
ws.Range(ws.Cells(1, 1), ws.Cells(1, len(smmry.columns))).Value = list(smmry.columns)
ws.Range(
    ws.Cells(2, 1), ws.Cells(len(smmry.index) + 1, len(smmry.columns))
).Value = smmry.values.tolist()

wb.Save()
wb.Close()

print(f"\n{timediff(start_time_free_cover, time.time())} roundtrip\n\n")

print("\n\n################################################")
print("#                                              #")
print("#      END 3/3 derv_checker_freecover.py X     #")
print("#                                              #")
print("################################################\n\n")
