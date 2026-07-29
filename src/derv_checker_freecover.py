#!/usr/bin/env python
# coding: utf-8

# # Prepare the Derivative Free Cover Sheet

print("\n\n################################################")
print("#                                              #")
print("#     START 4/4 derv_checker_freecover.py X    #")
print("#                                              #")
print("################################################\n\n")

# import libraries
print(
    "Importing libraries and setting up paths for the \
free cover report ..."
)

import time

start_time_free_cover = time.time()
start_time = time.time()

import pandas as pd
import os, sys
from constants import frcv_file, pthEXPORTS, pthSttlmnt
from utilities import timediff, prior_working_day, parn_de

# create a lookup table for fund UT status and investment team
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A:B")

print(
    f" {timediff(start_time, time.time())} importing libraries \
and setting up paths for the free cover report\n"
)

# get report date and selected summary sheet option

start_time = time.time()
print(
    "Getting the reporting date and the comparative \
prior reporting date ..."
)

fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()

# check if the required files have been downloaded, else continue
if not os.path.exists(fPARN) or not os.path.exists(fDE):
    sys.exit(
        f"Stopping: missing expected download(s):\n"
        f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
        f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}\n"
    )

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

# check if the required files have been downloaded, else
# if not os.path.exists(fPARN) or not os.path.exists(fDE):
#     sys.exit(
#         f"Stopping: missing expected download(s):\n"
#         f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
#         f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}"
#     )

print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds:",
    "\n",
    f"{(', ').join(funds.tolist())}",
)

prrDate = prior_working_day(rptDate)

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

# # check the column headings for column fc[0]
# print(fc.head(3))

# drop superfluous columns (positional: the first merge already dedupes the
# "Fund Code" key, so the prior day's duplicate fund code/UT/# land at 5,6,7)
keep_cols = [i for i in range(len(fc.columns)) if i not in (5, 6, 7)]
fc = fc.iloc[:, keep_cols]  # 6 columns

# print(fc.shape)

# # reset fc to include all columns except the second last one
# fc = fc.iloc[:, [i for i in range(fc.shape[1]) if i != fc.shape[1] - 2]]

# reset fc to include all columns except the "Fund Code" column
fc = fc.drop(columns="Fund Code")

# rename column headings
fc = fc.set_axis(
    [
        "Fund Code",
        "UT",
        "#",
        rf"Free Cover {rptDate.strftime('%a %d %b %Y')} (% NAV)",
        "Fund Name",
        rf"Free Cover {prrDate.strftime('%a %d %b %Y')} (% NAV)",
    ],
    axis=1,
)

# reorder columns
fc = fc.iloc[:, [0, 4, 3, 5, 1, 2]]

# merge with twoA to get fund long names
fc = pd.merge(fc, twoA, how="left", on="Fund Code")
fc = fc.rename(columns={"Fund Name_y": "Fund Name"}).drop(columns="Fund Name_x")
fc.insert(1, "Fund Name", fc.pop("Fund Name"))

# convert numbers to float type
fc[fc.columns[2:4]] = fc[fc.columns[2:4]].astype(float)

# identify funds below the dervthreshold
fc = fc[(fc[fc.columns[2]] < dervthreshold) & (fc["UT"] == "UT") & (fc["#"] != 0)]

# drop 'UT' and '#' columns
fc.drop(columns=["UT", "#"], inplace=True)  # 4 columns

# add an empty 'Comment' column https://www.geeksforgeeks.org/how-to-add-empty-column-to-dataframe-in-pandas/
fc["Comment"] = ""

# # view shape and columns of dataframe fc
# print(fc.head(3))

print(
    f"{timediff(start_time, time.time())} dataframing \
current and prior day derivative summaries\n"
)

# paste the comparative dataframe onto the 'Summary' sheet in Free Cover.xlsm
start_time = time.time()
print("Saving the newest Free Cover.xlsm file ...")

# Open the Excel application
import win32com.client as win32  # !pip install pywin32

excel = win32.gencache.EnsureDispatch("Excel.Application")
excel.Visible = False

# open the workbook and select the worksheet
wb = excel.Workbooks.Open(frcv_file)
ws = wb.Worksheets("Summary")

# clear the contents of the MSCI_ALL sheet
ws.Cells.ClearContents()

# write the dataframe headings to the first row of the Summary sheet
ws.Range(ws.Cells(1, 1), ws.Cells(1, len(fc.columns))).Value = list(fc)

# print(fc)
# write the dataframe Summary sheet
ws.Range(
    ws.Cells(2, 1), ws.Cells(len(fc.index) + 1, len(fc.columns))
).Value = fc.values.tolist()

# save and close the workbook and then quit the Excel application
wb.Save()
wb.Close()
# excel.Visible = True

print(
    f"{timediff(start_time, time.time())} saving the \
newest Free Cover.xlsm file\n"
)

print(
    f"{timediff(start_time_free_cover, time.time())} \
roundtrip time for free cover file"
)

print("\n\n################################################")
print("#                                              #")
print("#       END 4/4 derv_checker_freecover.py X    #")
print("#                                              #")
print("################################################\n\n")
