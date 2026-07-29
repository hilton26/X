#!/usr/bin/env python
# coding: utf-8

# # Merge the CS1 fund PARN sheets into one file to be used by cs1_reporting.ipynb

# https://stackoverflow.com/questions/20908018/import-multiple-excel-files-into-python-pandas-and-concatenate-them-into-one-dat

print("\n\n###################################################")
print("#                                                 #")
print("#         START cs1_PARN_merge_csv.py    X        #")
print("#                                                 #")
print("###################################################\n\n")

# libraries, libraries!
import time

start_time = time.time()
start_time_cs1_merge = time.time()

from datetime import datetime
import pandas as pd
from pathlib import Path
import os
from tqdm import tqdm, notebook  # notebook version of tqdm
from constants import pthPy, pth_dl, pthReports, pthTest
from utilities import timediff, prior_month_end, batch_list, osprey, r_classifier

# get inputs to pass to Eagle
start_time = time.time()
print("Collecting the report input data ...")

# get fund codes from r28_cs1 tab of the py_report.xlsm sheet
# df = pd.read_excel(pthPy, sheet_name="r28_cs1", usecols="A,F").dropna(subset = ['Fund'])
df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df.iloc[:, 0].apply(str.upper)

# get report date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=2)
k = df1.iloc[1, 0]
rptDate = k.date() if k == k else prior_month_end(datetime.today()).date()

# print inputs
print(
    f"\nPARN CS1 fund report{'' if len(funds) == 1 else 's'} to be merged as at {rptDate.strftime('%A %d %b %Y')}\n \
{len(funds)} fund{'' if len(funds) == 1 else 's'}: \n  {(',').join(funds)}"
)
print(f"\n{timediff(start_time, time.time())} collecting the report input data\n")

# ... dataframe the fund holdings in PARN format by looping over
# their files and appending to an initially empty dataframe
start_time = time.time()
print(f"Merging the PARN csv files into a dataframe ...")

num_batches = 2 if len(funds) > 1 else 1
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start=1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

holdings = pd.DataFrame()  # initialise an empty dataframe
for batch_filepath in batch_filepaths:
    data = pd.read_csv(batch_filepath)
    holdings = pd.concat([holdings, data])

# convert the holdings CS1 fund NAV columns to float64
cols_to_sum = ["Sum of Market Value Income", "Current Exposure"]
for col in cols_to_sum:
    holdings[col] = holdings[col].str.replace(",", "").astype("float64")

print(
    f" {timediff(start_time, time.time())} merging the PARN csv files with {len(holdings['Entity Name'].unique())} funds into a dataframe\n"
)

# get the fund NAVs and dataframe them
start_time = time.time()
s1 = "'s" if len(funds) == 1 else "s'"
s2 = "" if len(funds) == 1 else "s"
print(
    f"Getting the {len(funds)} CS1 fund{s1} NAV{s2} as at {rptDate.strftime('%A %d %B %Y')} with osprey() ...\n"
)

# derive the fund NAVs file name
name = "CS1"
navs_fln = os.path.join(
    Path.home(),
    "Downloads",
    f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)
print(
    f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv expected to be downloaded"
)

# if it doesn't yet exist, get the holdings ...
if os.path.isfile(navs_fln):
    print(f"\n  {navs_fln} exists\n")
    pass
else:
    sleep_time = 20
    count = 1
    number_of_tries = 10  # number of tries to download NAVs from Eagle
    while not os.path.isfile(navs_fln) and count < number_of_tries:
        print(f"\nDownloading the fund NAVs, try {count} of {number_of_tries} ...")
        osprey("fnav", (",").join(funds), rptDate, rptDate, name, "csv")
        count += 1
        print(f"Sleeping for {sleep_time} seconds before trying again ...\n")
        time.sleep(sleep_time)

# ... then dataframe the fund NAVs
navs_fln = os.path.join(
    Path.home(),
    "Downloads",
    f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)
navs = pd.read_csv(navs_fln)

# convert the Total column from object to float
navs["Total Net Assets"] = (
    navs["Total Net Assets"].str.replace(",", "").astype("float64")
)

print(f"\n{navs_fln}\n")

print(
    f" {timediff(start_time, time.time())} getting the {len(funds)} fund{s1} NAV{s2} as at {rptDate.strftime('%A %d %B %Y')} \
with osprey()"
)

# merge the CS1 fund holdings and NAVs, and compare their totals
start_time = time.time()
print(
    f"\nMerging and comparing the {len(funds)} CS1 fund holdings and NAV{s2} as at {rptDate.strftime('%A %d %B %Y')} ..."
)

holdings_totals = holdings.groupby("Entity ID", as_index=False).sum()[
    ["Entity ID", "Sum of Market Value Income", "Current Exposure"]
]
sums_cf = holdings_totals.merge(
    navs, how="left", left_on="Entity ID", right_on="NAV Entity ID"
)
sums_cf.drop(["Entity Name", "NAV Entity ID"], axis=1, inplace=True)
sums_cf["SoMVI-CE"] = (
    sums_cf["Sum of Market Value Income"] - sums_cf["Current Exposure"]
)
sums_cf["1-CE/SoMVI %"] = (
    1 - sums_cf["Current Exposure"] / sums_cf["Sum of Market Value Income"]
) * 100
sums_cf["SoMVI-NAV"] = abs(
    sums_cf["Sum of Market Value Income"] - sums_cf["Total Net Assets"]
)
sums_cf["1-NAV/SoMVI %"] = (
    1 - sums_cf["Total Net Assets"] / sums_cf["Sum of Market Value Income"]
) * 100
sums_cf = sums_cf.sort_values(by="SoMVI-NAV", ascending=False)
cols_order = [
    "Effective Date",
    "Entity ID",
    "Sum of Market Value Income",
    "Current Exposure",
    "Total Net Assets",
    "SoMVI-CE",
    "1-CE/SoMVI %",
    "SoMVI-NAV",
    "1-NAV/SoMVI %",
]
sums_cf = sums_cf[cols_order]

# set the number of decimals to present
cols_2dp = ["Sum of Market Value Income", "Current Exposure", "Total Net Assets"]
for col in cols_2dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.2f}")

cols_6dp = ["SoMVI-CE", "1-CE/SoMVI %", "SoMVI-NAV", "1-NAV/SoMVI %"]
for col in cols_6dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.6f}")

# print(sums_cf)

print(
    f" {timediff(start_time, time.time())} merging and comparing the {len(funds)} CS1 fund holdings and NAVs as at {rptDate.strftime('%A %d %B %Y')}"
)

# convert the PARN holdings into Reg 28 format with corresponding headings
start_time = time.time()
print(
    f"\nConverting the CS1 fund PARN holdings in readiness for Reg 28 classification ..."
)

s = "" if len(funds) == 1 else "s"
cs1_fname = os.path.join(
    pthTest, f"CS1 PARN holdings ({len(funds)}) {rptDate.strftime('%d%b%Y')}.xlsx"
)

hold_cols = [
    "Entity ID",
    "Investment Type",
    "i Issue Name",
    "PrimaryAssetID",
    "CCY",
    "Sum of Market Value Income",
    "% of Total Market Value",
    "Current Exposure",
]
hReg28 = holdings[hold_cols]  # identify the subset of holdings columns to be used
hReg28 = hReg28.rename(
    columns={
        "Entity ID": "Entity Name",
        "PrimaryAssetID": "Primary Asset ID",
        "Sum of Market Value Income": "End Market Value",
        "% of Total Market Value": "Percentage of Market Value",
        "Current Exposure": "Closing Exposure PA",
    }
)
hReg28.insert(
    5, "Reg28 Classification", ""
)  # insert the classification column as the new column 5
hReg28.insert(
    9, f"{rptDate.strftime('%d %b %Y')}", ""
)  # insert the report date as a header in the last column
hReg28.iloc[0, 9] = cs1_fname
hReg28.iloc[1, 9] = "CS1"
hReg28.reset_index(drop=True, inplace=True)

# hReg28.info()

print(
    f" {timediff(start_time, time.time())} converting the CS1 fund PARN holdings in readiness for Reg 28 classification\n"
)

# write the CS1 holdings dataframe to review it as a worksheet
start_time = time.time()
print(f"\nWriting the CS1 fund holdings dataframe and navs dataframe to a sheet ...")

with pd.ExcelWriter(cs1_fname, engine="xlsxwriter") as writer:
    hReg28.to_excel(writer, index=False, sheet_name="All")  # write the NAV sheet
    holdings.to_excel(writer, index=False, sheet_name="PARN")  # write the NAV sheet
    sums_cf.to_excel(
        writer, index=False, sheet_name="NAVs"
    )  # write the missing NAVs sheet

# print(f"  {cs1_fname}")

print(
    f" {timediff(start_time, time.time())} writing the CS1 fund holdings dataframe and navs dataframe to a sheet\n"
)

os.startfile(cs1_fname)

print(f" {timediff(start_time_cs1_merge, time.time())} merging CS! PARN reports\n")

# run r_classifier function from utilities.py
r_classifier("cs1", cs1_fname, rptDate)


print("\n\n###################################################")
print("#                                                 #")
print("#          END cs1_PARN_merge_csv.py    X         #")
print("#                                                 #")
print("###################################################\n\n")
