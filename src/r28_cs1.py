#!/usr/bin/env python
# coding: utf-8

# # Prepare CS1 Reports From PARN Download to Report Saving

print("\n\n#######################")
print("#   START r28_cs1.py  #")
print("#######################\n\n")

import time

start_time = time.time()
start_time_r28_cs1 = start_time

# libraries, libraries!
print("Importing libraries and setting up paths ...")
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
from tqdm import tqdm, notebook
from constants import pthPy, pthTest, pth_dl
from utilities import (
    timediff,
    last_working_day,
    prior_month_end,
    osprey,
    batch_list,
    r_classifier,
)
import subprocess

# get inputs to pass to Eagle
start_time = time.time()
print("Collecting input data ...")

# get report fund codes from the py_report.xlsm 'arc' sheet
funds = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()  # funds
funds.iloc[:, 0] = funds.iloc[:, 0].str.upper()  # capitalise fund codes
funds.columns = ["Fund"]  # rename column
print(funds)

# get report date
df = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows = 9)
k = df.iloc[1, 0]
rptDate = (
    k if k ==k else prior_month_end(datetime.today().date())
)  # prior month end or report date override; type is datetime()
print(rptDate, k)

# check inputs
s = "" if len(funds) == 1 else "s"
print(
    f"{len(funds)} PARN CS1 fund report{s} as at {rptDate.strftime('%A %d %b %Y')} to be downloaded:\n  {(', ').join(funds)}"
)
print(f"\n{timediff(start_time, time.time())} collecting input data\n")

# download the PARN reports in batches
start_time = time.time()
num_batches = 2 if len(funds) != 1 else 1
print(
    f"Downloading the {len(funds)} fund holdings for {rptDate.strftime('%a %d %b %Y')} in {num_batches} batches ..."
)

batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start=1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    print(f"Get {filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)
    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath}...\n")
        osprey("parn", (",").join(batch), rptDate, rptDate, fln, "csv")

print(
    f"\n{timediff(start_time, time.time())} downloading the {len(funds)} fund holdings for {rptDate.strftime('%a %d %b %Y')} in {num_batches} batches"
)

# join the downloaded holding reports into a dataframe
start_time = time.time()
print("Dataframing the fund holding reports for CS1 reporting\n")

holdings = pd.DataFrame()
for batch_filepath in batch_filepaths:
    # print(f" {batch_filepath}")
    df_new = pd.read_csv(batch_filepath)
    holdings = pd.concat([holdings, df_new])

# convert date columns from type object
date_cols = ["Next Coupon Date", "Maturity Date", "i Position Effective Date"]
for date_col in date_cols:
    holdings[date_col] = pd.to_datetime(holdings[date_col])

# convert value columns from type object to type float
value_cols = [
    "Original Nominal",
    "Clean Book Value",
    "Clean Market Value",
    "Accrued Income",
    "Dividend Receivable",
    "Sum of Market Value Income",
    "Market Value %",
    "Current Exposure",
]
for value_col in value_cols:
    holdings[value_col] = (
        holdings[value_col].astype(str).str.replace(",", "").astype(float)
    )

print(
    f" {len(holdings['Entity Name'].unique())} fund holdings \
as at {holdings['i Position Effective Date'].iloc[0].strftime('%d %b %Y')} in the dataframe"
)

print(
    f"\n{timediff(start_time, time.time())} dataframing the fund holding reports for CS1 reporting\n"
)

# get the fund NAVs
start_time = time.time()
print(
    f"Getting the {len(funds)} CS1 funds' NAVs as at {rptDate.strftime('%A %d %B %Y')} with osprey() ..."
)

name = "CS1"
navs_fln = os.path.join(
    pth_dl, f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
)

if os.path.exists(navs_fln):
    print(
        f" CS1 fund NAVs as at {rptDate.strftime('%a %d %b %Y')} already downloaded: {navs_fln}"
    )
    pass
else:
    osprey("fnav", (",").join(funds), rptDate, rptDate, name, "csv")

# dataframe the downloaded fund NAVs
navs = pd.read_csv(navs_fln)

# convert the Total column from object to float
navs["Total Net Assets"] = (
    navs["Total Net Assets"].str.replace(",", "").astype("float64")
)
# navs['Total Net Assets'] = navs['Total Net Assets'].apply(lambda x: f"{x:,.2f}") # present with thousands separator and to two decimals

print("\n", navs_fln)

print(
    f" {timediff(start_time, time.time())} getting the {len(funds)} funds' NAV{'s' if len(funds) != 1 else ''} as at {rptDate.strftime('%A %d %B %Y')} \
with osprey()"
)

# merge the CS1 fund holdings and NAVs, and compare their totals
start_time = time.time()
print(
    f"Merging and comparing the {len(funds)} CS1 fund holdings and NAVs as at {rptDate.strftime('%A %d %B %Y')} ..."
)

holdings_totals = holdings.groupby("Entity ID", as_index=False).sum(numeric_only=True)[
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

# sums_cf
# ...

print(
    f" {timediff(start_time, time.time())} merging and comparing the {len(funds)} \
CS1 fund holdings and NAVs as at {rptDate.strftime('%A %d %B %Y')}"
)

# convert the PARN holdings into Reg 28 format with correspodning headings
start_time = time.time()
print(
    f"Converting the CS1 fund PARN holdings in readiness for Reg 28 classification ..."
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
    f" {timediff(start_time, time.time())} converting the CS1 fund PARN holdings in readiness for Reg 28 classification"
)

# write the CS1 holdings dataframe to review it as a worksheet
start_time = time.time()
print("Writing the CS1 fund holdings dataframe and navs dataframe to a sheet ...")

writer = pd.ExcelWriter(cs1_fname, engine="xlsxwriter")  # instantiate a sheet writer
hReg28.to_excel(writer, index=False, sheet_name="All")  # write the NAV sheet
holdings.to_excel(writer, index=False, sheet_name="PARN")  # write the NAV sheet
sums_cf.to_excel(writer, index=False, sheet_name="NAVs")  # write the missing NAVs sheet
writer.close()  # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets
print(f" \n{cs1_fname}\n")

print(
    f" {timediff(start_time, time.time())} writing the CS1 fund holdings dataframe and navs dataframe to a sheet\n"
)

# run the CS1 reporting script
start_time = time.time()
print(f"Classifying the holdings for the CS1 reports\n")

r_classifier(
    "cs1",
    cs1_fname,
)

print(
    f" \n{timediff(start_time, time.time())} classifying the holdings for the CS1 reports\n"
)

# run the CS1 reporting script
start_time = time.time()
print(f"Generating the CS1 reports\n")

subprocess.run(
    ["python", "C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/cs1_reporting.py"]
)

print(f" \n{timediff(start_time, time.time())} generating the CS1 reports\n")

print(
    "\n",
    f"{timediff(start_time_r28_cs1, time.time())} roundtripping download, merge and Reg 28 CS1 reports \n",
)

print("\n\n#######################")
print("#    END r28_cs1.py   #")
print("#######################\n\n")
