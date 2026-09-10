#!/usr/bin/env python
# coding: utf-8

# # Pull Data from Eagle in a Pre-Set Format

# How to wait until Element is Visible in Selenium Python
#
# https://pythonexamples.org/python-selenium-wait-until-element-is-visible/

print("Importing libraries ...")
import time
from datetime import datetime
import os  # to open the Downloads folder when done
import pandas as pd
from pathlib import Path
from tqdm import tqdm
from constants import pthPy, report_types_dict, pth_dl, ptl_trades
from utilities import timediff, osprey, batch_list

start_time = time.time()
start_time_eagle_gen = start_time

print(f" {timediff(start_time, time.time())} importing libraries", "\n")

start_time = time.time()
print("Collecting input data ...")

# get fnd codes from the py_report.xlsm sheet and the close the sheet
df = pd.read_excel(pthPy, sheet_name="arc", usecols="AQ").dropna()
fnds_ = (",").join(df.iloc[:, 0].apply(str.upper))
s = "s" if len(fnds_.split(",")) > 1 else ""
print(f" {fnds_}")
# read the sheet into a dataframe

# get report parameters from the py_report.xlsm sheet and the close the sheet
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="AU", nrows=7)
df1
rpt_type = df1.iloc[4, 0]
date_from = df1.iloc[2, 0]
date_to = df1.iloc[0, 0]
ext = df1.iloc[5, 0]
num_batches = df1.iloc[6, 0]

print(report_types_dict)

# gather inputs
g = [
    report_types_dict[i][0]
    for i in report_types_dict
    if report_types_dict[i][0] == rpt_type
][0]  # get dictionary value (1st value)


# function to find the key given a dictionary and a value
def find_key_by_label(d, label):
    return next((k for k, v in d.items() if v[0] == label), None)


key = find_key_by_label(report_types_dict, g)
print(key.upper())


k = (
    f"on {date_from.strftime('%A %d %B %Y')}"
    if date_from == date_to
    else f"from {date_from.strftime('%A %d %B %Y')} to {date_to.strftime('%A %d %B %Y')}"
)  # get the report date

# check inputs
print(
    "\n",
    f"{[report_types_dict[i][0] for i in report_types_dict if report_types_dict[i][0] == rpt_type][0]} \
in {ext} format {k} for {len(fnds_.split(','))} fund{s} -",
    "\n",
    f"  {fnds_}",
    "\n",
)
print(f"{timediff(start_time, time.time())} collecting input data", "\n")

start_time = time.time()
print(f"Getting {g} in {ext} format {k} for the {len(fnds_.split(','))} fund{s} ...")
print(f" {fnds_}\n")
# get the report
rpt = [i for i in report_types_dict if report_types_dict[i][0] == rpt_type][
    0
]  # get dictionary key


# download the PARN reports in batches
batch_size = int(len(fnds_.split(",")) / num_batches)
batches = batch_list(
    fnds_.split(","), batch_size=min(len(fnds_.split(",")), batch_size)
)
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start=1)):
    fln = f"{index}_of_{len(batches)}_"
    filename = f"{key.upper()} {fln}({len(batch)}) {date_from.strftime('%d%b%Y')} to {date_to.strftime('%d%b%Y')}.csv"
    print(f"{filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}\n")
    # print(f" {len(batch)} files:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

    # print(f"\n\n\n\n{batch_filepath}\n{os.path.isfile(batch_filepath)}\n\n\n\n")

    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath}...\n")
        osprey(key, (",").join(batch), date_from, date_to, fln, "csv")

for batch_filepath in batch_filepaths:
    print(batch_filepath)

# concatenate the batches into a list
dfs = []
for batch_filepath in batch_filepaths:
    df = pd.read_csv(batch_filepath)
    dfs.append(df)

# concatenate the dataframes
if dfs:
    df_combined = pd.concat(dfs, ignore_index=True)
else:
    df_combined = pd.DataFrame()

df_combined.iloc[:, 19:40].info()

# convert date columns to datetime format
date_cols = ["Trade Date", "Settlement Date"]
df_combined[date_cols] = (
    df_combined[date_cols]
    .apply(pd.to_datetime, format="%m/%d/%Y")
    .apply(lambda x: x.dt.date)
)


# convert value columns from object to numeric format
value_cols = (
    list(range(5, 11))
    + list(range(17, 22))
    + list(range(25, 27))
    + list(range(28, 30))
    + [35]
    + list(range(37, 39))
)
df_combined.iloc[:, value_cols] = df_combined.iloc[:, value_cols].apply(
    lambda x: pd.to_numeric(x.replace(",", ""), errors="coerce")
)


# write the dataframe to an xlsx file
output_filename = (
    f"{key.upper()} {date_from.strftime('%d%b%Y')} to {date_to.strftime('%d%b%Y')}.xlsx"
)
output_filepath = os.path.join(ptl_trades, output_filename)
df_combined.to_excel(output_filepath, index=False)

os.startfile(ptl_trades)

print(
    f"{timediff(start_time, time.time())} getting {g} in {ext} format {k} for the {len(fnds_.split(','))} fund{s} -",
    "\n",
    f" {fnds_}\n",
)
print(f"Roundtrip time: {timediff(start_time_eagle_gen, time.time())}")

os.startfile(os.path.join(Path.home(), "Downloads"))
