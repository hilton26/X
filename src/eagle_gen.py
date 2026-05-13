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
from utilities import timediff, osprey
from constants import pthPy, report_types_dict

from dotenv import load_dotenv

load_dotenv()

start_time = time.time()
start_time_eagle_gen = start_time

print(f"{timediff(start_time, time.time())} importing libraries", "\n")

start_time = time.time()
print("Collecting input data ...")

# get fnd codes from the py_report.xlsm sheet and the close the sheet
df = pd.read_excel(pthPy, sheet_name="arc", usecols="X").dropna()
fnds_ = (",").join(df.iloc[:, 0].apply(str.upper))
s = "s" if len(fnds_.split(",")) > 1 else ""
print(fnds_)
# read the sheet into a dataframe

# get report parameters from the py_report.xlsm sheet and the close the sheet
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="AB")
df1
rpt_type = df1.iloc[4, 0]
date_from = df1.iloc[2, 0]
date_to = df1.iloc[0, 0]
ext = df1.iloc[5, 0]

# gather inputs
g = [
    report_types_dict[i][0]
    for i in report_types_dict
    if report_types_dict[i][0] == rpt_type
][0]  # get dictionary value (1st value)
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
print(f" {fnds_}", "\n")
# get the report
rpt = [i for i in report_types_dict if report_types_dict[i][0] == rpt_type][
    0
]  # get dictionary key
osprey(rpt, fnds_, date_from, date_to, "gen", ext)

print(
    f"{timediff(start_time, time.time())} getting {g} in {ext} format {k} for the {len(fnds_.split(','))} fund{s} -",
    "\n",
    f" {fnds_}",
    "\n",
)
print(f"Roundtrip time: {timediff(start_time_eagle_gen, time.time())}")

os.startfile(os.path.join(Path.home(), "Downloads"))
