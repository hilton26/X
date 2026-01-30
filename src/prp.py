#!/usr/bin/env python
# coding: utf-8

# # Proportion Reg 28 and Reg 30 Reports

# libraries, libraries!

import time

start_time_prpn = time.time()
start_time = time.time()
print(f"Importing libraries ...")

from datetime import datetime
import pandas as pd
import numpy as np
import os
from tqdm import tqdm
from constants import pthPy, pthReports
from utilities import timediff, prior_month_end

print(f"Importing libraries completed: {timediff(start_time, time.time())}\n")

# get report variables

start_time = time.time()
print("Getting report variables ...")

# report variables
sr = pd.read_excel(pthPy, sheet_name="arc", index_col=None, usecols="AL", nrows=4)
prn = str(sr.iat[2, 0]).upper()  # fund name
rpt_type = sr.iat[3, 0]
rpt_type  # Reg28 or Reg30 report
k = sr.iat[1, 0]  # report date
rptDate = (
    k.date() if isinstance(k, datetime) else prior_month_end(datetime.today()).date()
)
fnm = os.path.join(
    pthReports, f"{prn.upper()} {rpt_type} {rptDate.strftime('%d%b%Y')}.xlsx"
)
print(f" Fund: {prn}\n Report type: {rpt_type}\n Date: {rptDate}\n Path: {fnm}")

# recalculate percentages
src = pd.read_excel(fnm)
src["End Market Value"] = src["End Market Value"] / src["End Market Value"].sum()
src["Closing Exposure PA"] = (
    src["Closing Exposure PA"] / src["Closing Exposure PA"].sum()
)

# get investor fund details
fnds = pd.read_excel(pthPy, sheet_name="arc", index_col=None, usecols="AI:AK").dropna(
    subset=["Name"]
)

# fnds = sr[['Alternative Client Name', 'Name', 'Entity ID', 'Value (ZAR)']]
s = "" if len(fnds["Entity ID"]) == 1 else "s"
print(
    f"\n {len(fnds['Entity ID'])} investor fund{s}:\n  {(', ').join(list(fnds['Entity ID']))}\n"
)

print(f"Getting report variables completed: {timediff(start_time, time.time())}", "\n")

# create the client proportional reports

start_time = time.time()
print("Creating client proportional reports ...")

mdd0 = []
eed0 = []
mdd1 = []
eed1 = []
for i in tqdm(range(len(fnds))):
    # create the proportional client holdings
    client = src.copy()
    client["Entity Name"] = f"{fnds.iat[i, 1]} ({prn})"
    client["End Market Value"] = client["End Market Value"] * fnds.iat[i, 2]
    client["End Market Value"] = client["End Market Value"].round(decimals=2)
    client["Closing Exposure PA"] = client["Closing Exposure PA"] * fnds.iat[i, 2]
    client["Closing Exposure PA"] = client["Closing Exposure PA"].round(decimals=2)

    # get rounding differences before eliminating them
    mdd0.append(client["End Market Value"].sum())
    eed0.append(client["Closing Exposure PA"].sum())

    # eliminate rounding differences
    client.at[0, "End Market Value"] = (
        client.at[0, "End Market Value"]
        - client["End Market Value"].sum()
        + fnds.iat[i, 2]
    )
    client.at[0, "Closing Exposure PA"] = (
        client.at[0, "Closing Exposure PA"]
        - client["Closing Exposure PA"].sum()
        + fnds.iat[i, 2]
    )

    # get rounding differences after eliminating them
    mdd1.append(client["End Market Value"].sum())
    eed1.append(client["Closing Exposure PA"].sum())

    # save the client fund in the reporting folder

    mn = f"{fnds.iat[i, 0]} {fnds.iat[i, 1]} ({prn}) {rpt_type} {rptDate.strftime('%d%b%Y')}.xlsx"
    client.to_excel(
        os.path.join(pthReports, mn),
        index=False,
        sheet_name=f"{fnds.iat[i, 1]} {rpt_type} {rptDate.strftime('%d%b%Y')}",
    )

print(f"{timediff(start_time, time.time())} creating client proportional reports")

# open the reporting folder where the files are saved
os.startfile(os.path.realpath(pthReports))

print(os.path.join(pthReports, mn))
print(
    f"\n{timediff(start_time_prpn, time.time())} roundtrip time for {len(fnds['Entity ID'])} investor fund{s}"
)
