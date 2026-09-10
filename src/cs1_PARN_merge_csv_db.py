#!/usr/bin/env python
# coding: utf-8

# # Merge the CS1 fund PARN sheets into one file — DB version
#
# A drop-in replacement for cs1_PARN_merge_csv.py's one Eagle call —
# osprey('fnav', ...) for the fund NAVs — which queried Eagle directly
# with a retry loop. Everything else (merging the PARN batch files,
# which don't come from Eagle themselves, and the Reg 28 prep/workbook
# write) is unchanged. cs1_PARN_merge_csv.py itself is untouched.
#
# Also adds an "issuer_code" column (see hold_cols below) and calls
# r_classifier_db() instead of r_classifier(), which runs issuers_1_db.py
# instead of issuers_1.py — see that file for why.
#
# Reads the PARN batch files produced by cs1_PARN_download_csv_db.py.
# NOTE: this now specifically requires that DB-based downloader's output,
# not the original Eagle-based cs1_PARN_download_csv.py — this script
# pulls an "issuer_code" column out of the batch files (see hold_cols
# below), which only cs1_PARN_download_csv_db.py's holdings query adds;
# the Eagle-sourced PARN export never carried it.

print("\n\n#########################################")
print("#                                       #")
print("#   START cs1_PARN_merge_csv_db.py      #")
print("#                                       #")
print("#########################################\n\n")

import time

start_time = time.time()
start_time_cs1_merge = time.time()

from datetime import datetime
import pandas as pd
from pathlib import Path
import os
from sqlalchemy import create_engine, text
from constants import pthPy, pth_dl, pthTest, db_eagle_uri
from utilities import timediff, prior_month_end, batch_list, r_classifier_db

# get inputs — identical to cs1_PARN_merge_csv.py, not Eagle-related
start_time = time.time()
print("Collecting the report input data ...")

# get fund codes from 'arc' tab of the py_report.xlsm sheet
df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df.iloc[:, 0].apply(str.upper)

# get report date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=2)
k = df1.iloc[1, 0]
rptDate = k.date() if k == k else prior_month_end(datetime.today()).date()

print(
    f'\nPARN CS1 fund report{"" if len(funds) == 1 else "s"} to be merged as at \
{rptDate.strftime("%A %d %b %Y")}\n \
{len(funds)} fund{"" if len(funds) == 1 else "s"}: \n  {(",").join(funds)}'
)
print(f"\n{timediff(start_time, time.time())} collecting the report input data\n")

# ---------------------------------------------------------------------
# Merge the PARN batch files into one dataframe — unchanged, these
# files don't come from Eagle directly (either download script writes
# the same file names)
# ---------------------------------------------------------------------
start_time = time.time()
print(f"Merging the PARN csv files into a dataframe ...")

num_batches = 2 if len(funds) > 1 else 1
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), batch_size))
batch_filepaths = []
for index, batch in enumerate(batches, start=1):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

holdings = pd.DataFrame()
for batch_filepath in batch_filepaths:
    data = pd.read_csv(batch_filepath)
    holdings = pd.concat([holdings, data])

# convert the holdings CS1 fund NAV columns to float64
cols_to_sum = ["Sum of Market Value Income", "Current Exposure"]
for col in cols_to_sum:
    holdings[col] = holdings[col].astype(str).str.replace(",", "").astype("float64")

print(
    f' {timediff(start_time, time.time())} merging the PARN csv files with \
{len(holdings["Entity Name"].unique())} funds into a dataframe\n'
)

# ---------------------------------------------------------------------
# Fund NAVs (FNAV equivalent), from prime_eagle.tmp_eagle_aum — replaces
# the osprey('fnav', ...) retry loop
# ---------------------------------------------------------------------
start_time = time.time()
s1 = "'s" if len(funds) == 1 else "s'"
s2 = "" if len(funds) == 1 else "s"
print(
    f"Getting the {len(funds)} CS1 fund{s1} NAV{s2} as at \
{rptDate.strftime('%A %d %B %Y')} from the database ...\n"
)

name = "CS1"
navs_fln = os.path.join(Path.home(), "Downloads", f'FNAV {name}({len(funds)}) {rptDate.strftime("%d%b%Y")}.csv')
print(f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv expected")

if os.path.isfile(navs_fln):
    print(f"\n  {navs_fln} exists\n")
else:
    try:
        engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 15})
        q = text(
            """
            SELECT
                m.portfolio_name AS `Entity Name`,
                a.datestamp      AS `Effective Date`,
                a.portfolio_code AS `NAV Entity ID`,
                a.star_nav       AS `Total Net Assets`,
                a.status
            FROM tmp_eagle_aum a
            LEFT JOIN s_eagle_portfolio_meta m ON m.portfolio_code = a.portfolio_code
            WHERE a.portfolio_code IN :codes AND a.datestamp = :d
            """
        ).bindparams(codes=tuple(funds), d=rptDate)
        with engine.connect() as conn:
            result = conn.execute(q)
            fnav = pd.DataFrame(result.fetchall(), columns=list(result.keys()))
        engine.dispose()

        not_finalised = fnav.loc[fnav["status"] != "AUDIT", "NAV Entity ID"].tolist()
        if not_finalised:
            print(f"  NOTE: NAV not yet AUDIT-status for: {sorted(not_finalised)}")
        fnav = fnav.drop(columns=["status"])

        missing = set(funds) - set(fnav["NAV Entity ID"])
        if missing:
            print(f"  WARNING: no FNAV data for {rptDate}: {sorted(missing)}")

        fnav.to_csv(navs_fln, index=False)
        print(f"  saved {len(fnav)} rows to {navs_fln}")
    except Exception as e:
        print(e)

# ... then dataframe the fund NAVs — unchanged
navs = pd.read_csv(navs_fln)
navs["Total Net Assets"] = navs["Total Net Assets"].astype(str).str.replace(",", "").astype("float64")

print(
    f" {timediff(start_time, time.time())} getting the {len(funds)} fund\
{'s' if len(funds) != 1 else ''}' NAV{'s' if len(funds) != 1 else ''} as at \
{rptDate.strftime('%A %d %B %Y')} from the database"
)

# ---------------------------------------------------------------------
# Everything below is unchanged from cs1_PARN_merge_csv.py: merge and
# compare holdings vs NAVs, build the Reg 28-ready sheet, write and
# open the workbook, and run r_classifier()
# ---------------------------------------------------------------------
start_time = time.time()
print(
    f'\nMerging and comparing the {len(funds)} CS1 fund holdings and NAVs as at \
{rptDate.strftime("%A %d %B %Y")} ...'
)

holdings_totals = holdings.groupby("Entity ID", as_index=False).sum()[
    ["Entity ID", "Sum of Market Value Income", "Current Exposure"]
]
sums_cf = holdings_totals.merge(navs, how="left", left_on="Entity ID", right_on="NAV Entity ID")
sums_cf.drop(["Entity Name", "NAV Entity ID"], axis=1, inplace=True)
sums_cf["SoMVI-CE"] = sums_cf["Sum of Market Value Income"] - sums_cf["Current Exposure"]
sums_cf["1-CE/SoMVI %"] = (1 - sums_cf["Current Exposure"] / sums_cf["Sum of Market Value Income"]) * 100
sums_cf["SoMVI-NAV"] = abs(sums_cf["Sum of Market Value Income"] - sums_cf["Total Net Assets"])
sums_cf["1-NAV/SoMVI %"] = (1 - sums_cf["Total Net Assets"] / sums_cf["Sum of Market Value Income"]) * 100
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

cols_2dp = ["Sum of Market Value Income", "Current Exposure", "Total Net Assets"]
for col in cols_2dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.2f}")

cols_6dp = ["SoMVI-CE", "1-CE/SoMVI %", "SoMVI-NAV", "1-NAV/SoMVI %"]
for col in cols_6dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.6f}")

print(
    f' {timediff(start_time, time.time())} \
merging and comparing the {len(funds)} \
CS1 fund holdings and NAVs as at {rptDate.strftime("%A %d %B %Y")}'
)

# convert the PARN holdings into Reg 28 format with corresponding headings
start_time = time.time()
print(f"\nConverting the CS1 fund PARN holdings \
in readiness for Reg 28 classification ...")

cs1_fname = os.path.join(pthTest, f'CS1 PARN holdings ({len(funds)}) \
{rptDate.strftime("%d%b%Y")}.xlsx')

hold_cols = [
    "Entity ID",
    "Investment Type",
    "i Issue Name",
    "PrimaryAssetID",
    "issuer_code",
    "CCY",
    "Sum of Market Value Income",
    "% of Total Market Value",
    "Current Exposure",
]
hReg28 = holdings[hold_cols]
hReg28 = hReg28.rename(
    columns={
        "Entity ID": "Entity Name",
        "PrimaryAssetID": "Primary Asset ID",
        "Sum of Market Value Income": "End Market Value",
        "% of Total Market Value": "Percentage of Market Value",
        "Current Exposure": "Closing Exposure PA",
    }
)
hReg28.insert(6, "Reg28 Classification", "")
hReg28.insert(10, f'{rptDate.strftime("%d %b %Y")}', "")
hReg28.iloc[0, 10] = cs1_fname
hReg28.iloc[1, 10] = "CS1"
hReg28.reset_index(drop=True, inplace=True)

print(
    f" {timediff(start_time, time.time())} converting the CS1 fund PARN holdings \
in readiness for Reg 28 classification\n"
)

# write the CS1 holdings dataframe to review it as a worksheet
start_time = time.time()
print(f"\nWriting the CS1 fund holdings dataframe and navs dataframe to a sheet ...")

with pd.ExcelWriter(cs1_fname, engine="xlsxwriter") as writer:
    hReg28.to_excel(writer, index=False, sheet_name="All")
    holdings.to_excel(writer, index=False, sheet_name="PARN")
    sums_cf.to_excel(writer, index=False, sheet_name="NAVs")

print(
    f" {timediff(start_time, time.time())} writing the CS1 fund holdings \
dataframe and navs dataframe to a sheet\n"
)

os.startfile(cs1_fname)

print(f" {timediff(start_time_cs1_merge, time.time())} merging CS1 PARN reports\n")

# run r_classifier_db function from utilities.py — uses issuers_1_db.py,
# not issuers_1.py, since this workbook's extra "issuer_code" column
# shifts every subsequent column by one position
r_classifier_db("cs1", cs1_fname, rptDate)

print("\n\n#########################################")
print("#                                       #")
print("#    END cs1_PARN_merge_csv_db.py       #")
print("#                                       #")
print("#########################################\n\n")
