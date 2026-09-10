#!/usr/bin/env python
# coding: utf-8

# # Download the lookthroughs for Reg28 and Reg30 classification — DB version
#
# A drop-in replacement for lt_dl.py's Eagle/Selenium scrape (osprey('r28i', ...)).
# Queries prime_eagle directly instead, and writes the exact same batch file
# names/columns that lt_merge.py already expects, so that script — unchanged —
# can read them without knowing the difference. lt_dl.py itself is untouched.
#
# *** UNVERIFIED COLUMN MAPPING ***
# DB access to prime_eagle was unavailable while writing this script (Netskope
# wasn't routing pim-cpt-mysql-prod for this device at the time), so the four
# R28I-specific source columns below — reg28_classification, end_market_value,
# percentage_of_market_value, closing_exposure_pa — are best-effort guesses
# following the snake_case naming convention seen in
# cs1_PARN_download_csv_db.py's analytics_cols mapping (e.g. "Sum of Market
# Value Income" -> sum_of_market_value_income), NOT confirmed against the live
# schema. Run `SHOW COLUMNS FROM tmp_eagle_holdings_analytics` and correct
# analytics_cols below before trusting this script's output for an actual
# Reg 28 classification run.
#
# Same known gap as the other _db scripts this session: a fund whose holdings
# only exist in the simpler tmp_eagle_holdings table (no Reg28/valuation
# analytics columns) gets those columns as blank rather than an error —
# flagged in the console output when it happens.

print("\n\n#####################################")
print("#                                   #")
print("#         START lt_dl_db.py         #")
print("#                                   #")
print("#####################################\n\n")

import time

start_time = time.time()
start_time_lt_dl = time.time()

print("Importing libraries ...")
from datetime import datetime
import pandas as pd
import os
from sqlalchemy import create_engine, text
from constants import pthPy, pth_dl, db_eagle_uri
from utilities import timediff, prior_month_end, batch_list

print(f" {timediff(start_time, time.time())} importing libraries\n")

# get inputs — identical to lt_dl.py, not Eagle-related
start_time = time.time()
print("Collecting input data ...")

# get fund codes from 'arc' tab of the py_report.xlsm sheet
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df1.iloc[:, 0].str.upper()

# get report parameters date
df = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
k = df.iloc[1, 0]
rptDate = (
    k if isinstance(k, datetime) else prior_month_end(datetime.today().date())
)  # prior month end or report date override; type is datetime()
num_batches = df.iloc[2, 0]

s = "" if len(funds) == 1 else "s"
print(
    f"{len(funds)} fund lookthrough{s} as at {rptDate.strftime('%A %d %b %Y')} \
to be pulled from the database:\n  {(', ').join(funds)}"
)
print(f"\n{timediff(start_time, time.time())} collecting input data\n")

# ---------------------------------------------------------------------
# Holdings (R28I equivalent), from prime_eagle.tmp_eagle_holdings_analytics
# with a fallback to the simpler tmp_eagle_holdings for any fund not
# carried in the analytics table. One DB round trip for all funds, then
# split and saved into the SAME batch files lt_dl.py would have produced,
# so lt_merge.py needs no changes.
# ---------------------------------------------------------------------
start_time = time.time()
print("Getting the lookthrough holdings data from the database ...")

analytics_cols = {
    "portfolio_name": "Entity Name",
    "investment_type": "Investment Type",
    "instrument_name": "i Issue Name",
    "instrument_code": "Primary Asset ID",
    "currency": "CCY",
    "reg28_classification": "Reg28 Classification",  # UNVERIFIED
    "end_market_value": "End Market Value",  # UNVERIFIED
    "percentage_of_market_value": "Percentage of Market Value",  # UNVERIFIED
    "closing_exposure_pa": "Closing Exposure PA",  # UNVERIFIED
    "datestamp": "i Position Effective Date",
    "portfolio_code": "Entity ID",
}
r28i_column_order = list(analytics_cols.values())
analytics_select = ", ".join(f"`{src}` AS `{dst}`" for src, dst in analytics_cols.items())

# the simpler table has no Reg28 classification or exposure columns at all
holdings_cols = {
    "portfolio_name": "Entity Name",
    "investment_type": "Investment Type",
    "instrument_name": "i Issue Name",
    "instrument_code": "Primary Asset ID",
    "currency": "CCY",
    "datestamp": "i Position Effective Date",
    "portfolio_code": "Entity ID",
}
holdings_select = ", ".join(f"`{src}` AS `{dst}`" for src, dst in holdings_cols.items())

wbLT_db = pd.DataFrame(columns=r28i_column_order)
try:
    engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 10})
    analytics_query = text(
        f"""
        SELECT {analytics_select}
        FROM tmp_eagle_holdings_analytics
        WHERE portfolio_code IN :codes AND datestamp = :d
        """
    ).bindparams(codes=tuple(funds), d=rptDate)

    with engine.connect() as conn:
        result = conn.execute(analytics_query)
        hldgs = pd.DataFrame(result.fetchall(), columns=list(result.keys()))

        found_codes = set(hldgs["Entity ID"]) if not hldgs.empty else set()
        fallback_codes = set(funds) - found_codes
        fallback = pd.DataFrame()
        if fallback_codes:
            fallback_query = text(
                f"""
                SELECT {holdings_select}
                FROM tmp_eagle_holdings
                WHERE portfolio_code IN :codes AND datestamp = :d
                """
            ).bindparams(codes=tuple(fallback_codes), d=rptDate)
            fb_result = conn.execute(fallback_query)
            fallback = pd.DataFrame(fb_result.fetchall(), columns=list(fb_result.keys()))
    engine.dispose()

    wbLT_db = pd.concat([hldgs, fallback], ignore_index=True)
    wbLT_db = wbLT_db.reindex(columns=r28i_column_order)

    if not fallback.empty:
        print(
            "  NOTE: these funds only have the reduced column set "
            "(no Reg28 classification or exposure columns available in prime_eagle): "
            f"{sorted(fallback_codes)}"
        )
    still_missing = set(funds) - set(wbLT_db["Entity ID"].dropna())
    if still_missing:
        print(f"  WARNING: no holdings data at all for {rptDate}: {sorted(still_missing)}")
except Exception as e:
    print(e)

print(
    f"{timediff(start_time, time.time())} getting the lookthrough \
holdings data from the database\n"
)

# ---------------------------------------------------------------------
# Split into the same batches (and file names) lt_dl.py would have
# produced, so lt_merge.py needs no changes
# ---------------------------------------------------------------------
start_time = time.time()
print("Saving the holdings data into the expected batch files ...")

batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), max(batch_size, 1)))
batch_filepaths = []
for index, batch in enumerate(batches, start=1):
    fln = f"{index}_of_{len(batches)}"
    filename = f"R28I {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    s = "" if len(batch) == 1 else "s"
    print(f"{filename}, a batch of {len(batch)} file{s}:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        continue

    batch_df = wbLT_db[wbLT_db["Entity ID"].isin(batch)]
    batch_df.to_csv(batch_filepath, index=False)
    print(f"  saved {len(batch_df)} rows to {batch_filepath}")

print(
    f"{timediff(start_time, time.time())} saving the holdings data into the \
expected batch files\n"
)

print(
    f"\n {timediff(start_time_lt_dl, time.time())} \
pulling the lookthrough holdings from the database\n",
)

print("\n\n#####################################")
print("#                                   #")
print("#          END lt_dl_db.py          #")
print("#                                   #")
print("#####################################\n\n")
