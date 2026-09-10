#!/usr/bin/env python
# coding: utf-8

# # Download the CS1 fund PARN holdings — DB version
#
# A drop-in replacement for cs1_PARN_download_csv.py's Eagle/Selenium
# scrape (osprey('parn', ...)). Queries prime_eagle directly instead,
# and writes the exact same batch file names/columns that
# cs1_PARN_merge_csv.py already expects, so that script — unchanged —
# can read them without knowing the difference. cs1_PARN_download_csv.py
# itself is untouched.
#
# Same known gap as the other _db PARN scripts this session: a fund
# whose holdings only exist in the simpler tmp_eagle_holdings table
# (no bond/valuation analytics columns) gets those columns as blank
# rather than an error — flagged in the console output when it happens.

print("\n\n#############################################")
print("#                                           #")
print("#   START cs1_PARN_download_csv_db.py       #")
print("#                                           #")
print("#############################################\n\n")

import time

start_time = time.time()
start_time_cs1_PARN_download = start_time

print("Importing libraries ...")
from datetime import datetime
import pandas as pd
import os
from sqlalchemy import create_engine, text
from constants import pthPy, pth_dl, db_eagle_uri
from utilities import timediff, prior_month_end, batch_list

print(f" {timediff(start_time, time.time())} importing libraries\n")

# get inputs — identical to cs1_PARN_download_csv.py, not Eagle-related
start_time = time.time()
print("Collecting input data ...")

# get fund codes from 'arc' tab of the py_report.xlsm sheet
df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df.iloc[:, 0].apply(str.upper)

# get report date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
k = df1.iloc[1, 0]
rptDate = k.date() if k == k else prior_month_end(datetime.today()).date()

print(
    "\n",
    f'PARN CS1 fund report{"" if len(funds) == 1 else "s"} to be pulled from the \
database as at {rptDate.strftime("%A %d %b %Y")}',
    "\n",
    f"{len(funds)} fund{'' if len(funds) == 1 else 's'}: ",
    "\n",
    f"{(', ').join(funds)}",
)
print(f"\n{timediff(start_time, time.time())} collecting input data\n")

# ---------------------------------------------------------------------
# Holdings (PARN equivalent), from prime_eagle.tmp_eagle_holdings_analytics
# with a fallback to the simpler tmp_eagle_holdings for any fund not
# carried in the analytics table. One DB round trip for all funds, then
# split and saved into the SAME batch files the original script would
# have produced, so cs1_PARN_merge_csv.py needs no changes.
# ---------------------------------------------------------------------
start_time = time.time()
print("Getting the holdings data from the database ...")

analytics_cols = {
    "portfolio_name": "Entity Name",
    "valuation_first_level": "Valuation First Level",
    "valuation_second_level": "Valuation Second Level",
    "instrument_name": "i Issue Name",
    "instrument_code": "PrimaryAssetID",
    "issue_description": "Issue Description",
    "original_nominal": "Original Nominal",
    "clean_book_value": "Clean Book Value",
    "clean_market_value": "Clean Market Value",
    "accrued_income": "Accrued Income",
    "dividend_receivable": "Dividend Receivable",
    "sum_of_market_value_income": "Sum of Market Value Income",
    "market_price_/yield": "Market Price /Yield",
    "%_of_total_market_value": r"% of Total Market Value",
    "next_coupon_date": "Next Coupon Date",
    "maturity_date": "Maturity Date",
    "coupon": "Coupon",
    "duration": "Duration",
    "modified_duration": "Modified Duration",
    "investment_type": "Investment Type",
    "security_type": "Security Type",
    "sub_security_type": "Sub Security Type",
    "currency": "CCY",
    "sector": "Sector",
    "supersector": "Supersector",
    "industry": "Industry",
    "issuer_rating_lt": "Issuer Rating",
    "security_rating": "Security Rating",
    "naca_yield": "NACA Yield",
    "nacm_yield": "NACM Yield",
    "weighted_avg_naca_yield": "Weighted Avg NACA Yield",
    "weighted_avg_nacm_yield": "Weighted Avg NACM Yield",
    "market_value_%": "Market Value %",
    "current_exposure": "Current Exposure",
    "current_exposure_%": "Current Exposure %",
    "datestamp": "i Position Effective Date",
    "portfolio_code": "Entity ID",
    "weighted_average_nacs_yield": "Weighted Average NACS Yield",
    "weighted_average_coupon": "Weighted Average Coupon",
    "weighted_modified_duration": "Weighted Modified Duration",
    "issuer": "Issuer",
    "issuer_code": "issuer_code",
    "isin_code": "ISIN Code",
    "ticker": "Ticker",
}
parn_column_order = list(analytics_cols.values())
analytics_select = ", ".join(f"`{src}` AS `{dst}`" for src, dst in analytics_cols.items())

holdings_cols = {
    "portfolio_name": "Entity Name",
    "instrument_name": "i Issue Name",
    "instrument_code": "PrimaryAssetID",
    "investment_type": "Investment Type",
    "security_type": "Security Type",
    "currency": "CCY",
    "datestamp": "i Position Effective Date",
    "portfolio_code": "Entity ID",
    "ticker": "Ticker",
}
holdings_select = ", ".join(f"`{src}` AS `{dst}`" for src, dst in holdings_cols.items())

wbH_db = pd.DataFrame(columns=parn_column_order)
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

    wbH_db = pd.concat([hldgs, fallback], ignore_index=True)
    wbH_db = wbH_db.reindex(columns=parn_column_order)
    for col in ["Next Coupon Date", "Maturity Date", "i Position Effective Date"]:
        wbH_db[col] = wbH_db[col].apply(
            lambda d: d.strftime("%#m/%#d/%Y") if pd.notna(d) else d
        )

    if not fallback.empty:
        print(
            "  NOTE: these funds only have the reduced column set "
            "(no valuation/analytics columns available in prime_eagle): "
            f"{sorted(fallback_codes)}"
        )
    still_missing = set(funds) - set(wbH_db["Entity ID"].dropna())
    if still_missing:
        print(f"  WARNING: no holdings data at all for {rptDate}: {sorted(still_missing)}")
except Exception as e:
    print(e)

print(f"{timediff(start_time, time.time())} getting the holdings data from the database\n")

# ---------------------------------------------------------------------
# Split into the same batches (and file names) the original script
# would have produced, so cs1_PARN_merge_csv.py needs no changes
# ---------------------------------------------------------------------
start_time = time.time()
print("Saving the holdings data into the expected batch files ...")

num_batches = 2 if len(funds) > 1 else 1
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size=min(len(funds), batch_size))
batch_filepaths = []
for index, batch in enumerate(batches, start=1):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    s = "" if len(batch) == 1 else "s"
    print(f"{filename}, a batch of {len(batch)} file{s}:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        continue

    batch_df = wbH_db[wbH_db["Entity ID"].isin(batch)]
    batch_df.to_csv(batch_filepath, index=False)
    print(f"  saved {len(batch_df)} rows to {batch_filepath}")

print(
    f"{timediff(start_time, time.time())} saving the holdings data into the \
expected batch files\n"
)

print(
    f"\n {timediff(start_time_cs1_PARN_download, time.time())} \
downloading CS1 PARN holdings from the database\n",
)

print("\n\n#############################################")
print("#                                           #")
print("#    END cs1_PARN_download_csv_db.py        #")
print("#                                           #")
print("#############################################\n\n")
