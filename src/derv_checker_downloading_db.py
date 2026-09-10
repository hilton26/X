#!/usr/bin/env python
# coding: utf-8

# # Pull Holdings and Derivative Data — DB version
#
# A drop-in replacement for derv_checker_downloading.py's Eagle/Selenium
# scrape (PARN holdings, DERV derivatives, and the prior day's FNAV used
# for the "Change in NAV" column). Queries prime_eagle directly instead,
# and writes the exact same file names/columns that derv_checker_table.py
# and the rest of the pipeline already expect, so nothing downstream
# needs to change. derv_checker_downloading.py itself is untouched.
#
# NOT replaced here: the daily bank-recon step (reads Outlook, unrelated
# to Eagle) still runs via overdrafts.py, unchanged.
#
# Known gaps (see prints below when they're hit):
# - Market Value is written as 0.0 for all derivatives. That matches the
#   real report for ordinary listed futures/options, but a real report
#   can show a nonzero Market Value for CFDs and some foreign-listed
#   options (confirmed against DERV (185) 20Aug2026.csv) — prime_eagle's
#   tmp_eagle_derivatives table doesn't carry that value, so this DB
#   version cannot reproduce it yet.
# - Delta is stored as 0.0 for futures in tmp_eagle_derivatives (Eagle's
#   own report shows 1.0 for futures by convention) — corrected below.
# - The >300-fund half1/half2 batching in the original script doesn't
#   apply here — a single SQL query handles the whole fund list at once.

print("\n\n##################################################")
print("#                                                #")
print("#  START 1/4 derv_checker_downloading_db.py X    #")
print("#                                                #")
print("##################################################\n\n")

import time

start_time = time.time()
start_time_derivative_downloading = time.time()
print("\n\nImporting libraries ...\n")

import pandas as pd
import os, sys, subprocess
from datetime import datetime
from sqlalchemy import create_engine, text
from constants import pthEXPORTS, pth_dl, pthOverdrafts, db_eagle_uri
from utilities import timediff, parn_de, prior_working_day

# get report date and selected summary sheet option
fPARN, fDE, funds, rptDate, summ_yn, dervthreshold, batches = parn_de()
funds = list(funds)

print(f"{timediff(start_time, time.time())} importing libraries\n")

# derive file names — identical convention to derv_checker_downloading.py
start_time = time.time()
print("Deriving file names ...")

filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
fPARN = os.path.join(pth_dl, f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv")
fDE = os.path.join(pth_dl, f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv")
bank_file_sa = pthOverdrafts + rf"\{rptDate.strftime('%Y%m%d')}_overdrafts_sa.xlsx"

print(
    f"\n {rptDate.strftime('%A %d %b %Y')} \
for {len(funds)} funds:\n {(',').join(funds)}\n",
)

# ---------------------------------------------------------------------
# Prior day's FNAV (Total Net Assets), from prime_eagle.tmp_eagle_aum
# ---------------------------------------------------------------------
ystdy_date = prior_working_day(rptDate).strftime("%Y%m%d")
ystdy_path = pthEXPORTS + rf"\{ystdy_date}_derv_calc.xlsx"
y_date = prior_working_day(rptDate)
ystdy_dwnl = rf"FNAV ({len(funds)}) {y_date.strftime('%d%b%Y')}.csv"

if os.path.isfile(ystdy_path) and os.path.getsize(ystdy_path) > 10:
    print(f"\n {ystdy_date}_derv_calc.xlsx NAVs will be added to today's file\n")
elif os.path.isfile(os.path.join(pth_dl, ystdy_dwnl)):
    print(f" {ystdy_dwnl} already exists")
else:
    print(f"\n {ystdy_date} NAVs will be pulled from the database")
    try:
        engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 10})
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
        ).bindparams(codes=tuple(funds), d=y_date.date())
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
            print(f"  WARNING: no FNAV data for {y_date.date()}: {sorted(missing)}")

        fnav.to_csv(os.path.join(pth_dl, ystdy_dwnl), index=False)
        print(f"  saved {len(fnav)} rows to {ystdy_dwnl}")
    except Exception as e:
        print(e)

# ---------------------------------------------------------------------
# Derivatives (DERV equivalent), from prime_eagle.tmp_eagle_derivatives
# ---------------------------------------------------------------------
start_time = time.time()
print("\nDownloading and then saving derivative data ...")

if os.path.isfile(fDE):
    print(f"  {os.path.basename(fDE)} already exists")
else:
    try:
        engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 10})

        # Entity Name for DERV must come from the SAME source as the
        # holdings (PARN) section below, NOT s_eagle_portfolio_meta —
        # confirmed the two can disagree (e.g. PIMBAL: "Prescient
        # Balanced Fund" in s_eagle_portfolio_meta vs "PIM Balanced" in
        # the holdings tables). derv_checker_table.py merges holdings
        # and derivatives on Entity Name + PrimaryAssetID, so a mismatch
        # here would silently drop every derivative for the fund.
        name_q = text(
            """
            SELECT portfolio_code, portfolio_name FROM tmp_eagle_holdings_analytics
            WHERE portfolio_code IN :codes AND datestamp = :d
            UNION
            SELECT portfolio_code, portfolio_name FROM tmp_eagle_holdings
            WHERE portfolio_code IN :codes AND datestamp = :d
            """
        ).bindparams(codes=tuple(funds), d=rptDate)
        with engine.connect() as conn:
            fund_names = dict(conn.execute(name_q).fetchall())

        q = text(
            """
            SELECT
                d.portfolio_code        AS `Entity ID`,
                d.instrument_description AS `i Issue Name`,
                d.instrument_code       AS `Primary Asset ID`,
                d.nominal               AS `Nominal Holding`,
                d.delta                 AS `Delta`,
                d.effective_exposure    AS `Effective Exposure`,
                d.put_call
            FROM tmp_eagle_derivatives d
            WHERE d.portfolio_code IN :codes AND d.datestamp = :d
            """
        ).bindparams(codes=tuple(funds), d=rptDate)
        with engine.connect() as conn:
            result = conn.execute(q)
            derv = pd.DataFrame(result.fetchall(), columns=list(result.keys()))
        engine.dispose()

        # Eagle's own DERV export shows Delta = 1.0 for futures by
        # convention; tmp_eagle_derivatives stores 0.0 for them
        derv.loc[derv["put_call"] == "FUTURE", "Delta"] = 1.0
        derv = derv.drop(columns=["put_call"])
        derv["Entity Name"] = derv["Entity ID"].map(fund_names)
        no_name = derv.loc[derv["Entity Name"].isna(), "Entity ID"].unique()
        if len(no_name):
            print(f"  NOTE: no matching holdings-table name for: {sorted(no_name)}")
        derv = derv.drop(columns=["Entity ID"])
        # see the module docstring — Market Value isn't available from
        # this table (only matters for CFDs / some foreign options)
        derv["Market Value"] = 0.0
        derv = derv[
            [
                "Entity Name",
                "i Issue Name",
                "Primary Asset ID",
                "Nominal Holding",
                "Delta",
                "Market Value",
                "Effective Exposure",
            ]
        ]

        derv.to_csv(fDE, index=False)
        print(f"  saved {len(derv)} rows to {os.path.basename(fDE)}")
    except Exception as e:
        print(e)

print(
    f" {timediff(start_time, time.time())} downloading and then \
saving derivative data\n"
)

# ---------------------------------------------------------------------
# Holdings (PARN equivalent), from prime_eagle.tmp_eagle_holdings_analytics
# with a fallback to the simpler tmp_eagle_holdings for any fund not
# carried in the analytics table
# ---------------------------------------------------------------------
start_time = time.time()
print("Downloading and then saving holdings data ...")

if os.path.isfile(fPARN):
    print(f"  {os.path.basename(fPARN)} already exists")
else:
    try:
        engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 10})

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
            "%_of_total_market_value": "% of Total Market Value",
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
            "isin_code": "ISIN Code",
            "ticker": "Ticker",
        }
        parn_column_order = list(analytics_cols.values())
        analytics_select = ", ".join(
            f"`{src}` AS `{dst}`" for src, dst in analytics_cols.items()
        )
        analytics_query = text(
            f"""
            SELECT {analytics_select}
            FROM tmp_eagle_holdings_analytics
            WHERE portfolio_code IN :codes AND datestamp = :d
            """
        ).bindparams(codes=tuple(funds), d=rptDate)

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
        holdings_select = ", ".join(
            f"`{src}` AS `{dst}`" for src, dst in holdings_cols.items()
        )

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
                fallback = pd.DataFrame(
                    fb_result.fetchall(), columns=list(fb_result.keys())
                )
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
            print(
                f"  WARNING: no holdings data at all for {rptDate}: {sorted(still_missing)}"
            )

        wbH_db.to_csv(fPARN, index=False)
        print(f"  saved {len(wbH_db)} rows to {os.path.basename(fPARN)}")
    except Exception as e:
        print(e)

print(
    f"{timediff(start_time, time.time())} downloading and then \
saving holdings data\n"
)

print(
    f"\n {timediff(start_time_derivative_downloading, time.time())} \
downloading holdings, derivative, and NAV data from the database. \
Next step is compiling.\n",
)

# save daily bank reconciliation files — unchanged, unrelated to Eagle/DB
start_time = time.time()
print("\n\nSaving daily bank recons from Outlook")
bank = os.path.join(os.path.dirname(__file__), "overdrafts.py")
subprocess.run([sys.executable, bank])
print(f" {timediff(start_time, time.time())} saving daily bank recons from Outlook\n\n")

# test that holdings (fPARN) and derivatives (fDE) downloaded correctly
test_fPARN = False
if os.path.exists(fPARN):
    df_fPARN = pd.read_csv(fPARN)
    test_fPARN = df_fPARN.columns[1] == "Valuation First Level"

test_fDE = False
if os.path.exists(fDE):
    df_fDE = pd.read_csv(fDE)
    test_fDE = df_fDE.columns[6] == "Effective Exposure"

print(f"\n Expected downloads for {rptDate.strftime('%A %d %B %Y')}:")
print(
    f"  {os.path.basename(fPARN)} which {'exists' if test_fPARN else 'does not exist'}"
)
print(f"  {os.path.basename(fDE)} which {'exists' if test_fDE else 'does not exist'}")
print(
    f"  {os.path.basename(bank_file_sa)} which \
{'exists' if os.path.exists(bank_file_sa) else 'does not exist'}"
)

print("\n\n##################################################")
print("#                                                #")
print("#   END 1/4 derv_checker_downloading_db.py X     #")
print("#                                                #")
print("##################################################\n\n")
