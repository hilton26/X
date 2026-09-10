#!/usr/bin/env python
# coding: utf-8

# # Downloading PGF UCITS Hedge Share Class NAVs and Holdings — DB version
#
# A drop-in replacement for pgf_downloading.py's Eagle/Selenium scrape.
# Instead of logging into Eagle, this queries the read-only PIM database
# mirror (prime_eagle) directly for the same two reports, and writes them
# to the exact same file names/columns pgf_compiling.py already expects.
# pgf_downloading.py and pgf_compiling.py are unchanged — this script is
# meant to be run instead of pgf_downloading.py, before pgf_compiling.py.
#
# See src/docs/database-access/07_DATABASE_SETUP.md and the "What is in
# the databases" section there for background on prime_eagle.

print("\n\n##########################################")
print("#                                        #")
print("#  START 1/2 pgf_downloading_db.py X     #")
print("#                                        #")
print("##########################################\n\n")

import time

start_time_pgf_dl = time.time()
start_time = time.time()

# libraries, libraries!
from datetime import datetime
import pandas as pd
import os
from sqlalchemy import create_engine, text
from constants import pthPy, pth_dl, db_eagle_uri
from utilities import timediff, prior_working_day

print("Getting input variables ...")
# get report inputs (same arc-sheet columns G:I as pgf_downloading.py)
df = pd.read_excel(pthPy, sheet_name="arc", usecols="G:I")
rptDate = (
    df.iloc[0, 2].date()
    if isinstance(df.iloc[0, 2], datetime) and not pd.isna(df.iloc[0, 2])
    else prior_working_day(datetime.today())
)  # prior working day or report date override; has type datetime()

# get portfolio lists
navs_codes = df["pgf: UT prices"].dropna().tolist()  # for unit trust prices
hldgs_codes = df["pgf: PAR-N"].dropna().tolist()  # for class holdings

# derive names of the files to be written — identical convention to
# pgf_downloading.py / pgf_compiling.py, so pgf_compiling.py can read
# these files without any changes
utP_fln = f"UTPS PGF_UT_Prices({len(navs_codes)}) {rptDate.strftime('%d%b%Y')}.csv"
parN_fln = f"PARN PGF_Holdings({len(hldgs_codes)}) {rptDate.strftime('%d%b%Y')}.csv"
utP_nm = os.path.join(pth_dl, utP_fln)
parN_nm = os.path.join(pth_dl, parN_fln)

print(
    f" Report date:          {rptDate}",
    "\n",
    f"Class NAV codes:      {(',').join(navs_codes)}",
    "\n",
    f"Class holdings codes: {(',').join(hldgs_codes)}\n",
)
print(f"Expected output files:\n  {parN_fln}\n  {utP_fln}")
print(f"{timediff(start_time, time.time())} getting input variables\n")

# ---------------------------------------------------------------------
# Unit trust prices (UTPS equivalent), from prime_eagle.tmp_eagle_nav
# ---------------------------------------------------------------------
start_time = time.time()
print("Getting the unit trust prices report from the database ...")

if os.path.isfile(utP_nm):
    print(f"  {utP_nm} already exists")
else:
    try:
        engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 10})
        query = text(
            """
            SELECT
                entity_name  AS `Entity Name`,
                datestamp    AS `Effective Date`,
                entity_id    AS `NAV Entity ID`,
                jse_code     AS `JSE Code`,
                nav_price    AS `NAV Price`,
                clean_price  AS `Clean Price`,
                income_price AS `Income Price`,
                class_size   AS `Class Size`,
                units_in_class AS `Units in Class`,
                daily_flows  AS `Daily Flows`
            FROM tmp_eagle_nav
            WHERE entity_id IN :codes AND datestamp = :rptDate
            """
        ).bindparams(codes=tuple(navs_codes), rptDate=rptDate)
        with engine.connect() as conn:
            result = conn.execute(query)
            navs = pd.DataFrame(result.fetchall(), columns=list(result.keys()))
        engine.dispose()

        # three columns Eagle's UTPS export carries that tmp_eagle_nav
        # does not — fill with 0.0 to match the real export's behaviour
        for col in ["Distribution Rate", "Mgmt Fee", "Performance Fee"]:
            navs[col] = 0.0
        navs["Effective Date"] = navs["Effective Date"].apply(
            lambda d: d.strftime("%#m/%#d/%Y")
        )
        # enforce the exact column order of the real UTPS export
        navs = navs.reindex(
            columns=[
                "Entity Name",
                "Effective Date",
                "NAV Entity ID",
                "JSE Code",
                "NAV Price",
                "Clean Price",
                "Income Price",
                "Class Size",
                "Units in Class",
                "Daily Flows",
                "Distribution Rate",
                "Mgmt Fee",
                "Performance Fee",
            ]
        )

        found = set(navs["NAV Entity ID"])
        missing = set(navs_codes) - found
        if missing:
            print(f"  WARNING: no unit trust price data for {rptDate}: {sorted(missing)}")

        navs.to_csv(utP_nm, index=False)
        print(f"  saved {len(navs)} rows to {utP_nm}")
    except Exception as e:
        print(e)

print(
    f"{timediff(start_time, time.time())} getting the unit trust prices report\n"
)

# ---------------------------------------------------------------------
# Holdings (PARN equivalent), from prime_eagle.tmp_eagle_holdings_analytics
# with a fallback to the simpler tmp_eagle_holdings for any class not
# carried in the analytics table (e.g. classes with no bond/derivative
# analytics to report)
# ---------------------------------------------------------------------
start_time = time.time()
print("Getting the class holdings report from the database ...")

if os.path.isfile(parN_nm):
    print(f"  {parN_nm} already exists")
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
        # final column order must match the real PARN export exactly,
        # since pgf_compiling.py pastes this straight into a template
        parn_column_order = list(analytics_cols.values())

        analytics_select = ", ".join(
            f"`{src}` AS `{dst}`" for src, dst in analytics_cols.items()
        )
        analytics_query = text(
            f"""
            SELECT {analytics_select}
            FROM tmp_eagle_holdings_analytics
            WHERE portfolio_code IN :codes AND datestamp = :rptDate
            """
        ).bindparams(codes=tuple(hldgs_codes), rptDate=rptDate)

        # the simpler table only carries these — everything else stays NaN
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
            fallback_codes = set(hldgs_codes) - found_codes
            fallback = pd.DataFrame()
            if fallback_codes:
                fallback_query = text(
                    f"""
                    SELECT {holdings_select}
                    FROM tmp_eagle_holdings
                    WHERE portfolio_code IN :codes AND datestamp = :rptDate
                    """
                ).bindparams(codes=tuple(fallback_codes), rptDate=rptDate)
                fb_result = conn.execute(fallback_query)
                fallback = pd.DataFrame(
                    fb_result.fetchall(), columns=list(fb_result.keys())
                )
        engine.dispose()

        hldgs = pd.concat([hldgs, fallback], ignore_index=True)
        hldgs = hldgs.reindex(columns=parn_column_order)
        for col in ["Next Coupon Date", "Maturity Date", "i Position Effective Date"]:
            hldgs[col] = hldgs[col].apply(
                lambda d: d.strftime("%#m/%#d/%Y") if pd.notna(d) else d
            )

        if not fallback.empty:
            print(
                "  NOTE: these classes only have the reduced column set "
                "(no valuation/analytics columns available in prime_eagle): "
                f"{sorted(fallback_codes)}"
            )
            # Known cosmetic difference from the real Eagle PARN export:
            # a class with genuinely no holdings gets one blank "no data"
            # placeholder row from Eagle, but shows as several zero-value
            # cash-sweep rows here (from tmp_eagle_holdings directly).
            # Functionally equivalent (all values are zero either way) —
            # confirmed against real PGF PARN samples for PGBEUR, PGBGBP,
            # PGPRGBP, and PGPRUS0% on 2026-08-18.
        still_missing = set(hldgs_codes) - set(hldgs["Entity ID"].dropna())
        if still_missing:
            print(
                f"  WARNING: no holdings data at all for {rptDate}: "
                f"{sorted(still_missing)}"
            )

        hldgs.to_csv(parN_nm, index=False)
        print(f"  saved {len(hldgs)} rows to {parN_nm}")
    except Exception as e:
        print(e)

print(f"{timediff(start_time, time.time())} getting the class holdings report\n")
print(
    f"{timediff(start_time_pgf_dl, time.time())} roundtrip time to get "
    "unit trust prices and holdings reports from the database\n"
)

print(f"{rptDate.strftime('%a %d %B %Y')} downloads")
print(f"  {parN_nm} which {'exists' if os.path.exists(parN_nm) else 'does not exist'}")
print(f"  {utP_nm} which {'exists' if os.path.exists(utP_nm) else 'does not exist'}")

print("\n\n##########################################")
print("#                                        #")
print("#   END 1/2 pgf_downloading_db.py X      #")
print("#                                        #")
print("##########################################\n\n")
