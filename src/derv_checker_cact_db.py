#!/usr/bin/env python
# coding: utf-8

# # Cash Activity — DB "best-effort" version — NOT equivalent to real CACT
#
# derv_checker_cact.py downloads Eagle's CACT report: a fund's own cash
# LEDGER, categorized by transaction type (WITHDRAWAL, CONTRIBUTION,
# ADDSUBSCRIPTION, ADDREDEMPTION, DISTRIBUTION, CASHXFER, CASHDIV,
# MISCINC, MARGIN, COUPON, BUY, SELL, ...), with security-level detail
# (Primary Asset ID, ISIN, custody bank/account, FX rate, etc.).
#
# prime_eagle has NO table with that structure. Checked (this session):
# pit_cash_recon_pim, tmp_eagle_cash_projection, ts_eagle_flows,
# tmp_eagle_statement_m, s_creations_lookup, tmp_eagle_bank_details —
# none carry a fund-level categorized cash ledger.
#
# This script is a best-effort SUBSTITUTE using the closest available
# table, ts_eagle_flows — but that table is INVESTOR-LEVEL creation/
# liquidation activity (e.g. "Mr X bought units in Fund Y"), not the
# fund's own cash ledger. It's conceptually related to CACT's external
# subscription/redemption flows (BUY/SELL/SWITCH IN/SWITCH OUT there is
# analogous to CONTRIBUTION/WITHDRAWAL/ADDSUBSCRIPTION/ADDREDEMPTION in
# real CACT), but it is NOT the same data, and carries none of CACT's
# security-level trade/margin/coupon activity at all.
#
# Known gaps, deliberately not papered over:
# - ts_eagle_flows has one "datestamp", not separate Settlement/
#   Accounting Date columns — both output columns get the same value.
# - No currency column — Local Amount is just copied from Base Amount,
#   assumed ZAR. Not verified.
# - No Primary Asset ID / ISIN / custody / FX rate columns at all —
#   left blank.
# - Fund matching is by exact (case/whitespace-insensitive) name match
#   against s_eagle_portfolio_meta — any fund whose name doesn't match
#   exactly is dropped and reported as unmatched, not guessed at.
#
# Output is saved with a distinct "CACT_DB_PARTIAL" filename — deliberately
# NOT the same naming convention as the real CACT files, so it can never
# be silently picked up by derv_checker_table.py or any other script
# expecting real CACT data. This is for manual review only.

print("\n\n#####################################################")
print("#                                                   #")
print("#   START derv_checker_cact_db.py (PARTIAL, review-only)   #")
print("#                                                   #")
print("#####################################################\n\n")

import time

start_time = time.time()

import pandas as pd
import os
from sqlalchemy import create_engine, text
from constants import pth_dl, db_eagle_uri
from utilities import timediff, parn_de

print("Getting the reporting date and fund list ...")
fPARN, fDE, funds, rptDate, summ_yn, dervthreshold, batches = parn_de()
funds = list(funds)
print(f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds\n")

print("Querying ts_eagle_flows and matching fund names to codes ...")
engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 15})
q = text(
    """
    SELECT
        f.datestamp,
        m.portfolio_code,
        f.fund,
        f.transaction,
        f.type,
        f.product,
        f.requestedvalue,
        f.requestedunits
    FROM ts_eagle_flows f
    LEFT JOIN s_eagle_portfolio_meta m
        ON UPPER(TRIM(m.portfolio_name)) = UPPER(TRIM(f.fund))
    WHERE f.datestamp = :d
    """
).bindparams(d=rptDate)
with engine.connect() as conn:
    result = conn.execute(q)
    raw = pd.DataFrame(result.fetchall(), columns=list(result.keys()))
engine.dispose()

matched = raw[raw["portfolio_code"].isin(funds)].copy()
unmatched_funds = sorted(
    raw.loc[raw["portfolio_code"].isna(), "fund"].unique().tolist()
)
not_in_our_list = sorted(
    raw.loc[
        raw["portfolio_code"].notna() & ~raw["portfolio_code"].isin(funds), "fund"
    ].unique().tolist()
)

print(f"  {len(raw)} raw ts_eagle_flows rows for {rptDate}")
print(f"  {len(matched)} rows matched to a fund in our list")
if unmatched_funds:
    print(f"  WARNING: fund names with no match in s_eagle_portfolio_meta at all: {unmatched_funds}")
if not_in_our_list:
    print(f"  NOTE: matched funds not in our derivative-cover fund list (dropped): {not_in_our_list}")

# reshape to loosely resemble the real CACT column structure — see the
# module docstring for exactly which columns are genuinely available
# vs blank placeholders
cact_like = pd.DataFrame(
    {
        "Entity_Name": matched["fund"],
        "Settlement Date": matched["datestamp"],
        "Accounting Date": matched["datestamp"],  # same field, no separate date available
        "Transaction Type": matched["transaction"],
        "Investment Type": None,
        "Primary Asset ID": None,
        "Transaction description": matched["product"] + " | " + matched["type"],
        "Base Amount": matched["requestedvalue"],
        "Entity ID": matched["portfolio_code"],
        "Custody Bank": None,
        "Scrip Account": None,
        "Bank Account Number": None,
        "Security Type": None,
        "Base Currency": None,
        "Local Amount": matched["requestedvalue"],  # NOT verified — assumed same as Base Amount
        "Settlement Currency": None,
        "FX Rate": None,
        "ISIN Code": None,
        "Event ID": None,
        "Cash activity cancel flag": None,
    }
)

out_fln = f"CACT_DB_PARTIAL ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
out_path = os.path.join(pth_dl, out_fln)
cact_like.to_csv(out_path, index=False)
print(f"\n  saved {len(cact_like)} rows to {out_path}")
print(
    "  Reminder: this is a best-effort PARTIAL substitute for manual review, "
    "not equivalent to real CACT data — see the module docstring for exactly "
    "what's missing."
)

print(f"\n{timediff(start_time, time.time())} total time\n")

print("\n\n#####################################################")
print("#                                                   #")
print("#   END derv_checker_cact_db.py (PARTIAL, review-only)     #")
print("#                                                   #")
print("#####################################################\n\n")
