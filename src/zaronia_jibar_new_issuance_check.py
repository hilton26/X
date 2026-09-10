#!/usr/bin/env python
# coding: utf-8

# # Likely NEW Non-ZARONIA Issuance Since the "No New JIBAR" Milestone (1 May 2026)
#
# Builds on zaronia_jibar_check.py, which identifies floating/reference-
# rate-linked securities held in any PIM fund on or after 1 May 2026
# whose instrument description does NOT reference ZARONIA (i.e. JIBAR-
# or PRIME-referenced). That script's own caveat: being HELD since 1 May
# doesn't mean NEWLY ISSUED since 1 May — most JIBAR paper still held is
# legacy, pre-existing debt that continues trading until maturity, which
# is expected and not a compliance concern.
#
# This script adds the proxy check needed to separate those two cases:
# for each non-ZARONIA instrument held since the cutoff, check whether
# it EVER appeared in fund holdings BEFORE 1 May 2026.
#   - Appeared before  -> legacy paper, held through the cutover (expected)
#   - Never appeared before -> no evidence of pre-existing life; more
#     likely a new issuance after the cutoff, worth reviewing
#
# This is still a proxy, not a definitive answer — an instrument could
# have been issued shortly before 1 May and simply not yet been bought
# by a PIM fund at that point, and then bought soon after. It narrows
# the 324-instrument list down to the ones actually worth checking
# against real issue dates, rather than being a final verdict on its own.

import re
import pandas as pd
from sqlalchemy import create_engine, text
from constants import db_eagle_uri

CUTOFF = "2026-05-01"  # the "no new JIBAR" milestone date
engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 15})

print(f"Finding reference-rate instruments held since {CUTOFF} ...")
q = text(
    """
    SELECT
        instrument_code,
        instrument_name,
        sub_security_type,
        currency,
        MIN(datestamp) AS first_held_since_cutoff,
        MAX(datestamp) AS last_held,
        MAX(maturity_date) AS maturity_date,
        GROUP_CONCAT(DISTINCT portfolio_code ORDER BY portfolio_code) AS funds
    FROM tmp_eagle_holdings_analytics
    WHERE datestamp >= :cutoff
      AND instrument_name REGEXP 'ZARONIA|JIBAR|JB[0-9]|PRIME'
    GROUP BY instrument_code, instrument_name, sub_security_type, currency
    """
)
with engine.connect() as conn:
    result = conn.execute(q, {"cutoff": CUTOFF})
    df = pd.DataFrame(result.fetchall(), columns=list(result.keys()))


def classify(name):
    name_up = name.upper()
    if re.search(r"\bFWD\b|FORWARD", name_up):
        return "FX_FORWARD (excluded, not a rate-referenced instrument)"
    if "ZARONIA" in name_up:
        return "ZARONIA"
    if "JIBAR" in name_up:
        return "JIBAR"
    # allow optional whitespace around "+" (e.g. "JB3+ 205", "JB3 + 410")
    if re.search(r"(?<![A-Z0-9])JB[0-9]{1,2}\s*\+\s*[0-9]", name_up):
        return "JIBAR"
    if re.search(r"(?<![A-Z0-9])PRIME(?![A-Z0-9])", name_up):
        return "PRIME"
    return "OTHER/UNCLEAR"


df["reference_rate"] = df["instrument_name"].apply(classify)
not_zaronia = df[
    ~df["reference_rate"].str.startswith("FX_FORWARD") & (df["reference_rate"] != "ZARONIA")
].copy()

print(f"  {len(not_zaronia)} not-ZARONIA-referenced instruments found\n")

print(f"Checking each for holdings history before {CUTOFF} ...")
codes = tuple(not_zaronia["instrument_code"].unique())
pre_cutoff_q = text(
    """
    SELECT DISTINCT instrument_code
    FROM tmp_eagle_holdings_analytics
    WHERE instrument_code IN :codes AND datestamp < :cutoff
    """
).bindparams(codes=codes, cutoff=CUTOFF)
with engine.connect() as conn:
    seen_before = {row[0] for row in conn.execute(pre_cutoff_q).fetchall()}
engine.dispose()

not_zaronia["held_before_cutoff"] = not_zaronia["instrument_code"].isin(seen_before)
likely_new = not_zaronia[~not_zaronia["held_before_cutoff"]].sort_values(
    ["reference_rate", "first_held_since_cutoff"]
)
legacy = not_zaronia[not_zaronia["held_before_cutoff"]]

print(f"  {len(legacy)} have pre-cutoff history -> legacy paper, expected")
print(f"  {len(likely_new)} have NO pre-cutoff history -> review these\n")

cols = [
    "reference_rate",
    "instrument_code",
    "instrument_name",
    "currency",
    "maturity_date",
    "first_held_since_cutoff",
    "last_held",
    "funds",
]
print("=== Likely new non-ZARONIA issuance (no holdings before the cutoff) ===")
print(likely_new[cols].to_string(index=False))

out_path = "likely_new_non_zaronia_issuance_since_1May2026.csv"
likely_new[cols].to_csv(out_path, index=False)
print(f"\nSaved to {out_path}")

legacy_out_path = "legacy_non_zaronia_held_since_1May2026.csv"
legacy[cols].to_csv(legacy_out_path, index=False)
print(f"Legacy (pre-existing) list saved to {legacy_out_path}")
