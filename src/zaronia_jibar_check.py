#!/usr/bin/env python
# coding: utf-8

# # Non-ZARONIA Floating-Rate Holdings Since the "No New JIBAR" Milestone (1 May 2026)
#
# From 1 May 2026, new SA market contracts referencing JIBAR are meant to
# have stopped in favour of ZARONIA. This script identifies floating/
# reference-rate-linked securities held in any PIM fund on or after that
# date whose instrument description does NOT reference ZARONIA.
#
# Detection is text-based on instrument_name (per the request), matching
# any of: ZARONIA, the literal word JIBAR, the "JB<digit>" tenor
# shorthand actually used in these descriptions (e.g. "Jb3" = 3-month
# JIBAR), or PRIME. Plain fixed-rate bonds (no rate reference at all)
# are intentionally excluded — they aren't part of the JIBAR/ZARONIA
# question.
#
# Note: this identifies what's currently HELD since 1 May 2026, not
# whether a given bond was newly ISSUED after that date — the holdings
# data here doesn't carry an original issue date, so a JIBAR-referenced
# bond below could be pre-existing legacy paper (expected, fine) or a
# genuinely new post-cutover issuance (the actual compliance concern).
# That distinction needs checking against each instrument's issue date.

import pandas as pd
from sqlalchemy import create_engine, text
from constants import db_eagle_uri

CUTOFF = "2026-05-01"  # the "no new JIBAR" milestone date

engine = create_engine(db_eagle_uri, connect_args={"connect_timeout": 15})
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
engine.dispose()


import re


def classify(name):
    name_up = name.upper()
    # FX forward deal-reference codes can coincidentally contain a
    # "JB<digit>"-looking substring (e.g. "25JB1LB4KVDGDNRD") — these
    # aren't rate-referenced debt instruments at all, so exclude them
    # before checking for a genuine rate reference.
    if re.search(r"\bFWD\b|FORWARD", name_up):
        return "FX_FORWARD (excluded, not a rate-referenced instrument)"
    if "ZARONIA" in name_up:
        return "ZARONIA"
    if "JIBAR" in name_up:
        return "JIBAR"
    # genuine JIBAR tenor notation always appears as e.g. "Jb3+210" —
    # require the digit(s) not be glued to a preceding alphanumeric
    # character, and the "+spread" suffix, to avoid matching stray
    # substrings inside unrelated codes
    if re.search(r"(?<![A-Z0-9])JB[0-9]{1,2}\+[0-9]", name_up):
        return "JIBAR"
    if re.search(r"(?<![A-Z0-9])PRIME(?![A-Z0-9])", name_up):
        return "PRIME"
    return "OTHER/UNCLEAR"


df["reference_rate"] = df["instrument_name"].apply(classify)

excluded = df[df["reference_rate"].str.startswith("FX_FORWARD")]
rate_referenced = df[~df["reference_rate"].str.startswith("FX_FORWARD")]
not_zaronia = rate_referenced[rate_referenced["reference_rate"] != "ZARONIA"].sort_values(
    ["reference_rate", "instrument_code"]
)

print(f"Candidate matches since {CUTOFF}: {len(df)} total")
print(f"  of which excluded as FX forwards (coincidental code match, not rate-referenced): {len(excluded)}")
print()
print(f"Genuinely rate-referenced instruments: {len(rate_referenced)}")
print(rate_referenced["reference_rate"].value_counts().to_string())
print()
print(f"NOT ZARONIA-referenced: {len(not_zaronia)}\n")
print(
    not_zaronia[
        [
            "reference_rate",
            "instrument_code",
            "instrument_name",
            "currency",
            "maturity_date",
            "first_held_since_cutoff",
            "last_held",
            "funds",
        ]
    ].to_string(index=False)
)

out_path = "not_zaronia_referenced_since_1May2026.csv"
not_zaronia.to_csv(out_path, index=False)
print(f"\nSaved to {out_path}")
