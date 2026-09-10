#!/usr/bin/env python
# coding: utf-8

# # Derivative Cover Reporting — DB version
#
# Same orchestration as derv_checker.py, but calls
# derv_checker_downloading_db.py (queries prime_eagle directly) instead
# of derv_checker_downloading.py (Eagle/Selenium scrape). Every other
# stage — derv_checker_table.py, derv_checker_compiling.py,
# derv_checker_summarising.py, derv_checker_freecover.py, and
# derv_checker_cact.py — is unchanged and unaffected either way, since
# none of them call Eagle directly; they only read files the downloading
# stage produces. derv_checker.py itself is untouched — this is a
# separate entry point to use instead of it.
#
# derv_checker_cact.py (cash activity) has no clean database equivalent
# yet (checked: neither pit_cash_recon_pim nor tmp_eagle_cash_projection
# matches CACT's categorized withdrawal/contribution/distribution
# structure), so it still runs Eagle-based, unchanged, at the end.

# libraries, libraries!
import time
import os
import sys
import pandas as pd
import subprocess
from constants import (
    pthEXPORTS,
    frcv_file,
    dc_do_db,
    dc_co,
    dc_su,
    dc_fr,
    dc_tb,
    dc_ct,
)
from utilities import timediff, parn_de

start_time_roundtrip = time.time()


# define the report function
def derv_check_db():
    start_time = time.time()
    from datetime import datetime

    # import report variables with parn_de()
    (fPARN, fDE, funds, rptDate, summ_yn, dervthreshold, batches) = parn_de()
    # derive file names
    print("Deriving file names ...")
    df_frcv = pd.read_excel(
        frcv_file, sheet_name="Summary", header=None, usecols="C", nrows=1
    )
    filename = os.path.join(
        pthEXPORTS,
        f"Derv \
{rptDate.strftime('%d%b%Y')}.xlsx",
    )

    # test that derivatives (fDE) downloaded
    if (os.path.exists(fDE) and os.path.getsize(fDE) > 10 
        and os.path.exists(fPARN) and os.path.getsize(fPARN) > 10):
        if rptDate == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date():
            print(
                f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: \
{frcv_file} was completed at {time.ctime(os.path.getmtime(frcv_file))}"
            )

    else:
        try:
            print("\nStarting derv_checker_downloading_db\n")
            subprocess.run([sys.executable, dc_do_db])  # download hldgs & dervs from DB
            subprocess.run([sys.executable, dc_tb])  # summarised table format
            subprocess.run([sys.executable, dc_co])  # compile calc per fund
            subprocess.run([sys.executable, dc_su])  # summarise over all funds
            subprocess.run([sys.executable, dc_fr])  # prepare daily report

            if os.path.isfile(filename):
                print(
                    f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: \
{filename.removeprefix(pthEXPORTS)} was completed \
at {time.ctime(os.path.getmtime(filename))}\n",
                )
        except Exception as e:
            print(e)

        print(
            "\n",
            f"{timediff(start_time_roundtrip, time.time())} roundtrip \
time to download and complete derivative cover reports",
        )
    print(f"\n {timediff(start_time, time.time())} to run derv_check_db()")

    if os.path.exists(fPARN) and os.path.exists(fDE):
        subprocess.run([sys.executable, dc_ct])  # import cash flows (still Eagle-based)
    else:
        pass


if __name__ == "__main__":
    derv_check_db()
