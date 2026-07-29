#!/usr/bin/env python
# coding: utf-8

# # Derivative Cover Reporting

# ### Input:
#
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports, "downloader" tab
#
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_downloading.ipynb "dervs" and "creds" tabs
#
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_compiling.ipynb "dervs" and "creds" tabs

import time

start_time_roundtrip = time.time()

# libraries, libraries!
import os, sys
import pandas as pd
import subprocess
from constants import (
    pthEXPORTS,
    frcv_file,
    dc_do,
    dc_co,
    dc_su,
    dc_fr,
    dc_tb,
    dc_fc,
    dc_ct,
)  # downloading, compiling, summarising, freecover (ex), table, freec (new), cact
from utilities import timediff, parn_de, prior_working_day


# define the report function
def derv_check():
    start_time = time.time()
    from datetime import datetime
    
    # import report variables with parn_de()
    fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()

    # derive file names
    print("Deriving file names ...")
    df_frcv = pd.read_excel(
        frcv_file, sheet_name="Summary", header=None, usecols="C", nrows=1
    )
    filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
    
    # test that holdings (fPARN) downloaded
    test_fPARN = False
    if os.path.exists(fPARN):
        df_fPARN = pd.read_csv(fPARN)
        test_fPARN = (
            os.path.exists(fPARN) and df_fPARN.columns[1] == "Valuation First Level"
        )

    # test that derivatives (fDE) downloaded
    test_fDE = False
    if os.path.exists(fDE):
        df_fDE = pd.read_csv(fDE)
        test_fDE = os.path.exists(fDE) and df_fDE.columns[6] == "Effective Exposure"

    if (
        rptDate.date()
        == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date()
    ):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: {frcv_file} \
was completed at {time.ctime(os.path.getmtime(frcv_file))}"
        )    
    
    # elif not test_fPARN or not test_fDE:
    #     sys.exit(
    #         f"\n\n STOPPING: missing expected download(s):\n"
    #         f"  {fPARN} which {'exists' if test_fPARN else 'does not exist'}\n"
    #         f"  {fDE} which {'exists' if test_fDE else 'does not exist'}\n"
    #     )

    else:
        try:
            print("\nStarting derv_checker_downloading\n")
            subprocess.run([sys.executable, dc_do])  # download holdings and dervs
            subprocess.run([sys.executable, dc_tb])  # summarised table format
            # subprocess.run([sys.executable, dc_fc])  # NEW freecover sheet
            subprocess.run([sys.executable, dc_co])  # compiling calc for each fund
            subprocess.run([sys.executable, dc_su])  # collate a summary of all funds
            subprocess.run([sys.executable, dc_fr])  # prepare daily report
            subprocess.run([sys.executable, dc_ct])  # import cash flows

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
            f"{timediff(start_time_roundtrip, time.time())} roundtrip time to \
download and complete derivative cover reports",
        )
    print(f"\n {timediff(start_time, time.time())} to run derv_check()")

if __name__ == "__main__":
    derv_check()