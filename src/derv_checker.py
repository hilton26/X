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
import os
from pathlib import Path
import pandas as pd
import subprocess
from constants import pthPy, pthEXPORTS, frcv_file, pthDaily, dc_do, dc_co, dc_su, dc_fr
from utilities import timediff, prior_working_day


# define the report function
def derv_check():
    from datetime import datetime

    # check if the report already exists, if yes, exit the script
    df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
        subset=[0]
    )
    k = df.iloc[2, 1]
    rptDate = k if isinstance(k, datetime) else prior_working_day(datetime.today())
    filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
    df_frcv = pd.read_excel(
        frcv_file, sheet_name="Summary", header=None, usecols="C", nrows=1
    )

    if (
        rptDate.date()
        == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date()
    ):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: {frcv_file} \
        was completed at {time.ctime(os.path.getmtime(frcv_file))}"
        )
        return
    else:
        try:
            print("\nStarting derv_checker_downloading\n")

            subprocess.run(["python", dc_do])
            subprocess.run(["python", dc_co])
            subprocess.run(["python", dc_su])
            subprocess.run(["python", dc_fr])

            if os.path.isfile(filename):
                print(
                    f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: \
                {filename.removeprefix(pthEXPORTS)} was completed \
                    at {time.ctime(os.path.getmtime(filename))}",
                    "\n",
                )

        except Exception as e:
            print(e)

        print(
            "\n",
            f"{timediff(start_time_roundtrip, time.time())} roundtrip time to \
                download and complete derivative cover reports",
        )


derv_check()
