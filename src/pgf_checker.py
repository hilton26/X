#!/usr/bin/env python
# coding: utf-8

# # Compile PGF Hedge Share Class Sheet

# define the report function
def pgf_check():
    # libraries, libraries!
    import time
    from datetime import datetime
    import os, schedule
    import pandas as pd
    from pathlib import Path
    import subprocess
    from constants import pthPy, pthHdg, pth_dl, pg_do, pg_co
    from utilities import timediff, prior_working_day

    start_time_pgf = time.time()

    # get report date
    df = pd.read_excel(pthPy, sheet_name="arc", usecols="G:I")
    k = df.iloc[1, 2]
    rptDate = (
        k
        if isinstance(k, datetime) and not pd.isna(k)
        else prior_working_day(datetime.today())
    )  # prior working day or report date override

    # construct file names
    UTs_name = os.path.join(
        pth_dl,
        f"UTPS PGF_UT_prices({len(df['pgf: UT prices'].dropna())}) {rptDate.strftime('%d%b%Y')}.csv",
    )
    NAV_name = os.path.join(
        pth_dl,
        f"PARN PGF_Holdings({len(df['pgf: PAR-N'].dropna())}) {rptDate.strftime('%d%b%Y')}.csv",
    )

    # print("Expected file names before pgf_downloading.ipynb:","\n",UTs_name,"\n",NAV_name)

    filename = pthHdg + rf"\{rptDate.strftime('%Y%m%d')} PGF Share Class Hedges.xlsx"
    if os.path.isfile(filename):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: {filename.removeprefix(pthHdg)} \
was completed at {time.ctime(os.path.getmtime(filename))}"
        )
        # https://www.geeksforgeeks.org/python-os-path-getmtime-method/
        pass
    else:
        try:
            # run the hedge checker
            subprocess.run(["python", pg_do])
            if os.path.isfile(UTs_name) and os.path.isfile(NAV_name):
                # FileNotFoundError: [Errno 2] No such file or directory: 'C:\\Users\\hilton.netta\\Downloads\\UTPS PGF_UT_prices(7) 05Mar2025.csv'
                subprocess.run(["python", pg_co])

                if os.path.isfile(NAV_name):
                    print(f"Missing the UT prices file: {UTs_name}")
                elif os.path.isfile(UTs_name):
                    print(f"Missing the holdings file:  {NAV_name}")
                else:
                    print(f"Holdings file and UT prices file exist")
        except Exception as e:
            print(e)

        print(
            f"\n{timediff(start_time_pgf, time.time())} roundtrip \
time to complete the pgf report"
        )


# execute the report function
pgf_check()
