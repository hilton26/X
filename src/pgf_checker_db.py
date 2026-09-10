#!/usr/bin/env python
# coding: utf-8

# # Compile PGF Hedge Share Class Sheet — DB version
#
# Same orchestration as pgf_checker.py, but calls pgf_downloading_db.py
# (queries prime_eagle directly) instead of pgf_downloading.py (Eagle/
# Selenium scrape). pgf_compiling.py is unchanged and unaffected either
# way. pgf_checker.py itself is untouched — this is a separate entry
# point to use instead of it.


# define the report function
def pgf_check_db():
    # libraries, libraries!
    import time
    from datetime import datetime
    import os
    import pandas as pd
    import subprocess
    import sys
    from constants import pthPy, pthHdg, pth_dl, pg_do_db, pg_co
    from utilities import timediff, prior_working_day

    start_time_pgf = time.time()

    # get report date
    df = pd.read_excel(pthPy, sheet_name="arc", usecols="G:I")
    rptDate = (
        df.iloc[0, 2].date()
        if isinstance(df.iloc[0, 2], datetime) and not pd.isna(df.iloc[0, 2])
        else prior_working_day(datetime.today())
    )  # prior working day or report date override

    # construct file names
    UTs_name = os.path.join(
        pth_dl,
        f"UTPS PGF_UT_Prices({len(df['pgf: UT prices'].dropna())}) \
{rptDate.strftime('%d%b%Y')}.csv",
    )
    NAV_name = os.path.join(
        pth_dl,
        f"PARN PGF_Holdings({len(df['pgf: PAR-N'].dropna())}) \
{rptDate.strftime('%d%b%Y')}.csv",
    )

    print(
        "Expected file names before pgf_downloading_db.py:",
        "\n",
        UTs_name,
        "\n",
        NAV_name,
    )

    filename = (
        pthHdg
        + rf"\{rptDate.strftime('%Y%m%d')} PGF \
Share Class Hedges.xlsx"
    )
    if os.path.isfile(filename):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: \
{filename.removeprefix(pthHdg)} was completed \
at {time.ctime(os.path.getmtime(filename))}"
        )
        pass
    else:
        try:
            # run the DB-based hedge checker
            subprocess.run([sys.executable, pg_do_db])
            if os.path.isfile(UTs_name) and os.path.isfile(NAV_name):
                subprocess.run([sys.executable, pg_co])

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


if __name__ == "__main__":
    pgf_check_db()
