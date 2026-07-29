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

# define the report function
def derv_check():
    import time

    start_time_roundtrip = time.time()

    # libraries, libraries!
    import os, subprocess, sys, time, pandas as pd

    # import pandas as pd
    from datetime import datetime
    from constants import (
        pthPy,
        pthEXPORTS,
        frcv_file,
        pthDaily,
        dc_do,
        dc_co,
        dc_su,
        dc_fr,
    )
    from utilities import timediff, prior_working_day

    # check if the report already exists, if yes, exit the script
    df = pd.read_excel(pthPy, sheet_name="dervs", header=None, usecols="A,D:E").dropna(
        subset=[0]
    )
    k = df.iloc[0, 2]
    rptDate = (
        k if isinstance(k, datetime) else prior_working_day(datetime.today())
    )  # prior working day or report date override given today's date
    filename = os.path.join(pthEXPORTS, f"Derv {rptDate.strftime('%d%b%Y')}.xlsx")
    df_frcv = pd.read_excel(
        frcv_file, sheet_name="Summary", header=None, usecols="C", nrows=1
    )

    if (
        rptDate.date()
        == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date()
    ):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: {frcv_file.removeprefix(pthDaily)} \
was completed {time.ctime(os.path.getmtime(frcv_file))}"
        )
        return
    else:
        try:
            print("\n", "Starting derv_checker_downloading.ipynb", "\n")
            subprocess.run([sys.executable, dc_do])
            print(
                "\n",
                "derv_checker_downloading completed",
                "\n",
                "Starting derv_checker_compiling.ipynb",
                "\n",
            )

            subprocess.run([sys.executable, dc_co])
            print(
                "\n",
                "derv_checker_compiling completed",
                "\n",
                "Starting derv_checker_summarising",
                "\n",
            )

            subprocess.run([sys.executable, dc_su])
            print(
                "\n",
                "derv_checker_summarising completed",
                "\n",
                "Starting derv_checker_freecover",
                "\n",
            )

            subprocess.run([sys.executable, dc_fr])
            print("\n", "derv_checker_freecover completed", "\n")

            if os.path.isfile(filename):
                print(
                    f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: \
                {filename.removeprefix(pthEXPORTS)} was completed at {time.ctime(os.path.getmtime(filename))}",
                    "\n",
                )

        except Exception as e:
            print(e)

        print(
            "\n",
            f"{timediff(start_time_roundtrip, time.time())} roundtrip time to download and \
complete {rptDate.strftime('%d %b %Y')} derivative cover reports",
        )


# list the frequency of runs

from datetime import datetime, timedelta

# list of times every 15 minutes from 8:45 AM to 17:30
start = datetime.strptime("08:45", "%H:%M")
end = datetime.strptime("19:30", "%H:%M")
increment = 10  # every 15 minutes

times = []
current = start
while current <= end:
    times.append(current.strftime("%H:%M"))
    current += timedelta(minutes=increment)

print(f"{increment} minute cadence: {', '.join(times)}")


# scheduler    https://pypi.org/project/schedule/
import schedule, time

for t in times:
    schedule.every().monday.at(t).do(lambda: derv_check())
    schedule.every().tuesday.at(t).do(lambda: derv_check())
    schedule.every().wednesday.at(t).do(lambda: derv_check())
    schedule.every().thursday.at(t).do(lambda: derv_check())
    schedule.every().friday.at(t).do(lambda: derv_check())

while True:
    schedule.run_pending()
    time.sleep(10)
