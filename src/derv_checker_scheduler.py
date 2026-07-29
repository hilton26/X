#!/usr/bin/env python
# coding: utf-8

# # Scheduler for derv_checker.py
#
# Runs derv_checker.py every 15 minutes from 07:43 to 18:13 (Mon-Fri).
# Stops for the day once the completed output file - Free Cover.xlsm -
# has already been updated today.

import os
import subprocess
import sys
import time
from datetime import datetime, timedelta

import schedule

from constants import frcv_file, pth_gitrepo

DERV_CHECKER = pth_gitrepo + r"\derv_checker.py"


def already_completed_today():
    """True if Free Cover.xlsm has already been updated today."""
    return (
        datetime.fromtimestamp(os.path.getmtime(frcv_file)).date()
        == datetime.today().date()
    )


def run_derv_checker():
    if already_completed_today():
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: "
            f"{os.path.basename(frcv_file)} already updated today - stopping scheduler"
        )
        schedule.clear()
        return

    print(
        f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: running derv_checker.py"
    )
    subprocess.run([sys.executable, DERV_CHECKER])


# build run times: every 15 minutes from 07:43 to 18:13
start = datetime.strptime("07:43", "%H:%M")
end = datetime.strptime("18:13", "%H:%M")
increment = 15

times = []
current = start
while current <= end:
    times.append(current.strftime("%H:%M"))
    current += timedelta(minutes=increment)

print(f"Scheduled run times: {', '.join(times)}")

for t in times:
    schedule.every().monday.at(t).do(run_derv_checker)
    schedule.every().tuesday.at(t).do(run_derv_checker)
    schedule.every().wednesday.at(t).do(run_derv_checker)
    schedule.every().thursday.at(t).do(run_derv_checker)
    schedule.every().friday.at(t).do(run_derv_checker)

while schedule.get_jobs():
    schedule.run_pending()
    time.sleep(10)

print("Scheduler finished for today")
