#!/usr/bin/env python
# coding: utf-8

# # Scheduler for osprey_x_test.py
#
# Runs osprey_x_test.py every 10 minutes, Monday-Friday.

import subprocess
import sys
import time
from datetime import datetime
import schedule
import src.lt_1_ny_1 as lt_1_ny_1
from utilities import timediff

def run_osprey_x_test():
    if datetime.today().weekday() >= 5:  # skip Sat/Sun
        return

    start_time = time.time()
    print(
        f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: running osprey_x_test.py"
    )
    subprocess.run([sys.executable, lt_1_ny_1.__file__])
    print(f"{timediff(start_time, time.time())} running osprey_x_test.py")


schedule.every(10).minutes.do(run_osprey_x_test)

print("Scheduler running osprey_x_test.py every 10 minutes (Mon-Fri). Ctrl+C to stop.")

while True:
    schedule.run_pending()
    time.sleep(10)
