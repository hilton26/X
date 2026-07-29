#!/usr/bin/env python
# coding: utf-8

import time

# libraries, libraries!
import os
import subprocess
import sys
from constants import pre_issuers_1, issuers_1
from utilities import timediff


# from utilities import * for this sript to run in VSC
# https://stackoverflow.com/questions/72962306/how-to-use-run-command-to-execute-another-notebook-using-file-path
def gemsmed_rpt():
    try:
        print("gemsmed_pre_issuers_1.py is running ...")
        start_time_1 = time.time()
        subprocess.run([sys.executable, pre_issuers_1])
        print(f"{timediff(start_time_1, time.time())} for gemsmed_pre_issuers_1.py")

        start_time_2 = time.time()
        print("issuers_1.py is running ...")
        subprocess.run([sys.executable, issuers_1])
        print(f"{timediff(start_time_2, time.time())} to complete issuers_1.py")

    except Exception as e:
        print(e)


# execute the function
start_time_gemsmed = time.time()

gemsmed_rpt()

print("\n", f"{timediff(start_time_gemsmed, time.time())} roundtrip time")
