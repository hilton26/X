

'''
Runs cs1_parn_download.py and then
runs cs1_parn_merge.py which then also
runs rclasssifier()

'''

import time
start_time_roundtrip = time.time()

# libraries, libraries!
import os, sys
import pandas as pd
import subprocess
from constants import (
    pth_dl,
    cs1_dl,
    cs1_mg
)  # cs1 downloading, cs1 merging
from utilities import timediff, osprey

# define the report function
def cs1_reports():
    start_time = time.time()
    from datetime import datetime

    # get fund codes from r28_cs1 tab of the py_report.xlsm sheet
    df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
    funds = df.iloc[:, 0].apply(str.upper)

    # get report date
    df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=2)
    k = df1.iloc[1, 0]
    rptDate = k.date() if k == k else prior_month_end(datetime.today()).date()

    # print inputs
    fnav_name = os.path.join(pth_dl, )
    # test that derivatives (fDE) downloaded
    if (
        rptDate.date()
        == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date()
    ):
        print(
            f"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: {frcv_file} \
was completed at {time.ctime(os.path.getmtime(frcv_file))}"
        )    

    else:
        try:
            print("\nStarting CS1 download and reporting ...\n")
            subprocess.run([sys.executable, cs1_dl])  # download holdings and dervs
            subprocess.run([sys.executable, cs1_mg])  # summarised table format

    print(f"\n {timediff(start_time, time.time())} to \
run cs1() for {rptDate.strftime('%A %d %B %YY')}")

if __name__ == "__main__":
    cs1()