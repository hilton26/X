# # Derivative Cover Reporting

# ### Input:
#
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations
# \GRC\Compliance\Daily\py_reports, "downloader" tab
#
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo
# /derv_checker_downloading.ipynb "dervs" and "creds" tabs
#
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo
# /derv_checker_compiling.ipynb "dervs" and "creds" tabs


# define the report function
def derv_check():
    # libraries, libraries!
    import time
    import os
    import sys
    import pandas as pd
    import subprocess
    from constants import (
        pthEXPORTS,
        pthDaily,
        frcv_file,
        dc_do,  # can  be replicated
        dc_tb,
        dc_fc,
        dc_ct,
    )

    # downloading, compiling, summarising,
    # freecover (ex), table, freec (new), cact
    from utilities import timediff, parn_de

    start_time = time.time()
    start_time_roundtrip = time.time()
    from datetime import datetime

    # import report variables with parn_de()
    (fPARN, fDE, funds, rptDate, summ_yn, dervthreshold, batches) = parn_de()

    # dataframe Free Cover.xlsm
    print("Deriving file names ...")
    df_frcv = pd.read_excel(
        frcv_file, sheet_name="Summary", header=None, usecols="C", nrows=1
    )

    # derive file name, e.g., \Derv 12Apr2023.xlsx
    filename = os.path.join(pthEXPORTS, f"{rptDate.strftime(r'%Y%m%d')}_derv_calc.xlsx")
    # print(filename)

    # check if Free Cover.xlsm was completed for today
    if (
        rptDate.date()
        == datetime.strptime(df_frcv.iloc[0, 0][15:26], "%d %b %Y").date()
    ):
        frcv_rel = frcv_file.removeprefix(pthDaily)
        print(
            rf"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: "
            rf"{frcv_rel} for "
            rf"{rptDate.strftime('%a %d %b %Y')} was "
            rf"completed at {time.ctime(os.path.getmtime(frcv_file))}",
        )

    else:
        try:
            print("\nStarting derv_checker_downloading\n")
            subprocess.run([sys.executable, dc_do])  # download hldgs & dervs
            subprocess.run([sys.executable, dc_tb])  # summarised table format
            subprocess.run([sys.executable, dc_fc])  # freecover sheet
            # subprocess.run([sys.executable, dc_ct])  # import cash flows

            if os.path.isfile(filename):
                print(
                    rf"{datetime.now().strftime('%Hh%M:%Ss %a %d %b %Y')}: "
                    rf"{filename.removeprefix(pthEXPORTS)} was completed "
                    rf"at {time.ctime(os.path.getmtime(filename))}\n",
                )
        except Exception as e:
            print(e)

        print(
            rf"\n{timediff(start_time_roundtrip, time.time())} roundtrip "
            rf"time to download and complete derivative cover reports"
        )
    print(
        rf"\n {timediff(start_time, time.time())} to run derv_check(), "
        rf"now on to getting the cash transaction reports",
    )

    if os.path.exists(fPARN) and os.path.exists(fDE):
        subprocess.run([sys.executable, dc_ct])  # import cash flows
    else:
        pass


if __name__ == "__main__":
    derv_check()
