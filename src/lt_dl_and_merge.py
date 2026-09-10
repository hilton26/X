#!/usr/bin/env python
# coding: utf-8

# # Download and Merge Eagle Lookthrough Sheets

# ### Input:
#
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports, "downloader" and "creds" tabs
#
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_downloading.ipynb
#
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_merge.ipynb

print("\nDownloading and merging lookthrough reports ...\n")

import time

start_time = time.time()
start_time_download = time.time()
import subprocess
import sys
from pathlib import Path
import shutil
from constants import lt_dl, lt_merge, pth_dl, pthReports
from utilities import timediff

# download lookthrough holdiings based on funds in column N,
# and desired number of batches to be downloaded in column S
subprocess.run([sys.executable, lt_dl])


# derive name of downloaded lookthrough holdings file
df = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df.iloc[:, 0].str.upper()

# get report parameters date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
rptDate = (
    df1.iloc[1, 0].date()
    if isinstance(df1.iloc[1, 0], datetime)
    else prior_month_end(datetime.today().date())
)  # prior month end or report date override; type is datetime()

lt_fname = os.path.join(
    pthTest,
    f"LT holdings ({len(funds)}) \
{rptDate.strftime('%d%b%Y')}.xlsx",
)

rpts_done = os.path.join(
    pthTest,
    f"issuers_3_{rptDate.strftime('%d%b%Y')}.xlsx",
)

# if lt_fname.is_file:
#     subprocess.run([sys.executable, lt_merge])
# else:
#     print(f" Not all the fund lookthroughs were downloaded")
#     pass

if rpts_done.is_file:
    print(
        f"Transferring completed Reg 28 and Reg 30 \
reports to the reporting folder"
    )

    for pattern in (
        f"*Reg28 {rptDate.strftime('%d%b%Y')}.xlsx",
        f"*Reg30 {rptDate.strftime('%d%b%Y')}.xlsx",
    ):
        for file in Path(pth_dl).glob(pattern):
            shutil.copy(file, pthReports)

else:
    print(f" regulatory reports have not been created")
    pass

print(
    f"\n {timediff(start_time_download, time.time())} \
downloading and merging lookthrough reports"
)
