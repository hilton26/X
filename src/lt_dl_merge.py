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

print(
    "Downloading and merging lookthrough \
reports ...",
    "\n",
)

import time

start_time = time.time()
start_time_download = time.time()
import subprocess
from constants import lt_dl, lt_merge
from utilities import timediff

subprocess.run(["python", lt_dl])
subprocess.run(["python", lt_merge])

print(
    f"Downloading and merging lookthrough \
reports completed: \
{timediff(start_time_download, time.time())}"
)
