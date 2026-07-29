####### TEST r_classifier()
# def r_classifier(report_type, url_input, report_date=datetime.datetime.today()):
"""
receives a link to a sheet in excel format then
runs issuers_1.ipynb using the given inputs

"""

import time
from turtle import pd

start_time_classifier = time.time()

from datetime import datetime
import pandas as pd
import os
import sys
import subprocess
import xlwings as xw
from constants import pthPy, issuers_1
from utilities import timediff

# prep
from constants import pthTest
from utilities import prior_month_end

# prep 1. get report date
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=2)
k = df1.iloc[1, 0]
rptDate = k.date() if k == k else prior_month_end(datetime.today()).date()

# prep 2. get funds
df2 = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df2.iloc[:, 0].apply(str.upper)

cs1_fname = os.path.join(
    pthTest,
    f"CS1 PARN holdings \
({len(funds)}) {rptDate.strftime('%d%b%Y')}.xlsx",
)

# inputs from "arc" sheet to r_classifier(): report_type, url_input, report_date
report_type = "cs1"
url_input = pd.read_excel(pthPy, sheet_name="arc", usecols="V").iloc[6, 0]
report_date = rptDate

print(report_type, "\n", url_input, "\n", report_date, "\n", cs1_fname)

# import r_classifier() input variables
rptDate = report_date
rptType = "CS1 format only" if report_type == "cs1" else "Reg 28 and Reg 30 only"
url = url_input.replace('"', "")

print(report_date, "\n", datetime.today())

# updating "arc" sheet
xw.Book(pthPy).sheets("arc").range(
    "V4"
).value = rptType  # xw.Book(pthPy).sheets("classifier").range("M1").value = rptType
xw.Book(pthPy).sheets("arc").range(
    "V8"
).value = url  # xw.Book(pthPy).sheets("classifier").range("L2").value = url
xw.Book(pthPy).save()
xw.Book(pthPy).close()

# run issuers_1.ipynb
subprocess.run([sys.executable, issuers_1])

print(
    f"{timediff(start_time_classifier, time.time())} \
        executing issuers_1.ipynb"
)

####### END TEST r_classifier()
