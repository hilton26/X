#!/usr/bin/env python
# coding: utf-8

# # Compile and Summarise PGF Hedge Share Class Sheets

print("\n\n#####################################")
print("#                                   #")
print("#    START 2/2 pgf_compiling.py X   #")
print("#                                   #")
print("#####################################\n\n")

# Import libraries
import time

start_time = time.time()
start_time_pgf_compiling = start_time
print("Importing libraries ...")

from datetime import datetime
import pandas as pd
import xlwings as xw
import os, shutil
import re  # for regex
from re import search  # for regex
from pathlib import Path
from constants import pthPy, pthHdg, pth_dl, pth_hdg_tmpl
from utilities import timediff, prior_working_day

print(f"{timediff(start_time, time.time())} importing libraries", "\n")

# Get the names of the downloaded unit trust prices and the holdings files and create dataframes from themfiles just saved in the Downloads folder

start_time = time.time()
print("Creating unit trust prices and holdings dataframes ...")

# get the report date and fund clodes to be looked up
df = pd.read_excel(pthPy, sheet_name="arc", usecols="G:I")
k = df.iloc[1, 2]
rptDate = (
    k
    if isinstance(k, datetime) and not pd.isna(k)
    else prior_working_day(datetime.today())
)  # prior working day or report date override; has type datetime.datetime()
df_navs = df["pgf: UT prices"].dropna()
df_hldgs = df["pgf: PAR-N"].dropna()

# names of the downloaded unit trust prices and the holdings files
parN_nm = os.path.join(
    pth_dl, f"UTPS PGF_UT_prices({len(df_navs)}) {rptDate.strftime('%d%b%Y')}.csv"
)
utP_nm = os.path.join(
    pth_dl, f"PARN PGF_Holdings({len(df_hldgs)}) {rptDate.strftime('%d%b%Y')}.csv"
)
print("", rptDate, "\n", parN_nm, "\n", utP_nm)

# create dataframes of the holdings and unit trust sheets
wbU = pd.read_csv(parN_nm)  # class NAVs (13 columns)
wbH = pd.read_csv(utP_nm)  # portfolio holdings (43 columns)

print(
    f"{timediff(start_time, time.time())} creating unit trust prices and holdings dataframes",
    "\n",
)

# Convert numerical columns from str to float

start_time = time.time()
print(
    "Converting numerical holdings and derivative columns from type string to type float ..."
)

# convert derivative columns from str to float
# https://stackoverflow.com/questions/55557004/getting-attributeerror-float-object-has-no-attribute-replace-error-while
headsH = [
    "Original Nominal",
    "Sum of Market Value Income",
    "% of Total Market Value",
    "Market Value %",
    "Current Exposure",
    "Current Exposure %",
    r"Market Price /Yield",
]
for head in headsH:
    wbH[head] = [str(x).replace(",", "").replace("-", "-") for x in wbH[head]]
    wbH[head] = wbH[head].astype(float)

# holdings date column from type string to type datetime
wbH["i Position Effective Date"] = pd.to_datetime(wbH["i Position Effective Date"])

# holdings columns to be converted from str to float
try:
    headsU = [
        "NAV Price",
        "Clean Price",
        "Income Price",
        "Class Size",
        "Units in Class",
    ]
    for head in headsU:
        wbU[head] = [str(x).replace(",", "").replace("-", "-") for x in wbU[head]]
        wbU[head] = wbU[head].astype(float)

    # holdings date column from type string to type datetime
    wbU["Effective Date"] = pd.to_datetime(wbU["Effective Date"])
except:
    print(" Unit trust prices likely not yet available")

print(
    f"{timediff(start_time, time.time())} converting numerical holdings and derivative columns from type string to type float",
    "\n",
)

# get previous report date from extracted file names in the PGF hedge report folder
# https://www3.ntu.edu.sg/home/ehchua/programming/howto/Regexe.html

start_time = time.time()
print("Getting prior day file ...")

fls = [
    int(re.match("\d{8}", str(k)).group())
    for k in os.listdir(pthHdg)
    if re.match("\d{8}", str(k)) != None
]
fln = os.path.join(pthHdg, str(max(fls)) + " PGF Share Class Hedges.xlsx")
print(" ", fln)

# get the day before's values
pr_S = pd.read_excel(fln, sheet_name="Summary", header=0, nrows=11, usecols="I")

print(f"{timediff(start_time, time.time())} getting prior day file", "\n")

# populate the pgf template

start_time = time.time()
print(f"Creating new PGF sheet for {rptDate.strftime('%a %#d %b %Y')} ...")

# open the pgf template workbook and assign sheet variables
wb = xw.Book(pth_hdg_tmpl)  # open the pgf template workbook as an object
shtH = wb.sheets[
    "Portfolio Valuation"
]  # assign sheet containing holdings as obtained from Eagle
shtU = wb.sheets[
    "Class NAVs"
]  # assign sheet containing unit trust prices as obtained from Eagle
shtZ = wb.sheets["Summary"]  # assign summary sheet

# update the pgf holdings and unit price sheets in the pgf template
# excel.DisplayAlerts = False                                        # suppress Excel warning dialogues
shtH.clear()  # clear the receiving holdings sheet
shtH.range("A1").options(index=False).value = wbH  # paste fund holdings
shtU.clear()  # clear the receiving deltas sheet
shtU.range("A1").options(index=False).value = wbU  # paste fund unit trust prices
shtZ.range("J1").options(index=False).value = pr_S  # paste prior day's summary values
wb.save(
    os.path.join(
        pthHdg,
        f"{rptDate.strftime('%Y%m%d')} \
PGF Share Class Hedges.xlsx",
    )
)  # save the day's file
wb.save(os.path.join(pthHdg, "PGF Share Class Hedges.xlsm"))  # save the template
# excel.DisplayAlerts = True                                        # unsuppress Excel warning dialogues
wb.close()  # close the derv template workbook object

print(
    f"{timediff(start_time, time.time())} creating new PGF sheet for {rptDate.strftime('%a %#d %b %Y')}",
    "\n",
)

start_time = time.time()
print("Deleting holdings and unit trust price files ...")

if os.path.isfile(parN_nm):
    os.remove(parN_nm)

if os.path.isfile(utP_nm):
    os.remove(utP_nm)

print(
    f"{timediff(start_time, time.time())} deleting holdings and unit trust price files",
    "\n",
)
print(f"\n\n{timediff(start_time_pgf_compiling, time.time())} roundtrip time", "\n")
# # Open the pgf sheet for review

# start_time = time.time()
# print('Opening the Summary sheet for review ...')

# import win32com.client as win32 # library to convert xls to xlsx
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# excel.Workbooks.Open(os.path.join(pth, f'{rptDate.strftime("%Y%m%d")} PGF Share Class Hedges.xlsx'))
# excel.Quit()

# print(f'Opening the Summary sheet for review completed: {timediff(start_time, time.time())}', '\n')
# print(f'Roundtrip time for compiling the PGF hedge class sheet: {timediff(start_time_pgf_compiling, time.time())}', '\n')

print("\n\n#####################################")
print("#                                   #")
print("#     END 2/2 pgf_compiling.py X    #")
print("#                                   #")
print("#####################################\n\n")
