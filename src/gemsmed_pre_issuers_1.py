#!/usr/bin/env python
# coding: utf-8

# ## Compile GEMSMED weekly Reg 30 report


print("\n\n###########################################")
print("#                                         #")
print("#     START 1/2 gemsmed_pre_issuers_1  X  #")
print("#                                         #")
print("###########################################\n\n")


# import time
import time

start_time = time.time()
start_time_gemsmed_pre = start_time

# libraries, libraries!
from datetime import datetime
import pandas as pd
import numpy as np
import os
from pathlib import Path
from constants import pthPy, pth_dl, pthLOCAL
from utilities import prior_working_day, timediff, osprey, r_classifier

print("Importing libraries to get fund holdings ...")
# " ...any time you see a loop somewhere in your code in you can simply wrap it in either tdqm() or notebook.tqdm() in Jupyter"

print(
    f" {timediff(start_time, time.time())} importing libraries to get fund holdings",
    "\n",
)

# get report parameters

start_time = time.time()
print("Getting the report parameters ...\n")

df = pd.read_excel(pthPy, sheet_name="arc", usecols="K").dropna()  # portfolio codes
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="L", nrows=2)  # reporting dates
k = df1.iloc[1, 0]
rptDate = (
    k
    if isinstance(k, datetime) and not pd.isna(k)
    else prior_working_day(datetime.today())
)
funds = (",").join(df.iloc[:, 0])
suffix = "csv"
s = "" if len(df.iloc[:, 0]) == 1 else "s"
print(
    f" {rptDate.strftime('%a %d %b %Y')} for {len(df.iloc[:, 0])} fund{s}:\n  {funds}"
)
name = funds if len(df.iloc[:, 0]) == 1 else "gmunu"

# derive file names
holds_nm = os.path.join(pth_dl, f"R28I {name}(1) {rptDate.strftime('%d%b%Y')}.{suffix}")
dervs_nm = os.path.join(pth_dl, f"DERV {name}(1) {rptDate.strftime('%d%b%Y')}.{suffix}")
print(f"\nFiles expected:\n {holds_nm}\n {dervs_nm}")

print(f"\n{timediff(start_time, time.time())} getting the report parameters", "\n")

# get and save holdings and then derivative data

start_time = time.time()
print("Extracting and saving the fund holdings and derivatives ...\n")

# get holdings in PAR-N format
if os.path.exists(holds_nm):
    print(f"{holds_nm} exists")
else:
    osprey("r28i", funds, rptDate, rptDate, name, suffix)
    if os.path.exists(holds_nm):
        print(f"{holds_nm} downloaded")
    else:
        print(f"{holds_nm} did not download")
        pass

# get derivative data
if os.path.exists(dervs_nm):
    print(f"{dervs_nm} exists")
else:
    osprey("derv", funds, rptDate, rptDate, name, suffix)
    if os.path.exists(dervs_nm):
        print(f"{dervs_nm} downloaded")
    else:
        print(f"{dervs_nm} did not download")
        pass

print(
    f"\n {timediff(start_time, time.time())} extracting and saving the fund holdings and derivatives",
    "\n",
)

# prepare the combined holdings and derivatives dataframe

start_time = time.time()
print(f"Preparing the combined holdings and derivatives dataframes ...\n")

# make dataframes out of the holdings and derivative data files
holds = pd.read_csv(os.path.join(holds_nm))
dervs = pd.read_csv(os.path.join(dervs_nm))

# merge holdings and derivative values
a = holds.merge(dervs, how="left", on="Primary Asset ID", suffixes=("", "_2"))
# a = pd.concat([holds, dervs], ignore_index=True)

# convert numerical columns from str to float
start_time_conv = time.time()
print(" Converting numerical columns from type string to type float ...")

# convert derivative columns from str to float
# https://stackoverflow.com/questions/55557004/getting-attributeerror-float-object-has-no-attribute-replace-error-while
heads = [
    "End Market Value",
    "Percentage of Market Value",
    "Closing Exposure PA",
    "Effective Exposure",
]
for head in heads:
    a[head] = [str(x).replace(",", "").replace("-", "-") for x in a[head]]
    a[head] = a[head].astype(float)
print(
    f"  {timediff(start_time, time.time())} converting numerical columns from type string to type float"
)
print(
    f"\n{timediff(start_time, time.time())} preparing the combined holdings and derivatives dataframes",
    "\n",
)

# save the dataframe as a file

start_time = time.time()
print(f"Saving the file ...\n")

# insert effective exposure values in closing exposure column
# https://datascience.stackexchange.com/questions/56668/pandas-change-value-of-a-column-based-another-column-condition

a["Closing Exposure PA"] = np.where(
    a["Effective Exposure"].isnull(), a["End Market Value"], a["Effective Exposure"]
)

# make SYTH values negative
a["Closing Exposure PA"] = np.where(
    a["Investment Type"] == "SYTH", -a["Closing Exposure PA"], a["Closing Exposure PA"]
)

print(a)

# drop unneccesary columns
columns_to_drop = [
    "Entity Name_2",
    "i Issue Name_2",
    "Nominal Holding",
    "Delta",
    "Market Value",
    "Effective Exposure",
]
a.drop(columns_to_drop, axis=1, inplace=True)

# add an empty column with report date
a[f"{rptDate.strftime('%d %b %Y')}"] = ""

# drop rows with zero values
b = len(a)
a.drop(a[a["Closing Exposure PA"] == 0].index, inplace=True)
print("", f"{b - len(a)} zero rows dropped")
print(
    "", f"{a['End Market Value'].sum() - a['Closing Exposure PA'].sum()} NAV difference"
)

show = ["FT", "OP", "SYTH"]
a[a["Investment Type"].isin(show)]

# save as an excel file
a.to_excel(
    os.path.join(pthLOCAL, f"{name} Reg28 {rptDate.strftime('%d%b%Y')}.xlsx"),
    sheet_name="All",
    index=False,
)

print(f"Saving the file completed: {timediff(start_time, time.time())}")
print(
    "\n",
    os.path.join(pthLOCAL, f"{name} Reg28 {rptDate.strftime('%d%b%Y')}.xlsx"),
    "\n",
)

# update py_reports.xlsm classifier sheet with downloaded gemsmed_c file location

start_time = time.time()
print(f"Updating the classifier sheet with downloaded {name} file location ...\n")

import xlwings as xw

wb = xw.Book(pthPy)
# ws                   = wb.sheets('classifier')
# ws.range("L2").value = os.path.join(pthLOCAL, f'{name} Reg28 {rptDate.strftime("%d%b%Y")}.xlsx')
# ws.range("M1").value = 'Reg 28 and Reg 30 only' # alternative: 'CS1 format only'
ws = wb.sheets("arc")
ws.range("V4").value = "Reg 28 and Reg 30 only"  # alternative: 'CS1 format only'
ws.range("V8").value = os.path.join(
    pthLOCAL, f"{name} Reg28 {rptDate.strftime('%d%b%Y')}.xlsx"
)
wb.save()
wb.close()

# r_classifier('r28', os.path.join(pthLOCAL, f'{name} Reg28 {rptDate.strftime("%d%b%Y")}.xlsx'), rptDate):

print(
    f"Updating the classifier sheet with downloaded {name} file location completed: {timediff(start_time, time.time())}",
    "\n",
)
print(
    os.path.join(
        Path.home(),
        "Documents",
        "DervFiles",
        f"{name} Reg28 {rptDate.strftime('%d%b%Y')}.xlsx",
    ),
    "\n",
)
print(
    f"\n{timediff(start_time_gemsmed_pre, time.time())} roundtrip time to get {name} holdings",
    "\n",
)

print("\n\n###############################################")
print("#                                             #")
print("#        END 1/2 gemsmed_pre_issuers_1  X     #")
print("#                                             #")
print("###############################################\n\n")
