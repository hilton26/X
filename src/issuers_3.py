#!/usr/bin/env python
# coding: utf-8

# # Step 3: Changing % Column and Splitting out Reports

# libraries, libraries!

print("Importing libraries and setting paths for issuers_3 ...")
import time

start_time_issuers_3 = time.time()
start_time = time.time()

# general libraries
from datetime import datetime
import shutil  # to copy issuers_2 and issuers_3
import pandas as pd  # for dataframes
import numpy as np  # for np.NaN
import re  # for regex
from re import search  # for regex
import os
import math
from tqdm import tqdm
from constants import (
    pthPy,
    pthReports,
    pthTest,
    pth_struct,
    pthSttlmnt,
    yll,
    iss_2,
    iss_3,
)
from utilities import timediff

print(f" {timediff(start_time, time.time())} importing libraries and setting paths\n")

# list the funds
df_check = pd.read_excel(pthPy, sheet_name="arc", usecols="V", nrows=7)
url = df_check.iloc[6, 0].replace('"', "")
rpt = df_check.iloc[2, 0]
if url == url:  # ... if so, use the py_reports url ...
    py_input = pd.read_excel(url, engine="openpyxl")
    if isinstance(list(py_input)[9], str):
        rptDate = datetime.strptime(list(py_input)[9], "%d %b %Y")
    else:
        rptDate = list(py_input)[9]
else:  # ... prompt for a valid url
    print('Please provide a valid URL in cell "V8" of the "arc" tab')

fnds = py_input["Entity Name"].unique()
funds = pd.Series(fnds)
s = "" if len(fnds) == 1 else "s"
cs1 = df_check.iloc[2, 0]
print(type(funds), funds)

# (2) get report date - https://stackoverflow.com/questions/43544514/pandas-read-specific-excel-cell-value-into-a-variable
start_time = time.time()
print("", f"Getting report date and fund names ...")

# check if a url was given in the py_reports file and then ...
df_check = pd.read_excel(pthPy, sheet_name="arc", usecols="V", nrows=7)
url = df_check.iloc[6, 0].replace('"', "")
rpt = df_check.iloc[2, 0]
if url == url:  # ... if so, use the py_reports url ...
    py_input = pd.read_excel(url, engine="openpyxl")
    if isinstance(list(py_input)[9], str):
        rptDate = datetime.strptime(list(py_input)[9], "%d %b %Y")
    else:
        rptDate = list(py_input)[9]
else:  # ... prompt for a valid url
    print('Please provide a valid URL in cell "V8" of the "arc" tab')

fnds = py_input["Entity Name"].unique()
funds = pd.Series(fnds)
s = "" if len(fnds) == 1 else "s"
cs1 = df_check.iloc[2, 0]

suffix = f" for {len(funds)} funds" if cs1 == "CS1 format only" else ""
print(
    f'   Report date is {rptDate.strftime("%d %B %Y")} and report requested is "{cs1}"{suffix}'
)
print(f"   {len(funds)} fund{s}: \n     {(', ').join(list(funds))}")

print(
    f"\n  {timediff(start_time, time.time())} getting report date and fund names completed: {timediff(start_time, time.time())} \n"
)

# get accrual, margin, and settlement account inputs
start_time = time.time()
print("Getting accrual, margin, and settlement account inputs ...")

accr = pd.read_excel(pth_struct, sheet_name="accr", usecols=["accruals"]).dropna()
marg = pd.read_excel(pth_struct, sheet_name="accr", usecols=["margins"]).dropna()
sttlmnt = pd.read_excel(
    pthSttlmnt, sheet_name="Sttlmnt", usecols=["Fund", "Custodian", "SAFEX"]
).dropna(subset=["Fund"])

accr_list = [x for e in accr.values.tolist() for x in e]
margin_list = [x for e in marg.values.tolist() for x in e]

print(
    f" {timediff(start_time, time.time())} getting accrual, margin, and settlement account inputs\n"
)

# look up security issuers and classifications from merge merge the issuers_2 classified securities and input funds
start_time = time.time()
print("", f"Looking up security classifications ...")

# get merged (issuers_2.xlsx created in issuers_2.ipynb)
issuers_2 = pd.read_excel(
    iss_2,
    sheet_name="all",
    usecols=[
        "Primary Asset ID",
        "Issuer",
        "Reg 28 Classification",
        "Reg 30 Classification",
        "Infrastructure",
        "Derivative",
        "Counterparty",
        "margin",
    ],
)

#  get all_yall (example created in issuers_1.ipynb)
yall = pd.read_excel(yll, sheet_name="all").drop(
    axis="columns", columns=["Reg28 Classification"]
)

df = yall.merge(
    issuers_2, left_on="Primary Asset ID", right_on="Primary Asset ID", how="left"
)

f_codes = df[
    "Entity Name"
].unique()  # fund codes needed later to create individual fund reports

print(f" {timediff(start_time, time.time())} looking up security classifications\n")

# confirm all funds have a settlement bank account by merging df and sttlmnt
start_time = time.time()
print("Confirming all funds have a settlement account, else opening fund_codes.xlsx")

no_sttlmnt = df.merge(sttlmnt, left_on="Entity Name", right_on="Fund", how="left")
fnds_0 = no_sttlmnt[no_sttlmnt["Custodian"].isna()]["Entity Name"].unique()
s1 = "s" if len(fnds_0) != 1 else ""
print(
    f"  {len(fnds_0)} fund{s1} with no settlement account: \n    {(', ').join(list(fnds_0))}"
)

print(
    f"{timediff(start_time, time.time())} confirming all funds have a settlement account, else opening fund_codes.xlsx\n"
)

# set up classification function for remaining securities
start_time = time.time()
print("Setting up function to classify the remaining fund-specific securities ...")


# function to identify NaNs in floats
def is_nan(
    value,
):  # https://stackoverflow.com/questions/70272742/how-to-check-for-floatnan-in-python
    try:
        return math.isnan(float(value))
    except ValueError:
        return False


# function to classify remaining fund_specific securities
def classify2(
    row,
):  # https://stackoverflow.com/questions/23586510/return-multiple-columns-from-pandas-apply
    row["Issuer"] = (
        sttlmnt.loc[sttlmnt["Fund"] == row["Entity Name"]].iat[0, 2]
        if (row["Primary Asset ID"] in margin_list)
        else sttlmnt.loc[sttlmnt["Fund"] == row["Entity Name"]].iat[0, 1]
        if ((row["Primary Asset ID"] in accr_list) | (row["Investment Type"] == "SYTH"))
        else (sttlmnt.loc[sttlmnt["Fund"] == row["Entity Name"]].iat[0, 2])
        if (row["Primary Asset ID"] in margin_list)
        else row["Issuer"]
    )
    # if Primary Asset ID is an accrual or 'SYTH', assign the fund's settlement bank as issuer
    # else, if Primary Asset ID is a margin account, assign the fund's SAFEX counterparty as issuer
    # else reassign the issuer previously assigned

    row["Reg 28 Classification"] = (
        "1.1(c)"
        if (row["Primary Asset ID"] == "SAFEX")
        or ((row["margin"] == 1) and (row["CCY"] == "ZAR"))
        else "1.2(a)"
        if ((row["Investment Type"] == "SYTH") and (row["CCY"] != "ZAR"))
        else "1.2(c)"
        if ((row["margin"] == 1) and (row["CCY"] != "ZAR"))
        else "1.1(a)"
        if (
            (row["Investment Type"] == "SYTH")
            or (is_nan(row["Reg 28 Classification"]) and row["CCY"] == "ZAR")
        )
        else "1.2(a)"
        if ((row["CCY"] != "ZAR") and (row["Investment Type"] == "CASH"))
        else row["Reg 28 Classification"]
    )
    row["Reg 30 Classification"] = (
        "7(a)(ii)"
        if (
            (row["Primary Asset ID"] == "SAFEX")
            | (row["Primary Asset ID"] == "VARMARG")
            | (row["Investment Type"] == "FWD")
        )
        else "1(b)"
        if (row["CCY"] != "ZAR")
        else "1(a)(i)"
        if ((row["Investment Type"] == "SYTH") or is_nan(row["Reg 30 Classification"]))
        else row["Reg 30 Classification"]
    )
    return row


print(
    f" {timediff(start_time, time.time())} setting up function to classify the remaining fund-specific securities\n"
)

# classify the remaining fund-specific securities
start_time = time.time()
print(" Classifying the remaining fund-specific securities ...")

try:
    df = df.apply(classify2, axis=1)
except Exception as ex:
    print(
        f'  Exception: add {(", ").join(list(fnds_0))} to "Sttlmnt" tab in \n   {pthSttlmnt}'
    )

    # access Excel application to open the sheet
    import win32com.client as win32  # library to convert xls to xlsx

    excel_app = win32.gencache.EnsureDispatch(
        "Excel.Application"
    )  # to open excel application
    if len(fnds_0) != 0:
        excel_app.Workbooks.Open(pthSttlmnt)

    excel_app.Quit

print(
    f"  {timediff(start_time, time.time())} classifying the remaining fund-specific securities\n"
)

# write the dataframe to review it as a workbook
start_time = time.time()
print("Writing the dataframe to a file for review ...")

# identifying unassigned securities
noR28 = df[
    (df["Reg 28 Classification"] == "--- r28 ---")
    | df["Reg 28 Classification"].isnull()
]
noR30 = df[
    (df["Reg 30 Classification"] == "--- r30 ---")
    | df["Reg 30 Classification"].isnull()
]
noIssuer = df[(df["Issuer"] == "-xxx-") | df["Issuer"].isnull()]

# writing issuers_3.xlsx for review
writer = pd.ExcelWriter(iss_3, engine="xlsxwriter")  #!pip install xlsxwriter
df.to_excel(writer, index=False, sheet_name=f"all ({len(df)})")
noR28.to_excel(
    writer, index=False, sheet_name=f"no_r28 ({len(noR28)})"
)  # unique securities
noR30.to_excel(
    writer, index=False, sheet_name=f"no_r30 ({len(noR30)})"
)  # unique securities
noIssuer.to_excel(
    writer, index=False, sheet_name=f"no_issr ({len(noIssuer)})"
)  # unique securities
writer.close()

print(
    "  Issuers not assigned   :",
    len(noIssuer),
    "\n",
    " Reg 28 not classified  :",
    len(noR28),
    "\n",
    " Reg 30 not classified  :",
    len(noR30),
    "\n",
)

print(
    f" {timediff(start_time, time.time())} writing the dataframe to a file for review\n"
)

# assign report column headings
start_time = time.time()
print("", "Saving the reports else opening issuers_3.xlsx for review ...")

if len(noR28) + len(noR30) + len(noIssuer) > 0:
    excel.Workbooks.Open(iss_3)
else:
    R30 = df[
        [
            "Entity Name",
            "Investment Type",
            "i Issue Name",
            "Primary Asset ID",
            "CCY",
            "Reg 30 Classification",
            "End Market Value",
            "Percentage of Market Value",
            "Closing Exposure PA",
            "Issuer",
        ]
    ]
    R28 = df[
        [
            "Entity Name",
            "Investment Type",
            "i Issue Name",
            "Primary Asset ID",
            "CCY",
            "Reg 28 Classification",
            "End Market Value",
            "Percentage of Market Value",
            "Closing Exposure PA",
            "Issuer",
            "Infrastructure",
        ]
    ]
    R28CS1 = df[
        [
            "Entity Name",
            "Investment Type",
            "i Issue Name",
            "Primary Asset ID",
            "CCY",
            "Reg 28 Classification",
            "End Market Value",
            "Percentage of Market Value",
            "Closing Exposure PA",
            "Issuer",
            "Infrastructure",
            "Derivative",
            "Counterparty",
        ]
    ]

    if cs1 == "Reg 28 and Reg 30 only":
        # write the Reg 28 report to an xlsx file
        for f_code in tqdm(f_codes):
            fname = os.path.join(
                pthTest, f"{f_code} Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
            )
            R30[R30["Entity Name"] == f_code].to_excel(
                fname,
                sheet_name=f"{f_code} Reg30 {rptDate.strftime('%d%b%Y')}",
                index=False,
            )
        # write the Reg 30 report to an xlsx file
        for f_code in tqdm(f_codes):
            fname = os.path.join(
                pthTest, f"{f_code} Reg28 {rptDate.strftime('%d%b%Y')}.xlsx"
            )
            R28[R28["Entity Name"] == f_code].to_excel(
                fname,
                sheet_name=f"{f_code} Reg28 {rptDate.strftime('%d%b%Y')}",
                index=False,
            )
    else:
        # write the bulk Reg 28 CS1 report to an xlsx file
        # cs1_fname = os.path.join(pthTest, f'CS1 PARN holdings ({len(funds)}) {rptDate.strftime("%d%b%Y")}.xlsx')
        # cs1_rpt_name = os.path.join(pthTest, f'Reg28 CS1 reports ({len(f_codes)}) {rptDate.strftime("%d%b%Y")}.xlsx')
        cs1_rpt_name = os.path.join(
            pthTest, f"Reg28 CS1 reports {rptDate.strftime('%d%b%Y')}.xlsx"
        )
        writer = pd.ExcelWriter(
            cs1_rpt_name, engine="xlsxwriter"
        )  # instantiate a sheet writer
        R28CS1.to_excel(
            writer, index=False, sheet_name="CS1_All"
        )  # write the NAV sheet
        writer.close()  # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets
        print(" ", cs1_rpt_name)
        funds_not_classified = set(f_codes) ^ set(funds)
        print(
            " ",
            f"{len(funds_not_classified)} funds not classified with issuers_1.ipynb as at {rptDate.strftime('%d %B %Y')}:",
            "\n",
            f"   {(',').join(funds_not_classified)}",
        )

        # # write each Reg 28 CS1 report to an xlsx file
        # for f_code in tqdm(f_codes):
        #     fname = os.path.join(
        #         pthTest, f"{f_code} Reg28 CS1 {rptDate.strftime('%d%b%Y')}.xlsx"
        #     )
        #     R28CS1[R28CS1["Entity Name"] == f_code].to_excel(
        #         fname,
        #         sheet_name=f"{f_code} Reg28 CS1 {rptDate.strftime('%d%b%Y')}",
        #         index=False,
        #     )

    # open the reporting folder and the Reg_Tests folder
    os.startfile(os.path.realpath(pthReports))
    os.startfile(os.path.realpath(pthTest))

print(
    "",
    f"{timediff(start_time, time.time())} saving the reports else opening issuers_3.xlsx for review",
    "\n",
)

# save issuers_2 and issuers_3 as workbooks with reporting date appended
# https://stackoverflow.com/questions/123198/how-to-copy-files
start_time = time.time()
print("Saving issuers_2 and issuers_3 with month-end date ...")

# reset issuers_1 input sheet
import xlwings as xw

xw.Book(pthPy).sheets("arc").range("V4").value = "Reg 28 and Reg 30 only"
xw.Book(pthPy).save()
xw.Book(pthPy).close()

# get list of month-end 17 portfolio codes
rest = pd.read_excel(pthPy, sheet_name="funds", usecols=["Month-end 17"]).dropna()
me17 = rest["Month-end 17"].tolist()

# https://stackoverflow.com/questions/740287/how-to-check-if-one-of-the-following-items-is-in-a-list
# if the month-end 17 fund codes are in the list of funds being classified, save issuers_2 and issuers_3 files with reporting date appended
if cs1 == "Reg 28 and Reg 30 only":
    if len([i for i in f_codes if i in me17]) == len(me17) or len(f_codes) > 100:
        for k in range(2, 4):
            shutil.copy2(
                os.path.join(pthTest, f"issuers_{k}.xlsx"),
                os.path.join(pthTest, f"issuers_{k}_{rptDate.strftime('%d%b%Y')}.xlsx"),
            )
            # copy2() to preserve timestamp

print(
    f' {timediff(start_time, time.time())} saving issuers_2 and issuers_3 with month-end date unless "CS1" selected\n'
)

print(
    f" {timediff(start_time_issuers_3, time.time())} ISSUERS_3 COMPLETED\n===========================\n"
)
