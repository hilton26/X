#!/usr/bin/env python
# coding: utf-8

# # Step 2: Categorisation of securities
#
# ### Merge Eagle and structures dataframes and then assign
# Reg 28 and Reg 30 categories including "Infrastructure"

# libraries, libraries!

print("Importing libraries and setting paths for issuers_2 ...")
import time

start_time_issuers_2 = time.time()
start_time = time.time()

# general libraries
from datetime import datetime
import pandas as pd  # for dataframes
import numpy as np  # for np.NaN
import re  # for regex
from re import search  # for regex
from constants import pthPy, pthTest, pth_struct, iss_1, iss_2, mergd
from utilities import timediff

# utilise openpyxl tools to add excel features to results sheet
import openpyxl as px  # for adding sort filters to the excel sheet
from openpyxl.cell import Cell  # to format cells
from openpyxl.styles import (
    Alignment,
    Color,
    PatternFill,
    Font,
    Border,
)  # to format cells

#     # access Excel application to open results sheet
# import  win32com.client as win32                                          # library to convert xls to xlsx
# excel = win32.gencache.EnsureDispatch('Excel.Application')                # to open excel application

print(
    f" {timediff(start_time, time.time())} importing libraries and setting paths for issuers_2\n"
)

# create functions
start_time = time.time()
print("Setting up functions ...")

# utility function to check if a string contains a list element
# https://www.geeksforgeeks.org/python-test-if-string-contains-element-from-list/


def res(t_list, t_string):
    return bool([ele for ele in t_list if (ele in t_string)])


# define an inf(rastructure) function
# https://stackoverflow.com/questions/26886653/create-new-column-based-on-values-from-
# other-columns-apply-a-function-of-multi
def inf(row):
    if row["Infra"] == "i" and row["GovGuar"] != 1:
        return "11(b)"


# derivative asset class function
def derv_ft_op(row):
    if res(["OP", "FT"], row["Investment Type"]) and res(
        ["AU", "XU"], str(row["Commodity"]).upper()
    ):
        return "r"  # commodity
    if res(["OP", "FT"], row["Investment Type"]) and res(
        ["JADM", "ZAUS", "CURR", "ZAR", "USD/ZAR", "JPN YEN"],
        row["i Issue Name"].upper(),
    ):
        return "c"  # currency
    if res(["OP", "FT"], row["Investment Type"]) and res(
        ["ALSI", "FTSE", "UKX", "SPX", "HSCEI", "SPYQ", "SX5E", "NKY", "CSI", "SHSN"],
        row["i Issue Name"].upper(),
    ):
        return "e"  # equity
    if (row["Investment Type"] == "OP") and (
        (row["Derivative"] in "Structured Note") or (row["Derivative"] in "Linked Note")
    ):
        return "e"  # equity
    elif row["Investment Type"] == "FT" and row["Property"] == "P":
        return "p"  # property
    elif row["Investment Type"] == "DERV" and "TRS" in row["Primary Asset ID"].upper():
        return "e"  # equity
    elif row["Investment Type"] == "DERV" and (
        "CPI" in row["i Issue Name"].upper() or row["Term"] >= 396
    ):
        return "d"  # debt
    elif (row["Investment Type"] == "DERV") or (row["Investment Type"] == "FWD"):
        return "c"  # currency
    elif row["Investment Type"] == "FT" and (
        row["GPL"] == "G" or res(["BOND", "CBT", "NOTE"], row["i Issue Name"].upper())
    ):
        return "d"  # debt
    elif res(["OP", "FT"], row["Investment Type"]):
        return "e"  # equity # default to asset type 'equity' if none of the filters above apply


print(f" {timediff(start_time, time.time())} setting up functions\n")

# read in unique security dataframe from earlier process and saved as 'issuers_1.xlsx' to use as input
start_time = time.time()
print(f"Reading in unique securities which were assigned issuers earlier ...")

uniques_cols = [
    "Investment Type",
    "i Issue Name",
    "Primary Asset ID",
    "CCY",
    "Reg28 Classification",
    "End Market Value",
    "Percentage of Market Value",
    "Closing Exposure PA",
    "Issuer",
    "CLN",
    "FRN",
    "ILB",
    "BESA",
    "repo",
    "margin",
    "Date",
    "MedCirc",
    "GovGuar",
    "Term",
    "Derivative",
    "Counterparty",
]
# excl 'Entity Name', 'MedCirc062022', 'MedCirc122023'

uniques = pd.read_excel(iss_1, sheet_name="uniques", usecols=uniques_cols)

print(
    f" {timediff(start_time, time.time())} reading in unique securities which were assigned issuers earlier\n"
)

# import the structures file with security attributes
start_time = time.time()
print(f"Getting security attribute columns ...")

# specify columns to include in 'reg' dataframe
reg_cols = [
    "Issuer Name",
    "vs structures",
    "MCap",
    "Ticker",
    "Deb",
    "Exchange",
    "Foreign",
    "GPL",
    "Fund",
    "Bank",
    "Debt",
    "Equity",
    "Property",
    "Commodity",
    "Partg_emplr",
    "HL",
    "HF",
    "PEF",
    "fnd_typ",
    "Infra",
    "DI900",
]

# get the 'reg' sheet with market caps
reg = pd.read_excel(pth_struct, sheet_name="reg", usecols=reg_cols)

print(f" {timediff(start_time, time.time())} getting security attribute columns\n")

merged = uniques.merge(reg, left_on="Issuer", right_on="Issuer Name", how="left")
merged

# with all non-accrual and non-FWD instruments assigned an issuer, look up instrument attributes by a merge with the funds
start_time = time.time()
print(
    f'Merging issuer-dentified securities with security attributes from "structures.xlsm" ...'
)

# https://www.youtube.com/watch?v=AHS925L8JVk&t=9sb
merged = uniques.merge(reg, left_on="Issuer", right_on="Issuer Name", how="left")

# save the merged workbook to mergd
merged.to_excel(mergd, sheet_name="merged", index=False)

print(
    f' {timediff(start_time, time.time())} merging issuer-dentified securities with security attributes from "structures.xlsm"\n'
)

# add an inf(rastructure) column to the dataframe called 'merged'
start_time = time.time()
print(
    f"Adding an infrastructure column to the uniques + structures merged dataframe ..."
)

merged["Infrastructure"] = merged.apply(inf, axis=1)

print(
    f" {timediff(start_time, time.time())} adding an infrastructure column to the uniques + structures merged dataframe\n"
)

# Identify equities without market caps # https://www.geeksforgeeks.org/filter-pandas-dataframe-with-multiple-conditions/
start_time = time.time()
print(f"Identifying equities and securities without market caps ...\n")

eq_no_MCap = merged.loc[
    (merged["MCap"].isnull().values) & (merged["Investment Type"] == "EQ")
].drop_duplicates(subset="Primary Asset ID", keep="first")["Primary Asset ID"]

no_MCap = merged.loc[(merged["MCap"].isnull().values)].drop_duplicates(
    subset="Primary Asset ID", keep="first"
)["Primary Asset ID"]
s = "y" if len(eq_no_MCap) == 1 else "ies"
print(
    f" {len(eq_no_MCap)} equit{s} without market caps out of {len(no_MCap)} total equities: \n   {(', ').join(eq_no_MCap)} \n"
)

print(
    f"{timediff(start_time, time.time())} identifying equities and securities without market caps\n"
)

# (3) identify derivative asset classes
start_time = time.time()
print("Identifying derivative asset classes ...")

merged["dX"] = merged.apply(lambda row: derv_ft_op(row), axis=1)
der_no_dX = merged[
    merged["Investment Type"].isin(["DERV", "FT", "OP"]) & merged["dX"].isnull()
]
s = "" if len(der_no_dX) == 1 else "s"
print(
    f" {len(der_no_dX)} unclassified derivative instrument{s}: \n   {(', ').join(list(der_no_dX.iloc[:, 2]))}"
)
# print(f' {len(merged["dX"].unique())} derivative ("OP", "FT", "DERV") securities')

print(f"{timediff(start_time, time.time())} identifying derivative asset classes\n")

# ### TEST ###
# overrides  = pd.read_excel(pth_struct, sheet_name = 'overrides')  # security data overrides
# for instr in overrides['InstrCode']: #https://www.geeksforgeeks.org/different-ways-to-iterate-over-rows-in-pandas-dataframe/
#     #instr = overrides['InstrCode'][i]
#     print(i, instr, overrides['Date'][i], isinstance(overrides['Date'][i], datetime.date), overrides['Date'][i] == "")

# insert overridden instrument codes
start_time = time.time()
print("Inserting overridden instrument codes ...")

overrides = pd.read_excel(pth_struct, sheet_name="overrides")  # security data overrides
for i in range(
    len(overrides)
):  # https://www.geeksforgeeks.org/different-ways-to-iterate-over-rows-in-pandas-dataframe/
    instr = overrides["InstrCode"][i]

    if (
        merged["Primary Asset ID"].eq(overrides.loc[i, "InstrCode"]).any()
    ):  # if a fund instrument ID matches one in the overrides file ...
        # https://www.statology.org/pandas-check-if-column-contains-string/
        # https://www.geeksforgeeks.org/python-data-types/ they are str, int, float, complex, list, tuple

        # print(i, instr)
        # print(i, instr, overrides['Date'][i], isinstance(overrides['Date'][i], datetime.date), overrides['Date'][i] == "")

        # test for 'BESA'
        if isinstance(
            overrides["BESA"][i], str
        ):  # if the overrides contain 'B' for BESA ...
            merged.at[merged[merged["Primary Asset ID"] == instr].index[0], "BESA"] = (
                overrides["BESA"][i]
            )
            # ... overide the value in the BESA column with 'B', per the overrides file
            # https://datagy.io/pandas-get-row-number/
            # https://stackoverflow.com/questions/13842088/set-value-for-particular-cell-in-pandas-dataframe-using-index

        # test for 'Issuer'
        # if isinstance(overrides['Issuer'][i], str):
        # merged.at[merged[merged['Primary Asset ID'] == instr].index[0], 'Issuer'  ] = overrides['Issuer'  ][i]

        # test for 'Date'
        # if isinstance(overrides['Date'][i], datetime.date):
        if isinstance(overrides["Date"][i], datetime):
            merged.at[merged[merged["Primary Asset ID"] == instr].index[0], "Date"] = (
                overrides["Date"][i]
            )

        # test for 'Term'
        if isinstance(overrides["Term"][i], float):
            merged.at[merged[merged["Primary Asset ID"] == instr].index[0], "Term"] = (
                overrides["Term"][i]
            )

        # # test for 'MCap'
        # if isinstance(overrides['MCap'][i], float):
        #     merged.at[merged[merged['Primary Asset ID'] == instr].index[0], 'MCap'    ] = overrides['MCap'    ][i]

        # # test for 'Property'
        # if isinstance(overrides['Property'][i], str):
        #     merged.at[merged[merged['Primary Asset ID'] == instr].index[0], 'Property'] = overrides['Property'][i]

        # # test for 'Government guarantee'
        # if isinstance(overrides['GovGuar'][i], str):
        #     merged.at[merged[merged['Primary Asset ID'] == instr].index[0], 'GovGuar' ] = overrides['GovGuar' ][i]

print(f" {timediff(start_time, time.time())} inserting overridden instrument codes\n")

# function that creates a new column based on the values of other columns
# Python: Check if String Contains Substring https://stackabuse.com/python-check-if-string-contains-substring/

start_time = time.time()
print("Defining the Reg 28 classification function ...")


def classify_Reg28(row):
    # 1.1(a) Cash with SARB-registered bank
    if (
        (
            (
                res(["CASH", "FWD"], row["Investment Type"])
                or (row["dX"] == "c")
                or (row["dX"] == "d")
            )
            and (row["margin"] != 1)
            and (row["Primary Asset ID"] != "SAFEX")
            and (row["Bank"] == "s")
            and (row["CCY"] == "ZAR")
        )
        or (row["Investment Type"] == "SYTH")
        or (
            (row["Derivative"] == "Swap" or row["Derivative"] == "Currency Forward")
            and (row["CCY"] == "ZAR")
        )
    ):
        return "1.1(a)"

    # # 1.1(b) Money market instruments with SARB registered bank
    # elif res(['ST','FI'], row['Investment Type']) and (row['CCY'] == 'ZAR') and (row['Bank'] == 's') and (row['margin'] != 1)  or \
    # ((row['Fund'] == 'f') and (row['fnd_typ'] == 'c') and (row['CCY'] == 'ZAR')):
    #     return '1.1(b)'

    # 1.1(b) Money market instruments with SARB registered bank
    elif (
        (
            res(["ST", "FI"], row["Investment Type"])
            and (row["CCY"] == "ZAR")
            and (row["Bank"] == "s")
            and (row["margin"] != 1)
            and (row["Term"] < 396)
        )
        or ((row["Fund"] == "f") and (row["fnd_typ"] == "c") and (row["CCY"] == "ZAR"))
        or (row["FRN"] == 1)
        or ("CALL" in row["Primary Asset ID"])
    ):
        return "1.1(b)"

    # 1.1(c) Positive margin account balance with SARB registered bank
    elif (
        ((row["margin"] == 1) and (row["Bank"] == "s") and (row["CCY"] == "ZAR"))
        or (row["Primary Asset ID"] == "SAFEX")
        or (row["Primary Asset ID"] == "VARMARG")
    ):
        return "1.1(c)"

    # 1.2(a) Cash with a foreign bank
    elif (
        (
            res(["ST", "CASH", "FWD", "SYTH"], row["Investment Type"])
            and row["Bank"] == "b"
        )
        or (res(["OP", "FT"], row["Investment Type"]) and row["dX"] == "c")
        or (row["Primary Asset ID"] == "0649317")
        or ((row["Derivative"] == "Swap") or (row["Derivative"] == "Currency Forward"))
        and (row["CCY"] != "ZAR")
        and (row["Investment Type"] != "SYTH")
    ):
        return "1.2(a)"

    # 1.2(b) Cash with an African bank
    elif res(["ST", "CASH"], row["Investment Type"]) and row["Bank"] == "a":
        return "1.2(b)"

    # 1.2(c) Money market instruments with foreign bank
    # elif ((row['margin'] == 1) and (row['Bank'] != 's') and (row['CCY'] != 'ZAR')) or \
    elif (
        ((row["margin"] == 1) and (row["CCY"] != "ZAR"))
        or (
            row["Investment Type"] == "ST"
            or row["dX"] == "c"
            or ((row["Investment Type"] == "FI") and (row["Term"] < 396))
        )
        and row["Bank"] == "b"
        or ((row["Issuer"] == "JSE Yield-X") and (row["dX"] == "c"))
        or ((row["Fund"] == "f") and (row["fnd_typ"] == "c") and (row["CCY"] != "ZAR"))
    ):
        return "1.2(c)"

    # 2.1(a) SA government debt, ZAR-denominated
    elif (
        (row["Issuer"] == "Republic of South Africa" or row["GovGuar"] == 1)
        and (row["CCY"] == "ZAR")
        and (row["Investment Type"] != "SYTH")
    ):
        return "2.1(a)"

    # 2.1(b) Foreign government debt and foreign government guaranteed debt
    elif (
        (row["Issuer"] != "Republic of South Africa")
        and (row["GPL"] == "G")
        and (row["Investment Type"] != "SYTH")
    ):
        return "2.1(b)"

    # 2.1(c)(i) Debt issued by SARB registered bank with MCap > R20bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "s"
        and row["Term"] >= 396
        and row["MCap"] > 20
    ):
        return "2.1(c)(i)"

    # 2.1(c)(ii) Debt issued by SARB registered bank with R2bn < MCap < R20bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "s"
        and row["Term"] >= 396
        and row["MCap"] < 20
        and row["MCap"] > 2
    ):
        return "2.1(c)(ii)"

    # 2.1(c)(iii) Debt issued by SARB registered bank with MCap < R2bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "s"
        and row["Term"] >= 396
        and row["MCap"] < 20
        and row["MCap"] < 2
    ):
        return "2.1(c)(iii)"

    # 2.1(c)(iv) Debt issued by SARB registered, unlisted bank
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "s"
        and row["Term"] >= 396
        and (row["MCap"] == 0 or pd.isna(row["MCap"]))
    ):
        return "2.1(c)(iv)"

    # 2.2(a) SA government debt, non-ZAR-denominated
    elif row["Issuer"] == "Republic of South Africa" and row["CCY"] != "ZAR":
        return "2.2(a)"

    # 2.2(c)(i) Debt issued by foreign bank with MCap > R20bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "b"
        and row["Term"] >= 396
        and row["MCap"] > 20
    ):
        return "2.2(c)(i)"

    # 2.2(c)(ii) Debt issued by foreign bank with R2bn < MCap < R20bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "b"
        and row["Term"] >= 396
        and row["MCap"] < 20
        and row["MCap"] > 2
    ):
        return "2.2(c)(ii)"

    # 2.2(c)(iii) Debt issued by foreign bank with MCap < R2bn
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "b"
        and row["Term"] >= 396
        and row["MCap"] < 20
        and row["MCap"] < 2
    ):
        return "2.2(c)(iii)"

    # 2.2(c)(iv) Debt issued by unlisted foreign bank
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and row["Bank"] == "b"
        and row["Term"] >= 396
        and pd.isna(row["MCap"])
    ):
        return "2.2(c)(iv)"

    # 2.1(d)(i) Listed debt issued by PFMA entities and by listed corporates
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or (row["dX"] == "d"))
        and ((row["GPL"] == "p") or (row["MCap"] > 0))
        and (row["BESA"] == "B")
        and (row["CCY"] == "ZAR")
        or (
            (row["Fund"] == "f")
            and (row["fnd_typ"] == "d")
            and (row["CCY"] == "ZAR")
            and (row["Infra"] != "i")
        )
    ):
        return "2.1(d)(i)"

    # 2.1(d)(ii) Unlisted debt issued by PFMA entities and by listed corporates
    elif (
        (res(["FI", "ST"], row["Investment Type"]) or row["dX"] == "d")
        and (row["GPL"] == "p" or row["MCap"] > 0)
        and row["BESA"] != "B"
        and row["CCY"] == "ZAR"
    ):
        return "2.1(d)(ii)"

    # 2.1(e)(i) Listed debt issued by unlisted non-SOE corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and row["GPL"] != "p"
        and pd.isna(row["MCap"])
        and row["BESA"] == "B"
        and row["CCY"] == "ZAR"
    ):
        return "2.1(e)(i)"

    # 2.1(e)(ii) Unlisted debt issued by unlisted non-SOE corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and (row["GPL"] != "p")
        and (pd.isna(row["MCap"]) or (row["MCap"] == 0))
        and (row["BESA"] != "B")
        and (row["CCY"] == "ZAR")
        or (
            (row["Fund"] == "f")
            and (row["fnd_typ"] == "d")
            and (row["CCY"] == "ZAR")
            and (row["Infra"] == "i")
        )
    ):
        return "2.1(e)(ii)"

    # 2.2(d)(i) Listed foreign debt issued by listed corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and ((row["GPL"] == "p") or (row["MCap"] > 0))
        and (row["CCY"] != "ZAR")
        or ((row["Fund"] == "f") and (row["fnd_typ"] == "d") and (row["CCY"] != "ZAR"))
    ):
        return "2.2(d)(i)"

    # 2.2(d)(ii) Unlisted foreign debt issued listed corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and (row["GPL"] == "p" or row["MCap"] > 0)
        and row["BESA"] != "B"
        and row["CCY"] != "ZAR"
    ):
        return "2.2(d)(ii)"

    # 2.2(e)(i) Listed foreign debt issued by unlisted corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and row["GPL"] != "p"
        and (pd.isna(row["MCap"]) or row["MCap"] == 0)
        and row["BESA"] != "B"
        and row["CCY"] != "ZAR"
    ):
        return "2.2(e)(i)"

    # 2.2(e)(ii) Unlisted foreign debt issued by unlisted corporates
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and row["GPL"] != "p"
        and pd.isna(row["MCap"])
        and row["BESA"] != "B"
        and row["CCY"] != "ZAR"
    ):
        return "2.2(e)(ii)"

    # 3.1(a)(i) Listed ordinary and preference non-property shares, market cap >= R20bn
    elif (
        ((row["Investment Type"] == "EQ" and row["MCap"] >= 20) or (row["dX"] == "e"))
        and (row["CCY"] == "ZAR")
        and (row["Property"] != "P")
        and (row["Investment Type"] != "SYTH")
        or (
            (row["Fund"] == "f")
            and (row["fnd_typ"] == "e")
            and (row["CCY"] == "ZAR")
            and (row["Investment Type"] != "SYTH")
        )
    ):
        return "3.1(a)(i)"

    # 3.1(a)(ii) Listed ordinary and preference non-property shares, market cap >= R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] >= 2
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
    ):
        return "3.1(a)(ii)"

    # 3.1(a)(iii) Listed ordinary and preference non-property shares, market cap < R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] < 2
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
    ):
        return "3.1(a)(iii)"

    # 3.1(b) Unlisted ordinary and preference non-property shares
    elif (
        row["Investment Type"] == "EQ"
        and pd.isna(row["MCap"])
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
    ):
        return "3.1(b)"

    # 3.2(a)(i) Listed ordinary and preference non-property shares, market cap >= R20bn
    elif (
        ((row["Investment Type"] == "EQ") and (row["MCap"] >= 20) or (row["dX"] == "e"))
        and (row["CCY"] != "ZAR")
        and (row["Property"] != "P")
        and (row["Investment Type"] != "SYTH")
    ) or (
        (row["Fund"] == "f")
        and (row["fnd_typ"] == "e")
        and (row["CCY"] != "ZAR")
        and (row["Investment Type"] != "SYTH")
    ):
        return "3.2(a)(i)"

    # 3.2(a)(ii) Listed foreign ordinary and preference non-property shares, market cap >= R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] >= 2
        and row["CCY"] != "ZAR"
        and row["Property"] != "P"
    ):
        return "3.2(a)(ii)"

    # 3.2(a)(iii) Listed foreign ordinary and preference non-property shares, market cap < R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] < 2
        and row["CCY"] != "ZAR"
        and row["Property"] != "P"
    ):
        return "3.2(a)(iii)"

    # 3.2(b) Unlisted foreign ordinary and preference non-property shares
    elif (
        row["Investment Type"] == "EQ"
        and pd.isna(row["MCap"])
        and row["CCY"] != "ZAR"
        and row["Property"] != "P"
    ):
        return "3.2(b)"

    # 4.1(a)(i) Listed property ordinary and preference shares, market cap >= R20bn
    elif (
        ((row["Investment Type"] == "EQ" and row["MCap"] >= 10) or row["dX"] == "p")
        and row["CCY"] == "ZAR"
        and row["Property"] == "P"
        or ((row["Fund"] == "f") and (row["fnd_typ"] == "p") and (row["CCY"] == "ZAR"))
    ):
        return "4.1(a)(i)"

    # 4.1(a)(ii) Listed ordinary and preference property shares, market cap >= R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] >= 3
        and row["CCY"] == "ZAR"
        and row["Property"] == "P"
    ):
        return "4.1(a)(ii)"

    # 4.1(a)(iii) Listed ordinary and preference property shares, market cap < R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] < 3
        and row["CCY"] == "ZAR"
        and row["Property"] == "P"
    ):
        return "4.1(a)(iii)"

    # 4.1(b) Unlisted ordinary and preference property shares
    elif (
        row["Investment Type"] == "EQ"
        and pd.isna(row["MCap"])
        and row["CCY"] == "ZAR"
        and row["Property"] == "P"
    ):
        return "4.1(b)"

    # 4.2(a)(i) Listed foreign ordinary and preference property shares, market cap >= R20bn
    elif (
        ((row["Investment Type"] == "EQ" and row["MCap"] >= 10) or row["dX"] == "p")
        and row["CCY"] != "ZAR"
        and row["Property"] == "P"
        or ((row["Fund"] == "f") and (row["fnd_typ"] == "p") and (row["CCY"] != "ZAR"))
    ):
        return "4.2(a)(i)"

    # 4.2(a)(ii) Listed foreign ordinary and preference property shares, market cap >= R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] >= 3
        and row["CCY"] != "ZAR"
        and row["Property"] == "P"
    ):
        return "4.2(a)(ii)"

    # 4.2(a)(iii) Listed foreign ordinary and preference property shares, market cap < R2bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] < 3
        and row["CCY"] != "ZAR"
        and row["Property"] == "P"
    ):
        return "4.2(a)(iii)"

    # 4.2(b) Unlisted foreign ordinary and preference property shares
    elif (
        row["Investment Type"] == "EQ"
        and pd.isna(row["MCap"])
        and row["CCY"] != "ZAR"
        and row["Property"] == "P"
    ):
        return "4.2(b)"

    # 5.1(a) Local commodities - NewGold, NewPlats and NewPalladium
    # if search(  'NEWGOLD', row['i Issue Name'].upper()) and row['CCY'] == 'ZAR' and \
    # row['Primary Asset ID'].find('GLD') != -1:
    elif (row["CCY"] == "ZAR") and (row["Commodity"] == "au"):
        return "5.1(a)(i)"  # local commodity, gold
    elif (row["CCY"] == "ZAR") and (row["dX"] == "r") and (row["Commodity"] != "au"):
        return "5.1(a)(ii)"  # local commodity, non-gold

    # 5.2(a) Foreign commodities
    elif (row["CCY"] != "ZAR") and (row["Commodity"] == "au"):
        return "5.2(a)(i)"  # foreign commodity, gold
    elif (row["CCY"] != "ZAR") and (row["dX"] == "r") and (row["Commodity"] != "au"):
        return "5.1(a)(ii)"  # foreign commodity, non-gold

    # 8. Hedge funds
    elif row["HF"] == "hf":
        return "8.1(a)(i)"  # local fund of hedge funds
    elif row["HF"] == "fohf":
        return "8.1(a)(ii)"  # local hedge funds

    # 9. Private equity funds
    elif row["PEF"] == "pef":
        return "9.1(a)(i)"  # local private equity fund
    elif row["PEF"] == "fopef":
        return "9.1(a)(ii)"  # local fund of private equity funds

    # 10. 'Other' - purposely no script here

    # 11. Infrastructure
    elif row["Infrastructure"] == 1:
        return "11(b)"

    # uncategorised
    else:
        return "--- r28 ---"


print(
    f" {timediff(start_time, time.time())} defining the Reg 28 classification function\n"
)

# function that classifies securities in terms of Rergulation 30 based on the values of other columns
# Python: Check if String Contains Substring https://stackabuse.com/python-check-if-string-contains-substring/

start_time = time.time()
print("Defining the Reg 30 classification function ...")

reg30_2b = ["CLN915", "CLN932", "CLN945", "CLN947", "CLN994"]  # per Circ 3 of 2025


def classify_Reg30(row):
    # Give precedence to the five securities held at 31 Dec 2024 and reclassified by the CMS as '7(b)'
    if (
        (row["Primary Asset ID"] == "CLN915")
        | (row["Primary Asset ID"] == "CLN932")
        | (row["Primary Asset ID"] == "CLN945")
        | (row["Primary Asset ID"] == "CLN947")
        | (row["Primary Asset ID"] == "CLN994")
    ):
        return "2(b)"

    # # Take CLN classification from Council for Medical Schemes circular as the default
    # elif pd.notna(row['MedCirc']) and pd.isna(row['CLN']):
    #     return  row['MedCirc']

    # # 7(a)(ii) "Other" securities
    # """ if any of these conditions are met, categorise the security as '7(a)(ii)' """
    # # elif res(['FWD', 'FT', 'OP', 'DERV'], row['Investment Type']) \
    # # or (row['repo'] == 'RPC') \
    # # elif (row['Derivative'] is np.nan) \
    # elif res(['FWD', 'FT', 'OP', 'DERV'], row['Investment Type']) \
    #     or ((row['margin'] == 1) and (row['Bank'] == 's')) \
    #     or (row['repo'] == 'RPC') \
    #     or ((row['Investment Type'] == 'FI') \
    #         and (row['BESA'    ] != 'B') \
    #         and (row['Property'] != 'P') \
    #         and (row['Deb'     ] != 'd') \
    #         and (row['GPL'     ] not in ['p', 'G', 'la'])
    #         and (row['Bank'    ] != 's')):
    #     return '7(a)(ii)'

    # 7(a)(ii) "Other" securities
    elif ((row["margin"] == 1) and (row["Bank"] == "s")) or (
        (row["Derivative"] is not np.nan)
        and (row["CLN"] != 1)
        and (row["BESA"] != "B")
        and (row["Property"] != "P")
        and (row["Deb"] != "d")
        and (row["GPL"] not in ["p", "G", "la"])
        and (row["Investment Type"] != "SYTH")
        or (
            res(["FWD", "FT", "OP", "DERV"], row["Investment Type"])
            and (row["Investment Type"] != "SYTH")
        )
    ):
        return "7(a)(ii)"
    # and (row['Bank'] is not np.nan)
    # simplified filter based on '7(a)(ii)' as identified in issuers_2_31Jul2025 tab 'R30 7(a)(ii) 200'

    # 1(a)(i) Cash with SARB-registered bank with DI900 >= R5bn, including margin account balances
    elif (
        (
            row["FRN"] == 1
            or res(["CASH", "ST"], row["Investment Type"])
            or row["margin"] == 1
        )
        and row["Bank"] == "s"
        and row["DI900"] >= 5
    ) or (row["Investment Type"] == "SYTH"):
        return "1(a)(i)"

    # 1(a)(ii)  Cash with SARB-registered bank with DI900 >= 0.1, including margin account balances
    elif (
        (
            row["FRN"] == 1
            or res(["CASH", "ST", "SYTH"], row["Investment Type"])
            or row["margin"] == 1
        )
        and row["Bank"] == "s"
        and row["DI900"] >= 0.1
    ):
        return "1(a)(ii)"

    # 1(a)(iii)  Cash with SARB-registered bank, collateralised with RSA government debt under an ISMA
    elif row["Investment Type"] == "ST" and res(
        ["ISMA", "Collateralised"], row["Primary Asset ID"]
    ):
        return "1(a)(iii)"

    # 1(b) Cash with a foreign bank
    elif (
        (
            row["FRN"] == 1
            or res(["CASH", "ST", "SYTH"], row["Investment Type"])
            or row["dX"] == "c"
        )
        and (row["Bank"] == "a" or row["Bank"] == "b")
        or row["Primary Asset ID"] == "0649317"
    ):
        return "1(b)(i)"

    # 2(a)(i) SA government and government-guaranteed debt
    elif (
        (
            row["Issuer"] == "Republic of South Africa"
            or row["GovGuar"] == 1
            or row["dX"] == "c"
        )
        and (row["CCY"] == "ZAR")
        and (row["Investment Type"] != "SYTH")
    ):
        return "2(a)(i)"

    # 2(a)(ii) SA local authority debt
    elif row["CCY"] == "ZAR" and row["GPL"] == "la":
        return "2(a)(ii)"

    # 2(a)(iii) DBSA debt
    elif row["Issuer"] == "Development Bank of Southern Africa":
        return "2(a)(iii)"

    # 2(a)(iv) IDC debt
    elif row["Issuer"] == "Industrial Development Corporation of South Africa":
        return "2(a)(iv)"

    # 2(a)(v) INCA debt
    elif row["Issuer"] == "Infrastructure Finance Corporation Limited":
        return "2(a)(v)"

    # 2(a)(vi) Land Bank debt
    elif row["Issuer"] == "Land and Agricultural Development Bank of South Africa":
        return "2(a)(vi)"

    # 2(a)(vii) TCTA debt
    elif row["Issuer"] == "Trans-Caledon Tunnel Authority":
        return "2(a)(vii)"

    # 2(a)(viii) SANRAL debt
    elif row["Issuer"] == "SA National Roads Agency SOC Ltd":
        return "2(a)(viii)"

    # 2(a)(ix) Eskom debt
    elif row["Issuer"] == "Eskom Holdings SOC Ltd":
        return "2(a)(ix)"

    # 2(a)(x) Transnet debt
    elif row["Issuer"] == "Transnet SOC Ltd":
        return "2(a)(x)"

    # 2(a)(xi) Debt issued by SARB-registered bank with DI900 >= R5bn
    elif row["Bank"] == "s" and pd.isna(row["repo"]) and row["DI900"] >= 5:
        return "2(a)(xi)"

    # 2(a)(xii) Debt issued by SARB-registered bank with DI900 >= 0.1bn
    elif row["Bank"] == "s" and pd.isna(row["repo"]) and row["DI900"] >= 0.1:
        return "2(a)(xii)"

    # 2(a)(xiii) Corporate debt listed on BESA and included in OTHI or ALBI
    # elif res(['FI', 'ST'], row['Investment Type']) and row['BESA'] == 'B' and row['CCY'] == 'ZAR' and \
    # pd.isna(row['GPL']) and pd.isna(row['Bank']):
    # return '2(a)(xiii)'

    # elif res(['FI', 'ST'], row['Investment Type']) and row['CCY'] == 'ZAR' and row['Exchange'] != 'CTSE' \
    # and row['Deb'] != 'd' and row['repo'] != 'RPC' and row['Property'] != 'P':
    # return '2(a)(xiv)'

    elif (
        res(["FI", "ST"], row["Investment Type"])
        and row["CCY"] == "ZAR"
        and row["Deb"] != "d"
        and row["repo"] != "RPC"
        and row["Property"] != "P"
    ):
        return "2(a)(xiv)"  # removed and row['Exchange'] != 'CTSE' for Medical Schemes Circular 11 of 2024

    # 2(b)(i) Foreign debt
    elif (
        res(["FI", "ST"], row["Investment Type"])
        and (row["Investment Type"] != "SYTH")
        and (row["CCY"] != "ZAR")
    ) | any(k in row["Primary Asset ID"] for k in reg30_2b):
        return "2(b)(i)"
    # elif (res(['FI', 'ST'], row['Investment Type']) and row['CCY'] != 'ZAR') | row['Primary Asset ID'].isin(reg30_2b):
    #     return '2(b)(i)'

    # 3(a)(i) SA property
    elif (
        res(["EQ", "FI"], row["Investment Type"])
        and row["Property"] == "P"
        and row["CCY"] == "ZAR"
    ):
        return "3(a)(i)"

    # 3(b) Foreign property
    elif (
        res(["EQ", "FI"], row["Investment Type"])
        and row["Property"] == "P"
        and row["CCY"] != "ZAR"
    ):
        return "3(b)"

    # 4(a)(i) Unlisted equity and unlisted debt in the JSE's 'Development Capital and Venture Capital' (DCVC) sector
    elif (
        row["Investment Type"] == "EQ"
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
        and row["Exchange"] == "DCVC"
    ):
        return "4(a)(i)"

    # 4(a)(ii)(i) JSE Listed equity with market cap > R50bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] > 50
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
    ):
        return "4(a)(ii)(i)"

    # 4(a)(ii)(ii) JSE Listed equity with market cap > R5bn
    elif (
        row["Investment Type"] == "EQ"
        and row["MCap"] >= 5
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
    ):
        return "4(a)(ii)(ii)"

    # 4(a)(ii)(iii) JSE Listed equity with market cap < R5bn
    elif (
        ((row["Investment Type"] == "EQ" and row["MCap"] < 5) or pd.isna(row["MCap"]))
        and row["CCY"] == "ZAR"
        and row["Property"] != "P"
        and row["Deb"] != "d"
    ):
        return "4(a)(ii)(iii)"

    # 4(a)(iii)(i) JSE Listed ETFs linked to JSE ALSI
    elif (
        row["CCY"] == "ZAR"
        and row["Exchange"] == "XJSE"
        and row["Fund"] == "f"
        and "ALSI" in row["Issuer Name"]
    ):
        return "4(a)(iii)(i)"

    # 4(a)(iii)(ii) JSE Listed ETFs not linked to JSE ALSI
    elif (
        row["CCY"] == "ZAR" and row["Fund"] == "f" and "ALSI" not in row["Issuer Name"]
    ):
        return "4(a)(iii)(ii)"

    # 4(a)(iv)(i) (Unlisted) CISes linked to the ALSI
    elif (
        row["CCY"] == "ZAR"
        and pd.isna(row["Exchange"])
        and row["Fund"] == "f"
        and "ALSI" in row["Issuer Name"]
    ):
        return "4(a)(iv)(i)"

    # 4(a)(iv)(ii) (Unlisted) CISes linked to the ALSI
    elif (
        row["CCY"] == "ZAR"
        and pd.isna(row["Exchange"])
        and row["Fund"] == "f"
        and "ALSI" not in row["Issuer Name"]
    ):
        return "4(a)(iv)(ii)"

    # 4(a)(v)(i) Insurance policies linked to the JSE ALSI
    elif (
        row["CCY"] == "ZAR"
        and pd.isna(row["Exchange"])
        and row["Fund"] == "p"
        and "ALSI" in row["Issuer Name"]
    ):
        return "4(a)(v)(i)"

    # 4(a)(v)(ii) Insurance policies linked to the JSE ALSI
    elif (
        row["CCY"] == "ZAR"
        and pd.isna(row["Exchange"])
        and row["Fund"] == "p"
        and "ALSI" not in row["Issuer Name"]
    ):
        return "4(a)(v)(ii)"

    # 4(b)
    elif (row["Investment Type"] == "EQ" and row["Property"] != "P") and row[
        "CCY"
    ] != "ZAR":
        return "4(b)"

    # 5(a) SA corporate debentures
    elif row["CCY"] == "ZAR" and row["Deb"] == "d":
        return "5(a)"

    # 5(b) Foreign corporate debentures
    elif row["CCY"] != "ZAR" and row["Deb"] == "d":
        return "5(b)"

    # 6(a)(i) SA unlinked insurance policies
    elif row["CCY"] == "ZAR" and pd.isna(row["Exchange"]) and row["Fund"] == "p":
        return "6(a)(i)"

    # 6(a)(ii) SA Linked insurance policies
    elif row["CCY"] == "ZAR" and pd.isna(row["Exchange"]) and row["Fund"] == "p":
        return "6(a)(ii)"

    # 6(b) Foreign linked insurance policies
    elif row["CCY"] != "ZAR" and pd.isna(row["Exchange"]) and row["Fund"] == "p":
        return "6(b)"

    # 7(a)(i) SA inventories carried at lesser of book value and realisable value
    elif row["CCY"] == "ZAR" and pd.isna(row["Exchange"]) and row["Fund"] == "i":
        return "7(a)(i)"

    # 7(b) Foreign 'other' securities
    elif (row["CCY"] != "ZAR" and pd.isna(row["Exchange"]) and row["Fund"] == "i") or (
        (row["margin"] == 1) and (row["Bank"] == "b")
    ):
        return "7(b)"

    # (As a last resort) Take CLN classification from Council for Medical Schemes circular as the default
    elif pd.notna(row["MedCirc"]) and pd.isna(row["CLN"]):
        return row["MedCirc"]

    # uncategorised
    else:
        return "--- r30 ---"


print(
    f" {timediff(start_time, time.time())} defining the Reg 30 classification function\n"
)

# (10) assign Reg 28 and reg 30 classifications to each security
start_time = time.time()
print("Classifying each security ...")

merged["Reg 28 Classification"] = merged.apply(lambda row: classify_Reg28(row), axis=1)
merged["Reg 30 Classification"] = merged.apply(lambda row: classify_Reg30(row), axis=1)

print(f" {timediff(start_time, time.time())} classifying each security\n")

# access Excel application to open results sheet
import win32com.client as win32  # library to convert xls to xlsx

excel = win32.gencache.EnsureDispatch("Excel.Application")  # to open excel application

# (15) identify securities with absent classifications
start_time = time.time()
print(
    'Writing newly classified securities to "issuers_2.xlsx" and summarising them ...'
)

# drop duplicate 'no issuer' rows https://www.interviewqs.com/ddi-code-snippets/drop-duplicate-rows-pandas
no_cat_r28 = merged[
    (merged["Reg 28 Classification"] == "--- r28 ---")
    | (merged["Reg 28 Classification"].isnull())
].drop_duplicates(subset="Primary Asset ID", keep="first")
no_cat_r30 = merged[
    (merged["Reg 30 Classification"] == "--- r30 ---")
    | (merged["Reg 30 Classification"].isnull())
].drop_duplicates(subset="Primary Asset ID", keep="first")
no_derv_ft_op = merged[
    (
        (merged["Investment Type"] == "DERV")
        | (merged["Investment Type"] == "FT")
        | (merged["Investment Type"] == "OP")
    )
    & pd.isna(merged["dX"])
].drop_duplicates(subset="Primary Asset ID", keep="first")
term_neg = merged[merged["Term"] < 0].drop_duplicates(
    subset="Primary Asset ID", keep="first"
)
uniques = merged.drop_duplicates(subset="Primary Asset ID", keep="first")

# Securities without market caps # https://www.geeksforgeeks.org/filter-pandas-dataframe-with-multiple-conditions/
no_MCap = merged[(merged["MCap"].isnull().values)].drop_duplicates(
    subset="Primary Asset ID", keep="first"
)

# Equities without market caps # https://www.geeksforgeeks.org/filter-pandas-dataframe-with-multiple-conditions/
eq_no_MCap = merged[
    (merged["MCap"].isnull().values)
    & (merged["Investment Type"] == "EQ")
    & (merged["Fund"] != "f")
].drop_duplicates(subset="Primary Asset ID", keep="first")

# (16A)identifying securities assigned to RSA government
rsa = merged[(merged["Issuer"] == "Republic of South Africa")].drop_duplicates(
    subset="Primary Asset ID", keep="first"
)  # RSA government issuer

# (16B)identifying securities without named issuers
uiss = merged[(merged["Issuer"] == "-xxx")].drop_duplicates(
    subset="Primary Asset ID", keep="first"
)  #

# (17) write the dataframe to review it as a workbook
# issuers_2 = pd.ExcelWriter(r'P:\Working Folders\Hilton\W\Reg_Tests\issuers_2.xlsx', engine  = 'xlsxwriter')

# merged.to_excel(       issuers_2, index = False,  sheet_name =  'all'                )             # categorised securities
# uniques.to_excel(      issuers_2, index = False,  sheet_name = f'uniques ({len(uniques)})')        # unique securities
# no_MCap.to_excel(      issuers_2, index = False,  sheet_name = f'no MCaps ({len(no_MCap)})')       # unassigned issuers
# eq_no_MCap.to_excel(   issuers_2, index = False,  sheet_name = f'no eq MCaps ({len(eq_no_MCap)})') # unassigned issuers
# term_neg.to_excel(     issuers_2, index = False,  sheet_name = f'neg term ({len(term_neg)})')      # -ve terms to maturity
# no_derv_ft_op.to_excel(issuers_2, index = False,  sheet_name = f'no dX ({len(no_derv_ft_op)})')    # unassigned derivatives
# no_cat_r28.to_excel(   issuers_2, index = False,  sheet_name = f'no R28 ({len(no_cat_r28)})')      # Reg 28 not assigned
# no_cat_r30.to_excel(   issuers_2, index = False,  sheet_name = f'no R30 ({len(no_cat_r30)})')      # Reg 30 not assigned
# rsa.to_excel(          issuers_2, index = False,  sheet_name = f'rsa ({len(rsa)})')                # RSA government issuer
# uiss.to_excel(         issuers_2, index = False,  sheet_name = f'issuers ({len(uiss)})')           # unnamed issuers

# items = np.sort(merged['Reg 28 Classification'].unique()) # get numpy array of items and sort them alphabetically
# for item in items:
#     t = merged[(merged['Reg 28 Classification'] == item)].drop_duplicates(subset='Primary Asset ID', keep = 'first') # get a subset of the 'merged' dataframe
#     t.to_excel(issuers_2, index = False,  sheet_name = f'R28 {item} ({len(t)})')        # show the subset as a tab in 'issuers_2.xlsx'

# items = np.sort(merged['Reg 30 Classification'].unique()) # get numpy array of items and sort them alphabetically
# for item in items:
#     t = merged[(merged['Reg 30 Classification'] == item)].drop_duplicates(subset='Primary Asset ID', keep = 'first') # get a subset of the 'merged' dataframe
#     t.to_excel(issuers_2, index = False,  sheet_name = f'R30 {item} ({len(t)})')        # show the subset as a tab in 'issuers_2.xlsx'

# issuers_2.close()

# (17) write the dataframe to review it as a workbook
with pd.ExcelWriter(iss_2, engine="xlsxwriter") as writer:
    merged.to_excel(writer, index=False, sheet_name="all")  # categorised securities
    uniques.to_excel(
        writer, index=False, sheet_name=f"uniques ({len(uniques)})"
    )  # unique securities
    no_MCap.to_excel(
        writer, index=False, sheet_name=f"no MCaps ({len(no_MCap)})"
    )  # unassigned issuers
    eq_no_MCap.to_excel(
        writer, index=False, sheet_name=f"no eq MCaps ({len(eq_no_MCap)})"
    )  # unassigned issuers
    term_neg.to_excel(
        writer, index=False, sheet_name=f"neg term ({len(term_neg)})"
    )  # -ve terms to maturity
    no_derv_ft_op.to_excel(
        writer, index=False, sheet_name=f"no dX ({len(no_derv_ft_op)})"
    )  # unassigned derivatives
    no_cat_r28.to_excel(
        writer, index=False, sheet_name=f"no R28 ({len(no_cat_r28)})"
    )  # Reg 28 not assigned
    no_cat_r30.to_excel(
        writer, index=False, sheet_name=f"no R30 ({len(no_cat_r30)})"
    )  # Reg 30 not assigned
    rsa.to_excel(
        writer, index=False, sheet_name=f"rsa ({len(rsa)})"
    )  # RSA government issuer
    uiss.to_excel(
        writer, index=False, sheet_name=f"issuers ({len(uiss)})"
    )  # unnamed issuers

    items = np.sort(
        merged["Reg 28 Classification"].unique()
    )  # get numpy array of items and sort them alphabetically
    for item in items:
        t = merged[(merged["Reg 28 Classification"] == item)].drop_duplicates(
            subset="Primary Asset ID", keep="first"
        )  # get a subset of the 'merged' dataframe
        t.to_excel(
            writer, index=False, sheet_name=f"R28 {item} ({len(t)})"
        )  # show the subset as a tab in 'issuers_2.xlsx'

    items = np.sort(
        merged["Reg 30 Classification"].unique()
    )  # get numpy array of items and sort them alphabetically
    for item in items:
        t = merged[(merged["Reg 30 Classification"] == item)].drop_duplicates(
            subset="Primary Asset ID", keep="first"
        )  # get a subset of the 'merged' dataframe
        t.to_excel(
            writer, index=False, sheet_name=f"R30 {item} ({len(t)})"
        )  # show the subset as a tab in 'issuers_2.xlsx'

writer.close()

# present results
print(
    "  Total securities                   :",
    len(merged),
    "\n",
    " Unique securities                  :",
    len(uniques),
    "\n",
    " All securities without market caps :",
    len(no_MCap),
    "\n",
    " Negative term securities           :",
    len(term_neg),
    "\n",
    " Equities without market caps       :",
    len(eq_no_MCap),
    "\n",
    " RSA government as issuer           :",
    len(rsa),
    "\n",
    "\n",
    " Unassigned derivative asset class  :",
    len(no_derv_ft_op),
    "\n",
    " Unnamed issuers                    :",
    len(uiss),
    "\n",
    " Reg 28 uncategorised securities    :",
    len(no_cat_r28),
    "\n",
    " Reg 30 uncategorised securities    :",
    len(no_cat_r30),
)

print(
    f'{timediff(start_time, time.time())} writing newly classified securities to "issuers_2.xlsx" and summarising them \n'
)

start_time = time.time()
print(
    "Opening issuers_2 if derivatives or issuers unassigned or securities not categorised, ..."
)

if len(no_derv_ft_op) + len(uiss) + len(no_cat_r28) + len(no_cat_r30) > 0:
    excel.Workbooks.Open(iss_2)

print(
    f" {timediff(start_time, time.time())} opening issuers_2.xlsx if derivatives or issuers unassigned or securities not categorised",
    "\n",
)

print(
    f" {timediff(start_time_issuers_2, time.time())} ISSUERS_2 COMPLETED \n ================================"
)
