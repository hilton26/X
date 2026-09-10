#!/usr/bin/env python
# coding: utf-8

# # Assigning issuer names to securities in a portfolio — DB version

# A copy of issuers_1.py, adapted for the extra "issuer_code" column
# cs1_PARN_merge_csv_db.py now adds to the CS1 workbook. Two changes
# from issuers_1.py:
#
# 1. That extra column shifts every column after "Primary Asset ID" by
#    one position, so the fixed "A:J" read range and the "date column
#    is at position 9" assumptions below would silently misread the
#    file (or grab the wrong header as the report date) if left as-is.
#    Fixed to "A:K" / position 10.
# 2. classify1() now checks the database-sourced "issuer_code" first
#    (after the existing CLN/repo/structured-note special cases, which
#    are deliberate overrides and still take precedence), ahead of the
#    regex-based issuer_did() text matching. issuer_code is a short
#    code (e.g. "NED", "ABSA", "INVESTECLTD"), not the full formatted
#    name issuer_did() returns (e.g. "Nedbank Ltd") — kept as its own
#    value rather than mapped to a full name, since no reliable
#    short-code -> full-name table was already loaded here to map
#    through safely.
#
# Source: cs1_PARN_merge_csv_db.py's "All" sheet, columns: 0. 'Entity
# Name', 1. 'Investment Type', 2. 'i Issue Name', 3. 'Primary Asset ID',
# 4. 'issuer_code', 5. 'CCY', 6. 'Reg28 Classification', 7. 'End Market
# Value', 8. 'Percentage', 9. 'Closing Exposure PA', 10. (report date)
#
# On "TypeError: This COM object can not automate the makepy process - please run makepy manually for this object", \
# close the Excel dialogue box and run this script again.')


print("\n\n####################################")
print("#                                  #")
print("#      START issuers_1_db.py      #")
print("#                                  #")
print("####################################\n\n")


# libraries, libraries!
print("Importing libraries for issuers_1_db ...\n")
import time

start_time_issuers_1 = time.time()
start_time = time.time()

# general libraries
from datetime import datetime
import pandas as pd  # for dataframes

import numpy as np  # for np.NaN
import re  # for regex
from re import search  # for regex
from datetime import datetime  # for script run durations
import os  # for BESA folder contents
import sys
from tqdm import tqdm
import subprocess
from constants import (
    pthPy,
    pthBESA,
    pth_struct,
    pthCLNs,
    pthMedCirc2026,
    pthSttlmnt,
    pthTest,
    yll,
    iss_1_db,
    issuers_2,
    issuers_3,
)
from utilities import timediff

print(
    f" {timediff(start_time, time.time())} \
importing libraries for issuers_1_db\n"
)

# (1) read in the lookthrough holdings file,
# from the 'arc' sheet of py_reports.xlsm, as a dataframe

start_time = time.time()
print("Reading in classifier input file ...")

# check if a url was given in the py_reports file and then ...
df_check = pd.read_excel(pthPy, sheet_name="arc", usecols="V", nrows=7)
url = df_check.iloc[6, 0].replace('"', "")
rpt = df_check.iloc[2, 0]
if url == url:  # ... if so, use the py_reports url ...
    # the "All" sheet now has 11 columns (0-10): the extra "issuer_code"
    # column shifts the report-date header from column J (index 9) to
    # column K (index 10) — see module docstring
    py_input = pd.read_excel(url, usecols="A:K")
    if isinstance(list(py_input)[10], str):
        rptDate = datetime.strptime(list(py_input)[10], "%d %b %Y")
    else:
        rptDate = list(py_input)[10]
else:  # ... prompt for a valid url
    print(
        'Please provide a valid URL \
in cell "L2" of the "classifier" tab'
    )

print(f" {url if isinstance(url, str) else 'No url to look-through holdings'}\n")

fnds = py_input["Entity Name"].unique()
funds = (", ").join(fnds)
s = "" if len(fnds) == 1 else "s"
print(
    f"{rptDate.strftime('%a %d %b %Y')} instrument \
classifications for {len(fnds)} fund{s}:\n {funds}"
)

# get BESA data
besa_fnames = [
    int(s[re.search(r"\d{8}", s).span()[0] : re.search(r"\d{8}", s).span()[0] + 8])
    for s in os.listdir(pthBESA)
    if "." in s
]
besa_fdate = max(besa_fnames)  # BESA file date as an integer
bsaDate = datetime.strptime(str(besa_fdate), "%Y%m%d")  # BESA file date as a datetime

print(f" Report date    : {rptDate.strftime('%a %d %b %Y')}")
print(f" Report type    : {rpt}")
print(f" BESA file date : {bsaDate.strftime('%a %d %b %Y')}")
print(
    f" {len(py_input) - 1:,} securities, {len(py_input['Primary Asset ID'].unique()):,} \
(1/{(len(py_input) - 1) / len(py_input['Primary Asset ID'].unique()):.1f} times or \
{len(py_input['Primary Asset ID'].unique()) / (len(py_input) - 1) * 100:.1f}%) of which are unique"
)

print(
    f"\n {timediff(start_time, time.time())} \
reading in classifier input file\n"
)

# get input data
start_time = time.time()
print(
    "Reading in input data incl regex, CLNs, med schemes, \
settlement, BESA, accruals, realty, and margins ..."
)

res = list(filter(lambda x: str(besa_fdate) in x, os.listdir(pthBESA)))[0]
bsa = os.path.join(pthBESA, res)
besa_data = pd.read_csv(
    bsa, skiprows=4, usecols=["Bond Code"]
).dropna()  # BESA-listed security codes
lstds = pd.read_excel(
    pth_struct, sheet_name="listed", usecols=["Bond Code"]
).dropna()  # Listed debt override
dfrgx = pd.read_excel(
    pth_struct,
    sheet_name="dates",
    usecols=["date_regex", "format", "alternative format"],
).dropna(subset=["format"])  # regex formats
sa_hols = pd.read_excel(
    pth_struct, sheet_name="hols", usecols=["sa_hols"]
).dropna()  # import SA holidays
accr = pd.read_excel(pth_struct, sheet_name="accr", usecols=["accruals"]).dropna()
fnd_typ = pd.read_excel(pth_struct, sheet_name="fnd", usecols=["description"]).dropna()
marg = pd.read_excel(pth_struct, sheet_name="accr", usecols=["margins"]).dropna()
issrgx = pd.read_excel(
    pth_struct,
    sheet_name="issuers",
    usecols=[
        "description",
        "id",
        "issuer name",
    ],
).dropna(subset=["issuer name"])
indx1 = pd.read_excel(
    pth_struct, sheet_name="issuers", usecols=["description", "issuer name", "ticker"]
).dropna(subset=["issuer name"])
bx_re = pd.read_excel(
    pth_struct, sheet_name="issuers", usecols=["issuer name", "property"]
).dropna(subset=["issuer name"])
realty = bx_re[bx_re["property"] == "P"].drop("property", axis=1).reset_index(drop=True)
bond_data = pd.read_excel(
    pth_struct, sheet_name="guar", usecols=["Bond Code", "Guarantee Type"]
).dropna()  # govt guarantee
clns = pd.read_excel(
    pthCLNs,
    sheet_name="CLN",
    usecols=["Code", "CLN?", "Counterparty Long Name", "Issuer Long Name"],
).dropna(subset=["Code"])
med_circ = pd.read_excel(
    pthMedCirc2026, sheet_name="ListedDebtDec2025", usecols=["Bond Code", "2026Circ7"]
).dropna(subset=["Bond Code"])
sttlmnt = pd.read_excel(
    pthSttlmnt, sheet_name="Sttlmnt", usecols=["Fund", "Custodian", "SAFEX"]
).dropna(subset=["Fund"])
indx = indx1[indx1["ticker"] == "Index"].drop("ticker", axis=1).reset_index(drop=True)

accr_list = [x for e in accr.values.tolist() for x in e]
margin_list = [x for e in marg.values.tolist() for x in e]

print(
    f" {timediff(start_time, time.time())} reading \
in input data incl regex, CLNs, med schemes, \
settlement, BESA, accruals, realty, and margins\n"
)


# 'descr' + 'id' columns
start_time = time.time()
print(
    'Joining "description" and "id" columns from \
the "issuers" tab of pth_struct.xlsm into a dataframe ...'
)

cols_to_convert = ["description", "id"]
for col in cols_to_convert:
    issrgx[col] = issrgx[col].fillna("").astype(str)

print(issrgx[["description", "id"]].info())

issrgx["descid"] = issrgx[["description", "id"]].agg(
    lambda row: "|".join(v for v in row if v), axis=1
)

print(
    f' {timediff(start_time, time.time())} joining \
"description" and "id" columns from the "issuers" \
tab of pth_struct.xlsm into a dataframe\n'
)


# Determine funds with nil effective exposure
start_time = time.time()
print(
    "Determining funds with zero effective exposure \
and funds with no settlement account"
)

names = fnds
zero_EE = []
for name in names:
    if py_input[py_input["Entity Name"] == name]["Closing Exposure PA"].sum() == 0:
        zero_EE.append(name)

print(
    f" {len(zero_EE)} funds with empty effective exposure \
column: \n  {(', ').join(list(zero_EE))} \n"
)

# confirm all funds have a settlement bank account by merging df and sttlmnt
no_sttlmnt = py_input.merge(sttlmnt, left_on="Entity Name", right_on="Fund", how="left")
fnds_0 = no_sttlmnt[no_sttlmnt["Custodian"].isna()]["Entity Name"].unique()
s1 = "s" if len(fnds_0) != 1 else ""
print(
    f" {len(fnds_0)} fund{s1} with no settlement account: \n  {(', ').join(list(fnds_0))}"
)

# determine securities with integer IDs (SA government securities without "R" prefixed)
is_int_mask = py_input["Primary Asset ID"].apply(lambda x: isinstance(x, int))
print(
    f'{len(py_input[is_int_mask])} SA government \
securities without "R" prefixed'
)

py_input["Primary Asset ID"] = py_input["Primary Asset ID"].apply(
    lambda x: "R" + str(x) if isinstance(x, int) else x
)

print(
    f"{timediff(start_time, time.time())} determining fund with \
zero effective exposure and funds with no settlement account \n"
)

# create text pattern functions

start_time = time.time()
print("Setting up functions ...")


def cln(txt):
    if txt in list(clns["Code"]):
        return clns[clns["Code"] == txt]["CLN?"].iloc[0]
    else:
        return None


# function to identify FRNs and NCDs
def frn(txt):
    pattern = "STEP UP|STEP-UP|STEP_UP|STEPUP|NCD|FRN"
    if re.search(pattern, str(txt).upper()):
        return 1


# function to identify ILBs
def ilb(txt):
    pattern = "INFLATION|CPI|ILB"
    if re.search(pattern, str(txt).upper()):
        return 1


# function to identify funds
def fnd(txt):
    import re
    from re import search

    pattern1 = r"\b(?:FUND(?!\s*MANAGE)|FUND(?!S)|UCIT|ETF|ISHARES)\b"
    patternc = r"\b(PHYSICAL GOLD|GOLD ETC|COMMODITY|PHYSICAL SILVER|SILVER ETC|PLATINUM|PALLADIUM)\b"
    patternm = r"\b(MONEY)\b"
    patternd = r"\b(BONDS|$ TIP|$TIP|DURATION|YIELD|INCOME|INTEREST|POSITIVE RETURN)\b"
    patterne = r"\b(EQUITY|WORLD|FEEDER|\sPLUS\s|MSCI|SMALL CAP|LARGE CAP|GLOBAL|OPPORTUNITY|BIN YUAN|BALANCED|INTL)\b"
    patternp = r"\b(\sREAL\s|REALTY|REIT|PROPERTY|REAL ESTATE|HOMES|FAIRVEST|HYPROP)\b"
    if re.search(pattern1, str(txt).upper()) and re.search(patternc, str(txt).upper()):
        return "fc"
    elif re.search(pattern1, str(txt).upper()) and re.search(
        patternm, str(txt).upper()
    ):
        return "fm"
    elif re.search(pattern1, str(txt).upper()) and re.search(
        patternd, str(txt).upper()
    ):
        return "fd"
    elif re.search(pattern1, str(txt).upper()) and re.search(
        patterne, str(txt).upper()
    ):
        return "fe"
    elif re.search(pattern1, str(txt).upper()) and re.search(
        patternp, str(txt).upper()
    ):
        return "fp"
    else:
        return None


def property(txt):
    import re
    from re import search

    pattern = r"\b(\sREAL\s|REALTY|REIT|PROPERTY|REAL ESTATE|HOMES|FAIRVEST|HYPROP)\b"
    if re.search(pattern, str(txt).upper()):
        return "p"


# function to discern property tickers; ## Claude, 9 Mar 2026
def classify_re(row):
    isin_match = row["ISIN"] in isins["ISIN"].values
    return "P" if isin_match or name_match else np.nan


# function to identify pribvate equity funds
def pef(txt):
    import re
    from re import search

    pattern1 = r"\b(PRIVATE EQUITY)\b"
    pattern2 = r"\b(PRIVATE EQUITY FUND OF FUNDS|PRIVATE FOF)\b"
    if re.search(pattern1, str(txt).upper()):
        return "pef"
    elif re.search(pattern2, str(txt).upper()):
        return "pefof"
    else:
        return None


# function to identify pribvate equity funds
def hf(txt):
    import re
    from re import search

    pattern1 = r"HEDGE FUND(?! OF FUNDS)"
    pattern2 = r"\b(HEDGE FUND OF FUNDS|HEDGE FOF)\b"
    if re.search(pattern1, str(txt).upper()):
        return "hf"
    elif re.search(pattern2, str(txt).upper()):
        return "hfof"
    else:
        return None


def commodity(txt):
    import re
    from re import search

    pattern1 = r"\b(?:PHYSICAL GOLD|GOLD ETC)\b"
    pattern2 = r"\b(?:COMMODITY|PHYSICAL SILVER|SILVER ETC|PLATINUM|PALLADIUM)\b"
    if re.search(pattern1, str(txt).upper()):
        return "au"
    elif re.search(pattern2, str(txt).upper()):
        return "xu"
    else:
        return None


# function to find index
def inx(txt):
    if (indx["description"].eq(txt)).any():
        if (
            indx.loc[indxbond_data["Bond Code"] == txt].iat[0, 1]
            == "GOVERNMENT GUARANTEE"
        ):
            return 1


# function to derive margin account
def mrg(txtA, txtB):
    pattern1 = r"\sMARGIN|(?!\d\d)MARG(?!\d\d)"  # pattern to test txtA in 'i Issue Name' column
    pattern2 = "VARMAR"  # pattern to test txtB in 'Primary Asset ID' column
    if re.search(pattern1, str(txtA).upper()) or re.search(
        pattern2, str(txtB).upper()
    ):
        return 1


# function to indicate if a security is government guaranteed
def gvg(txt):
    if (bond_data["Bond Code"].eq(txt)).any():
        if (
            bond_data.loc[bond_data["Bond Code"] == txt].iat[0, 1]
            == "GOVERNMENT GUARANTEE"
        ):
            return 1


# function that indicates if bond security is BESA-listed
def besa(txt):
    if (besa_data["Bond Code"].eq(txt)).any() or (lstds["Bond Code"].eq(txt)).any():
        return "B"


# function to identify strings of text starting with 3 or 4 capital letters and ending with two or three digits
# to identify candidate BESA-listed securities
def besa_maybe(txt):
    pattern = r"^[A-Z]{3,4}\d{2,3}$"
    if re.search(pattern, str(txt).upper()):
        return 1


# function to identify a repo
def repo(txt):
    pattern = "RPCO|RPMT|RPCA"
    if re.search(pattern, str(txt).upper()):
        return txt[0:3].upper()


# function to get days remaining to maturity
def term(date_string):
    try:
        if len(date_string) == 0:
            return None
        else:
            return datetime.strptime(date_string, "%d%b%Y") - rptDate

    # manage exceptions:
    except TypeError as te:
        return None
    except ValueError as ve:
        return None


# function to assign med scheme category based on current Circular 11 of 2024 from the CMS
def medcirc(txt):
    if med_circ["Bond Code"].eq(txt).any():
        return med_circ.loc[med_circ["Bond Code"] == txt].iat[0, 1]


# utility function to check if a string contains an element in a given list
def res(t_list, t_string):
    return bool([ele for ele in t_list if (ele in t_string)])


# function that extracts date from string
def datex(txt):
    try:
        for pattern in dfrgx["date_regex"]:
            if re.search(pattern, txt.title()):
                break
        return datetime.strptime(
            re.search(pattern, txt.title()).group(),
            dfrgx.loc[dfrgx["date_regex"] == pattern].iat[0, 1],
        ).strftime("%d%b%Y")

    # manage exceptions:
    except ValueError as ve:
        return None
    except TypeError as te:
        return None
    except AttributeError as ae:
        return None


# function to derive issuer name based on search for a pattern in instrument description and instrument id
def issuer_did(txt):
    for pattern in issrgx["descid"]:
        if re.search(pattern, str(txt).upper()):
            return (
                issrgx.loc[issrgx["descid"] == pattern]
                .reset_index(drop=True)
                .loc[0, "issuer name"]
            )
            break


# function to identify a derivative counterparty
def counterparty(row):
    if row["Investment Type"] == "SYTH":
        return np.nan
    elif (clns["Code"].eq(row["Primary Asset ID"])).any():
        return clns.loc[clns["Code"] == row["Primary Asset ID"]].iat[0, 2]
    elif (
        (row["Investment Type"] == "OP")
        and (
            (row["Derivative"] in "Structured Note")
            or (row["Derivative"] in "Linked Note")
        )
        and (row["CCY"] == "ZAR")
    ):
        return issuer_did(row["i Issue Name"])
    elif ((row["Investment Type"] == "OP") or (row["Investment Type"] == "FT")) and (
        row["CCY"] == "ZAR"
    ):
        return "JSE"
    elif ((row["Investment Type"] == "OP") or (row["Investment Type"] == "FT")) and (
        row["CCY"] != "ZAR"
    ):
        return "Exchange"
    elif (row["Investment Type"] == "EQ") or (row["Investment Type"] == "DERV"):
        return row["Issuer"]
    else:
        return row["Issuer"]


# function to identify a derivative for Reg 28 CS1
def derivative(row):
    pattern1 = "STRUCTURED"
    pattern2 = "LINKED"
    if row["Investment Type"] == "SYTH":
        return np.nan
    elif re.search(pattern1, row["i Issue Name"].upper()):
        return "Structured Note"
    elif re.search(pattern2, row["i Issue Name"].upper()):
        return "Linked Note"
    elif row["CLN"] == 1:
        return "Credit-linked Note"
    elif row["repo"] == "RPC":
        return "Repo Trade"
    elif row["Investment Type"] == "DERV":
        return "Swap"
    elif row["Investment Type"] == "FT":
        return "Futures"
    elif row["Investment Type"] == "OP":
        return "Listed Option"
    elif row["Investment Type"] == "FWD":
        return "Currency Forward"


# function to derive index name name based on search for a pattern in instrument description
def dexin(txt):
    for pattern in indx["description"]:
        if re.search(pattern, str(txt).upper()):
            return indx.loc[indx["description"] == pattern].iat[0, 1]
            break


print(f" {timediff(start_time, time.time())} setting up functions", "\n")


# function to assign an issuer
start_time = time.time()
print("Setting up issuer identifier function ...")


def classify1(row):
    t = issuer_did(row["i Issue Name"])  # temp, so function only gets called once
    g = issuer_did(row["Primary Asset ID"])  # temp, so function only gets called once

    if cln(row["Primary Asset ID"]) == 1:  # if cln is included in list then ...
        return clns.loc[clns["Code"] == row["Primary Asset ID"]].iat[
            0, 1
        ]  # ... look up reference entity

    elif repo(row["Primary Asset ID"]) == "RPC":  # repo, bank legs RPCO and RPCA
        return "Absa Bank Ltd"

    elif repo(row["Primary Asset ID"]) == "RPMT":  # repo, government bond leg
        return "Republic of South Africa"

    elif (row["Investment Type"] == "OP") and (
        (row["Derivative"] in "Structured Note") or (row["Derivative"] in "Linked Note")
    ):  # repo, gov leg
        return dexin(row["i Issue Name"])

    elif isinstance(row.get("issuer_code"), str) and row["issuer_code"].strip():
        # database-sourced issuer_code, authoritative where present —
        # takes priority over the regex text matching below, but not
        # over the CLN/repo/structured-note overrides above
        return row["issuer_code"]

    elif isinstance(g, str):  # INSTRUMENT ID is a hit ...
        return g

    elif isinstance(t, str):  # INSTRUMENT DESCRIPTION is a hit ...
        return t

    else:  # signifies no issuer assigned
        return "-xxx-"


print(
    f" {timediff(start_time, time.time())} setting up issuer identifier function", "\n"
)


# (2) remove blank Market Value rows and zero-value Effective Exposure rows
start_time = time.time()
print("Removing NaN and zero value rows ...")

before = py_input.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
before["Diff"] = before["End Market Value"] - before["Closing Exposure PA"]

rowsNaN = len(py_input[py_input["End Market Value"].isnull()])
py_input = py_input[py_input["End Market Value"].notnull()]  # remove NaN MV column rows
df_input = py_input[
    (round(py_input["End Market Value"], 2) != 0)
    | (round(py_input["Closing Exposure PA"], 2) != 0)
]

# drop the last column (the report-date column, now at position 10 —
# see module docstring)
df_input = df_input.drop(df_input.columns[10], axis=1)

# delete rows where MV and EE are zero to two decimals
print(
    f" {rowsNaN:,} NaN rows and {len(py_input) - len(df_input):,} zero effective exposure rows \
removed, {len(df_input):,} rows remain"
)

after = df_input.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
after["Diff"] = after["End Market Value"] - after["Closing Exposure PA"]

print(f"{timediff(start_time, time.time())} removing NaN and zero value rows", "\n")


# (3) change the 'Percentage of Market Value' column
start_time = time.time()
print('Changing "Percentage of Market Value" column ...')

navs = df_input.groupby("Entity Name")["End Market Value"].sum()
nav = navs.to_dict()

newTMV = []
for i, row in tqdm(df_input.iterrows(), total=df_input.shape[0]):
    if nav[row["Entity Name"]] == 0:
        fndpct = 100
    else:
        fndpct = row["End Market Value"] / nav[row["Entity Name"]] * 100
    newTMV.append(fndpct)

df_input["Percentage of Market Value"] = newTMV

print(
    f'{timediff(start_time, time.time())} changing "Percentage of Market Value" column',
    "\n",
)


# (4) save 'df_input' dataframe including ALL funds as a workbook to be used later
start_time = time.time()
print('Saving "df_input" dataframe as a workbook called "yall" ...')

with pd.ExcelWriter(yll, engine="xlsxwriter") as writer:
    df_input.to_excel(writer, index=False, sheet_name="all")
writer.close()

print(
    f' {timediff(start_time, time.time())} saving \
"df_input" dataframe as a workbook called "yall"\n'
)


uniques = df_input.drop_duplicates(subset="Primary Asset ID", keep="first")
print(list(uniques), uniques.shape)


# (5) find unique instruments and identify their instrument attributes
start_time = time.time()
print("Isolating unique securities ...")

uniques = df_input.drop_duplicates(subset="Primary Asset ID", keep="first")

uniques.drop(["Entity Name"], axis=1, inplace=True, errors="ignore")

uniques = uniques[
    ~uniques["Primary Asset ID"].isin(accr_list)
    & ~uniques["Primary Asset ID"].isin(margin_list)
]

uniques.reset_index(drop=True, inplace=True)

print(f" {len(uniques)} unique securities")

print(
    f"   {timediff(start_time, time.time())} \
isolating unique securities\n"
)


# (6) identify the unique instruments' attributes
start_time = time.time()
print("Appending instrument attributes to the unique securities ...")

uniques[["FundX", "PEFX", "HFX", "CommodityX"]] = None

uniques["CLN"] = uniques["Primary Asset ID"].map(cln)
uniques["FRN"] = uniques["i Issue Name"].map(frn)
uniques["ILB"] = uniques["i Issue Name"].map(ilb)
uniques["Date"] = uniques["i Issue Name"].map(datex)  # security maturity date
uniques["FundX"] = uniques["i Issue Name"].map(fnd)
uniques["PropertyX"] = uniques["i Issue Name"].map(property)
uniques["PEFX"] = uniques["i Issue Name"].map(pef)
uniques["HFX"] = uniques["i Issue Name"].map(hf)
uniques["CommodityX"] = uniques["i Issue Name"].map(commodity)  # commodity
uniques["MedCirc"] = uniques["Primary Asset ID"].map(medcirc)
uniques["GovGuar"] = uniques["Primary Asset ID"].map(gvg)
uniques["repo"] = uniques["Primary Asset ID"].map(repo)
uniques["BESA"] = uniques["Primary Asset ID"].map(besa)
uniques["BESA_MAYBE"] = uniques["Primary Asset ID"].map(besa_maybe)
uniques["margin"] = uniques.apply(
    lambda x: mrg(x["i Issue Name"], x["Primary Asset ID"]), axis=1
)
uniques["Derivative"] = uniques.apply(derivative, axis=1)
uniques["Term"] = (pd.to_datetime(uniques["Date"], format="%d%b%Y") - rptDate).dt.days

print(
    f" {timediff(start_time, time.time())} appending \
instrument attributes to the unique securities\n"
)


# (7) identify issuers
start_time = time.time()
print(
    f"Identifying issuers over {len(uniques):,} unique securities \
for {len(fnds)} fund{s} at {rptDate.strftime('%a %d %b %Y')} using \
{len(issrgx['issuer name'].dropna()):,} regex patterns and the database \
issuer_code column where available ..."
)

issuers = []
for index, row in tqdm(uniques.iterrows(), total=uniques.shape[0]):
    issuers.append(classify1(row))

uniques["Issuer"] = issuers

n_from_issuer_code = sum(
    1
    for i, row in uniques.iterrows()
    if isinstance(row.get("issuer_code"), str)
    and row["issuer_code"].strip()
    and row["Issuer"] == row["issuer_code"]
)
print(
    f" {n_from_issuer_code:,} of {len(uniques):,} issuers assigned \
directly from the database issuer_code column"
)

print(f" {timediff(start_time, time.time())} identifying issuers\n")


# (8) identify derivative counterparties for Reg 28 CS1 of 2023
start_time = time.time()
print("Identifying derivative counterparties ...")

uniques["Counterparty"] = uniques.apply(counterparty, axis=1)

print(
    f" {timediff(start_time, time.time())} \
identifying derivative counterparties\n"
)


# (9) identify securities with absent issuers
start_time = time.time()
print("Identifying securities with absent issuers ...")

no_issuer = uniques[
    (uniques["Issuer"] == "-xxx-") | uniques["Issuer"].isnull()
].drop_duplicates(subset="Primary Asset ID", keep="first")

no_CLN_issuer = uniques[
    (uniques["CLN"] == 1) & (uniques["Issuer"].isna())
].drop_duplicates(subset="Primary Asset ID", keep="first")

print(f" Unallocated issuers : {str(len(no_issuer.Issuer))}")
print(f" Unallocated CLNs    : {str(len(no_CLN_issuer.CLN))}")

print(
    f" {timediff(start_time, time.time())} \
identifying securities with absent issuers\n"
)


# (10) write the dataframe to review it as a workbook
start_time = time.time()
print("Writing the dataframe to a sheet for review ...")

iss1_xl = pd.ExcelWriter(iss_1_db, engine="xlsxwriter")
uniques.to_excel(iss1_xl, index=False, sheet_name="uniques")
no_issuer.to_excel(iss1_xl, index=False, sheet_name=f"no issuers ({len(no_issuer)})")
no_CLN_issuer.to_excel(
    iss1_xl, index=False, sheet_name=f"no CLN issuers ({len(no_CLN_issuer)})"
)

iss1_xl.close()

print(
    f" {timediff(start_time, time.time())} \
writing the dataframe to a sheet for review\n"
)


# (11) prettify the sheets using openpyxl
start_time = time.time()
print("Giving the review sheet structure with openpyxl ...")

import openpyxl as px
from openpyxl.cell import Cell
from openpyxl.styles import (
    Alignment,
    Color,
    PatternFill,
    Font,
    Border,
)

wb = px.load_workbook(iss_1_db)
for sheet in wb.worksheets:
    sheet.auto_filter.ref = sheet.dimensions
    sheet.freeze_panes = sheet["J2"]
    for row in sheet.iter_rows(min_row=1, max_row=1):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.fill = PatternFill(
                start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
            )

# add a hyperlink to the 'pth_struct' sheet
sht = wb[f"no issuers ({len(no_issuer)})"]
sht["P1"].hyperlink = pth_struct
sht["P1"].value = "pth_struct.xlsm"
sht["P1"].font = Font(bold=True, underline="single", color="0000EE")
sht["P1"].alignment = Alignment(horizontal="left", vertical="top")

wb.save(iss_1_db)
wb.close()

print(
    f" {timediff(start_time, time.time())} giving the \
review sheet structure with openpyxl\n",
)
print(
    f"\n {timediff(start_time_issuers_1, time.time())} \
ISSUERS_1_DB COMPLETED\n===============================\n"
)


print("\n\n####################################")
print("#                                  #")
print("#      END issuers_1_db.py         #")
print("#                                  #")
print("####################################\n\n")


# (12) run issuers_2, and _3.ipynb if all securities
# have an assigned issuer, else open issuers_1_db.xlsx
start_time = time.time()
print(
    "Running issuers_2, and _3.ipynb if all securities \
have an assigned issuer, else opening issuers_1_db.xlsx, ..."
)

print(
    f"Exceptions: \n {len(no_issuer)} unnamed issuers, \
and,\n {len(no_CLN_issuer)} unnamed CLN issuers"
)

if (len(no_issuer.Issuer) == 0) and (len(no_CLN_issuer.CLN) == 0):
    subprocess.run([sys.executable, issuers_2])
    subprocess.run([sys.executable, issuers_3])
else:
    os.startfile(iss_1_db)
    print(r"Check the issuers and CLNs in the \issuers_1_db.xlsx file")

print(
    f"\n{timediff(start_time_issuers_1, time.time())} \
running issuers_1_db, _2, and _3.ipynb\n"
)

print("\n\n##########################################")
print("#                                        #")
print("#     END issuers_1_db, _2, and _3.py    #")
print("#                                        #")
print("##########################################\n\n")
