#!/usr/bin/env python
# coding: utf-8

# # Month-End 17 Holdings Reports to External Clients

# ### Dependencies
#
# issuers_2 and issuers_3 files for the reporting month suffixed with _ddMmmyyyy.xlsx

print("\n\n##########################")
print("#   START monthend17.py  #")
print("##########################\n\n")

import time

start_time_me17 = time.time()

from datetime import datetime
import pandas as pd
import shutil  # for creating folders and copying files
from tqdm import tqdm
from constants import (
    pth_me17,
    pthPy,
    pthTest,
    pthSttlmnt,
    pth_struct,
    pth_m_reports,
    pth_EC,
    pth_PPSBAL,
    pthReports,
    fldr_PManco,
)
from utilities import timediff, prior_month_end

# dataframe the issuers_2_ and issuers_3_ sheets
start_time = time.time()

# df = pd.read_excel(pth_me17, sheet_name="Process", usecols="C", nrows=1)
df = pd.read_excel(pthPy, sheet_name="arc", usecols="AG", nrows=2)
k = df.iloc[1, 0]
rptDate = k if k == k else prior_month_end().date()
print(
    f" {rptDate.strftime('%A %d %b %Y')} \
reporting date from 'arc' sheet"
)

# get the 'all' sheets from issuers_2,xlsx and issuers_3.xlsx and merge them
issuers_2 = pthTest + rf"\issuers_2_{rptDate.strftime('%d%b%Y')}.xlsx"
is2 = pd.read_excel(issuers_2, sheet_name=None)
is2 = pd.read_excel(issuers_2, sheet_name=list(is2.keys())[0])

issuers_3 = pthTest + rf"\issuers_3_{rptDate.strftime('%d%b%Y')}.xlsx"
is3 = pd.read_excel(issuers_3, sheet_name=None)
is3 = pd.read_excel(issuers_3, sheet_name=list(is3.keys())[0])

rptDate_issuers_2 = datetime.strptime(issuers_2[-14:].replace(".xlsx", ""), "%d%b%Y")

print(
    f" {rptDate_issuers_2.strftime('%A %d %b %Y')} reporting date from issuers_2.xlsx"
)

print(
    f"{timediff(start_time, time.time())} dataframing the issuers_2 and issuers_3 sheets used for month-end reporting\n"
)

# create the funds dataframe with the merge of issuers_2 and issuers_3 and fund long names
start_time = time.time()
print("Merging issuers_2, issuers_3, and fund long names")

# merge issuers_3 with issuers_2 to retain instrument classifications - # https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.merge.html
cols = [
    "Primary Asset ID",
    "i Issue Name",
    "Investment Type",
    "Reg 28 Classification",
    "Reg 30 Classification",
    "Percentage of Market Value",
    "End Market Value",
    "Closing Exposure PA",
    "margin",
    "Derivative",
    "Issuer",
    "Counterparty",
    "Infrastructure",
    "CCY",
]
funds = is3.merge(
    is2, how="left", left_on=cols, right_on=cols, suffixes=["_is3", "_is2"]
)

# check if any NaNs in the classification column - https://stackoverflow.com/questions/29530232/how-to-check-if-any-value-is-nan-in-a-pandas-dataframe
# print(issr['Reg 28 Classification'].isnull().values.sum(), '\n\n', funds.shape, '\n\n', list(funds))

# merge issuers_2 + issuers_3 and names to arrive at the funds dataframe
names = pd.read_excel(
    pthSttlmnt, sheet_name="Funds", usecols=["Fund Code", "Fund Name"]
)

funds = funds.merge(names, how="left", left_on="Entity Name", right_on="Fund Code")
print(f" {funds.shape[0]:,.0f} rows and {funds.shape[1]} columns in 'funds' dataframe")

# get currency descriptions
crrncy = pd.read_excel(
    pth_struct, sheet_name="curr", usecols=["Curr", "CurrencyDescription"]
)
# crrncy.isnull().values.sum() # check if any NaNs in the dataframe

# merge funds with currency descriptions
funds = funds.merge(crrncy, how="left", left_on="CCY", right_on="Curr")

print(
    f"{timediff(start_time, time.time())} merging issuers_2, issuers_3, and fund long names"
)

# add new columns - https://stackoverflow.com/questions/29517072/add-column-to-dataframe-with-constant-value

start_time = time.time()
print("Adding columns to the dataframe")
import numpy as np

funds["PortfolioCode"] = funds["Entity Name"].apply(lambda x: x + "_EXP")
funds["Fund Name"] = funds["Fund Name"].apply(lambda x: str(x) + " (Exploded)")
funds["Book Value"] = funds["End Market Value"]
funds["Holding"] = funds["Book Value"].apply(np.floor)
funds["Level"] = "Instrument"
funds["% Selection"] = funds["Percentage of Market Value"]
funds["% Category"] = np.nan
funds["Unrealised P/L"] = funds["End Market Value"] - funds["Book Value"]
funds["ScripOnLoan"] = 0
funds["EmbargoHolding"] = 0
funds["Instrument_CounterpartyCode"] = funds["Counterparty"]
funds["Instrument_CurrencyCode"] = funds["CCY"]
funds["Instrument_InstrumentTypeCode"] = funds["Investment Type"]
funds["Instrument_IssuerCode"] = funds["Issuer"]
funds["Instrument_MaturityDate"] = funds["Date"]
funds["Instrument_PropertyIndicator"] = list(funds["property"] == "P")

# new column based on an if-else statement - https://www.dataquest.io/blog/tutorial-add-column-pandas-dataframe-based-on-if-else-condition/
funds["Issuer_Infrastructure"] = np.where(funds["infra"] == "i", True, False)


# https://stackoverflow.com/questions/40953914/python-return-multiple-values-and-check-for-return-false
def listedtest(a):
    if a[:6] == "3.1(a)" or a[:6] == "4.1(a)" or a[:6] == "3.2(a)" or a[:6] == "4.2(a)":
        return True


funds["Instrument_ListedIndicator"] = list(
    map(listedtest, funds["Reg 28 Classification"]) or (funds["BESA"] == "B")
)

print(f" {timediff(start_time, time.time())} adding columns to the dataframe \n")

# rename column headings - https://stackoverflow.com/questions/11346283/renaming-column-names-in-pandas

start_time = time.time()
print("Reordering and renaming dataframe columns (semi roundtrip)")
# reorder column headings
# dictionary for the renamed column headings
dict = {
    "Fund Name": "PortfolioDescription",
    "Primary Asset ID": "InstrumentCode",
    "i Issue Name": "InstrumentDescription",
    "CCY": "Currency",
    "CurrencyDescription": "Currency Description",
    "End Market Value": "Market Value",
    "Closing Exposure PA": "Exposure",
    "Percentage of Market Value": "% Portfolio",
    "Reg 28 Classification": "Instrument_Reg28Classification",
}

funds.rename(columns=dict, inplace=True)
# print(list(funds))

# set the column order
col_order = [
    "Level",
    "PortfolioCode",
    "PortfolioDescription",
    "InstrumentCode",
    "InstrumentDescription",
    "Currency",
    "Currency Description",
    "Holding",
    "Book Value",
    "Market Value",
    "Exposure",
    "% Portfolio",
    "% Selection",
    "% Category",
    "Unrealised P/L",
    "ScripOnLoan",
    "EmbargoHolding",
    "Instrument_CounterPartyCode",
    "Instrument_CountryCode",
    "Instrument_CurrencyCode",
    "Instrument_InstrumentTypeCode",
    "Instrument_IssuerCode",
    "Instrument_ListedIndicator",
    "Instrument_MaturityDate",
    "Instrument_PropertyIndicator",
    "Instrument_Reg28Classification",
    "Issuer_Infrastructure",
]

funds = funds.reindex(
    col_order, axis=1
)  # https://docs.kanaries.net/topics/Pandas/pandas-reorder-columns
# "It's important to note that you need to pass axis=1 to reindex() method to specify that you're reordering columns, not rows.'

# print(len(list(funds)), list(funds))

print(
    f" {timediff(start_time, time.time())} Reordering and renaming dataframe columns (semi roundtrip) \n"
)

# create the month end reporting folders if they don't yet exist - https://flexiple.com/python/python-make-directory

start_time = time.time()
print(f"Creating the month-end reporting folders, if they don't yet exist\n")

import os

# create the month-end_17 reporting folder
pth = os.path.join(pth_m_reports, rptDate.strftime("%Y"), rptDate.strftime("%Y %m"))
if not os.path.exists(pth):
    os.makedirs(pth)

# create the month reporting folder for ECICBALC
fldr_ECICBALC = os.path.join(pth_EC, f"{rptDate.strftime('%Y%m')}")
if not os.path.exists(fldr_ECICBALC):
    os.makedirs(fldr_ECICBALC)

# create the month-end reporting folder for PPSBAL:
fldr_PPSBAL = os.path.join(pth_PPSBAL, f"{rptDate.strftime('%Y%m')}")
if not os.path.exists(fldr_PPSBAL):
    os.makedirs(fldr_PPSBAL)

print(
    f" {timediff(start_time, time.time())} creating the month-end reporting folders, if they didn't yet exist\n"
)

# # TEST
# print('', pth_m_reports,'\n', fldr_ECICBALC)


# copy PSIF Reg30 to PManco folder and
# copy ECICBALC Reg28 and Reg28 Table 2 to ECICBALC folder
# https://www.freecodecamp.org/news/python-copy-file-copying-files-to-another-directory/
start_time = time.time()
print(
    "Copying PSIF, PLMED, GACASH, GAEMBF, and ECICBALC reports to their respective reporting folders ..."
)

# from folders and file names
pthReports = r"P:\Investment Operations\GRC\Compliance\Reg28 and Reg30 Reporting"
fln_GAEM_Reg30 = f"GAEMBF Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
fln_GACA_Reg30 = f"GACASH Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
fln_ECIC_Reg28 = f"ECICBALC Reg28 {rptDate.strftime('%d%b%Y')}.xlsx"
fln_ECIC_Tbl2 = f"ECICBALC Reg28 Table2 {rptDate.strftime('%d%b%Y')}.xlsx"
fln_ECIC_SchIB = f"ECICBALC Reg28 SchIB {rptDate.strftime('%d%b%Y')}.xlsx"
fln_PSIF_Reg30 = f"PSIF Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
fln_PLMED_Reg30 = f"PLMED Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"

# to folders
# fldr_EC        = r'P:\Investment Operations\Segregated Clients\Active Clients\Export Credit Insurance Corporation\ECICBAL\Reporting\Monthly Reports'
# pth_m_reports  = r'P:\Investment Operations\GRC\Compliance\Reporting Requirements\Monthly Reports'

# copy GACASH to new folder
frm = os.path.join(pthReports, fln_GACA_Reg30)
fl_new = f"{names[names['Fund Code'] == 'GACASH'].iloc[0, 1]} (GACASH) Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
to = os.path.join(
    pth_m_reports, rptDate.strftime("%Y"), rptDate.strftime("%Y %m"), fl_new
)
if os.path.exists(frm):
    shutil.copyfile(frm, to)

# copy GAEMBF to new folder
frm = os.path.join(pthReports, fln_GAEM_Reg30)
fl_new = f"{names[names['Fund Code'] == 'GAEMBF'].iloc[0, 1]} (GAEMBF) Reg30 {rptDate.strftime('%d%b%Y')}.xlsx"
to = os.path.join(
    pth_m_reports, rptDate.strftime("%Y"), rptDate.strftime("%Y %m"), fl_new
)
if os.path.exists(frm):
    shutil.copyfile(frm, to)

# copy PSIF to new folder
frm = os.path.join(pthReports, fln_PSIF_Reg30)
to = os.path.join(fldr_PManco, f"{rptDate.strftime('%Y%m%d')} PSIF Reg30.xlsx")
if os.path.exists(frm):
    shutil.copyfile(frm, to)

# copy PLMED to month-end folder
frm = os.path.join(pthReports, fln_PLMED_Reg30)
to = os.path.join(
    pth_m_reports, rptDate.strftime("%Y"), rptDate.strftime("%Y %m"), fln_PLMED_Reg30
)
if os.path.exists(frm):
    shutil.copyfile(frm, to)

# copy ECICBALC Reg28 to new folder
filenames = ["fln_ECIC_Reg28", "fln_ECIC_Tbl2"]
for filename in filenames:
    frm = os.path.join(pthReports, filename)
    to = os.path.join(fldr_ECICBALC, filename)
    if os.path.exists(frm):
        shutil.copyfile(frm, to)

print(
    f" {timediff(start_time, time.time())} copying PSIF, PLMED, GAEMBF, and ECICBALC reports to their respective reporting folders competed \n"
)

# get list of month-end 17 portfolio codes
res = pd.read_excel(pthPy, sheet_name="arc", usecols="AF").dropna()
me17_list = [s + "_EXP" for s in res["Month-end 17"]]
print(f"{len(me17_list)} funds: \n   {(', ').join(me17_list)}")


# create the individual look-through holdings report workbooks
start_time = time.time()
print("Creating and then saving individual look-through holdings report workbooks ...")

# open xlwings to write the reports
import xlwings as xw

app = xw.App(visible=False)  # no visible sheet updating

# for fund in tqdm(funds['PortfolioCode'].unique()):
for fund in tqdm(me17_list):
    # open the derv template and assign holdings and data sheets
    # print(fund)
    wb = app.books.add()  # open a new workbook
    sh = wb.sheets["Sheet1"]  # get the sheet
    sh.name = (
        f"{fund.replace('_EXP', '')} {rptDate.strftime('%d%b%Y')}"  # rename the sheet
    )

    sh.range("A1").options(index=False).value = funds[
        funds["PortfolioCode"] == fund
    ]  # paste fund holdings

    sh.range("A2:AA2").insert(shift="down")  # insert a heading row
    # - https://stackoverflow.com/questions/75377059/insert-row-at-the-top-of-an-excel-table-using-xlwings
    sh["A2"].value = "Portfolio"
    sh["B2:C2"].value = sh[
        "B3:C3"
    ].value  # copy fund code and fund description from row below
    sh["F2"].value = "ZAR"  # fund currency code and ...
    sh["G2"].value = "South African Rand"  # ... fund currency description

    hds = [
        "Holding",
        "Book Value",
        "Market Value",
        "Exposure",
        "% Portfolio",
        "% Selection",
    ]
    # print(fund)
    for index, hd in enumerate(hds):  # sum of the value columns
        sh.range(2, index + 8).value = funds[funds["PortfolioCode"] == fund][hd].sum()

    flname = f"{sh['C2'].value.replace(' (Exploded)', '')} ({sh['B2'].value.replace('_EXP', '')}) {rptDate.strftime('%d%b%Y')}"
    wb.save(
        os.path.join(pth, flname) + ".xlsx"
    )  # save the file as .xlsx because xlwings can't save it as .csv
    wb.close()  # close the fund's workbook

    df = pd.read_excel(
        os.path.join(pth, flname) + ".xlsx", header=0
    )  # pandas read the newly created .xlsx file
    df.to_csv(os.path.join(pth, flname) + ".csv", index=False)  # save the file as .csv
    try:
        os.remove(os.path.join(pth, flname) + ".xlsx")
    except:
        print(f"{os.path.join(pth, flname)}.xlsx not found")

print(
    f"Creating and then saving individual look-through holdings report workbooks competed: {timediff(start_time, time.time())}"
)
print("\n", f"Roundtrip time: {timediff(start_time_me17, time.time())}")

app.quit()

print("\n\n##########################")
print("#    END monthend17.py   #")
print("##########################\n\n")
