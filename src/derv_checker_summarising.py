#!/usr/bin/env python
# coding: utf-8

# # Summarise Derivative Cover Check Sheets - new format

print("\n\n##############################################")
print("#                                            #")
print("#   START 3/4 derv_checker_summarising.py X  #")
print("#                                            #")
print("##############################################\n\n")

# Import libraries

import time

start_time = time.time()
start_time_0 = start_time
print("Importing libraries to summarise the derivative calcs ...")

import pandas as pd
import os, sys, shutil
from utilities import timediff, parn_de
from tqdm import tqdm
from constants import (
    pthEXPORTS,
    pthMandates,
    pthOverdrafts,
    pthSttlmnt,
    pthDaily,
    pthLOCAL,
)

print(
    f" {timediff(start_time, time.time())} importing libraries \
to summarise the derivative calcs\n",
)

# Set location paths for the sheets to be used

start_time = time.time()
print("Setting up paths ...")

# create a lookup table for fund UT status and investment team
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A, D:E")

print(f" {timediff(start_time, time.time())} setting up paths", "\n")

# Get report date and selected summary sheet option
start_time = time.time()
print("Getting the reporting date and names of completed derivative files ...")

fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()

# check if the required files have been downloaded, else continue
if not os.path.exists(fPARN) or not os.path.exists(fDE):
    sys.exit(
        f"Stopping: missing expected download(s):\n"
        f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
        f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}\n"
    )

# df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
#     subset=[0]
# )
# k = df.iloc[2, 1]
# rptDate = (
#     k if isinstance(k, datetime) else prior_working_day(datetime.today())
# )  # prior working day or report date override; has type datetime()
# summ_yn = df.iloc[3, 1]
# funds = df[0].iloc[1:]

# # derive holdings and derivatives file paths
# fPARN = os.path.join(
#     pth_dl,
#     f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )

# fDE = os.path.join(
#     pth_dl,
#     f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )

# # check if the required files have been downloaded, else
# if not os.path.exists(fPARN) or not os.path.exists(fDE):
#     sys.exit(
#         f"Stopping: missing expected download(s):\n"
#         f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
#         f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}"
#     )

print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds:\n",
    f"{(', ').join(funds.tolist())}",
)

print(
    f" {timediff(start_time, time.time())} getting the reporting \
date and latest downloaded holdings and derivatives files\n",
)

# rptDate    = datetime(2025, 3, 6)    # TEST
y = os.scandir(pthEXPORTS)
pattern = f"Derv Calc {rptDate.strftime('%d%b%Y')}.xlsx"
start_y = time.time()
print(
    "",
    f"Get scandir() of derivative calcs \
folder for {rptDate.strftime('%a %d %B %Y')}",
)
funds_cmpl = []
for s in tqdm(
    y,
    desc=f"Getting the names of completed derivative \
calculation files for {rptDate.strftime('%A %d %B %Y')} ...",
):
    if s.name.endswith(pattern):  # all 'XXXX Derv  Calc ddMmmYYYY.xlsx' files
        funds_cmpl.append(
            s.name[:-25]
        )  # fund codes of all files with current report dates
print("", f"{len(funds_cmpl)} funds")
print(
    "",
    f" {timediff(start_y, time.time())} to get scandir() of \
derivative calcs folder for {rptDate.strftime('%a %d %B %Y')}",
)

print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} \
funds: {(', ').join(funds.tolist())}\n",
)
print(
    f" {'No' if summ_yn == 'No' else 'A'} summary \
sheet is required\n"
)

print(
    f" {len(funds_cmpl)} completed derivative calculation \
files at {rptDate.strftime('%A %#d %B %Y')}"
)
print(
    f"\n  {timediff(start_time, time.time())} getting the reporting \
date and names of completed derivative files\n",
)

# Create a dataframe with the necessary columns

start_time = time.time()
print("Creating a dataframe with the necessary columns ...")

summary = pd.DataFrame(funds_cmpl, columns=["Fund Code"])

cols = [
    "UT?",
    "#",
    "Cash Cover",
    "Cash Cover 2",
    "Incl CLNs & longer-dated debt",
    "Incl Underlying UTs",
    "SA Equity Indices",
    "Ex-SA Equity Indices",
    "SA Bond Indices",
    "Ex-SA Bond Indices",
    "Currency Futures",
    "Currency Forwards",
    "Swaps",
    "FRAs",
    "SA Bond Cover",
    "Foreign Equity Cover",
    "Local Equity Cover",
    "Total Equity",
    "Total Foreign",
    "Global Exposure",
    "Net Effective Exposure",
    "Fund Rules",
    "Team",
    "PIM Overdrafts",
    "Leverage (Gross)",
    "Leverage (Net)",
]

for col in cols:  # create an empty summary dataframe with the given column headings
    summary[col] = pd.Series([""] * len(summary), dtype=object)
    # https://www.reddit.com/r/learnpython/comments/n1ee17/how_to_add_multiple_empty_columns_into_my_data/

print(
    f"  {timediff(start_time, time.time())} creating a \
dataframe with the necessary columns\n",
)

# dataframe the fund derivative calc summaries

start_time = time.time()
# read each fund's 'xxxx Derv Calc ddmmmyyyy.xlsx' sheet into the summary dataframe
if summ_yn != "No":
    for index, fund in enumerate(
        tqdm(
            funds_cmpl,
            desc=f"Populating the summary \
dataframe with {len(funds_cmpl)} funds for \
{rptDate.strftime('%A %#d %B %Y')} ...",
        )
    ):  # https://stackoverflow.com/questions/522563/how-to-access-the-index-value-in-a-for-loop
        fn = (
            pthEXPORTS + rf"\{fund}" + f" Derv Calc {rptDate.strftime('%d%b%Y')}.xlsx"
        )  # access each fund's 'xxxx Derv Calc ddmmmyyyy.xlsx' sheet
        ddfSm = pd.read_excel(
            fn, header=None, sheet_name="Summary"
        )  # access data on the Summary sheet
        summary.iat[index, 1] = twoA[twoA.iloc[:, 0] == fund].iloc[
            0, 2
        ]  # column 'B' of summary sheet = "G3", has value 'UT' or '≠UT'
        summary.iat[index, 2] = ddfSm.iat[0, 4].astype(
            int
        )  # 'E1'  'number of derivatives'
        summary.iat[index, 3] = round(
            ddfSm.iat[37, 2], 3
        )  # 'C38' 'In/Adequate derivative cover for derivatives'
        summary.iat[index, 4] = round(
            ddfSm.iat[37, 4], 3
        )  # column 'E' of summary sheet = cash cover 2
        summary.iat[index, 5] = round(
            ddfSm.iat[37, 2], 3
        )  # column 'F' of summary sheet = 'C10'
        summary.iat[index, 6] = round(
            ddfSm.iat[35, 2], 3
        )  # column 'G' of summary sheet = 'C34' 'Cash from the underlying UTs'
        summary.iat[index, 7] = round(
            ddfSm.iat[1, 7] * 100, 3
        )  # column 'H' of summary sheet = 'H2' SA equity indices
        summary.iat[index, 8] = round(ddfSm.iat[4, 7] * 100, 3)
        summary.iat[index, 9] = round(ddfSm.iat[7, 7] * 100, 3)
        summary.iat[index, 10] = round(ddfSm.iat[10, 7] * 100, 3)
        summary.iat[index, 11] = round(ddfSm.iat[13, 7] * 100, 3)
        summary.iat[index, 12] = round(ddfSm.iat[16, 7] * 100, 3)
        summary.iat[index, 13] = round(ddfSm.iat[19, 7] * 100, 3)  # "H20" 'Swaps'
        summary.iat[index, 14] = round(ddfSm.iat[22, 7] * 100, 3)  # 'H23' 'FRAs'
        summary.iat[index, 15] = round(ddfSm.iat[5, 4], 3)  # "E6"  'SA bond cover'
        summary.iat[index, 16] = round(
            ddfSm.iat[9, 4], 3
        )  # "E10" 'Foreign equity cover'
        summary.iat[index, 17] = round(
            ddfSm.iat[13, 4], 3
        )  # "E14" 'Local equity cover'
        # summary.iat[index, 18] = round(ddfSm.iat[44, 2] * 100, 3)  # "C45" 'Total equity, incl property equity'
        # summary.iat[index, 19] = round(ddfSm.iat[58, 2]      , 3)  # "C59" 'Total foreign'
        # summary.iat[index, 20] = round(ddfGE.iat[ 1, 3] * 100, 3)
        # summary.iat[index, 21] = round(ddfGE.iat[ 2, 3] * 100, 3)
        summary.iat[index, 22] = fund  # f'{fund} calc sheet'
        # summary.iat[index, 23] = twoA.loc[twoA['Fund Code'] == f'{fund}'].iat[0,1] # investment team lookup on 2AXX sheet
        summary.iat[index, 23] = twoA[twoA.iloc[:, 0] == fund].iloc[
            0, 1
        ]  # investment team lookup on 2AXX sheet
        summary.iat[index, 24] = round(ddfSm.iat[3, 1], 2)  # fund NAV
        summary.iat[index, 25] = round(ddfSm.iat[18, 4], 1)  # fund leverage (gross)
        summary.iat[index, 26] = round(ddfSm.iat[19, 4], 1)  # fund leverage (net)

        # Using at[] and iat[] instead of loc[] and iloc[]
        # https://stackoverflow.com/questions/28757389/pandas-loc-vs-iloc-vs-at-vs-iat
        # at and iat are meant to access a scalar, that is, a single element in the dataframe,
        # while loc and iloc are meant to access several elements at the same time,
        # potentially to perform vectorized operations
        # https://medium.com/codex/dont-use-loc-iloc-with-loops-in-python-instead-use-this-f9243289dde7

        # sort and reindex the summary dataframe
        # https://stackoverflow.com/questions/17141558/how-to-sort-a-pandas-dataframe-by-two-or-more-columns
        # https://stackoverflow.com/questions/33165734/update-index-after-sorting-data-frame
        summary.reset_index(inplace=True, drop=True)
        # summary.sort_values(by = ['Cash Cover', '#'], ascending = [True, False], ignore_index = True) # inplace = True, )
        # df = df.reset_index().sort_values(by=['Date', 'index']).drop(['index'], axis=1)
        # https://stackoverflow.com/questions/48066933/pandas-sorting-days-whilst-preserving-order

print(
    f"  {timediff(start_time, time.time())} populating \
the summary dataframe with {len(funds_cmpl)} funds \
for {rptDate.strftime('%A %#d %B %Y')}\n"
)

summary = summary.sort_values(
    by="Cash Cover", ascending=True
)  # sort the cover calc dataframe by 'Cash Cover' in ascending order
# summary

# Get report date and selected summary sheet option and then populate the summaries in a dataframe

start_time = time.time()
print(
    f"Sorting and then saving the summary \
dataframe with {len(funds_cmpl)} funds for \
{rptDate.strftime('%A %#d %B %Y')} ..."
)

# sort the summary dataframe and save it to a new Excel file
ut_types = ["UT", "≠UT", "UCITS", "SAA", "TAA", "ICAV"]
sorted_summary = pd.DataFrame([])  # empty dataframe
for ut_type in ut_types:  # stack the > 0 derivative funds first ...
    summary_subset = summary[(summary["UT?"] == ut_type) & (summary["#"] != 0)]
    sorted_summary = pd.concat([sorted_summary, summary_subset])
    # sorted_summary = sorted_summary.sort_values(by = 'Cash Cover', ascending = False) # sort the summary dataframe

for ut_type in ut_types:  # ... then stack the no derivative funds
    summary_subset = summary[(summary["UT?"] == ut_type) & (summary["#"] == 0)]
    sorted_summary = pd.concat([sorted_summary, summary_subset])
    # sorted_summary = sorted_summary.sort_values(by = 'Cash Cover', ascending = False) # sort the summary dataframe

sorted_summary.reset_index(inplace=True, drop=True)

print(
    f" {timediff(start_time, time.time())} sorting and then\
saving the summary dataframe with {len(funds_cmpl)} \
funds for {rptDate.strftime('%A %#d %B %Y')}"
)

start_time = time.time()
print("Writing the dataframe to a sheet ...")

# TEST ++++++++++++++++

# # save the summary dataframe to a new Excel file as 'Derv ddmmyyyy.xlsx'
# sorted_summary.to_excel(pthEXPORTS + f'\Derv {rptDate}.xlsx', index = False, sheet_name = 'Summary')
# print(' ', pthEXPORTS + f'\Derv {rptDate}.xlsx')


# dataframe the PARN and Derv reports from the local Downloads folder
wbH = pd.read_csv(fPARN)
wbD = pd.read_csv(fDE)

# rename 'UNKNOWNs' as 'SWAPS' where they are not SYTH or empty portfolio holdings
# https://stackoverflow.com/questions/36909977/update-row-values-where-certain-condition-is-met-in-pandas
unknowns_filter = (
    (wbH["Valuation First Level"] == "UNKNOWN")
    & (wbH["Sub Security Type"] == "TRS")
    & (wbH["Investment Type"] != "SYTH")
    & pd.notna(wbH["Investment Type"])
)
wbH.loc[unknowns_filter, ["Valuation First Level", "Valuation Second Level"]] = "SWAPS"
unknowns = wbH.loc[unknowns_filter]
print(
    f'  {len(unknowns)} "UNKNOWN" securities \
found and amended: \
{(", ").join(unknowns["PrimaryAssetID"].tolist())}'
)

# identify "No Data found for this Entity" funds
no_data_filter = wbH["i Issue Name"] == "No Data found for this Entity"
no_data = wbH.loc[no_data_filter]
print(
    f'  {len(no_data)} "No data" \
fund{"s" if len(no_data["Entity ID"]) != 1 else ""}: \
{(", ").join(no_data["Entity ID"].tolist())}'
)

# # Change the '% of Total Market Value" column

# start_time = time.time()
# print("Recalculating and saving fund Total Market Value percentages ...")

# # change the '% of Total Market Value" column (N) to the fund-specific % based on 'Sum of Market Value Income' column (M)
# navs = wbH.groupby("Entity ID")[
#     "Sum of Market Value Income"
# ].sum()  # (column N) this has type 'pandas.core.series.Series'
# nav = navs.to_dict()  # nav series changed to dictionary to make it lookupable

# # recalc the '% of Total Market Value' column per fund and then ...
# newTMV = []
# for i, row in tqdm(wbH.iterrows()):
#     if nav[row["Entity ID"]] == 0:
#         fndpct = 100
#     else:
#         fndpct = row["Sum of Market Value Income"] / nav[row["Entity ID"]] * 100
#     newTMV.append(fndpct)

# # ... replace the '% of Total Market Value' with the values in the new list
# wbH["% of Total Market Value"] = (
#     newTMV  # wbH['% of Total Market Value'].sum(), check, should equal number of funds
# )

# # recalc the 'Current Exposure %' column per fund and then ...
# newCEp = []  # new Current Exposure % column
# for i, row in tqdm(
#     wbH.iterrows(), total=wbH.shape[0]
# ):  # https://stackoverflow.com/questions/47087741/use-tqdm-progress-bar-with-pandas
#     if nav[row["Entity ID"]] == 0:
#         currentexposurepct = 1
#     else:
#         currentexposurepct = row["Current Exposure"] / nav[row["Entity ID"]] * 100
#     newCEp.append(currentexposurepct)

# # ... replace the '% of Total Market Value' with the values in the new list
# wbH["Current Exposure %"] = (
#     newCEp  # wbH['% of Total Market Value'].sum(), check, should equal number of funds
# )

# print(f" {timediff(start_time, time.time())} recalculating and saving fund Total Market Value percentages","\n")


# convert date columns to datetime format
date_cols = ["i Position Effective Date", "Maturity Date", "Next Coupon Date"]
for date_col in date_cols:
    wbH[date_col] = pd.to_datetime(wbH[date_col])

# convert holdings numerical columns to numbers
num_cols = [
    "Original Nominal",
    "Clean Book Value",
    "Clean Market Value",
    "Accrued Income",
    "Dividend Receivable",
    "Sum of Market Value Income",
    "Market Price /Yield",
    r"% of Total Market Value",
    "Coupon",
    "Duration",
    "Modified Duration",
    "NACA Yield",
    "NACM Yield",
    "Weighted Avg NACA Yield",
    "Weighted Avg NACM Yield",
    "Market Value %",
    "Current Exposure",
    "Current Exposure %",
    "Weighted Average NACS Yield",
    "Weighted Average Coupon",
    "Weighted Modified Duration",
]
for num_col in num_cols:
    wbH[num_col] = wbH[num_col].astype(str).str.replace(",", "").astype(float)

# convert deltas numerical columns to numbers
num_cols_dervs = ["Nominal Holding", "Delta", "Market Value", "Effective Exposure"]
for num_col_derv in num_cols_dervs:
    wbD[num_col_derv] = wbD[num_col_derv].astype(str).str.replace(",", "").astype(float)

# write the summary, fund holdings, and derivative deltas to a workbook
summary_name = (
    pthEXPORTS + rf"\Derv {rptDate.strftime('%d%b%Y')}.xlsx"
)  # assign the file name
writer = pd.ExcelWriter(
    summary_name, engine="xlsxwriter"
)  # instantiate a sheet writer with file name
sorted_summary.to_excel(
    writer, index=False, sheet_name="Summary"
)  # write the summary sheet
wbH.to_excel(
    writer, index=False, sheet_name="Holdings"
)  # write the fund holdings sheet
wbD.to_excel(
    writer, index=False, sheet_name="Derv"
)  # write the derivative deltas sheet
writer.close()  # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets

# # write the summary, fund holdings, and derivative deltas to a workbook
# summary_name = pthEXPORTS + f'\Derv {rptDate.strftime("%d%b%Y")}.xlsx'      # assign the file name
# with pd.ExcelWriter(summary_name, engine = 'xlsxwriter') as writer:         # instantiate a sheet writer with file name
#     sorted_summary.to_excel(writer, sheet_name = 'Summary',  index = False) # write the summary sheet
#     wbH.to_excel(           writer, sheet_name = 'Holdings', index = False) # write the fund holdings sheet
#     wbD.to_excel(           writer, sheet_name = 'Derv',     index = False) # write the derivative deltas sheet
# writer.close() # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets
# # C:\Users\hilton.netta\AppData\Local\anaconda3\Lib\site-packages\xlsxwriter\workbook.py:368: UserWarning: Calling close() on already closed file.
# #  warn("Calling close() on already closed file.")

print(" ", summary_name)
# TEST ++++++++++++++++

print(
    f"\n  {timediff(start_time, time.time())} \
writing the dataframe to a sheet\n"
)

# sorted_summary

# prettify the summary sheet and add hyperlinks with xlwings

print(
    f"\nPrettifying and adding links to \
the summary sheet with xlwings ...\n"
)
start_time = time.time()

import xlwings as xw

if summ_yn != "No":
    wbS = xw.Book(pthEXPORTS + rf"\Derv {rptDate.strftime('%d%b%Y')}.xlsx")
    shtS = wbS.sheets["Summary"]  # derivative cover summary sheet
    # xl.DisplayAlerts = False                              # suppress Excel warning dialogues

    # format the headings row of the Summary file
    shtS["A1"].add_hyperlink(
        pthEXPORTS, f"Derivative Cover Calcs {rptDate.strftime('%d%b%Y')}"
    )
    shtS.range("A:A").column_width = 20.71
    shtS.range("B:B").column_width = 5.29
    shtS["W1"].add_hyperlink(
        r"P:\Investment Operations\GRC\Compliance\Client Mandates", "Fund Mandate"
    )
    shtS.range("W:W").column_width = 13.71
    shtS["X1"].add_hyperlink(pthSttlmnt, "Team")
    shtS.range("X:X").column_width = 11.71
    shtS["Y1"].value = "NAV"
    shtS["Y:Y"].number_format = "#,##0.00"
    shtS["Z1"].add_hyperlink(pthOverdrafts, "PIM Overdrafts")
    shtS["Z1"].api.HorizontalAlignment = -4108  # xlCenter
    shtS["Z1"].api.VerticalAlignment = -4160  # xlTop
    shtS[
        "A1:AA1"
    ].api.WrapText = True  # https://docs.xlwings.org/en/stable/missing_features.html
    shtS[
        "A1:AA1"
    ].font.bold = True  # https://docs.xlwings.org/en/stable/missing_features.html

    # add investment team names and links to fund mandates and calculation sheets
    start_time_links = time.time()

    for index, row in tqdm(
        sorted_summary[["Fund Code"]].iterrows(),
        total=sorted_summary[["Fund Code"]].shape[0],
        desc="Adding investment team names and \
links to fund mandates and calculation sheets ...",
    ):  # iterate over the funds
        # shtS['A' + str(index + 2)].value = f'{shtS["A" + str(index + 2)]} calc sheet'                                         # fund code
        shtS["A" + str(index + 2)].add_hyperlink(
            rf"{pthEXPORTS}\{row['Fund Code']} Derv Calc {rptDate.strftime('%d%b%Y')}.xlsx",
            f"{row['Fund Code']}",
        )  # link to calc sheet
        shtS["W" + str(index + 2)].add_hyperlink(
            rf"{pthMandates}\{row['Fund Code']} Rules.docx", f"{row['Fund Code']}"
        )  # link to fund mandate
    print(
        " ",
        f" {timediff(start_time_links, time.time())} adding \
investment team names and links to fund \
mandates and calculation sheets",
    )

    # add conditional formating for values that are negative or exceed 100% of NAV
    start_time_format = time.time()
    for a_cell in tqdm(
        shtS["D2:V2"].expand("down"),
        desc=f"Adding conditional formats for values < 0% or > 100% \
of NAV; {len(funds_cmpl) * 19:,.0f} \
= {len(funds_cmpl)} funds x 19 columns",
    ):
        if type(a_cell.value) in [float, int]:
            if a_cell.value < 0:
                a_cell.font.color = (
                    255,
                    0,
                    0,
                )  # red is (255,0,0) in RGB or #FF0000 in Hex
            if a_cell.value > 100 or a_cell.value < -100:
                a_cell.color = (255, 197, 255)  # light pink for cell colour
    print(
        " ",
        f" {timediff(start_time_format, time.time())} adding \
conditional formats for values that \
are negative or exceed 100% of NAV",
    )

if summ_yn != "No":
    wbS.save()
    wbS.close()

print(" ", pthEXPORTS + rf"\Derv {rptDate.strftime('%d%b%Y')}.xlsx")
print(
    f"\n {timediff(start_time, time.time())} prettifying \
and adding links to the summary sheet with xlwings\n",
)

# Save the summary dataframe to the derv_summary.xlsx template using xlwings

start_time = time.time()
print(
    f"Saving the summary dataframe to \
the derv_summary.xlsx template; \
{(len(funds_cmpl) + 1) * 19:,.0f} = ({len(funds_cmpl)} \
funds_cmpl + 1) x 19 columns"
)

# open the derv_summary.xlsx derv template and assign values to the holdings and deltas sheets
import xlwings as xw

with xw.App(visible=False) as app:
    # populate the summary sheet with derivative cover calc values for each fund
    wb = xw.Book(
        os.path.join(pthDaily, "derv_summary.xlsx")
    )  # open the derv calc workbook as an object
    shtS_S = wb.sheets[
        "Summary"
    ]  # assign sheet containing the funds derivative cover summary
    shtS_S.clear()  # clear the receiving holdings sheet
    shtS_S.range("A1").options(
        index=False
    ).value = sorted_summary  # paste fund holdings

    # conditional formatting
    for a_cell in tqdm(
        shtS_S["D1:V1"].expand("down"),
        desc=f"Adding conditional formats \
for {len(funds_cmpl)} funds",
    ):
        if type(a_cell.value) in [
            float,
            int,
        ]:  # ensure the cell bveing formatted is a float or an integer
            if abs(a_cell.value) >= 100:
                a_cell.color = (
                    255,
                    204,
                    255,
                )  # (255, 204, 255) or #FFCCFF is light pink
            elif a_cell.value < 0:
                a_cell.font.color = (255, 0, 0)  # (255,   0,   0) or #FF0000 is red

    # add fund derv calc hyperlinks
    for a_cell in tqdm(
        shtS_S["A2:A2"].expand("down"),
        desc=f"Adding hyperlinks to derivative \
calculation files for {len(funds_cmpl)} funds",
    ):
        a_cell.add_hyperlink(
            os.path.join(
                pthEXPORTS,
                f"{a_cell.value} Derv Calc \
{rptDate.strftime('%d%b%Y')}.xlsx",
            ),
            a_cell.value,
            screen_tip=None,
        )

    # add fund mandate hyperlinks
    for a_cell in tqdm(
        shtS_S["W2:W2"].expand("down"),
        desc=f"Adding hyperlinks to fund \
mandate files for {len(funds_cmpl)} funds",
    ):
        a_cell.add_hyperlink(
            os.path.join(pthMandates, f"{a_cell.value} rules.docx"),
            a_cell.value,
            screen_tip=None,
        )

    # add heading hyperlinks
    shtS_S["A1"].add_hyperlink(
        pthEXPORTS, f"{rptDate.strftime('%a %d %b %Y')} Calcs", screen_tip=None
    )
    shtS_S["W1"].add_hyperlink(pthMandates, "Fund Rules", screen_tip=None)
    shtS_S["Y1"].add_hyperlink(pthOverdrafts, "PIM Overdrafts", screen_tip=None)

    # make hyperlinked headings bold
    shtS_S["A1"].font.bold = True
    shtS_S["W1"].font.bold = True
    shtS_S["Y1"].font.bold = True

    # wrap text
    shtS_S["A1"].WrapText = True
    shtS_S["W1"].WrapText = True
    shtS_S["Y1"].WrapText = True

    # save the summary sheet and then close it
    wb.save(os.path.join(pthDaily, "derv_summary.xlsx"))  # save the file
    wb.close()

print(" ", os.path.join(pthDaily, "derv_summary.xlsx"))

print(
    f" {timediff(start_time, time.time())} saving \
the summary dataframe to the derv_summary.xlsx template\n",
)

# Delete contents of the temporary local folder
start_time = time.time()

local_folder_delete = "yes"
if local_folder_delete == "yes":
    for filename in tqdm(
        os.listdir(pthLOCAL),
        desc="Deleting \
contents of the local temporary folder",
    ):
        file_path = os.path.join(pthLOCAL, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Couldn't delete {file_path} because {e}")

print(
    f" {timediff(start_time, time.time())} deleting contents of \
the local temporary folder completed \n"
)

print(
    f" {timediff(start_time_0, time.time())} roundtrip time to \
summarise derivative calcs \n"
)

print("\n\n##############################################")
print("#                                            #")
print("#    END 3/4 derv_checker_summarising.py X   #")
print("#                                            #")
print("##############################################\n\n")
