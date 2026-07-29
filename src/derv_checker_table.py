#!/usr/bin/env python
# coding: utf-8

# # Compile the Derivative Cover Check Sheets

print("\n\n#######################################################")
print("#                                                     #")
print("#            START 2/4 derv_checker_table.py   X      #")
print("#                                                     #")
print("#######################################################\n\n")

# Libraries, libraries!

import time
from turtle import left

start_time = time.time()
start_time_derv_compiling = start_time
print("Importing libraries and setting paths ...")

import pandas as pd
import os
from datetime import timedelta
from tqdm import tqdm
from constants import pth_dl, pthSttlmnt, pthEXPORTS, pthOverdrafts
from utilities import timediff, prior_working_day, parn_de

print(
    f" {timediff(start_time, time.time())} importing libraries \
and setting paths\n"
)

# 1) Get report date, paths to the holdings and derivative
# metric files and selected summary sheet option

start_time = time.time()
print(
    "Getting the reporting date and latest downloaded \
holdings and derivatives files...\n"
)

# df = pd.read_excel(pthPy, sheet_name="arc", header=None, usecols="A,E").dropna(
#     subset=[0]
# )
# k = df.iloc[2, 1]
# rptDate = (
#     k if isinstance(k, datetime) else prior_working_day(datetime.today())
# )  # prior working day or report date override; has type datetime()
# summ_yn = df.iloc[3, 1]
# funds = df[0].iloc[1:]  # type is pandas Series

# # get paths to the expected holdings and derivative metric files
# fPARN = os.path.join(
#     pth_dl,
#     f"PARN ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )
# fDE = os.path.join(
#     pth_dl,
#     f"DERV ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
# )

# # check if the required files have been downloaded, else exit
# if not os.path.exists(fPARN) or not os.path.exists(fDE):
#     sys.exit(
#         f"\n\nStopping: missing expected download(s):\n"
#         f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
#         f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}\n\n"
#     )

fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()

# check if the required files have been downloaded, else continue
if not os.path.exists(fPARN) or not os.path.exists(fDE):
    sys.exit(
        f"Stopping: missing expected download(s):\n"
        f"  {os.path.basename(fPARN)} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
        f"  {os.path.basename(fDE)} which {'exists' if os.path.exists(fDE) else 'does not exist'}\n"
    )

# present report date and funds
print(
    f" {rptDate.strftime('%A %d %b %Y')} for {len(funds)} funds:\n",
    f"{(', ').join(funds.tolist())}",
)

td = rptDate + timedelta(days=397)  # to determine < 13-month maturities
twoA = pd.read_excel(pthSttlmnt, sheet_name="Funds", usecols="A:B, D:E")
# columns are 'Fund Code', 'Fund Name', 'UT', 'Team'

print(f" A summary sheet is{' not' if summ_yn == 'No' else ''} required\n")
print(
    f" {timediff(start_time, time.time())} getting the reporting \
date and latest downloaded holdings and derivatives files\n",
)

# 2) Dataframe the holdings and deltas csv files
# and convert numerical columns from str to float

start_time = time.time()
print("Creating holdings and deltas dataframes ...")

# dataframe of the holdings and deltas csv files
# wbH = pd.read_csv(parN_nm)
wbH = pd.read_csv(fPARN)
wbD = pd.read_csv(fDE)

# 4) Get the fund names and fund codes
# from the holdings dataframe
fnames = wbH["Entity Name"].unique()
fcodes = wbH["Entity ID"].unique()

# 5) Convert derivative dataframe columns from str to float
# https://stackoverflow.com/questions/55557004/getting-attributeerror-float-object-has-no-attribute-replace-error-while
headsD = ["Nominal Holding", "Delta", "Market Value", "Effective Exposure"]
for head in headsD:
    wbD[head] = [
        str(x).replace(",", "").replace("-", "-") for x in wbD[head]
    ]  # KeyError: 'Nominal Holding'
    wbD[head] = wbD[head].astype(float)

# 6) Convert holdings dataframe columns from str to float
headsH = [
    "Current Exposure",
    "Sum of Market Value Income",
    "Original Nominal",
    r"Market Price /Yield",
]
for head in headsH:
    wbH[head] = [str(x).replace(",", "").replace("-", "-") for x in wbH[head]]
    wbH[head] = wbH[head].astype(float)

# 7) Convert holdings date column
# from type string to type datetime
wbH["i Position Effective Date"] = pd.to_datetime(wbH["i Position Effective Date"])

# 8) Convert maturity date column from type object to
# type datetime and 'NaT' to a long date in datetime format
wbH["Maturity Date"] = pd.to_datetime(wbH["Maturity Date"])
# print(wbH.columns)


# function to convert 'NaT' to the report date in datetime format
def convert_NaT_to_report_date(dte):
    if str(dte) == "NaT":
        # return datetime(2099, 12, 31)
        return rptDate
    else:
        return dte


# 9a) For SYTHs, change 'UNKNOWN' in Valuation First Level and
# Valuation Second Level to 'SYTH' where Investment Type is 'SYTH'
syth_unkown_filter = (wbH["Valuation Second Level"] == "UNKNOWN") & (
    wbH["Investment Type"] == "SYTH"
)
wbH.loc[syth_unkown_filter, ["Valuation First Level", "Valuation Second Level"]] = (
    "SYTH"
)

# 9b) For TRSes, change 'UNKNOWN' in Valuation First Level to
# 'SWAPS' where they are not SYTH or empty portfolio holdings
trs_unknown_filter = (wbH["Valuation First Level"] == "UNKNOWN") & (
    wbH["Sub Security Type"] == "TRS"
)
trs_unknowns = wbH.loc[trs_unknown_filter]
print(
    f'  {len(trs_unknowns)} "UNKNOWN" \
TRSes: {(", ").join(trs_unknowns["PrimaryAssetID"].tolist())}'
)
wbH.loc[trs_unknown_filter, ["Valuation First Level", "Valuation Second Level"]] = (
    "SWAPS"
)
print(
    f'  {len(trs_unknowns)} "UNKNOWN" \
TRSes: {(", ").join(trs_unknowns["PrimaryAssetID"].tolist())}'
)

# 9c) For derivatives, change 'UNKNOWN' in Valuation Second Level to
# 'SWAPS' where they are not SYTH or empty portfolio holdings
valuation_2nd_level_clean = (
    wbH["Valuation Second Level"].astype(str).str.strip().str.upper()
)
sub_security_type_clean = wbH["Sub Security Type"].astype(str).str.strip().str.upper()
eqty_derv_unknown_filter = (valuation_2nd_level_clean == "UNKNOWN") & (
    sub_security_type_clean.isin(["EQUITY - FUTURE", "CFD"])
)
eqty_derv_unknowns = wbH.loc[eqty_derv_unknown_filter].copy()
# print(
#     f'  {len(eqty_derv_unknowns)} "UNKNOWN" \
# Equity Derivatives: {(", ").join(eqty_derv_unknowns["PrimaryAssetID"].tolist())}'
# )
wbH.loc[eqty_derv_unknown_filter, ["Valuation Second Level"]] = "Equity Derivatives"
# print(
#     f'  {len(eqty_derv_unknowns)} "UNKNOWN" \
# Equity Derivatives: {(", ").join(eqty_derv_unknowns["PrimaryAssetID"].tolist())}'
# )

bond_derv_unknown_filter = (wbH["Valuation Second Level"] == "UNKNOWN") & (
    wbH["Sub Security Type"] == "BNDFUT"
)
wbH.loc[bond_derv_unknown_filter, ["Valuation Second Level"]] = "Bond Derivatives"

# 9d) Identify all remaining "UNKNOWN" securities
all_unknown_filter = (wbH["Valuation First Level"] == "UNKNOWN") | (
    wbH["Valuation Second Level"] == "UNKNOWN"
)
all_unknowns = wbH.loc[all_unknown_filter]
all_defaults = wbH.loc[wbH["PrimaryAssetID"] == "Default"]
print(
    f'  {len(all_unknowns) - len(all_defaults)} "UNKNOWN" \
securities after including {len(all_defaults)} "Default" \
securities in {(", ").join(all_defaults["Entity Name"].tolist())}'
)

# 9e) Sort holdings
wbH.sort_values(
    by=["Entity Name", "Valuation First Level", "Valuation Second Level", "CCY"],
    inplace=True,
)

# 9f) Within each Entity Name, move repo holdings ("RPCA", "RPCO", "RPMT"
# PrimaryAssetID prefixes) to the top, grouped alphabetically by PrimaryAssetID
repo_prefixes = ("RPCA", "RPCO", "RPMT")
is_repo = wbH["PrimaryAssetID"].astype(str).str.startswith(repo_prefixes)
repo_sort_key = (~is_repo).astype(int)  # repo rows (0) sort before the rest (1)
repo_id_key = wbH["PrimaryAssetID"].where(
    is_repo, ""
)  # alphabetical within repo rows only

wbH["_repo_sort_key"] = repo_sort_key
wbH["_repo_id_key"] = repo_id_key
wbH.sort_values(
    by=["Entity Name", "_repo_sort_key", "_repo_id_key"],
    kind="stable",
    inplace=True,
)
wbH.drop(columns=["_repo_sort_key", "_repo_id_key"], inplace=True)

# 10) Convert holdings date columns to datetime format
date_cols = ["i Position Effective Date", "Maturity Date", "Next Coupon Date"]
for date_col in date_cols:
    wbH[date_col] = pd.to_datetime(wbH[date_col])

# 11) Convert holdings numerical columns to numbers
num_col_names = [
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
for num_col_name in num_col_names:
    wbH[num_col_name] = wbH[num_col_name].astype(str).str.replace(",", "").astype(float)

# 12) Convert derivative deltas numerical columns to numbers
num_col_names_dervs = ["Nominal Holding", "Delta", "Market Value", "Effective Exposure"]
for num_col_name_derv in num_col_names_dervs:
    wbD[num_col_name_derv] = (
        wbD[num_col_name_derv].astype(str).str.replace(",", "").astype(float)
    )

print(
    f" {timediff(start_time, time.time())} creating \
holdings and deltas dataframes\n",
)

# 13) change the '% of Total Market Value" column (N) to the fund-
# specific % based on 'Sum of Market Value Income' column (M)
start_time = time.time()
navs = wbH.groupby("Entity ID")[
    "Sum of Market Value Income"
].sum()  # (column N) this has type 'pandas.core.series.Series'
nav = navs.to_dict()  # nav series changed to dictionary to make it lookupable

# 14) Recalc the '% of Total Market Value' column per fund and then ...
newTMV = []
for i, row in tqdm(
    wbH.iterrows(),
    desc="Recalculating Total Market Value percentages",
    total=wbH.shape[0],
):
    if nav[row["Entity ID"]] == 0:
        fndpct = 100
    else:
        fndpct = row["Sum of Market Value Income"] / nav[row["Entity ID"]] * 100
    newTMV.append(fndpct)

# 15) ... replace the '% of Total Market Value' with the values in the new list
wbH[r"% of Total Market Value"] = newTMV

# 16) Recalc the 'Current Exposure %' column per fund and then ...
newCEp = []  # new Current Exposure % column
for i, row in tqdm(
    wbH.iterrows(),
    desc="Recalculating Current Exposure percentages",
    total=wbH.shape[0],
):
    if nav[row["Entity ID"]] == 0:
        currentexposurepct = 1
    else:
        currentexposurepct = row["Current Exposure"] / nav[row["Entity ID"]] * 100
    newCEp.append(currentexposurepct)

# 17) ... replace the '% of Total Market Value' with the values in the new list
wbH["Current Exposure %"] = newCEp

print(
    f" {timediff(start_time, time.time())} recalculating \
and saving fund Total Market Value percentages\n"
)

start_time_compiling = time.time()

# ######### TEST #########
# funds = ["PABS"]  # funds = ["PIMBAL", "PIPF", "GTCWP2", "GAEMBF"]
# for index, fund in enumerate(
#     tqdm(
#         funds,
#         desc=f"Compiling derivative calculations \
# for {len(funds)} funds at {rptDate.strftime('%d %b %Y')} ...",
#     )
# ):
#     # fcode = fund
#     # per fund, get holdings and derivatives sub-dataframe
#     ftyp = twoA[twoA["Fund Code"] == fund].iloc[0, 3]
#     fund_eagle_name = wbH[wbH["Entity ID"] == fund].iloc[0, 0]
#     hold = wbH[wbH["Entity ID"] == fund]
#     delt = wbD[wbD["Entity Name"] == fund_eagle_name]

#     fund_name = twoA[twoA["Fund Code"] == fund].iloc[0, 1]
#     print(
#         f"\n\n{index}: {fund_eagle_name}\n{fund_name} ({fund}), a {ftyp}\n{hold.shape}\n{delt.shape}\n"
#     )
# ######### TEST #########

# create an empty dataframe to hold the summary data if required
dfSummary = pd.DataFrame()
summary_cols = ["Fund Code"]
# print(dfSummary.columns)
# print(type(funds), funds, len(funds))

s = "" if len(funds) <= 1 else "s"
for index, fund in enumerate(
    tqdm(
        funds,
        desc=f"Compiling derivative calculations \
for {len(funds)} fund{s} at \
{rptDate.strftime('%a %d %b %Y')}",
    )
):
    start_time = time.time()

    # get fund static
    fund_eagle_name = wbH[wbH["Entity ID"] == fund].iloc[0, 0]
    hold = wbH[wbH["Entity ID"] == fund]
    delt = wbD[wbD["Entity Name"] == fund_eagle_name]
    ftyp = twoA[twoA["Fund Code"] == fund].iloc[0, 3]
    fund_name = twoA[twoA["Fund Code"] == fund].iloc[0, 1]
    # print(
    #     f"\n\n{index}: {fund_eagle_name}\n{fund_name} ({fund}), a {ftyp}\n{hold.shape}\n{delt.shape}\n"
    # )

    dervs = (
        (hold["Valuation First Level"] == "DERIVATIVES").sum()
        + (hold["Valuation First Level"] == "FORWARDS").sum()
        + (hold["Valuation First Level"] == "FORWARD RATE AGREEMENT").sum()
        + (hold["Security Type"] == "SWP").sum() / 3
        + (hold["Security Type"] == "SWAP").sum() / 3
        + hold["PrimaryAssetID"].str.startswith("RPMT").sum()
    )  # number of derivatives, catering for the three legs of a FRA

    nav = hold["Sum of Market Value Income"].sum()  # fund NAV

    cash = hold[
        (hold["Valuation First Level"] == "CASH")
        & (hold["Security Type"] != "CIS")
        & (hold["Sub Security Type"] != "REPO")
        & (~hold["i Issue Name"].str.upper().str.contains("MARGIN"))
    ][
        "Current Exposure"
    ].sum()  # cash incl accruals, excl MMFs, excl margins, excl repo P/L

    # mmfs = hold[
    #     (hold["PrimaryAssetID"] == "PRMFB3") | (hold["PrimaryAssetID"] == "PCMMB3")
    # ]["Current Exposure"].sum()  # + \
    # #

    mmfs = hold[hold["PrimaryAssetID"].isin(["PRMFB3", "PCMMB3"])][
        "Current Exposure"
    ].sum()
    #

    non_mmfs = (
        hold[
            (~hold["PrimaryAssetID"].isin(["PRMFB3", "PCMMB3", "PIMEVOA", "PIMIDFA"]))
            & (hold["Security Type"].isin(["CIS", "ETF"]))
        ]["Current Exposure"].sum()
        if ftyp not in ("UT", "ETF")
        else 0
    )

    mmis = (
        hold[
            (hold["Valuation First Level"] == "MONEY MARKET")
            & (hold["Maturity Date"] < td)
            & (hold["Sub Security Type"] != "CLN")
        ]["Current Exposure"].sum()
        if ftyp in ("UT", "ETF")
        else hold[hold["Valuation First Level"] == "MONEY MARKET"][
            "Current Exposure"
        ].sum()
    )  # money market instruments

    # mmis_non_UT = (
    #     hold[
    #         (hold["Valuation First Level"] == "MONEY MARKET")
    #         & (hold["Sub Security Type"] != "CLN")
    #     ]["Current Exposure"].sum()
    # )  # money market instruments for non-UTs

    clns = (
        hold[hold["Sub Security Type"] == "CLN"]["Current Exposure"].sum()
        if ftyp not in ("UT", "ETF")
        else 0
    )

    bonds = (
        hold[
            (hold["Valuation First Level"] == "BONDS")
            & (~hold["Sub Security Type"].isin(["CLN", "REPO"]))
            & (hold["Maturity Date"] < td)
        ]["Current Exposure"].sum()
        if ftyp in ("UT", "ETF")
        else hold[
            (hold["Valuation First Level"] == "BONDS")
            & (~hold["Sub Security Type"].isin(["REPO"]))
        ]["Current Exposure"].sum()
    )

    marg_jse = (
        hold[hold["PrimaryAssetID"] == "SAFEX"]["Current Exposure"].sum()
        + hold[hold["PrimaryAssetID"] == "VARMARG"]["Current Exposure"].sum()
    )  # JSE SAFEX and JSE VARMARG

    marg_otc = hold[
        (hold["i Issue Name"].str.upper().str.contains("MARGIN"))
        & (hold["PrimaryAssetID"] != "SAFEX")
        & (hold["PrimaryAssetID"] != "VARMARG")
    ]["Current Exposure"].sum()  # other margins, non-JSE margin accounts

    crry_derv = hold[
        hold["Valuation Second Level"].str.upper().str.contains("CURRENCY DERIVATIVES")
    ]["Current Exposure"].sum()

    # dervs

    # other_UTs = (
    #     0
    #     if twoA[twoA["Fund Code"] == fund]["UT"].iloc[0] == "UT"
    #     else hold[
    #         (hold["Security Type"] == "CIS")
    #         & (hold["Sub Security Type"] != "CSH")
    #         & (~hold["PrimaryAssetID"].isin(["PIMEVOA", "PIMIDFA"]))
    #     ]["Current Exposure"].sum()
    # )  # other, non-MMF, UTs

    # count_other_UTs = (
    #     0
    #     if twoA[twoA["Fund Code"] == fund]["UT"].iloc[0] == "UT"
    #     else hold[
    #         (hold["Security Type"] == "CIS")
    #         & (hold["Sub Security Type"] != "CSH")
    #         & (~hold["PrimaryAssetID"].isin(["PIMEVOA", "PIMIDFA"]))
    #     ]["Current Exposure"].count()
    # )  # other, non-MMF, UTs

    # other_ETFs = (
    #     0
    #     if twoA[twoA["Fund Code"] == fund]["UT"].iloc[0] == "UT"
    #     else hold[
    #         (hold["Security Type"] == "ETF")
    #         & (hold["Sub Security Type"] != "CSH")
    #     ]["Current Exposure"].sum()
    # )  # other, non-MMF, ETFs

    # count_other_ETFs = (
    #     0
    #     if twoA[twoA["Fund Code"] == fund]["UT"].iloc[0] == "UT"
    #     else hold[
    #         (hold["Security Type"] == "ETF")
    #         & (hold["Sub Security Type"] != "CSH")
    #     ]["Current Exposure"].count()
    # )  # other, non-MMF, ETFs

    ##################################################
    # define the derivative cover components

    fwds = hold[hold["Valuation Second Level"].str.upper().str.contains("FORWARDS")][
        "Current Exposure"
    ].sum()

    swps = hold[
        (hold["Valuation Second Level"] == "SWAPS")
        & (hold["Sub Security Type"] != "TRS")
    ]["Current Exposure"].sum()  # total return swaps

    repo = (
        hold[hold["PrimaryAssetID"].str.upper().str.startswith("RPCO", na=False)][
            "Current Exposure"
        ].sum()
        + hold[hold["PrimaryAssetID"].str.upper().str.startswith("RPCA", na=False)][
            "Current Exposure"
        ].sum()
    )
    # repo_gain = max(0, repo)  # net profit on repos
    # repo_loss = min(0, repo)  # net loss on repos

    fras = (
        hold[hold["Valuation First Level"] == "FORWARD RATE AGREEMENT"][
            "Original Nominal"
        ]
        .fillna(0)
        .dot(
            hold[hold["Valuation First Level"] == "FORWARD RATE AGREEMENT"][
                r"Market Price /Yield"
            ].fillna(0)
        )
    )  # FRAs as a dot product of nominals and prices

    trs = hold[(hold["Sub Security Type"] == "TRS")]["Current Exposure"].sum()
    # net negative mtm on OTC derivativesotcs      = fwds + min(0, otc_mtm) + repo_loss + fras

    otc_mtm = min(0, fwds) + min(0, swps) + min(0, repo) + min(0, fras) + min(0, trs)

    eqty_futs_sa = hold[
        (hold["Valuation Second Level"] == "Equity Derivatives")
        & (hold["CCY"] == "ZAR")
    ]["Current Exposure"].sum()  # SA equity futures

    eqty_futs_frgn = hold[
        (hold["Valuation Second Level"] == "Equity Derivatives")
        & (hold["CCY"] != "ZAR")
    ]["Current Exposure"].sum()  # ex_SA equity futures

    eqty_fut_frgn_mtm = hold[
        (hold["Valuation Second Level"] == "Equity Derivatives")
        & (hold["CCY"] != "ZAR")
    ]["Sum of Market Value Income"].sum()  # mark-to-market p&l on equity futures

    bond_futs_sa = hold[
        (hold["Valuation Second Level"] == "Bond Derivatives") & (hold["CCY"] == "ZAR")
    ]["Current Exposure"].sum()  # SA bond futures

    bond_futs_frgn = hold[
        (hold["Valuation Second Level"] == "Bond Derivatives") & (hold["CCY"] != "ZAR")
    ]["Current Exposure"].sum()  # ex-SA bond futures

    bond_futs_frgn_mtm = hold[
        (hold["Valuation Second Level"] == "Bond Derivatives") & (hold["CCY"] != "ZAR")
    ]["Sum of Market Value Income"].sum()  # mark-to-market p&l on bond futures

    crry_futs = hold[(hold["Valuation Second Level"] == "Currency Derivatives")][
        "Current Exposure"
    ].sum()  # currency futures

    ailf = (
        cash + mmfs + mmis + bonds + clns + max(0, repo) + marg_jse + marg_otc
    )  # total assets in liquid form

    # total OTC derivatives

    eqty_futs = hold[(hold["Valuation Second Level"] == "Equity Derivatives")][
        "Current Exposure"
    ].sum()  # equity futures

    bond_futs = hold[(hold["Valuation Second Level"] == "Bond Derivatives")][
        "Current Exposure"
    ].sum()  # ex-SA bond futures

    frgn_futs = 0

    otcs = fwds + min(0, swps) + min(0, repo) + min(0, fras) + min(0, trs)

    # frgn_trs  = hold[(hold['Sub Security Type'] == 'TRS') & (hold['CCY'] != 'ZAR')]['Current Exposure'].sum() # profit or loss on total return swaps

    lstd_drvs = eqty_futs + bond_futs + eqty_fut_frgn_mtm + bond_futs_frgn_mtm

    eqty_fut_frgn_mtm = -hold[
        (hold["Valuation Second Level"] == "Equity Derivatives")
        & (hold["CCY"] != "ZAR")
    ]["Sum of Market Value Income"].sum()  # mark-to-market p&l on equity futures

    bond_futs_frgn_mtm = -hold[
        (hold["Valuation Second Level"] == "Bond Derivatives") & (hold["CCY"] != "ZAR")
    ]["Sum of Market Value Income"].sum()  # mark-to-market p&l on bond futures

    # leverage
    excl = ["CASH", "MONEY MARKET", "UNKNOWN", "SYTH"]
    lvg_g = (
        hold[~hold["Valuation First Level"].isin(excl)]["Current Exposure"].abs().sum()
        / nav
        * 100
    )  # leverage gross
    lvg_c = (
        hold[~hold["Valuation First Level"].isin(excl)]["Current Exposure"].sum()
        / nav
        * 100
    )  # leverage commitment (net)

    eqty_cover_sa = hold[
        (hold["Valuation First Level"] == "EQUITIES") & (hold["CCY"] == "ZAR")
    ]["Current Exposure"].sum()  # directly held SA equity

    eqty_cover_frgn = hold[
        (hold["Valuation First Level"] == "EQUITIES") & (hold["CCY"] != "ZAR")
    ]["Current Exposure"].sum()  # directly held foreign equity

    bonds_cover_sa = hold[
        (hold["Valuation First Level"] == "BONDS") & (hold["CCY"] == "ZAR")
    ]["Current Exposure"].sum()  # directly held SA bonds

    bonds_cover_frgn = hold[
        (hold["Valuation First Level"] == "BONDS") & (hold["CCY"] != "ZAR")
    ]["Current Exposure"].sum()  # directly held foreign bonds

    net_eff_exp = -min(
        0, hold[hold["Investment Type"] == "SYTH"]["Current Exposure"].sum()
    )
    short_cash = -min(
        0,
        hold[
            (
                (hold["Valuation First Level"] == "CASH")
                & (~hold["PrimaryAssetID"].isin(["PRMFB3", "PCMMB3"]))
            )
        ]["Current Exposure"].sum(),
    )  # cash excl MMFs < 0

    otc_mtm = min(0, fwds) + min(0, swps) + min(0, repo) + min(0, fras) + min(0, trs)

    listed_dervs = (
        max(0, eqty_futs_sa)
        + max(0, eqty_futs_frgn)
        + eqty_fut_frgn_mtm
        + max(0, bond_futs_sa)
        + max(0, bond_futs_frgn)
        + bond_futs_frgn_mtm
        + max(0, crry_futs)
    )

    # assigning values to the summary dataframe
    dfSummary.at[index, "Fund Code"] = fund
    dfSummary.at[index, "UT?"] = ftyp
    dfSummary.at[index, "#"] = dervs
    dfSummary.at[index, "Cash Cover for UT"] = ailf + otc_mtm - listed_dervs + non_mmfs
    dfSummary.at[index, "AiLF"] = ailf
    dfSummary.at[index, "Cash"] = cash
    dfSummary.at[index, "MMFs"] = mmfs
    dfSummary.at[index, "MMIs"] = mmis
    dfSummary.at[index, "Bonds"] = bonds
    dfSummary.at[index, "CLNs"] = clns
    dfSummary.at[index, "Repo Net Gain"] = max(0, repo)
    dfSummary.at[index, "Margin JSE"] = marg_jse
    dfSummary.at[index, "Margin OTC"] = marg_otc
    dfSummary.at[index, "Forwards"] = min(0, fwds)
    dfSummary.at[index, "Swaps"] = min(0, swps)
    dfSummary.at[index, "Repo Net Loss"] = min(0, repo)
    dfSummary.at[index, "FRAs"] = min(0, fras)
    dfSummary.at[index, "TRSes"] = min(0, trs)
    dfSummary.at[index, "\u2211 OTC Derivatives"] = otc_mtm
    dfSummary.at[index, "Cover for OTC Derivatives"] = ailf + otc_mtm
    dfSummary.at[index, "Equity Futures SA"] = -max(0, eqty_futs_sa)
    dfSummary.at[index, "Equity Futures ex-SA"] = -max(0, eqty_futs_frgn)
    dfSummary.at[index, "Equity Futures ex-SA MtM P&L"] = -eqty_fut_frgn_mtm
    dfSummary.at[index, "Bond Futures SA"] = -max(0, bond_futs_sa)
    dfSummary.at[index, "Bond Futures ex-SA"] = -max(0, bond_futs_frgn)
    dfSummary.at[index, "Bond Futures ex-SA MtM P&L"] = -bond_futs_frgn_mtm
    dfSummary.at[index, "Currency Futures"] = -max(0, crry_futs)
    dfSummary.at[index, "\2211 Listed Derivatives"] = listed_dervs
    dfSummary.at[index, "Cover for All Derivatives"] = ailf + otc_mtm - listed_dervs
    dfSummary.at[index, "Incl CLNs & longer-dated debt"] = bonds + clns
    dfSummary.at[index, "Non_MMFs"] = non_mmfs
    dfSummary.at[index, "SA Equity Indices"] = eqty_futs_sa
    dfSummary.at[index, "Ex-SA Equity Indices"] = eqty_futs_frgn
    dfSummary.at[index, "SA Bond Indices"] = bond_futs_sa
    dfSummary.at[index, "Ex-SA Bond Indices"] = bond_futs_frgn
    dfSummary.at[index, "SA Bond Cover"] = bond_futs_sa
    dfSummary.at[index, "Ex-SA Bond Cover"] = bond_futs_frgn
    dfSummary.at[index, "Foreign Equity Cover"] = eqty_cover_frgn
    dfSummary.at[index, "Foreign Bond Cover"] = bonds_cover_frgn
    dfSummary.at[index, "Local Equity Cover"] = eqty_cover_sa
    dfSummary.at[index, "Total Equity"] = eqty_cover_sa + eqty_cover_frgn
    dfSummary.at[index, "Total Foreign"] = eqty_cover_frgn + bonds_cover_frgn
    dfSummary.at[index, "Cash Cover for non-UT"] = (
        ailf - otc_mtm - listed_dervs + non_mmfs
    )
    dfSummary.at[index, "Global Exposure"] = ailf + non_mmfs
    dfSummary.at[index, "Net Effective Exposure"] = net_eff_exp + short_cash
    dfSummary.at[index, "Fund Mandate"] = fund
    dfSummary.at[index, "Team"] = twoA[twoA["Fund Code"] == fund]["Team"].iloc[0]
    dfSummary.at[index, "Leverage (Gross)"] = lvg_g
    dfSummary.at[index, "Leverage (Net)"] = lvg_c
    dfSummary.at[index, "NAV"] = nav

    # normalise the summary metrics to the fund's NAV
    excl_cols = [
        "Fund Code",
        "UT?",
        "#",
        "Fund Mandate",
        "Team",
        "PIM Overdrafts",
        "NAV",
    ]
    for col in dfSummary.columns:
        if col not in excl_cols:
            dfSummary.at[index, col] = dfSummary.at[index, col] / nav * 100

# print(dfSummary.columns, "\n", dfSummary.shape)

# look up prior day's NAV
ystdy_date = prior_working_day(rptDate).strftime("%Y%m%d")
ystdy_filepath = pthEXPORTS + rf"\{ystdy_date}_derv_calc.xlsx"
ystdy = pd.read_excel(ystdy_filepath, usecols=["Fund Code", "NAV"]).dropna()
# print(ystdy.columns, "\n", ystdy.shape)
ystdy = ystdy.drop_duplicates(subset=["Fund Code"])
ystdy.head(15)
print(ystdy.columns, "\n", ystdy.shape)
# print(dfSummary.columns, "\n", dfSummary.shape)

# insert yesterday's NAV in the summary derivative calc dataframe

dfSummary = pd.merge(
    dfSummary, ystdy, on="Fund Code", how="left", suffixes=("_x", "_y")
)
# print(dfSummary.columns, "\n", dfSummary.shape)  # length changes from 153 to 281

y_date = prior_working_day(rptDate).strftime("%a %d %b %Y")
dfSummary = dfSummary.rename(columns={"NAV_x": "NAV", "NAV_y": f"NAV {y_date}"})
# print(dfSummary.columns, "\n", dfSummary.shape)  # length changes from 153 to 281

# prior NAV column name
prior_nav = dfSummary[f"NAV {y_date}"]
# print(prior_nav.shape)

# add a fund name column
dfSummary = pd.merge(
    dfSummary, twoA[["Fund Code", "Fund Name"]], on="Fund Code", how="left"
)
# print(dfSummary.columns, "\n", dfSummary.shape)

# suffix fund code to fund name
dfSummary["Fund Name"] = dfSummary["Fund Name"] + "(" + dfSummary["Fund Code"] + ")"
# print(dfSummary.head(3))

# if it's available, adjoin bank balances to dfSummary
bank_file_sa = pthOverdrafts + rf"\{rptDate.strftime('%Y%m%d')}_overdrafts_sa.xlsx"
if not os.path.isfile(bank_file_sa):
    print(
        "Bank balances not saved so will not \
be appended to the derivative summary"
    )
    pass
else:
    print(
        "Bank balances will be appended \
to the derivative summary"
    )
    bank = pd.read_excel(
        bank_file_sa, usecols="A,C,D,G", sheet_name=0, header=0, skiprows=range(11)
    )
    bank = bank[bank["Currency"] == "ZAR"]
    # print(bank.columns, "\n", bank.shape)
    bank["Value Date"] = pd.to_datetime(bank["Value Date"])
    bank_cols = {"Client ID": "Fund Code", "Closing Balance BNK": "Bank Bal"}
    bank.rename(columns=bank_cols, inplace=True)
    # print(bank.columns, "\n", bank.shape)
    dfSummary = dfSummary.merge(
        bank[["Fund Code", "Bank Bal"]], on="Fund Code", how="left"
    )
    # print(dfSummary.columns, "\n", dfSummary.shape)

# if it's available, adjoin cash activities to dfSummary
fCACT = os.path.join(
    pth_dl,
    f"CACT ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)
if not os.path.isfile(fCACT):
    print(f"  {fCACT} does not exist")
    pass
else:
    # 1. Load cash activities dataframe
    wbC = pd.read_csv(fCACT)

    # 2. Convert date and cash columns to
    # date and float respectively
    wbC_date_cols = ["Settlement Date", "Accounting Date"]
    for col in wbC_date_cols:
        wbC[col] = pd.to_datetime(wbC[col]).dt.date

    wbC_cash_cols = ["Base Amount", "Local Amount"]
    # wbC[wbC_date_cols] = pd.to_datetime(wbC[wbC_date_cols])
    for col in wbC_cash_cols:
        wbC[col] = wbC[col].astype(str).str.replace(",", "").astype(float)

    # 3. Sum Base (i.e., ZAR) Amount per fund for external flows
    flow_types = [
        "WITHDRAWAL",
        "CONTRIBUTION",
        "ADDSUBSCRIPTION",
        "ADDREDEMPTION",
        "DISTRIBUTION",
        "CASHXFER",
        "CASHDIV",
        "MISCINC",
    ]

    # create a filter for accounting date to be report date and for flow types
    mask_wbC = (
        pd.to_datetime(wbC["Accounting Date"]).dt.date == pd.to_datetime(rptDate).date()
    ) & (wbC["Transaction Type"].isin(flow_types))

    # apply the filter for accounting date and flow types
    cfl = wbC[mask_wbC]

    # sum the flow types per fund
    cfl = cfl.groupby(["Entity ID"], as_index=False)["Base Amount"].sum()
    # print(cfl.shape,"\n",cfl.columns)

    # TEST for sum
    cfl[cfl["Entity ID"] == "PETFIP"][["Entity ID", "Base Amount"]]

    # get the unique "Entity ID" entries from wbH
    wbH_unq = wbH.drop_duplicates(subset=["Entity ID"])[["Entity Name", "Entity ID"]]

    # print(wbC.shape, "\n", wbH_unq.shape)

    # 4. Merge to get Entity ID and fund name
    cfl = pd.merge(
        cfl,
        wbH_unq,
        on="Entity ID",
        how="right",
    ).rename(columns={"Entity ID": "Fund Code", "Base Amount": "Net In(Out) Flow"})

    # # TEST, should come to (#funds, 3)
    # print(wbC.columns,"\n", wbC.shape)

    # 4. Add a Net In(Out) Flow xolumn to the summary dataframe
    dfSummary = dfSummary.merge(
        cfl[["Fund Code", "Net In(Out) Flow"]],
        on="Fund Code",
        how="left",
    )

    # # TEST, should come to (#funds, 45)
    # print(dfSummary.columns, "\n", dfSummary.shape)

    # calculate % in(out)flow
    dfSummary["Net In(Out) Flow (%)"] = (
        (dfSummary["Net In(Out) Flow"] / prior_nav) * 100
    ).mask(prior_nav.isna(), "fresh flow")
    # # TEST, should come to (#funds, 45)
    # print(dfSummary.columns, "\n", dfSummary.shape)

    # add a new column showing change in NAV
    dfSummary["Net In(Out) Flow"] = dfSummary["Net In(Out) Flow"].fillna(0)
    dfSummary["\u0394 NAV (%)"] = (
        (dfSummary["NAV"] - dfSummary["Net In(Out) Flow"]) / prior_nav - 1
    ) * 100

    # move "\u0394 NAV (%)" and "Net In(Out)flow (%)" to sit just before the "AiLF" column
    cols = [
        c
        for c in dfSummary.columns
        if c not in ("\u0394 NAV (%)", "Net In(Out) Flow (%)")
    ]  # "\u0394" makes a delta sumbol
    ailf_pos = cols.index("AiLF")
    cols[ailf_pos:ailf_pos] = ["\u0394 NAV (%)", "Net In(Out) Flow (%)"]
    dfSummary = dfSummary[cols]

# sort the summary dataframe and save it to a new Excel file
start_time = time.time()

summary = dfSummary.sort_values(
    by="Cash Cover for UT", ascending=True
)  # sort the cover calc dataframe by 'Cash Cover for UT' in ascending order

ut_types = ["UT", "≠UT", "UCITS", "SAA", "TAA", "ICAV"]
sorted_summary = pd.DataFrame([])  # empty dataframe
for ut_type in ut_types:  # stack the > 0 derivative funds first ...
    summary_subset = summary[(summary["UT?"] == ut_type) & (summary["#"] != 0)]
    sorted_summary = pd.concat([sorted_summary, summary_subset])
    # sorted_summary = sorted_summary.sort_values(by = 'Cash Cover for UT', ascending = False) # sort the summary dataframe

for ut_type in ut_types:  # ... then stack the no derivative funds
    summary_subset = summary[(summary["UT?"] == ut_type) & (summary["#"] == 0)]
    sorted_summary = pd.concat([sorted_summary, summary_subset])
    # sorted_summary = sorted_summary.sort_values(by = 'Cash Cover for UT', ascending = False) # sort the summary dataframe

sorted_summary.reset_index(inplace=True, drop=True)

print(
    f" {timediff(start_time, time.time())} sorting and then \
saving the summary dataframe with {len(funds)} \
funds for {rptDate.strftime('%A %#d %B %Y')}\n"
)

# merge wbH and wbD on Entity Name and PrimaryAssetID / Primary Asset ID,
# retaining all wbH columns and appending columns unique to wbD
wbD_extra_cols = [
    col for col in wbD.columns if col not in wbH.columns and col != "Primary Asset ID"
]
wbD_for_merge = wbD[["Entity Name", "Primary Asset ID"] + wbD_extra_cols].rename(
    columns={"Primary Asset ID": "PrimaryAssetID"}
)
wbHD = wbH.merge(wbD_for_merge, on=["Entity Name", "PrimaryAssetID"], how="left")

# ##### TEST #####
# # print intermediate values
# print(f"{fund}, {ftyp}:\n")
# print(
#     f"{cash:,.2f}\n{mmfs:,.2f}\n{mmis:,.2f}\n{clns:,.2f}\n{bonds:,.2f}\n{min(0, repo):,.2f}\n{marg_jse:,.2f}\n{marg_otc:,.2f}\n{non_mmfs:,.2f}\n\n"
# )
# ##### TEST #####

# get a dataframe of funds with derivatives but no effective exposures
funds_missing_exposure = wbHD.loc[
    wbHD["Nominal Holding"].notna()
    & (wbHD["Nominal Holding"] != 0)
    & (wbHD["Investment Type"] != "SYTH")
    & (wbHD["Effective Exposure"].isna() | (wbHD["Effective Exposure"] == 0)),
    [
        "Entity ID",
        "Ticker",
        "Nominal Holding",
        "Delta",
        "Effective Exposure",
        "Current Exposure",
    ],
]

derv_calc_filepath = os.path.join(
    pthEXPORTS, f"{rptDate.strftime('%Y%m%d')}_derv_calc.xlsx"
)  # YYYYmmdd_derv_calc.xlsx

# write the sorted summary to Excel
sorted_summary.to_excel(derv_calc_filepath, sheet_name="Summary", index=False)
with pd.ExcelWriter(derv_calc_filepath, mode="a", engine="openpyxl") as writer:
    funds_missing_exposure.to_excel(writer, sheet_name="deltas_missing", index=False)

# preent funds and derivatives without effective exposures
s1 = "" if len(funds_missing_exposure["Entity ID"].unique()) == 1 else "s"
s2 = "" if funds_missing_exposure["Ticker"].nunique() == 1 else "s"
print(
    f"\n{len(funds_missing_exposure['Entity ID'].unique())} fund{s1} with \
derivatives lacking effective exposures:\
\n  {(', ').join(funds_missing_exposure['Entity ID'].unique().tolist())}"
)
print(
    f"over {funds_missing_exposure['Ticker'].nunique()} \
derivative{s2}:\n {(', ').join(funds_missing_exposure['Ticker'].unique().tolist())}"
)

start_time = time.time()
print(
    "\n\nSaving holdings, derivative deltas, and \
cash activities to a sheet ..."
)

# save holdings + deltas and trades to a file
holdings_filename = f"{rptDate.strftime('%Y%m%d')}_holdings.xlsx"
xlsx_path = os.path.join(pthEXPORTS, "Holdings", holdings_filename)

wbHD.to_excel(xlsx_path, sheet_name="holdings", index=False)
if not os.path.isfile(fCACT):
    print(
        f"\nCash activities for {rptDate.strftime('%a %d %b %Y')} \
will not be appended to the derivative summary sheet\n"
    )
    pass
else:
    print(
        f"\nCash activities for {rptDate.strftime('%a %d %b %Y')} \
will be appended to the derivative summary sheet\n"
    )
    with pd.ExcelWriter(xlsx_path, mode="a", engine="openpyxl") as writer:
        wbC.to_excel(writer, sheet_name="trades", index=False)
        funds_missing_exposure.to_excel(
            writer, sheet_name="deltas_missing", index=False
        )

print(
    f" {timediff(start_time, time.time())} saving holdings, \
derivative deltas, and cash activities to a sheet\n"
)

print(
    f"\n {timediff(start_time_derv_compiling, time.time())} total \
time to compile and save {len(fnames)} \
derivative calculations for {rptDate.strftime('%d %b %Y')}\n\n"
)

print("\n\n######################################################")
print("#                                                    #")
print("#           END 2/4 derv_checker_table.py    X       #")
print("#                                                    #")
print("######################################################\n\n")
