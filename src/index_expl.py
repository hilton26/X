#!/usr/bin/env python
# coding: utf-8
"""
Input:

 (i) Fund code in column 'AN' of sheet 'arc' and
 (ii) date of report in column 'AO'
 (iii) fund look-through url to be inserteted in column 'O'
 if it was not already downloaded in prepartion
 for the CS1 reports A csv file with look-through df_h

Output:

 An .xlsx file with look-through df_h including
 look-through to indices underlying derivatives

"""

import time

start_time_index_expl = time.time()

from datetime import datetime
import pandas as pd
import os
import re
from constants import pthPy, pth_struct, pth_dl, pth_BX, pthTest
from utilities import timediff, osprey, tmv, prior_month_end, r_classifier
from tqdm import tqdm

# gather report data
start_time_index_LT = time.time()
start_time = time.time()

# get name of fund(s) to be exploded
df = pd.read_excel(pthPy, sheet_name="arc", usecols="AN").dropna()
funds = list(df.iloc[:, 0])

# get URL of fund(s) to be exploded
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="AO", nrows=3)
url = df1.iloc[2, 0].replace('"', "")
# k = df1.iloc[1, 0]
rptDate = (
    df1.iloc[1, 0].date()
    if isinstance(df1.iloc[1, 0], datetime)
    else prior_month_end(datetime.today()).date()
)
print("", rptDate, "\n", (", ").join(funds))
print(f" {url if url == url else 'No URL'}")

print(
    f"\n {timediff(start_time, time.time())} \
getting report data"
)


# function to derive securities underlying the derivative
def ndx(txt):
    for pattern in idx.iloc[:, 0]:
        if re.search(pattern, str(txt).upper()):
            return idx.loc[idx["keyword"] == pattern].iat[
                0, 2
            ]  # note 'break' within the for loop
            break


# dataframe the BX indices
idx = pd.read_excel(pth_struct, sheet_name="dervs", usecols=[0, 1, 3, 4]).dropna(
    subset=["keyword"]
)

# locate the holdings else download them

start_time = time.time()
print(
    f"Downloading and dataframing {len(funds)} \
fund holdings for {rptDate.strftime('%d%b%Y')}"
)

if url == url:
    df_h = pd.read_excel(url, usecols="A:I").dropna(subset=["Entity Name"])
else:
    if len(funds) < 30:
        name = ""
        holdings_file_name = os.path.join(
            pth_dl, f"R28I {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
        )
        if os.path.exists(holdings_file_name):
            print(
                f"{holdings_file_name} \
already exists"
            )
            pass
        else:
            osprey("r28i", (",").join(funds), rptDate, rptDate, f"{name}", "csv")

        df_h = pd.read_csv(holdings_file_name)

    else:
        try:
            df_h = pd.DataFrame()  # initialise an empty dataframe
            for fund in tqdm(funds):
                name = fund
                fl_name = os.path.join(
                    pth_dl, f"R28I {name}(1) {rptDate.strftime('%d%b%Y')}.csv"
                )
                osprey("r28i", fund, rptDate, rptDate, f"{name}", "csv")
                data = pd.read_csv(fl_name)
                df_h = pd.concat([df_h, data])
                time.sleep(1)
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
        else:
            print(
                "Holdings files were downloaded \
and dataframed"
            )
        finally:
            print("Cleanup operations completed")

df_holdings_spare = df_h.copy(deep=True)
df_h = df_holdings_spare.copy(deep=True)
print((",").join(df_h["Entity Name"].unique()), "\n")
# print(list(df_h))

funds = list(df_h["Entity Name"].unique())

print(
    "\n",
    f"{timediff(start_time, time.time())} downloading and dataframing {len(funds)} \
fund holdings for {rptDate.strftime('%d%b%Y')}:\n   {(', ').join(funds)}",
)

# end of downloads
# identify derivatives and condense TRSes to one row

# (0) exclude zeroes from the holdings dataframe and condense the holdings

start_time_index_LT_processing = time.time()
start_time = time.time()

df_h = df_holdings_spare.copy(deep=True)  # spare copy of holdings

shape_before = df_h.shape

# recalculate percentage columns using tmv(df, by_col,
# num_cols_curr, num_cols_pct) as template
cols = ["End Market Value", "Closing Exposure PA"]
tmv(df_h, "Entity Name", cols, ["Percentage of Market Value", ""])

# remove NaN market value rows and zero-value effective exposure rows
df_h = df_h[df_h[cols[0]].notnull()]  # delete NaN MV column rows
df_h = df_h[
    (round(df_h[cols[0]], 2) != 0) | (round(df_h[cols[1]], 2) != 0)
]  # delete rows where MV=0 and EE=0 to 2 decimals

# reset holdings dataframe index to avoid ValueError:
# cannot reindex on an axis with duplicate labels
df_h = df_h.reset_index(drop=True)  # reset index and drop previous index

print(
    f"{df_h.shape[0]:,.0f} holdings dataframe rows \
remain over all {len(funds)} funds after \
deleting {shape_before[0] - df_h.shape[0]:,.0f} \
zero-value rows\n"
)

# combine all derivatives into a list

# (1) identify linked securities
substring_L = "linked"
linkeds = df_h[
    df_h["i Issue Name"].str.contains(f"{substring_L}", case=False, na=False)
]["Primary Asset ID"]

# (1a) identify structured notes
substring_S = "structured"
structs = df_h[
    df_h["i Issue Name"].str.contains(f"{substring_S}", case=False, na=False)
]["Primary Asset ID"]

# (2) identify futures and options
substring_FT_OP = ["FT", "OP"]
substring_not_FT_OP = [
    "LONG GILT",
    "US 2YR",
    "US 5YR",
    "US 10YR",
    "EURO-BUND",
    "ULTRA BOND",
]  # exclude bond options
ft_op = df_h[
    (
        df_h["Investment Type"].str.contains(
            "|".join(substring_FT_OP), case=False, na=False
        )
    )
    & ~df_h["i Issue Name"].str.contains("|".join(substring_not_FT_OP))
]["Primary Asset ID"]

# (3) identify securities with "TRS" in their names and make a unique column list of them
substring_TRS = "TRS"
trses = df_h[
    (df_h["Investment Type"] == "DERV")
    & (df_h["i Issue Name"].str.contains(f"{substring_TRS}", case=False))
]["Primary Asset ID"]  # case-insensitive
trses = pd.Series(trses.unique().tolist())

# (4) combine all derivatives into a list
all_dervs = pd.concat([linkeds, structs, ft_op, trses]).tolist()
print(
    f"{len(all_dervs)} derivatives (= {len(ft_op)} \
futures and options + {len(linkeds)} ELNs \
+ {len(structs)} structured notes + {len(trses)} \
TRSes) and\n{len(df_h):,.0f} non-zero line items \
over {len(funds)} funds: \n   {(', ').join(funds)}\n"
)

# (5) create a new column with names of assets underlying the derivatives
df_h["indx"] = df_h["i Issue Name"].apply(
    ndx
)  # create a new column with derivative names
df_h.loc[(df_h["Investment Type"] == "SYTH") & (~df_h["indx"].isnull()), "indx"] = (
    None  # set SYTH index names to None
)

# treatment of TRSes

# check sums before treatment of TRSes
df_navs = df_h.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
df_navs["Diff"] = (
    df_h.groupby("Entity Name")["End Market Value"].sum()
    - df_h.groupby("Entity Name")["Closing Exposure PA"].sum()
)
df_navs = df_navs.rename(
    columns={
        "End Market Value": "MV",
        "Percentage of Market Value": "%MV",
        "Closing Exposure PA": "EE",
    }
)

print(f"Check column totals before deleting second TRS legs:")
print(df_navs)

# (6) add a column to identify first ("0")
# and second ("1") instances of TRSes
df_trses = df_h.loc[
    df_h["Primary Asset ID"].isin(trses)
]  # subset dataframe of all TRSes
df_h["occ"] = df_trses.groupby(
    "Primary Asset ID"
).cumcount()  # insert a (new) duplicate counter column
df_h[df_h["Primary Asset ID"].isin(trses)]

# (7) loop through each fund to update values of TRSes
for fund in funds:
    for trs in trses:
        # print(fund, trs)
        if (
            len(df_h[(df_h["Entity Name"] == fund) & (df_h["Primary Asset ID"] == trs)])
            == 0
        ):
            # print(f"{trs} is not in {fund}")
            pass
        else:
            # values to be assigned to first leg of the TRS
            issuer1_nm = df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == trs)
                & (df_h["occ"] % 2 == 0)
            ]["i Issue Name"].iat[0][:-2]  # remove last two digits of TRS name, '_P'
            sum1_emv = df_h.loc[
                (df_h["Entity Name"] == fund) & (df_h["Primary Asset ID"] == trs)
            ]["End Market Value"].sum()
            sum1_pomv = df_h.loc[
                (df_h["Entity Name"] == fund) & (df_h["Primary Asset ID"] == trs)
            ]["Percentage of Market Value"].sum()
            sum1_cepa = 0

            # update values in first leg of the TRS
            df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == trs)
                & (df_h["occ"] % 2 == 0),
                "i Issue Name",
            ] = issuer1_nm
            df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == trs)
                & (df_h["occ"] % 2 == 0),
                "End Market Value",
            ] = sum1_emv
            df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == trs)
                & (df_h["occ"] % 2 == 0),
                "Percentage of Market Value",
            ] = sum1_pomv
            df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == trs)
                & (df_h["occ"] % 2 == 0),
                "Closing Exposure PA",
            ] = sum1_cepa

rows_before_TRS_legs_delete = df_h.shape[0]

# (8) delete second legs of TRSes, i.e., rows
# where occurrence_num is odd; for TRSes, MV = EE
df_h = df_h[(df_h["occ"] % 2 == 0) | (df_h["occ"].isnull())]

# (9) delete TRS occurrence counter column
df_h = df_h.drop("occ", axis=1)  # axis = 1 = column-wise operation

# (10) reset holdings dataframe index to avoid ValueError: cannot
# reindex on an axis with duplicate labels; drop pevious index
df_h = df_h.reset_index(drop=True)

print(
    f"\n\n{df_h.shape[0]:,.0f} rows remain over all {len(funds)} funds after \
deleting {rows_before_TRS_legs_delete - df_h.shape[0]:,.0f} TRS legs\n"
)

# check sums after treatment of TRSes
df_navs = df_h.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
df_navs["Diff"] = (
    df_h.groupby("Entity Name")["End Market Value"].sum()
    - df_h.groupby("Entity Name")["Closing Exposure PA"].sum()
)
df_navs = df_navs.rename(
    columns={
        "End Market Value": "MV",
        "Percentage of Market Value": "%MV",
        "Closing Exposure PA": "EE",
    }
)

print(f"Check column totals after deleting second TRS legs:")
print(df_navs, "\n")

print("\n", f"{timediff(start_time, time.time())} runtime")

# get index constituents

# (11) get a list of the full set of underlying
# indices and set up the index constituents

start_time = time.time()
print(f"Dataframing the index constituents for {rptDate.strftime('%d %b %Y')} ...")

# (12) determine which indices are available
# in the current BX file
bx = pth_BX + rf"\BX {rptDate.strftime('%d%b%Y')}.xlsx"
df_msci = pd.read_excel(
    bx, sheet_name="MSCI", usecols=[0, 1, 4, 10, 14]
).dropna()  # msci index constituents
df_jse = pd.read_excel(
    bx, sheet_name="JSE_indices", usecols=[0, 1, 4, 10, 14]
).dropna()  # jse index constituents
df_bsk = pd.read_excel(
    bx, sheet_name="BSK", usecols=[0, 1, 4, 10, 14]
).dropna()  # bskxxx index constituents, e.g., BSK141
df_tsy = pd.read_excel(
    bx, sheet_name="Treasuries", usecols=[0, 1, 4, 10, 14]
).dropna()  # treasuries
msci = df_msci["Index"].unique()  # a numpy.ndarray
jse = df_jse["Index"].unique()  # a numpy.ndarray
bsk = df_bsk["Index"].unique()  # a numpy.ndarray
tsy = df_tsy["Index"].unique()  # a numpy.ndarray

# (13) combine msci and jse index dataframes and
# add a percentage column based on market cap
ix = pd.concat([df_msci, df_jse, df_bsk, df_tsy], axis=0).reset_index(
    drop=True
)  # reset index and drop the old index
ix["%"] = ix.loc[:, "MC (ZAR)"] / ix.groupby("Index")["MC (ZAR)"].transform(
    sum
)  # calculate constituent weights
ix = ix.drop("MC (ZAR)", axis=1)  # drop the market cap column
ix["Name"] = ix["Name"].str.title()  # decapitalise all-caps text
# sum check
# print(ix.groupby('Index')[['%']].sum())

# (14) rename and then reorder the index columns
# to match those of the Reg 28 holdings report

# (14.1) create new ix columns
ix["Entity Name"] = None
ix["Investment Type"] = "EQ"
ix["Reg28 Classification"] = None
ix["End Market Value"] = 0.00
ix["Percentage of Market Value"] = 0.00
ix["Closing Exposure PA"] = 0.00
ix["indx"] = None

# (14.2) rename and then reorder the index columns
# to match those of the Reg 28 holdings report
new_col_names = {
    "Name": "i Issue Name",
    "Currency": "CCY",
    "PIM Ticker": "Primary Asset ID",
}
new_col_order = [
    "Entity Name",
    "Investment Type",
    "i Issue Name",
    "Primary Asset ID",
    "CCY",
    "Reg28 Classification",
    "End Market Value",
    "Percentage of Market Value",
    "Closing Exposure PA",
    "Index",
    "%",
    "indx",
]
ix = ix.rename(columns=new_col_names)
ix = ix[new_col_order]

print(
    f" {timediff(start_time, time.time())} \
dataframing the index constituents for \
{rptDate.strftime('%d %b %Y')}\n"
)

# index check: identify indices (i) in the funds, and (ii) their constituents in the market cap file, BX

start_time = time.time()
print(
    f"Identifying indices in the funds and \
confirming they are available for \
{rptDate.strftime('%d %b %Y')} ...\n"
)

# (15) identify derivatives without associated underltying assets in BX
no_index = df_h.loc[
    (df_h["Primary Asset ID"].isin(all_dervs))
    & (df_h["Investment Type"] != "SYTH")
    & (df_h["indx"].isnull())
]
print(
    f"{len(no_index)} derivatives with no index \
assigned in the {len(funds):,.0f} funds:\n   \
{'[null set]' if len(no_index) == 0 else (', ').join(list(no_index['Primary Asset ID']))}"
)

# (16) list indices in the funds
indices = pd.DataFrame(df_h[df_h["indx"].notnull()]["indx"].unique(), columns=["indx"])
print(
    f"\n{len(indices)} underlying indices \
in the {len(funds):,.0f} funds:\n  {(', ').join(indices['indx'].to_list())}"
)

# (17) confirm that each underlying index is available in the current BX file
# if not available, HALT THE LOOP in order for the missing index / indices to be made available
bond_indices = ["EUB", "UKB", "USB", "ZAB"]  # to not be expanded, for now
mask = (
    ~indices.isin(msci)
    & ~indices.isin(jse)
    & ~indices.isin(bsk)
    & ~indices.isin(bond_indices)
)  # boolean dataframe of indices not in BX
not_in_BX = indices[mask["indx"]]
print(
    f"\n{len(not_in_BX)} underlying indices not \
represented in BX file of \
{rptDate.strftime('%d %b %Y')}: \
\n  {'[null set]' if len(not_in_BX) == 0 else (', ').join(not_in_BX['indx'].to_list())}\n"
)
print(
    rf"Populate the constituents and weights \
of {(', ').join(not_in_BX['indx'].to_list())} in \
file P:...structures.xlsm"
)

print(
    f" {timediff(start_time, time.time())} \
identifying indices in the funds and confirming \
they are available for {rptDate.strftime('%d %b %Y')}\n"
)

# insert assets underlying the derivatives into the holdings dataframe

# (18) insert assets underlying the derivatives into the holdings dataframe

start_time = time.time()
# print(df1.shape)

# # NAVs before inserting index constituents
# df_navs = df_h.groupby('Entity Name')[['End Market Value','Percentage of Market Value','Closing Exposure PA']].sum()
# df_navs['Diff'] = df_h.groupby('Entity Name')['End Market Value'].sum() - df_h.groupby('Entity Name')['Closing Exposure PA'].sum()
# df_navs = df_navs.rename(columns = {'End Market Value': 'MV', 'Percentage of Market Value': '%MV' ,'Closing Exposure PA': 'EE'})
# print(f"\nNAVs before inserting index constituents:\n{df_navs}\n")

rows_before = df_h.shape[0]

# create a holdings subset of the derivative line items
dervs_only = df_h.loc[
    (df_h["Primary Asset ID"].isin(all_dervs)) & (df_h["Investment Type"] != "SYTH")
]

# change Investment Type 'EQ' to 'FI' for the linked and structured notes
dervs_only.loc[dervs_only["Investment Type"] == "EQ", "Investment Type"] = "FI"

# isolate the derivative instrument codes
pa_ids = dervs_only.loc[:, "Primary Asset ID"]

for fund in funds:
    sum_total_dervs = 0
    for pa_id in pa_ids.unique():  # for each unique derivative
        # print(fund, pa_id)
        pa_id_slice = df_h.loc[
            (df_h["Entity Name"] == fund)
            & (df_h["Primary Asset ID"] == pa_id)
            & (df_h["Investment Type"] != "SYTH")
        ]
        if len(pa_id_slice) == 0:
            # print(f"{pa_id} is not in {fund}")
            pass
        else:
            # check holdings row length
            df_rows_before_indices = df_h.shape[0]
            # print(f"{df_rows_before_indices:,.0f} holdings rows before adding assets underlying indices")

            # slice the current derivative row
            mv_ee = pa_id_slice
            # print(mv_ee)

            # determine the effective exposure obatined from the derivative
            effective_exposure = (
                mv_ee.iloc[0, 6] if mv_ee.iloc[0, 8] == 0 else mv_ee.iloc[0, 8]
            )
            # print(effective_exposure)

            # zero the effective exposure value of the derivative
            df_h.loc[
                (df_h["Entity Name"] == fund)
                & (df_h["Primary Asset ID"] == pa_id)
                & (df_h["Investment Type"] != "SYTH"),
                "Closing Exposure PA",
            ] = 0.0

            # get a slice of index dataframe for the index associated with the current derivative
            ix_slice = ix.loc[ix["Index"] == mv_ee.iloc[0, 9]].copy(deep=True)
            ix_slice["Closing Exposure PA"] = ix_slice["%"] * effective_exposure
            ix_slice["Entity Name"] = fund
            ix_slice.drop(["Index", "%"], axis=1, inplace=True)
            # print(f"{ix_slice.shape[0]:,.0f} ix_slice rows in {pa_id} summing to {ix_slice['Closing Exposure PA'].sum():,.2f}")

            # TEST sum the slice
            sum_total_dervs = sum_total_dervs + effective_exposure

            # join the derivative effective exposure to the holdings dataframe
            df_h = pd.concat([df_h, ix_slice], axis=0, ignore_index=True)
            # print(f"{df_h.shape[0] - df_rows_before_indices:,.0f} rows added after indexing")

# NAVs after inserting index
df_navs = df_h.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
df_navs["Diff"] = (
    df_h.groupby("Entity Name")["End Market Value"].sum()
    - df_h.groupby("Entity Name")["Closing Exposure PA"].sum()
)
df_navs = df_navs.rename(
    columns={
        "End Market Value": "MV",
        "Percentage of Market Value": "%MV",
        "Closing Exposure PA": "EE",
    }
)
print(
    f"\nNAVs after inserting index \
constituents:\n{df_navs}\n"
)

# (19) reset holdings dataframe index to avoid ValueError: cannot reindex on an axis with duplicate labels; drop pevious index
df_h = df_h.reset_index(drop=True)

# (20) delete TRS occurrence counter column
df_h = df_h.drop("indx", axis=1)  # axis = 1 specifies column-wise operation

print(
    f"\n{df_h.shape[0]:,.0f} holdings rows \
after {df_h.shape[0] - rows_before:,.0f} added\n"
)

print(
    f" {timediff(start_time, time.time())} adding {df_h.shape[0] - rows_before:,.0f} \
index constituents to the {len(funds):,.0f} funds for {rptDate.strftime('%d %b %Y')}\n"
)

# (21) condense duplicate securities into single rows

start_time = time.time()
print(
    f"Condensing duplicate \
securities into single rows"
)

df_z = pd.DataFrame()
for fund in funds:
    # hard copy each fund holdings set
    df_f = df_h.loc[df_h["Entity Name"] == fund].copy(deep=True)
    df_f = df_f.reset_index(drop=True)

    # # TEST check sums before aggregation
    # df_f_navs = df_f[['End Market Value','Percentage of Market Value','Closing Exposure PA']].sum()
    # diff = pd.Series(df_f['End Market Value'].sum() - df_f['Closing Exposure PA'].sum(), index = ['diff'])
    # print(f"Check column totals before aggregation:\n{pd.concat([df_f_navs, diff])}\n")

    # df_f['cif'] = df_f.groupby('Primary Asset ID').cumcount() # insert a (new) duplicate counter column
    # pa_ids = df_f.loc[(df_f['cif'] == 1) & (~df_f['Investment Type'].isin(['SYTH', 'DERV'])), 'Primary Asset ID'] # exclude the swaps under 'DERV'
    df_f["cif"] = (
        df_f.loc[~df_f["Investment Type"].isin(["SYTH", "DERV"])]
        .groupby("Primary Asset ID")
        .cumcount()
    )  # insert a duplicate counter column
    pa_ids = df_f.loc[df_f["cif"] == 1, "Primary Asset ID"]
    for pa_id in pa_ids:
        # sum over each instance of the duplicate ticker
        cepa = (
            df_f.loc[(df_f["Primary Asset ID"] == pa_id) & (~df_f["cif"].isnull())][
                "Closing Exposure PA"
            ]
            .sum()
            .round(2)
        )

        # replace the effective exposure value of the zeroth instance of the duplicate ticker
        df_f.loc[
            (df_f["Primary Asset ID"] == pa_id) & (df_f["cif"] == 0),
            "Closing Exposure PA",
        ] = cepa

        # # TEST view isolated instrument codes
        # print(pa_id, f"{cepa:,.2f}")

    # delete the non-zero duplicate ticker counter rows
    df_f = df_f[(df_f["cif"] == 0) | (df_f["cif"].isnull())]

    # delete the duplicate ticker counter column
    df_f = df_f.drop("cif", axis=1)  # axis = 1 specifies column-wise operation

    # sort the fund holdings dataframe by 'Investment Type', in order of CASH, EQ, FI, FT, etc.
    df_f = df_f.sort_values(
        by=["Investment Type", "End Market Value"], ascending=[True, False]
    )

    # # TEST check sums after aggregation
    # df_f_navs = df_f[['End Market Value','Percentage of Market Value','Closing Exposure PA']].sum()
    # diff = pd.Series(df_f['End Market Value'].sum() - df_f['Closing Exposure PA'].sum(), index = ['diff'])
    # print(f"\nCheck column totals after aggregation:\n{pd.concat([df_f_navs, diff])}\n")

    # merge all the funds' holdings into a single dataframe
    df_z = pd.concat([df_z, df_f], axis=0, ignore_index=True)

# add a new, empty, column for the report date
dt = rptDate.strftime("%d %b %Y")
df_z[dt] = None

# NAVs
df_navs = df_z.groupby("Entity Name")[
    ["End Market Value", "Percentage of Market Value", "Closing Exposure PA"]
].sum()
df_navs["Diff"] = (
    df_z.groupby("Entity Name")["End Market Value"].sum()
    - df_z.groupby("Entity Name")["Closing Exposure PA"].sum()
)
df_navs = df_navs.rename(
    columns={
        "End Market Value": "MV",
        "Percentage of Market Value": "%MV",
        "Closing Exposure PA": "EE",
    }
)
print(f"\nNAVs:\n{df_navs}\n")

print(
    f" {timediff(start_time, time.time())} \
condensing duplicate securities into single rows"
)

# (22) write the expanded fund holdings to Excel
file_name = (
    f"{pthTest}"
    + r"\Reg28_Index_LT_("
    + f"{len(funds)}"
    + ")_"
    + f"{rptDate.strftime('%d%b%Y')}"
    + ".xlsx"
)
df_z.to_excel(file_name, sheet_name="All", index=False)

print(
    f"{timediff(start_time_index_LT_processing, time.time())} \
processing time for {len(funds)} funds"
)
print(f"\n{file_name}")

start_time = time.time()
print(
    f"Running classifier script on the {len(funds)} \
funds as at {rptDate.strftime('%d %b %Y')}"
)
r_classifier(
    "Reg28",
    file_name,
)

s = "" if len(funds) == 1 else "s"
print(
    f" {timediff(start_time, time.time())} running \
classifier script on the {len(funds)} fund{s} \
as at {rptDate.strftime('%d %b %Y')}"
)
print(
    f"\n{timediff(start_time_index_LT, time.time())} \
total roundtrip time for {len(funds)} fund{s}"
)
