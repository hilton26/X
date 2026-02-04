print("\n\n##########################")
print("#   START lt_merge.py    #")
print("##########################\n\n")

# libraries, libraries!
import time
from datetime import datetime
import pandas as pd
import os, re
from pathlib import Path
from tqdm import tqdm
from constants import pthPy, pth_dl, pthTest
from utilities import (
    timediff,
    prior_month_end,
    osprey,
    r_classifier,
)
import subprocess

start_time = time.time()
start_time_merge_dl = time.time()

# get inputs
start_time = time.time()
print("Collecting input data ...\n")

# get fund codes from 'arc' tab of the py_report.xlsm sheet
df1 = pd.read_excel(pthPy, sheet_name="arc", usecols="N").dropna()
funds = df1.iloc[:, 0].str.upper()

# get report parameters date
df = pd.read_excel(pthPy, sheet_name="arc", usecols="S", nrows=3)
k = df.iloc[1, 0]
rptDate = (
    k if isinstance(k, datetime) else prior_month_end(datetime.today().date())
)  # prior month end or report date override; type is datetime()
num_batches = df.iloc[2, 0]

# check inputs
s = "" if len(funds) == 1 else "s"
print(
    f"{len(funds)} fund lookthrough{s} as at {rptDate.strftime('%A %d %b %Y')} \
to be downloaded:\n  {(', ').join(funds)}"
)
print(f"\n {timediff(start_time, time.time())} collecting input data\n")

# dataframe the lookthrough holdings saved as csvs in the local downloads folder
start_time = time.time()
print("Dataframing the lookthrough holdings ...\n")

# pattern2 = re.compile(fr"^R28I.*31Dec2025\.csv$")
pattern2 = re.compile(fr"^R28I.*{rptDate.strftime('%d%b%Y')}\.csv$")
folder = Path(pth_dl)
matching_files = [
    f.name for f in folder.iterdir() if f.is_file() and pattern2.match(f.name)
]

holdings = pd.DataFrame()
for matching_file in matching_files:
    df_new = pd.read_csv(os.path.join(pth_dl, matching_file))
    holdings = pd.concat([holdings, df_new])

# delete columns after the 9th column
holdings = holdings.iloc[:, :9]

# determine which fund holdings have not been downloaded
list_arc = df1.iloc[:,0].unique()
list_holdings = holdings.iloc[:,0].unique()
diff_funds = [item for item in list_arc if item not in list_holdings]
diff_funds
s = "'s" if len(diff_funds) == 1 else "s'"
print(f"{len(diff_funds)} fund{s} holdings not downloaded:\n {(',').join(diff_funds)}")

# convert value columns to type float
holdings_val_cols = ["End Market Value", "Closing Exposure PA"]

holdings[holdings_val_cols] = (
    holdings[holdings_val_cols].replace(",", "", regex=True).astype(float)
)
# holdings.info()

# add the report date at top of the tenth column, column "J"
holdings[f"{rptDate.strftime('%d %b %Y')}"] = None

print(
    f"\n {timediff(start_time, time.time())} dataframing \
the {len(holdings['Entity Name'].unique())} lookthrough holdings"
)

# get the fund NAVs
start_time = time.time()
print(
    f"Getting the {len(funds)} lookthrough funds' NAVs as at {rptDate.strftime('%A %d %B %Y')} with osprey() ..."
)

name = "LT"
navs_fln = os.path.join(
    pth_dl,
    f"FNAV {name}({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
)

if os.path.exists(navs_fln):
    print(
        f" Lookthrough fund NAVs as at {rptDate.strftime('%a %d %b %Y')} already downloaded: {navs_fln}"
    )
    pass
else:
    osprey("fnav", (",").join(funds), rptDate, rptDate, name, "csv")

# dataframe the downloaded fund NAVs
navs = pd.read_csv(navs_fln)

# determine which fund NAVs have not been downloaded
list_arc  =  df1.iloc[:,0].unique()
list_navs = navs.iloc[:,2].unique()
diff_navs = [item for item in list_arc if item not in list_navs]
diff_navs
s = "'s" if len(diff_navs) == 1 else "s'"
print(f"\n{len(diff_navs)} fund{s} NAVs not downloaded:\n {(',').join(diff_navs)}")

# convert navs date column to type datetime
navs["Effective Date"] = pd.to_datetime(navs["Effective Date"])

# convert the Total column from object to float
navs["Total Net Assets"] = (
    navs["Total Net Assets"].str.replace(",", "").astype("float64")
)

print(f"\n{navs_fln}\n")

print(
    f" {timediff(start_time, time.time())} getting {len(navs)} funds' NAV{'s' if len(navs) != 1 else ''} as at {rptDate.strftime('%A %d %B %Y')} \
with osprey()"
)

# merge the lookthrough holdings and NAVs, and compare their totals
start_time = time.time()
print(
    f"Merging and comparing the {len(funds)} lookthroughs and NAVs as at {rptDate.strftime('%A %d %B %Y')} ..."
)

holdings_totals = holdings.groupby("Entity Name").sum(numeric_only=True)[
    holdings_val_cols
]

len(holdings_totals)

holdings_totals.info()
navs.info()

sums_cf = holdings_totals.merge(
    navs, how="left", left_on="Entity Name", right_on="NAV Entity ID"
)

sums_cf["TNA-CE"] = sums_cf["Total Net Assets"] - sums_cf["Closing Exposure PA"]
sums_cf["1-CE/TNA %"] = (
    1 - sums_cf["Closing Exposure PA"] / sums_cf["Total Net Assets"]
) * 100
sums_cf["TNA-EMV"] = abs(sums_cf["Total Net Assets"] - sums_cf["End Market Value"])
sums_cf["1-EMV/TNA %"] = (
    1 - sums_cf["End Market Value"] / sums_cf["Total Net Assets"]
) * 100
sums_cf = sums_cf.sort_values(by="TNA-EMV", ascending=False)
cols = {
    "End Market Value": "EMV",
    "Closing Exposure PA": "CE",
    "Total Net Assets": "TNA",
}
sums_cf = sums_cf.rename(columns=cols)

list(sums_cf)

cols_order = [
    "Effective Date",
    "Entity Name",
    "NAV Entity ID",
    "EMV",
    "CE",
    "TNA",
    "TNA-EMV",
    "1-EMV/TNA %",
    "TNA-CE",
    "1-CE/TNA %",
]
sums_cf = sums_cf[cols_order]

sums_cf.info()

print(
    f" {timediff(start_time, time.time())} merging and comparing the {len(funds)} \
lookthroughs and NAVs as at {rptDate.strftime('%A %d %B %Y')}"
)

# convert the lookthrough holdings into Reg 28 format with corresponding headings
start_time = time.time()
print(f"Converting the lookthrough holdings in readiness for Reg 28 classification ...")

s = "" if len(funds) == 1 else "s"
lt_fname = os.path.join(
    pthTest, f"LT holdings ({len(funds)}) {rptDate.strftime('%d%b%Y')}.xlsx"
)

print("Writing the lookthrough holdings dataframe and navs dataframe to a sheet ...")

writer = pd.ExcelWriter(lt_fname, engine="xlsxwriter")  # instantiate a sheet writer
holdings.to_excel(
    writer, index=False, sheet_name="All"
)  # write the look-through holdings sheet
navs.to_excel(writer, index=False, sheet_name="NAVs")  # write the NAVs sheet
sums_cf.to_excel(
    writer, index=False, sheet_name="Compare"
)  # write the NAV comparison sheet
writer.close()  # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets

print(f" \n{lt_fname}\n")

print(
    f" {timediff(start_time, time.time())} writing the \
lookthrough holdings dataframe and navs dataframe to a sheet\n"
)

# ready for the Reg 28 reporting script

print(
    "\n",
    f"{timediff(start_time_merge_dl, time.time())} roundtrip to \
download and merge lookthroughs\n",
)

# call the classifier function
r_classifier("r28",lt_name, rptDate):

print("\n\n#######################")
print("#   END lt_merge.py   #")
print("#######################\n\n")
