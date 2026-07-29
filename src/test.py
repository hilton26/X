


print("\n\n(1) setting input variables for testing osprey()")
from datetime import datetime
import pandas as pd
from constants import pthPy, pth_dl
from utilities import prior_working_day
from osprey2 import osprey

# (1) inputs to function osprey() for testing purposes
# (1a) set report inputs
rpt_type = "parn"
name = f"PGF_Holdings"  # "PGF_UT_prices" for utps (column "G"), and "PGF_Holdings" for parn (column "H")
sfx = "csv"

# (1b) get funds list
cols_to_use = "H"
df = pd.read_excel(
    pthPy,
    sheet_name="arc",
    usecols=cols_to_use,
    header=None,
    skiprows=1,
).dropna()
df.columns = [0]  # assigning header number to 0 for the code below to work, e.g., df[0]
funds = ",".join(df[0].astype(str).tolist())
nf = len(df)

# (1c) get report dates and remaining inputs
cols_to_use = "I"
df = pd.read_excel(
    pthPy, sheet_name="arc", usecols=cols_to_use, header=None, skiprows=1, nrows=2
)
df.columns = [0]  # assigning header number to 0 for the code below to work, e.g., df[0]
k = df.iloc[1, 0]
rptDate = (
    k.date()
    if isinstance(k, datetime) and not pd.isna(k)
    else prior_working_day(datetime.today()).date()
)  # prior working day or report date override; has type datetime()
d_to = rptDate
d_from = d_to

# new_file_name = f"{rpt_type.upper()} {name}({n}) {d_to.strftime('%d%b%Y')}"
new_file_name = f"{rpt_type.upper()} {name}({nf}) {d_from.strftime('%d%b%Y')}{(' to ' + d_to.strftime('%d%b%Y') if d_to != d_from else '')}"

print(f"Expected file name per test.py: {new_file_name}.{sfx}\n\n")

if (pth_dl / f"{new_file_name}.{sfx}").exists():
    print(f"{new_file_name}.{sfx} already exists")
    # continue

print(
    f"\nInputs: \n rpt_type={rpt_type},\n name={name},\n sfx={sfx},\n d_from={d_from},\n d_to={d_to},\n funds={funds}\n\n"
)

#######################################
#                                     #
#    Applying the osprey() function   #
#                                     #
#######################################

osprey(rpt_type=rpt_type, funds=funds, d_from=d_from, d_to=d_to, name=name, sfx=sfx)
