# construct the derivative calculation for a fund

import time

start_time = time.time()

import pandas as pd
import os
from datetime import datetime
import xlwings as xw
from constants import pthPy, pthEXPORTS, derv_check_tmpl
from utilities import prior_working_day, timediff
from tqdm import tqdm

# find the column with the py script name
col_ref = "derv_portfolio"
arc_heads = pd.read_excel(pthPy, sheet_name="arc", header=None, nrows=1).iloc[0]
# print(arc_heads)
py_name = arc_heads[arc_heads == col_ref].index[0]
# print(py_name)

# set number of fund sheets to open after doing the calc
num_sheets_to_open = 4

# dataframe of funds and report date
df = pd.read_excel(pthPy, sheet_name="arc", usecols=[0, 4, py_name], header=0)
df = df.dropna(subset=[df.columns[0]])

# funds
funds = df[df.iloc[:, -1] == 1].iloc[:, 0]
# print(funds)

# report date
k = df.iloc[0, 1]
rptDate = (
    k.date() if isinstance(k, datetime) else prior_working_day(datetime.today()).date()
)  # prior working day or report date override; has type datetime()

# make a subset including the funds from the dail;y holdings set
filename = os.path.join(
    pthEXPORTS, "Holdings", f"{rptDate.strftime('%Y%m%d')}_holdings.xlsx"
)
print(f"\n{os.path.basename(filename)} will be used as the source\n")
all_holdings = pd.read_excel(filename, sheet_name="holdings")

s1 = "'s" if len(funds) == 1 else "s'"
s2 = "" if len(funds) == 1 else "s"
print(
    f"\n{len(funds)} fund{s1} derivative calc{s2} as at \
{rptDate.strftime('%a %d %b %Y')}:\n {(', ').join(funds.tolist())}\n"
)

# single hidden Excel instance reused for every fund, in manual calc mode, to
# avoid "Excel ran out of resources while attempting to calculate" errors caused
# by the formula-heavy template auto-recalculating on every write
app = xw.App(visible=False)
app.display_alerts = False
app.calculation = "manual"

try:
    # loop over each fund
    for fund in tqdm(funds):
        holdings = all_holdings[all_holdings["Entity ID"] == fund]

        # create a sheet for each fund derivative calculation
        wb = app.books.open(derv_check_tmpl)
        ws = wb.sheets["Data"]
        ws.clear_contents()
        ws["A1"].value = holdings.columns.tolist()
        ws["A2"].value = holdings.values.tolist()

        wb.app.calculate()  # one deliberate full recalculation, not one per paste

        out_path = os.path.join(
            pthEXPORTS, f"{rptDate.strftime('%Y%m%d')}_{fund}_derv_check.xlsx"
        )
        wb.save(out_path)
        wb.close()

        if len(funds) <= num_sheets_to_open:
            os.startfile(out_path)

finally:
    app.quit()

if len(funds) > num_sheets_to_open:
    os.startfile(pthEXPORTS)

print(
    f"\n\n {timediff(start_time, time.time())} total time to \
compile and save the {len(funds)} \
derivative check sheets for {rptDate.strftime('%d %b %Y')}"
)
