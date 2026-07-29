import time

start_time = time.time()
start_time_derivative_downloading = time.time()
print("\n\nImporting libraries ...\n")

# load libraries
import pandas as pd
import os, sys, subprocess
from constants import pth_dl
from utilities import timediff, parn_de, osprey

print("\nGetting the reporting date and names derivative files ...")

# get report inputs
fPARN, fDE, funds, rptDate, summ_yn, dervthreshold = parn_de()
cact_sets = 6

# check if the required files have been downloaded, else continue
if not os.path.exists(fPARN) or not os.path.exists(fDE):
    sys.exit(
        f"Stopping: missing expected download(s):\n"
        f"  {fPARN} which {'exists' if os.path.exists(fPARN) else 'does not exist'}\n"
        f"  {fDE} which {'exists' if os.path.exists(fDE) else 'does not exist'}\n"
    )

print(type(funds), "\n", (",").join(funds))
# print(funds, "\n", type(funds))

# derive cash activities file path
cact_name = f"CACT ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv"
fCACT = os.path.join(pth_dl, cact_name)

# get cash activities
start_time = time.time()
print("\n\nDownloading and then saving cash activities ...\n")

if os.path.exists(fCACT):
    print(f"   {cact_name} already downloaded")
    pass
else:
    if cact_sets != 1:  # if more than 1 set of CACT files is indicated...
        # ... split the fund list into cact_sets roughly equal-sized subsets
        bounds = [round(i * len(funds) / cact_sets) for i in range(cact_sets + 1)]
        cact_subsets = [funds[bounds[i] : bounds[i + 1]] for i in range(cact_sets)]
        cact_names = [
            f"CACT set{i}_of_{cact_sets}({len(subset)}) {rptDate.strftime('%d%b%Y')}.csv"
            for i, subset in enumerate(cact_subsets, start=1)
        ]

        for i, (subset, cact_i_name) in enumerate(
            zip(cact_subsets, cact_names), start=1
        ):
            start_time_i = time.time()
            cact_i = (",").join(subset)

            # check if the file was already downloaded before running osprey()
            if (
                os.path.exists(os.path.join(pth_dl, cact_i_name))
                and os.path.getsize(os.path.join(pth_dl, cact_i_name)) > 0
            ):
                print(f"  {cact_i_name} already exists\n")
                pass
            else:
                print(
                    f"\n\n  Downloading subset {i} of {cact_sets} of cash \
activities for {len(subset)} funds:\n   {(',').join(subset)}\n"
                )
                osprey(
                    "cact", cact_i, rptDate, rptDate, f"set{i}_of_{cact_sets}", "csv"
                )
                print(
                    f"  {timediff(start_time_i, time.time())} downloading \
subset {i} of {cact_sets} of cash activities"
                )

        dfs = [
            pd.read_csv(os.path.join(pth_dl, cact_i_name)) for cact_i_name in cact_names
        ]
        df_cact = pd.concat(dfs)
        # print([len(d) for d in dfs], sum(len(d) for d in dfs), len(df_cact))

        # write the combined dataframe to a csv file in the Downloads folder
        df_cact.to_csv(
            os.path.join(
                pth_dl,
                f"CACT ({len(funds)}) {rptDate.strftime('%d%b%Y')}.csv",
            ),
            index=False,
        )

    else:  # else get all the holdings in one go
        # check if the file was already downloaded before running osprey()
        # if os.path.isfile(pth_dl + r'\\' + full_name):
        if (
            os.path.exists(os.path.join(pth_dl, cact_name))
            and os.path.getsize(os.path.join(pth_dl, cact_name)) > 0
        ):
            print(f"  {cact_name} already exists\n")
            pass
        else:
            start_time_h = time.time()
            osprey("cact", (",").join(funds), rptDate, rptDate, "", "csv")
            print(
                f"  {timediff(start_time_h, time.time())} downloading \
all cash activities"
            )

    # print("", cact_name, "\n")

# if the cash activitities were downloaded,
# add them to the daily derivative summary
if os.path.exists(os.path.join(pth_dl, cact_name)):
    print("Appending trades to the derivative report")
    subprocess.run([sys.executable, r"src\derv_checker_table.py"])
else:
    print("Trades not available to be added to the derivative report")

print(
    f"\n\n {timediff(start_time, time.time())} total time \
downloading and then saving cash activities\n"
)
