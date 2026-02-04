#!/usr/bin/env python
# coding: utf-8

# ### Deletes a file given its path

import os
import pandas as pd
from constants import pthPy, pthHdg, frcv_file

df = pd.read_excel(pthPy, sheet_name="arc", usecols="AW", nrows=9)


if pd.isna(df.iloc[2, 0]):
    if df.iloc[0, 0] == "Derv":
        target = frcv_file
    else:
        target = (
            pthHdg
            + "\\"
            + df.iloc[1, 0].strftime("%Y%m%d")
            + " PGF Share Class Hedges.xlsx"
        )
else:
    target = df.iloc[2, 0]

print(target)

try:
    os.remove(target)
    print(f"File '{target}' deleted successfully.")
except FileNotFoundError:
    print(f"File '{target}' not found.")
except Exception as e:
    print(f"An error occurred: {e}")
