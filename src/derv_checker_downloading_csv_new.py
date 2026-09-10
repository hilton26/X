#!/usr/bin/env python
# coding: utf-8

# # Pull Holdings and Derivative Data from Eagle

# How to wait until Element is Visible in Selenium Python  
# 
# https://pythonexamples.org/python-selenium-wait-until-element-is-visible/

# In[13]:


# get utilities.ipynb which imports datetime, time
get_ipython().run_line_magic('run', 'utilities.ipynb')
start_time = time.time()
start_time_derivative_downloading = time.time()
print("Importing libraries and setting paths for derv_checker_downloading_csv.ipynb ...")

# load libraries
import pandas as pd
import numpy as np
import os
from pathlib import Path

# set paths
pthPy      = r'P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm'  # report variables are stored here
pthEXPORTS = r'P:\Investment Operations\GRC\Compliance\Derivative Cover'
df = pd.read_excel(pthPy, sheet_name="dervs", header=None, usecols="A,D:E").dropna(subset = [0])
p  = df.iloc[0, 2]
rptDate = p if isinstance(p, datetime) else prior_working_day() # prior working day or report date override; has type datetime()
filename = os.path.join(pthEXPORTS, f'Derv {rptDate.strftime("%d%b%Y")}.xlsx')
# os.path.isfile(filename)

print(f"{timediff(start_time, time.time())} importing libraries and setting paths for derv_checker_downloading_csv.ipynb","\n")


# In[14]:


start_time = time.time()
print(f"Setting up paths ...")

# set paths to the driver, urls, and report parameters
pthPy    = r"P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm"  # report variables are stored here
pth_dl   = os.path.join(Path.home(), "Downloads")
pthLOCAL = os.path.join(Path.home(), "Documents", "DervFiles")

print(f"{timediff(start_time, time.time())} setting up paths", "\n")


# In[15]:


# Get report date and selected summary sheet option

start_time = time.time()
print("Getting the reporting date and fund codes ...")

df = pd.read_excel(pthPy, sheet_name="dervs", header=None, usecols="A,D:E").dropna(subset = [0])
p  = df.iloc[0, 2]
rptDate = p if isinstance(p, datetime) else prior_working_day() # prior working day or report date override; has type datetime()
summ_yn = df.iloc[1, 1]
full    = df[0].iloc[1:]
funds   = (",").join(full.tolist())

df2 = pd.read_excel(pthPy, sheet_name="creds", header=None, usecols="A", nrows=2)
aladdin = df2.iloc[0, 0]
sesame  = df2.iloc[1, 0]

print("\n", f" {rptDate.strftime('%A %d %b %Y')} for {len(df[0][1:])} funds:","\n",f" {funds}", "\n")
print(f" {'No' if summ_yn == 'No' else 'A'} summary sheet is required", "\n")
print(f"{timediff(start_time, time.time())} getting the reporting date and fund codes","\n")


# In[16]:


# get the fund holdings in "portfolio analytics review - new" format

start_time = time.time()
print("Downloading and then saving holdings data ...")

if len(full) > 100: # if more than 100 funds are in the list ...
    full = df[0].iloc[1:]
    hlf  = int(len(full)/2)
    # ... get holdings for the first half of funds in the list, and ...
    half1 = (',').join(full[:hlf])
    half1_name = f'PARN half1({len(full[:hlf])}) {rptDate.strftime("%d%b%Y")}.csv'
    osprey("parn", half1, rptDate, rptDate, "half1", "csv", aladdin, sesame)

    # ... get holdings for the second half of funds in the list
    half2 = (',').join(full[hlf:])
    half2_name = f'PARN half2({len(full[hlf:])}) {rptDate.strftime("%d%b%Y")}.csv'
    osprey("parn", half2, rptDate, rptDate, "half2", "csv", aladdin, sesame)
    print('', half1_name, '\n', half2_name, '\n')

    # concat() dataframes of the two halves https://pandas.pydata.org/docs/user_guide/merging.html
    df1 = pd.read_csv(os.path.join(pth_dl, half1_name))
    df2 = pd.read_csv(os.path.join(pth_dl, half2_name))
    df3 = pd.concat([df1, df2])
    # print(len(df1), len(df2), len(df1) + len(df2), len(df3))
    
    # write the combined dataframe to a csv file in the Downloads folder
    df3.to_csv(os.path.join(pth_dl, f'PARN ({len(full[hlf:]) + len(full[:hlf])}) {rptDate.strftime("%d%b%Y")}.csv'), index = False)

else: # else get all the holdings in one go
    full_name = f'PARN ({len(full)}) {rptDate.strftime("%d%b%Y")}.csv'
    osprey("parn", funds, rptDate, rptDate, "", "csv", aladdin, sesame)
    print('', full_name, '\n')

print(f"{timediff(start_time, time.time())} downloading and then saving holdings data","\n")


# In[ ]:


# get the derivative metrics

start_time = time.time()
print('Downloading and then saving derivative data ...')

# def osprey(rpt_type = 'r28i', funds = 'PABS, SMMAIF' as string, d_from as datetime, d_to as datetime, name, sfx = 'csv' as string,
osprey("derv", funds, rptDate, rptDate, "", "csv", aladdin, sesame)

print(f'{timediff(start_time, time.time())} downloading and then saving derivative data','\n')
# C:\Users\hilton.netta\Downloads


# In[ ]:


print(f'{timediff(start_time_derivative_downloading, time.time())} downloading. Next step is compiling.', '\n')


# In[ ]:


get_ipython().system('jupyter nbconvert --to script derv_checker_downloading_csv_new.ipynb # convert from .ipynb to .py')


# In[ ]:




