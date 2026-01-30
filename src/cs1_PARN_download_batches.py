#!/usr/bin/env python
# coding: utf-8

# In[4]:


print('###################################')
print('# START cs1_PARN_download_batches #')
print('###################################')


# In[5]:


import time
start_time                   = time.time()
start_time_cs1_PARN_download = start_time

# libraries, libraries!
print('Importing libraries and setting up paths ...')
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
from tqdm import tqdm, notebook
from utilities import timediff, last_working_day, prior_month_end, osprey, batch_list

# set paths to the driver, urls, and report parameters
pth_py = r'P:\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm' # variables stored here
pth_dl = str(Path.home() / 'Downloads')

print(f' {timediff(start_time, time.time())} importing libraries and setting up paths', '\n')


# In[7]:


# get inputs to pass to Eagle
start_time = time.time()
print('Collecting input data ...')

# get fund codes from r28_cs1 tab of the py_report.xlsm sheet
df      = pd.read_excel(pth_py, sheet_name="r28_cs1", usecols="A,F").dropna(subset = ['Fund'])
funds   = df['Fund'].apply(str.upper)

# get report date
k       = df.iloc[0,1]
rptDate = k if isinstance(k, datetime) else prior_month_end(datetime.today()) # prior month end or report date override; type is datetime()

# get credentials
df      = pd.read_excel(pth_py, sheet_name = 'creds', header = None, usecols = 'A', nrows = 2)
aladdin = df.iloc[0, 0]
sesame  = df.iloc[1, 0]

# check inputs
s = "" if len(funds) == 1 else "s"
print(f'{len(funds)} PARN CS1 fund report{s} to be downloaded as at {rptDate.strftime("%A %d %b %Y")}:\n  {(", ").join(funds)}')
print(f'\n{timediff(start_time, time.time())} collecting input data\n')


# In[80]:


# download the PARN reports in batches
start_time = time.time()
num_batches = 2
print(f"Downloading the {len(funds)} fund holdings for {rptDate.strftime('%a %d %b %Y')} in {num_batches} batches ...")

batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size = min(len(funds),batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start = 1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    print(f"{filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}")    
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)
    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath}...\n")
        osprey('parn', (',').join(batch), rptDate, rptDate, fln, 'csv', aladdin, sesame)

print(f"\n{timediff(start_time, time.time())} downloading the {len(funds)} fund holdings for {rptDate.strftime('%a %d %b %Y')} in {num_batches} batches")


# In[81]:


# join the downloaded holding reports into a dataframe
start_time = time.time()
print('Dataframing the fund holding reports for CS1 reporting\n')

df = pd.DataFrame()
for batch_filepath in batch_filepaths:
    # print(f" {batch_filepath}")
    df_new = pd.read_csv(batch_filepath)
    df = pd.concat([df,df_new])

# convert date columns from type object
date_cols = ['Next Coupon Date', 'Maturity Date', 'i Position Effective Date']
for date_col in date_cols:
    df[date_col] = pd.to_datetime(df[date_col])

# convert cvalue columns from object to float
value_cols = ['Original Nominal', 'Clean Book Value', 'Clean Market Value', 'Accrued Income', 
              'Dividend Receivable', 'Sum of Market Value Income', 'Market Value %', 'Current Exposure']
for value_col in value_cols:
    df[value_col] = df[value_col].str.replace(",","").astype(float)

print(f" {len(df['Entity Name'].unique())} fund holdings as at {df['i Position Effective Date'].iloc[0].strftime('%d %b %Y')} in the dataframe")

print(f'\n{timediff(start_time, time.time())} dataframing the fund holding reports for CS1 reporting\n')


# In[ ]:


print('####################################')
print('#   END cs1_PARN_download_batches  #')
print('####################################')


# In[ ]:


# !jupyter nbconvert --to script cs1_PARN_download_csv.ipynb # convert from .ipynb to .py

