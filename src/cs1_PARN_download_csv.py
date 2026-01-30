#!/usr/bin/env python
# coding: utf-8

# In[1]:


print('###############################')
print('# START cs1_PARN_download_csv #')
print('###############################')


# In[2]:


import time
start_time                   = time.time()
start_time_cs1_PARN_download = start_time

# libraries, libraries!
print('Importing libraries ...')
from datetime import datetime
import pandas as pd
import os
from pathlib import Path
from tqdm import tqdm, notebook
from constants import pthPy, pth_dl
from utilities import timediff, last_working_day, prior_month_end, osprey, batch_list

print(f' {timediff(start_time, time.time())} importing libraries\n')


# In[3]:


# get inputs to pass to Eagle
start_time = time.time()
print('Collecting input data ...')

# get fund codes from r28_cs1 tab of the py_report.xlsm sheet
df      = pd.read_excel(pthPy, sheet_name="r28_cs1", usecols="A,F").dropna(subset = ['Fund'])
funds   = df['Fund'].apply(str.upper)

# get report date
k       = df.iloc[0,1]
rptDate = k if isinstance(k, datetime) else prior_month_end(datetime.today()) # prior month end or report date override; type is datetime()

# check inputs
print('\n',f'PARN CS1 fund report{"" if len(funds) == 1 else "s"} to be downloaded as at {rptDate.strftime("%A %d %b %Y")}', '\n', \
f'{len(funds)} fund{"" if len(funds) == 1 else "s"}: ','\n',f"{(', ').join(funds)}")
print('\n', f'{timediff(start_time, time.time())} collecting input data', '\n')


# In[4]:


# download the PARN reports in batches
num_batches = 2
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size = min(len(funds),batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start = 1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    print(f"{filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}\n")    
    # print(f" {len(batch)} files:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)
    if os.path.isfile(batch_filepath):
        print(f"\n{batch_filepath} exists\n")
        pass
    else:
        print(f"Downloading batch {index} of {len(batches)} as {batch_filepath}...\n")
        osprey('parn', (',').join(batch), rptDate, rptDate, fln, 'csv')

for batch_filepath in batch_filepaths:
    print(batch_filepath)


# In[6]:


# # determine which files still to be downloaded

# # list and then pick out the .csv files from the Downloads folder and then ...
# start_time = time.time()
# print(f'Identifying the PARN csv files dated {rptDate.strftime("%d %b %Y")} already in the local Downloads folder ...\n')

# # pick out the csv files in the Downloads folder
# import re
# pattern   = r"^PARN.*\(1\) " + f'{datetime.strftime(rptDate, "%d%b%Y")}' + "\.csv$"
# parn_csvs = [s for s in os.listdir(pth_dl) if re.match(pattern, s)]

# # fund PARN reports still to be downloaded
# done = pd.Series(parn_csvs).apply(lambda x: x.replace('PARN ','').replace(f'(1) {rptDate.strftime("%d%b%Y")}.csv',''))
# undone = sorted(list(set(funds) ^ set(done)))

# print(f' {len(parn_csvs)} downloaded + {len(undone)} to be downloaded = {len(funds)} in \
# py_reports.xlsm for {rptDate.strftime("%d %b %Y")}:\n {(", ").join(parn_csvs)}')

# print('\n', f'{timediff(start_time, time.time())} the PARN csv files dated {rptDate.strftime("%d %b %Y")} already in the local Downloads folder', '\n')

# # iterate individually through the list of reports to be downloaded from Eagle

# start_time = time.time()
# print(f'Downloading the {len(undone)} fund{"" if len(undone) == 1 else "s"} remaining fund PARN holdings for CS1 reporting \
# as at {rptDate.strftime("%A %d %B %Y")} ...', '\n')

# for fund in notebook.tqdm(undone):
#     # def osprey(rpt_type = 'r28i', funds = 'PABS,PIMBAL', d_from as datetime, d_to as datetime, name, sfx = 'xls', al, xe):
#     osprey('parn', fund, rptDate, rptDate, fund, 'csv')

# # open the Downloads folder
# os.startfile(os.path.realpath(Path.home() / 'Downloads'))

# print('\n\n', f'{timediff(start_time, time.time())} downloading the {len(undone)} fund{"" if len(undone) == 1 else "s"} remaining fund PARN holdings \
# for CS1 reporting as at {rptDate.strftime("%d %B %Y")} ({(time.time() - start_time_cs1_PARN_download) / (len(funds)):,.1f}sec/fund)')

# # C:\Users\hilton.netta\Downloads


# In[7]:


print('################################')
print('#   END cs1_PARN_download_csv  #')
print('################################')


# In[8]:


# !jupyter nbconvert --to script cs1_PARN_download_csv.ipynb # convert from .ipynb to .py


# In[ ]:




