#!/usr/bin/env python
# coding: utf-8

# # Merge the CS1 fund PARN sheets into one file to be used by cs1_reporting.ipynb

# https://stackoverflow.com/questions/20908018/import-multiple-excel-files-into-python-pandas-and-concatenate-them-into-one-dat

# In[ ]:


print('###############################')
print('#  START cs1_PARN_merge_csv   #')
print('###############################')


# In[19]:


# libraries, libraries!
import time
start_time      = time.time()
start_time_parn = time.time()

from datetime import datetime
import pandas as pd
from pathlib import Path
import os
from tqdm import tqdm, notebook # notebook version of tqdm
from constants import pthPy, pth_dl, pthReports, pthTest
from utilities import timediff, prior_month_end, batch_list, osprey


# In[8]:


# get inputs to pass to Eagle
start_time = time.time()
print('Collecting the report input data ...')

# get fund codes from r28_cs1 tab of the py_report.xlsm sheet
df = pd.read_excel(pthPy, sheet_name="r28_cs1", usecols="A,F").dropna(subset = ['Fund'])
funds = df['Fund'].apply(str.upper)

# get report date
k = df.iloc[0,1]
rptDate = k if isinstance(k, datetime) else prior_month_end(datetime.today()) # prior month end or report date override; type is datetime()

# check inputs
print(f'\nPARN CS1 fund report{"" if len(funds) == 1 else "s"} to be merged as at {rptDate.strftime("%A %d %b %Y")}\n \
{len(funds)} fund{"" if len(funds) == 1 else "s"}: \n{(",").join(funds)}')
print(f'\n{timediff(start_time, time.time())} collecting the report input data\n')


# In[17]:


# ... dataframe the fund holdings in PARN format by looping over their files and appending to an initially empty dataframe
start_time = time.time()
print(f'Merging the PARN csv files into a dataframe ...')

num_batches = 2
batch_size = int(len(funds) / num_batches)
batches = batch_list(funds, batch_size = min(len(funds),batch_size))
batch_filepaths = []
for index, batch in tqdm(enumerate(batches, start = 1)):
    fln = f"{index}_of_{len(batches)}_CS1"
    filename = f"PARN {fln}({len(batch)}) {rptDate.strftime('%#d%b%Y')}.csv"
    # print(f"{filename}, a batch of {len(batch)} files:\n   {(', ').join(batch)}\n")    
    # print(f" {len(batch)} files:\n   {(', ').join(batch)}\n")
    batch_filepath = os.path.join(pth_dl, filename)
    batch_filepaths.append(batch_filepath)

holdings = pd.DataFrame() # initialise an empty dataframe
for batch_filepath in batch_filepaths:
    # print(batch_filepath)
    data       = pd.read_csv(batch_filepath)
    holdings   = pd.concat([holdings, data])

# # check
# len(holdings['Entity Name'].unique())

# convert the holdings CS1 fund NAV columns to float64
cols_to_sum = ['Sum of Market Value Income', 'Current Exposure']
for col in cols_to_sum:
    holdings[col] = holdings[col].str.replace(',', '').astype('float64')

print(f' {timediff(start_time, time.time())} merging the PARN csv files with {len(holdings["Entity Name"].unique())} funds into a dataframe\n')


# In[20]:


# get the fund NAVs
start_time = time.time()
print(f"Getting the {len(funds)} CS1 funds' NAVs as at {rptDate.strftime('%A %d %B %Y')} with osprey() ...")

# get credentials
df      = pd.read_excel(pthPy, sheet_name = 'creds', header = None, usecols = 'A', nrows = 2)
aladdin = df.iloc[0, 0]
sesame  = df.iloc[1, 0]

name = 'CS1'
osprey('fnav', (',').join(funds), rptDate, rptDate, name, 'csv')

# dataframe the fund NAVs
navs_fln = os.path.join(Path.home(), "Downloads", f'FNAV {name}({len(funds)}) {rptDate.strftime("%d%b%Y")}.csv')
navs     = pd.read_csv(navs_fln)

# # convert the Total column from object to float
navs['Total Net Assets'] = navs['Total Net Assets'].str.replace(',', '').astype('float64')
# navs['Total Net Assets'] = navs['Total Net Assets'].apply(lambda x: f"{x:,.2f}") # present with thousands separator and to two decimals


print(f'\n{navs_fln}\n')

print(f" {timediff(start_time, time.time())} getting the {len(funds)} funds' NAV{'s' if len(funds) != 1 else ''} as at {rptDate.strftime('%A %d %B %Y')} \
with osprey()")


# In[21]:


# merge the CS1 fund holdings and NAVs, and compare their totals 
start_time = time.time()
print(f'\nMerging and comparing the {len(funds)} CS1 fund holdings and NAVs as at {rptDate.strftime("%A %d %B %Y")} ...')

holdings_totals = holdings.groupby('Entity ID', as_index = False).sum()[['Entity ID','Sum of Market Value Income','Current Exposure']]
sums_cf = holdings_totals.merge(navs, how = 'left', left_on = 'Entity ID', right_on = 'NAV Entity ID')
sums_cf.drop(['Entity Name', 'NAV Entity ID'], axis = 1, inplace = True)
sums_cf['SoMVI-CE']  = sums_cf['Sum of Market Value Income'] - sums_cf['Current Exposure']
sums_cf['1-CE/SoMVI %']  = (1 - sums_cf['Current Exposure']/sums_cf['Sum of Market Value Income']) * 100
sums_cf['SoMVI-NAV'] = abs(sums_cf['Sum of Market Value Income'] - sums_cf['Total Net Assets'])
sums_cf['1-NAV/SoMVI %']  = (1 - sums_cf['Total Net Assets']/sums_cf['Sum of Market Value Income']) * 100
sums_cf = sums_cf.sort_values(by='SoMVI-NAV', ascending=False)
cols_order = ['Effective Date','Entity ID','Sum of Market Value Income','Current Exposure','Total Net Assets',\
              'SoMVI-CE','1-CE/SoMVI %','SoMVI-NAV','1-NAV/SoMVI %']
sums_cf = sums_cf[cols_order]

# set the number of decimals to present
cols_2dp = ['Sum of Market Value Income','Current Exposure','Total Net Assets']
for col in cols_2dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.2f}")

cols_6dp = ['SoMVI-CE', '1-CE/SoMVI %','SoMVI-NAV','1-NAV/SoMVI %']
for col in cols_6dp:
    sums_cf[col] = sums_cf[col].apply(lambda x: f"{x:,.6f}")

# sums_cf

print(f' {timediff(start_time, time.time())} merging and comparing the {len(funds)} CS1 fund holdings and NAVs as at {rptDate.strftime("%A %d %B %Y")}')


# In[22]:


# convert the PARN holdings into Reg 28 format with correspodning headings
start_time = time.time()
print(f'\nConverting the CS1 fund PARN holdings in readiness for Reg 28 classification ...')

s = "" if len(funds) == 1 else "s"
cs1_fname = os.path.join(pthTest, f'CS1 PARN holdings ({len(funds)}) {rptDate.strftime("%d%b%Y")}.xlsx')

hold_cols = ['Entity ID', 'Investment Type','i Issue Name','PrimaryAssetID','CCY','Sum of Market Value Income',\
                           '% of Total Market Value','Current Exposure']
hReg28 = holdings[hold_cols] # identify the subset of holdings columns to be used
hReg28 = hReg28.rename(columns={'Entity ID': 'Entity Name', 'PrimaryAssetID': 'Primary Asset ID', 'Sum of Market Value Income': 'End Market Value',
                               '% of Total Market Value': 'Percentage of Market Value', 'Current Exposure': 'Closing Exposure PA'})
hReg28.insert(5, 'Reg28 Classification','') # insert the classification column as the new column 5
hReg28.insert(9, f'{rptDate.strftime("%d %b %Y")}','') # insert the report date as a header in the last column
hReg28.iloc[0,9] = cs1_fname
hReg28.iloc[1,9] = "CS1"
hReg28.reset_index(drop=True, inplace=True)

# hReg28.info()

print(f'\n {timediff(start_time, time.time())} converting the CS1 fund PARN holdings in readiness for Reg 28 classification\n')


# In[23]:


# write the CS1 holdings dataframe to review it as a worksheet
start_time = time.time()
print(f'\nWriting the CS1 fund holdings dataframe and navs dataframe to a sheet ...\n')

writer = pd.ExcelWriter(cs1_fname, engine = 'xlsxwriter')     # instantiate a sheet writer
hReg28.to_excel(  writer, index = False, sheet_name = 'All')  # write the NAV sheet
holdings.to_excel(writer, index = False, sheet_name = 'PARN')  # write the NAV sheet
sums_cf.to_excel( writer, index = False, sheet_name = 'NAVs') # write the missing NAVs sheet
writer.close() # https://pandas.pydata.org/docs/reference/api/pandas.ExcelWriter.html   class for writing DataFrame objects into excel sheets
print(f'\n{cs1_fname}\n')

print(f'\n {timediff(start_time, time.time())} writing the CS1 fund holdings dataframe and navs dataframe to a sheet\n')


# In[9]:


# update the classifier sheet in py_reports.xlsm in preparation for eagle_r28_r30_classifications.ipynb
start_time = time.time()
print(f'\nUpdating py_reports.xlsm "classifier" sheet with path to the look-through csv holdings file ...\n')
# https://stackoverflow.com/questions/13381384/modify-an-existing-excel-file-using-openpyxl-in-python

import xlwings as xw
wb                   = xw.Book(pthPy)
ws                   = wb.sheets('classifier')
ws.range("L2").value = cs1_fname
ws.range("M1").value = 'CS1 format only' # alternative: 'CS1 format only'
wb.save()
wb.close()

print(f'\n {timediff(start_time, time.time())} updating py_reports.xlsm "classifier" sheet with path to the look-through csv holdings file\n')


# In[10]:


# Reg 28 & 30  P:\Investment Operations\GRC\Compliance\Reg28 and Reg30 Reporting
# folder       P:\Working Folders\Hilton\W\Reg_Tests


# In[ ]:


print('###############################')
print('#   END cs1_PARN_merge_csv    #')
print('###############################')


# In[1]:


# !jupyter nbconvert --to script cs1_PARN_merge_csv.ipynb # convert from .ipynb to .py

