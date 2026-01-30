#!/usr/bin/env python
# coding: utf-8

# ### Update reporting dates on tabs r28_tbl2, r28ib, and r28_cs1 and forex rate on r28_cs1 in preparation for running reg28_tbl2_+_schib_+_me17.ipynb and cs1_r28.ipynb
# 

# In[3]:


print('##############################################')
print('#        START month_end_updates.ipynb       #')
print('##############################################')


# In[4]:


# update the classifier sheet in py_reports.xlsm in preparation for reg28_tbl2_+_schib_+_me17.ipynb
print("============================== \n STARTING report date updates \n==============================")
import time
start_time = time.time()
from utilities import timediff
import pandas as pd
from datetime import datetime

print('Updating py_reports.xlsm "classifier" sheet with month-end date and ZAR/USD exchange rate ...')

# get report date and corresponding ZAR/USD exchnage rate
pthPy    = r'\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm'
df = pd.read_excel(pthPy, sheet_name = 'funds_list', usecols = 'V,X', nrows = 1) # month-end reporting date and ZAR/USD exchange rate
rptDate = df.columns[0].date()
fxRate  = df.columns[1]
print('',rptDate, fxRate)

# get list of Reg 28 funds
df2  = pd.read_excel(pthPy, sheet_name = 'downloader', usecols = 'A').dropna() # list of Reg 28 and Reg 30 funds
df_m = pd.read_excel(pthPy, sheet_name = 'funds',      usecols = 'N').dropna() # med scheme funds to be excluded from Reg 28 set
# print(df2.shape, df_m.shape)
mask   = ~df2['Entity Name'].isin(df_m.iloc[:,0].tolist()) # create a boolean mask: True for rows NOT in the excluded list
df_r28 = df2[mask]# filter the DataFrame using the mask
print(f' {len(df_r28)} Reg 28 funds = {df2.shape[0]} total funds - {df_m.shape[0]} Reg 30 funds')
last_row = len(df_r28) + 1
r28_tr = list(map(list, zip(df_r28.iloc[:,0].tolist()))) # transposed

import xlwings as xw

wb_py = xw.Book(pthPy) # visible=False runs Excel in the background
visible = False

# update report date and forex rate
# ['r28_tbl2', 'r28ib', 'r28_cs1']

# update r28_tbl2 sheet with reporting date and list of Reg 28 funds
ws = wb_py.sheets('r28_tbl2')
ws.range("F2").value = rptDate
ws.range("D2").value = 'Reg28'
src = ws.range('A1').expand('down').value # get number of contents of the "A" column
ws.range("A2:A" + f"{str(len(src))}").clear_contents() # clear contents all the way down to the last element
ws.range("A2:A" + f"{last_row}").value = r28_tr # use zip to transpose a list - r29_tbl2

# update USD/ZAR rate on the funds tab
ws = wb_py.sheets('funds')
ws.range("W2").value = rptDate
ws.range("Y2").value = rptDate
ws.range("AA2").value = fxRate

# update r28ib sheet with reporting date and list of Reg 28 funds
ws = wb_py.sheets('r28ib')
ws.range("C3").value = rptDate
src = ws.range('A1').expand('down').value # get number of conetnetsof the "A" column
ws.range("A2:A" + f"{str(len(src))}").clear_contents() # clear contents all the way down to the last element
ws.range("A2:A" + f"{last_row}").value = r28_tr # use zip to transpose a list - r28ib

# update r28_cs1 sheet with reporting date, ZARUSD rate and list of Reg 28 funds
ws = wb_py.sheets('r28_cs1')
ws.range("F2").value = rptDate
ws.range("H2").value = fxRate
src = ws.range('A1').expand('down').value # get number of conetnetsof the "A" column
ws.range("A2:A" + f"{str(len(src))}").clear_contents() # clear contents all the way down to the last element
ws.range("A2:A" + f"{last_row}").value = r28_tr # use zip to transpose a list - r28_cs

# update 'classifier' sheet to 'CS1 format only'
ws = wb_py.sheets('classifier')
ws.range("M1").value = 'CS1 format only'
pth_lt = ws.range("L2").value

# save and close py_reports.xlsm workbook
wb_py.save()
wb_py.close()

# update reporting date in MonthEnd17.xlsm
pthME17 = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Reporting Requirements\Monthly Reports\MonthEnd17.xlsm"
wb_py = xw.Book(pthME17) # visible=False runs Excel in the background
visible = False

# update r28_tbl2 sheet with reporting date and list of Reg 28 funds
ws = wb_py.sheets('Process')
ws.range("C2").value = rptDate

# save and close py_reports.xlsm workbook
wb_py.save()
wb_py.close()

print('', pth_lt)

# confirm that issuers_2_31Aug2025.xlsx and issuers_3_31Aug2025.xlsx exist
# ...

print(f'{timediff(start_time, time.time())} updating py_reports.xlsm "classifier" sheet with month-end date and ZAR/USD exchange rate \n')
print("\n=============================== \n FINISHING report date updates \n===============================")


# In[5]:


print('##############################################')
print('#        END month_end_updates.ipynb         #')
print('##############################################')


# In[7]:


# !jupyter nbconvert --to script month_end_updates.ipynb

