#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import libraries
import pandas as pd
import numpy as np
import schedule
from datetime import datetime
import time
from time_diff_function import timediff


# In[563]:


start_time  = datetime.now()
print(f'Amending \eagle\.. files before StatPro run ...')


# In[564]:


# get instruments.csv and Holdings.csv data in eagle folder
eagle_path  = r'\\Pim-cpt-statpro\Plugins\Profiles\PIM\DataStage\Data\eagle'
holdings    = pd.read_csv(eagle_path + r'\Holdings.csv')
instruments = pd.read_csv(eagle_path + r'\Instruments.csv', on_bad_lines = 'skip')
# in case there's a tokenization error:
 #https://saturncloud.io/blog/how-to-fix-python-pandas-error-tokenizing-data/#:~:text=Fixing%20the%20Error,
 #-Here%20are%20some&text=The%20first%20step%20is%20to,it%20can%20be%20read%20properly


# In[515]:


# get lookups for missing Instruments.csv and Holdings.csv data
py_reports   = r'\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports.xlsm'
currency     = pd.read_excel(py_reports, sheet_name = 'statpro', usecols = "A:B")
# for "NA" passing as NaN - https://stackoverflow.com/questions/41417214/prevent-pandas-from-reading-na-as-nan

# get issuer names and shorter issuer name replacements
issuer_name  = pd.read_excel(py_reports, sheet_name = 'statpro', usecols = "E:F").dropna(axis = 0, how = 'all')


# In[516]:


# function to get country bigramme given currency trigramme
def cntry(curr):
    if currency['CURRENCY_CODE'].isin([curr]).any():
        return currency[currency['CURRENCY_CODE'] == curr].iat[0, 1]
    else:
        return 'US'
    
cntry('NAD') # test the function


# In[517]:


# function to shorten issuer name to 50 characters
def issuer(txt):
    if issuer_name['ISSUER_LONG_NAME'].isin([txt]).any():
        return issuer_name[issuer_name['ISSUER_LONG_NAME'] == txt].iat[0,1]
    else:
        return txt[:50]

# test the function
text1 = 'TOM BURKE COMMUNITY TRUST INVESTMENT SPV (PTY) LTD RF'
text2 = '123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890'
print(issuer(text1), len(text1), len(issuer(text1)))
print(issuer(text2), len(text2), len(issuer(text2)))


# In[518]:


# identify the empty entries in the "COUNTRY_CODE" column ("NA" for Namibia returns a NaN)
k = instruments.loc[(instruments['COUNTRY_CODE'].isnull())  &
                    (instruments['CURRENCY_CODE'] != 'NAD') &
                    (instruments['ISSUE_NAME']    != 'NAMIBIA')]

# identify the > 50 ISSUE_NAME securities in Instruments.csv
p = instruments[instruments['ISSUE_NAME'].str.len() > 50]

# identify the > 50 ISSUERCODE securities in Holdings.csv
q = holdings[holdings['ISSUERCODE'].str.len() > 50]

# summarise missing data
print('(1) ' + str(len(k)) + ' instruments with empty COUNTRY_CODE')
print('(2) ' + str(len(p)) + ' instruments with ISSUE_NAME over 50')
print('(3) ' + str(len(q)) + ' holding with ISSUERCODE over 50')


# In[519]:


# (1) update missing instruments COUNTRY_CODE
if len(k) > 0:
    for row_number in range(len(k)):
        instruments.at[k.index[row_number], 'COUNTRY_CODE'] = cntry(instruments.at[k.index[row_number], 'CURRENCY_CODE'])
    # check that COUNTRY_CODE was updated
    instruments[k.index[0]:k.index[len(k) - 1] + 1]


# In[520]:


# (2) update long, i.e., > 50 characters, ISSUE_NAME in instruments dataframe
if len(p) > 0:
    for row_number in range(len(p)):
        instruments.at[p.index[row_number], 'ISSUE_NAME'] = issuer(instruments.at[p.index[row_number], 'ISSUE_NAME'])
    # check that ISSUE_NAME in instruments was updated
    instruments[p.index[0]:p.index[len(p) - 1] + 1]


# In[521]:


# (3) update long, i.e., > 50 characters, ISSUERNAME in holdings dataframe
if len(q) > 0:
    for row_number in range(len(q)):
        holdings.at[q.index[row_number], 'ISSUERCODE'] = issuer(holdings.at[q.index[row_number], 'ISSUERCODE'])
    # check that ISSUERCODE in holdings was updated
    holdings[q.index[0]:q.index[len(q) - 1] + 1]


# In[522]:


# save the datframes over the Instruments.csv and Holdings.csv files in the eagle folder
#https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.to_csv.html
instruments.to_csv(eagle_path + r'\Instruments.csv', index  = False)
holdings.to_csv(   eagle_path + r'\Holdings.csv'   , index  = False)


# In[ ]:


print(f'Amending \eagle\.. files before StatPro run ...: {timediff(start_time, time.time())}')

