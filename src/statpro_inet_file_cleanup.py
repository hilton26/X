#!/usr/bin/env python
# coding: utf-8

# In[2]:


import pandas as pd
import numpy as np


# In[3]:


import os, shutil, time, datetime
import re


# In[4]:


inet_eod  = r'\\Pim-cpt-statpro\Plugins\Profiles\PIM\DataStage\Data\inet-eod'
inet_icdf = r'\\Pim-cpt-statpro\Plugins\Profiles\PIM\DataStage\Data\inet-icdf'


# In[5]:


# get lname of latest icdf file and where it will be backed up
icdf_fnames  = [int(s[re.search('\d{8}',s).span()[0]:re.search('\d{8}',s).span()[0]+8]) for s in 
                os.listdir(inet_icdf + r'\copy') if "." in s]
latest_icdf = inet_icdf + r'\copy\ICDFXML'   + str(max(icdf_fnames)) + '-2200.XML'
backup_icdf = inet_icdf + r'\Backup\ICDFXML' + str(max(icdf_fnames)) + '-2200.XML'

print('', latest_icdf, '\n', backup_icdf)


# In[6]:


# save the latest icdf file to Backup folder - https://www.geeksforgeeks.org/python-shutil-copy2-method/
shutil.copy2(latest_icdf, backup_icdf)


# In[7]:


# get lname of latest eod files and where they will be backed up
eod_fnames  = [int(s[re.search('0\d{8}',s).span()[0]:re.search('0\d{8}',s).span()[0]+9]) for s in 
                os.listdir(inet_eod + r'\copy') if "." in s]
latest_dcc = inet_eod + r'\copy\DCC0'   + str(max(eod_fnames)) + '.dat'
latest_dfc = inet_eod + r'\copy\DFC0'   + str(max(eod_fnames)) + '.dat'
latest_dpc = inet_eod + r'\copy\DPC0'   + str(max(eod_fnames)) + '.dat'

backup_dcc = inet_eod + r'\backup\DCC0' + str(max(eod_fnames)) + '.dat'
backup_dfc = inet_eod + r'\backup\DFC0' + str(max(eod_fnames)) + '.dat'
backup_dpc = inet_eod + r'\backup\DPC0' + str(max(eod_fnames)) + '.dat'

print('', latest_dcc, '\n', latest_dfc, '\n', latest_dpc, '\n\n', backup_dcc, '\n', backup_dfc, '\n', backup_dpc)


# In[8]:


# save the latest eod files to Backup folder - https://www.geeksforgeeks.org/python-shutil-copy2-method/
shutil.copy2(latest_dcc, backup_dcc)
shutil.copy2(latest_dfc, backup_dfc)
shutil.copy2(latest_dpc, backup_dpc)


# In[21]:


# delete files outside \inet_icdf\copy\ folder
for file in os.listdir(inet_icdf):
    if os.path.isfile(os.path.join(inet_icdf, file)):
        print(    os.path.join(inet_icdf, file))
        os.remove(os.path.join(inet_icdf, file))

# delete files inside \inet_icdf\copy\ folder
for file in os.listdir(inet_icdf + r'\copy'):
    if os.path.isfile(os.path.join(inet_icdf, 'copy', file)):
        print(    os.path.join(inet_icdf, 'copy', file))
        os.remove(os.path.join(inet_icdf, 'copy', file))
        
# copy the latest inet_icdf file back to the \inet_icdf\copy\ folder
shutil.copy2(backup_icdf, latest_icdf)


# In[25]:


# delete files outside \inet_eod\copy\ folder
for file in os.listdir(inet_eod):
    if os.path.isfile(os.path.join(inet_eod, file)):
        print(    os.path.join(inet_eod, file))
        os.remove(os.path.join(inet_eod, file))

# delete files inside \inet_icdf\copy\ folder
for file in os.listdir(inet_eod + r'\copy'):
    if os.path.isfile(os.path.join(inet_eod, 'copy', file)):
        print(    os.path.join(inet_eod, 'copy', file))
        os.remove(os.path.join(inet_eod, 'copy', file))
        
# copy the latest inet_eod files back to the \inet_eod\copy\ folder
shutil.copy2(backup_dcc, latest_dcc)
shutil.copy2(backup_dfc, latest_dfc)
shutil.copy2(backup_dpc, latest_dpc)


# In[ ]:




