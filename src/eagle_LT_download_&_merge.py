#!/usr/bin/env python
# coding: utf-8

# # Download and Merge Eagle Lookthrough Sheets

# ### Input:
# 
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports, "downloader" and "creds" tabs
# 
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_downloading.ipynb
# 
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_merge.ipynb

# create a time difference function
import time, datetime

print('Downloading and merging lookthrough reports ...', '\n')
start_time          = time.time()

# https://stackoverflow.com/questions/20186344/importing-an-ipynb-file-from-another-ipynb-file   
get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_downloading_csv.py"')
get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/eagle_LT_merge.py"')

print(f'Downloading and merging lookthrough reports completed: {timediff(start_time0, time.time())}')