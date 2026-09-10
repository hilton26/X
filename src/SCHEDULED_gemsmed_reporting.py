#!/usr/bin/env python
# coding: utf-8

# In[14]:


def gemsmed_rpt():
    import time
    import os
    from datetime import datetime
    start_time_gemsmed = time.time()
    import pandas as pd
    from constants import pthPy, pthTest, issuers_1_nb, gemsmed_report_nb
    from utilities import timediff, prior_working_day

    # get report input data
    df       = pd.read_excel(pthPy, sheet_name = "arc", usecols = "K").dropna()
    df1      = pd.read_excel(pthPy, sheet_name = "arc", usecols = "L", nrows = 2)
    k        = df1.iloc[1,0]
    rptDate  = k if k == k else prior_working_day(datetime.today())

    # derive report file name
    filename = os.path.join(pthTest, f'GEMSMEDC Reg30 {rptDate.strftime("%d%b%Y")}.xlsx')

    # check if the report already exists, if yes, exit the script
    if os.path.isfile(filename):
        print(f'{filename.removeprefix(pthTest)} was completed {time.ctime(os.path.getmtime(filename))}')
        return
    else:
        try:
            get_ipython().run_line_magic('run', '"$gemsmed_report_nb" # downloading')
            get_ipython().run_line_magic('run', '"$issuers_1_nb" # classifying and reporting')

        except Exception as e:
            print(e)

    print("\n", f"{timediff(start_time_gemsmed, time.time())} roundtrip time")


# In[15]:


# scheduler    https://pypi.org/project/schedule/
import schedule, time

times = ['09:10', '09:40', '10:10', '10:40', \
         '11:10', '11:40', '12:10', '12:40', \
         '13:10', '13:40', '14:10', '14:40', \
         '15:10', '15:40', '16:10', '16:40']
for t in times:
    schedule.every().monday.at(t).do(lambda: gemsmed_rpt())

while True:
    schedule.run_pending()
    time.sleep(10)


# In[ ]:


# !jupyter nbconvert --to script SCHEDULED_gemsmed_reporting.ipynb # convert from .ipynb to .py

