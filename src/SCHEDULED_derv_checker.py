#!/usr/bin/env python
# coding: utf-8

# # Derivative Cover Reporting

# ### Input:
# 
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports, "downloader" tab
# 
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_downloading.ipynb "dervs" and "creds" tabs
# 
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_compiling.ipynb "dervs" and "creds" tabs

# In[1]:


# define the report function
def derv_check():
    import time
    start_time_roundtrip = time.time()
    
    # libraries, libraries!
    import os, schedule
    import pandas as pd
    import subprocess
    from datetime import datetime
    from constants import pthPy, pthEXPORTS, frcv_file, pthDaily
    from utilities import timediff, prior_working_day
    
    #check if the report already exists, if yes, exit the script
    df         = pd.read_excel(pthPy, sheet_name="dervs", header=None, usecols="A,D:E").dropna(subset = [0])
    k          = df.iloc[0, 2]
    rptDate    = k if isinstance(k, datetime) else prior_working_day(datetime.today()) # prior working day or report date override given today's date
    filename   = os.path.join(pthEXPORTS, f'Derv {rptDate.strftime("%d%b%Y")}.xlsx')
    df_frcv    = pd.read_excel(frcv_file, sheet_name = "Summary", header = None, usecols = "C", nrows  = 1)
    
    if rptDate.date() == datetime.strptime(df_frcv.iloc[0, 0][15:26],"%d %b %Y").date():
        print(f'{datetime.now().strftime("%Hh%M:%Ss %a %d %b %Y")}: {frcv_file.removeprefix(pthDaily)} \
was completed {time.ctime(os.path.getmtime(frcv_file))}')
        return
    else:
        try:
            print('\n','Starting derv_checker_downloading.ipynb','\n')
            get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_downloading.ipynb"')
            # subprocess.run(["python", "C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_downloading.ipynb"])
            print('\n','derv_checker_downloading completed','\n','Starting derv_checker_compiling.ipynb','\n')
            
            get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_compiling.ipynb"')
            # subprocess.run(["python", "C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_compiling.ipynb"])
            print('\n','derv_checker_compiling completed','\n','Starting derv_checker_summarising','\n')
            
            get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_summarising.ipynb"')
            # subprocess.run(["python", "C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_summarising.ipynb"])
            print('\n','derv_checker_summarising completed','\n','Starting derv_checker_freecover','\n')    
            
            get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_freecover.ipynb"')
            # subprocess.run(["python", "C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_freecover.ipynb"])
            print('\n','derv_checker_freecover completed','\n')     
            
            if os.path.isfile(filename):
                print(f'{datetime.now().strftime("%Hh%M:%Ss %a %d %b %Y")}: \
                {filename.removeprefix(pthEXPORTS)} was completed at {time.ctime(os.path.getmtime(filename))}','\n')

        except Exception as e:
            print(e)

        print("\n", f"{timediff(start_time_roundtrip, time.time())} roundtrip time to download and \
complete {rptDate.strftime('%d %b %Y')} derivative cover reports")


# In[ ]:


# scheduler    https://pypi.org/project/schedule/
import schedule, time

times = ['08:45', '09:15', '09:45', '10:00', '10:15', '10:30', '11:00', '11:30', \
         '12:00', '12:30', '13:15', '13:30', '14:00', '14:30', '15:00', '15:40', \
         '16:00']
for t in times:
    schedule.every().monday.   at(t).do(lambda: derv_check())
    schedule.every().tuesday.  at(t).do(lambda: derv_check())
    schedule.every().wednesday.at(t).do(lambda: derv_check())
    schedule.every().thursday. at(t).do(lambda: derv_check())
    schedule.every().friday.   at(t).do(lambda: derv_check())

while True:
    schedule.run_pending()
    time.sleep(10)


# In[ ]:


# P:\Investment Operations\GRC\Compliance\Derivative Cover


# In[ ]:


# !jupyter nbconvert --to script SCHEDULED_derv_checker.ipynb # convert from .ipynb to .py

