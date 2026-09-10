#!/usr/bin/env python
# coding: utf-8

# # Compile PGF Hedge Share Class Sheet

# ### Inputs
# \\\\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports.xlsx "hdgs" tab

# ### Files and Folders
# \\\\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges

# In[1]:


# define the report function
def pgf_check():
    import time
    start_time_pgf = time.time()
    
    # libraries, libraries!
    import schedule, os
    import pandas as pd
    from pathlib import Path
    from datetime import datetime
    from utilities import timediff, prior_working_day
    from constants import pthPy, pthHdg
    
    # get report date
    df       = pd.read_excel(pthPy, sheet_name = 'hdgs', usecols = 'A:B, D')
    k        = df.iloc[0,2]
    rptDate  = k if isinstance(k, datetime) else prior_working_day(datetime.today()) # prior working day or report date override
    
    # construct file names
    UTs_name = os.path.join(Path.home(), 'Downloads', f'UTPS PGF_UT_prices({len(df["Unit Trust prices"].dropna())}) {rptDate.strftime("%d%b%Y")}.csv')
    NAV_name = os.path.join(Path.home(), 'Downloads',  f'PARN PGF_Holdings({len(df["PAR-N"            ].dropna())}) {rptDate.strftime("%d%b%Y")}.csv')
    
    filename = pthHdg + fr'\{rptDate.strftime("%Y%m%d")} PGF Share Class Hedges.xlsx'
    if os.path.isfile(filename):
        print(f'{datetime.now().strftime("%Hh%M:%Ss %a %d %b %Y")}: {filename.removeprefix(pthHdg)} \
was completed at {time.ctime(os.path.getmtime(filename))}')
        # https://www.geeksforgeeks.org/python-os-path-getmtime-method/
        return
    else:
        try:
            # run the hedge checker
            get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/pgf_downloading.ipynb"')
            if os.path.isfile(UTs_name) and os.path.isfile(NAV_name):
                # FileNotFoundError: [Errno 2] No such file or directory: 'C:\\Users\\hilton.netta\\Downloads\\UTPS PGF_UT_prices(7) 05Mar2025.csv'
                get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/pgf_compiling.ipynb"')
            else:
                if os.path.isfile(NAV_name):
                    print(f"Missing the UT prices file: {UTs_name}")
                else:
                    print(f"Missing the holdings file:  {NAV_name}")
    
            # # open the reporting folder
            # pth_pgf = r"\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges"
            # os.startfile(os.path.realpath(pth_pgf))
        
        except Exception as e:
            print(e)

        print("\n", f"{timediff(start_time_pgf, time.time())} roundtrip time complete the {rptDate.strftime('%a %d %b %Y')} pgf report")


# In[ ]:


# scheduler    https://pypi.org/project/schedule/
import schedule, time

times = ['13:00', '13:15', '13:30', '13:45', '14:00', \
         '14:30', '15:00', '15:30', '16:00', '16:30']
for t in times:
    schedule.every().monday.   at(t).do(lambda: pgf_check())
    schedule.every().tuesday.  at(t).do(lambda: pgf_check())
    schedule.every().wednesday.at(t).do(lambda: pgf_check())
    schedule.every().thursday. at(t).do(lambda: pgf_check())
    schedule.every().friday.   at(t).do(lambda: pgf_check())

while True:
    schedule.run_pending()
    time.sleep(10)


# In[ ]:


# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges


# In[ ]:


# !jupyter nbconvert --to script SCHEDULED_pgf_checker.ipynb # convert from .ipynb to .py

