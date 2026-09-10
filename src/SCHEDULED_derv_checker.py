#!/usr/bin/env python
# coding: utf-8

# # Derivative Cover Reporting

# ### Input:
# 
# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports, "downloader" tab
# 
# ### Dependencies:
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_downloading.ipynb "dervs" tab
# 
# C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/derv_checker_compiling.ipynb "dervs" tab

# In[13]:


# define the report function
def derv_run():
    import time
    start_time_roundtrip = time.time()
    
    # libraries, libraries!
    import os, schedule
    import pandas as pd
    import subprocess
    from datetime import datetime
    from constants import derv_checker #pthPy, pthEXPORTS, frcv_file, pthDaily, dc_do_nb, dc_co_nb, dc_su_nb, dc_fr_nb
    from utilities import timediff #, prior_working_day

    get_ipython().run_line_magic('run', '"$derv_checker" # downloading')

    #print(rf"\n{timediff(start_time_roundtrip, time.time())} roundtrip time to download and complete {rptDate.strftime('%d %b %Y')} derivative cover reports")


# In[14]:


derv_run()


# In[ ]:





# In[ ]:





# In[ ]:





# In[2]:


# list the frequency of runs

from datetime import datetime, timedelta

# list of times every 15 minutes from 8:45 AM to 17:30
start = datetime.strptime("08:23", "%H:%M")
end = datetime.strptime("18:30", "%H:%M")
increment = 15   # every 15 minutes

times = []
current = start
while current <= end:
    times.append(current.strftime("%H:%M"))
    current += timedelta(minutes=increment)

print((", ").join(times))


# In[3]:


# scheduler    https://pypi.org/project/schedule/
import schedule, time
         
for t in times:
    schedule.every().monday.   at(t).do(lambda: derv_run())
    schedule.every().tuesday.  at(t).do(lambda: derv_run())
    schedule.every().wednesday.at(t).do(lambda: derv_run())
    schedule.every().thursday. at(t).do(lambda: derv_run())
    schedule.every().friday.   at(t).do(lambda: derv_run())

while True:
    schedule.run_pending()
    time.sleep(10)

