#!/usr/bin/env python
# coding: utf-8

# # Compile PGF Hedge Share Class Sheet

# ### Inputs
# \\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\Daily\py_reports.xlsx "hdgs" tab

# ### Files and Folders
# \\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges

# In[5]:


# import libraries
import schedule, os
from datetime import datetime
from time_diff_function import timediff


# In[6]:


start_time_0 = datetime.now()

def pgf_check():
    try:
        # run the hedge checker
        get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/pgf_downloading.ipynb"')
        get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/pgf_compiling.ipynb"')

        # open the reporting folder
        pth_pgf = r'\\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges'
        os.startfile(os.path.realpath(pth_pgf))

        print('\n', f'Hedge checker roundtrip time: {timediff(start_time_0, datetime.now())}', '\n')
        
    except Exception as e:
        print(e)


# In[7]:


pgf_check()


# In[8]:


# \\PIM-CPT-FS.prescient.local\PIM-Documents$\Investment Operations\GRC\Compliance\PGF UCITS Share Class Hedges

