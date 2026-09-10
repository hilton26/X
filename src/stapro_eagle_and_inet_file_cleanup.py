#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import libraries
import schedule
from datetime import datetime
import time
from time_diff_function import timediff


# In[2]:


def statpro_eagle_and_inet_file_cleanup():
    try:
        #https://stackoverflow.com/questions/72962306/how-to-use-run-command-to-execute-another-notebook-using-file-path
        get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/statpro_eagle_file_cleanup.py"')
        get_ipython().run_line_magic('run', '"C:/Users/hilton.netta/OneDrive - Prescient/py/gitrepo/statpro_inet_file_cleanup.py"')
                
    except Exception as e:
        print(e)


# In[ ]:


# \\Pim-cpt-statpro\Plugins\Profiles\PIM\DataStage\Data\eagle

