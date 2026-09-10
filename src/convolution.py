#!/usr/bin/env python
# coding: utf-8

# # Convolution
# 
# #### But what is a convolution?
# #### https://www.youtube.com/watch?v=KuXjwB4LzSA

# In[6]:


import numpy as np
import timeit

arr1 = np.random.random(100000) # 100k random values
arr2 = np.random.random(100000) # 100k random values


# In[ ]:

%%timeit
np.convolve(arr1, arr2)


# In[ ]:


get_ipython().run_cell_magic('timeit', '', 'scipy.signal.fftconvolve*arr1, arr2)\n')


# In[ ]:


get_ipython().system('jupyter nbconvert --to script convolution.ipynb # convert from .ipynb to .py')


# In[ ]:




