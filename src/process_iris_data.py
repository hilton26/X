#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
 
# the types of data within the iris dataset
col_names = ["sepal_length", "sepal_width", "petal_length", "petal_width", "subspecies"]

pthX = r'C:\Users\hilton.netta\OneDrive - Prescient\py\gitrepo'
iris_data = pd.read_csv(r'C:\Users\hilton.netta\OneDrive - Prescient\py\gitrepo\iris.csv', 
                        sep=',', header=None, names=col_names, index_col=None)


# In[2]:


#%%
iris_setosa = iris_data[iris_data['subspecies']=='Iris-setosa']
iris_setosa=iris_setosa.drop(['subspecies'], axis=1)
 
iris_versicolor = iris_data[iris_data['subspecies']=='Iris-versicolor']
iris_versicolor=iris_versicolor.drop(['subspecies'], axis=1)
 
iris_virginica = iris_data[iris_data['subspecies']=='Iris-virginica']
iris_virginica=iris_virginica.drop(['subspecies'], axis=1)
 
#%%
mean_setosa = iris_setosa.mean(axis=0)
std_setosa = iris_setosa.std(axis=0)
 
mean_versicolor = iris_versicolor.mean(axis=0)
std_versicolor = iris_versicolor.std(axis=0)
 
mean_virginica = iris_virginica.mean(axis=0)
std_virginica = iris_virginica.std(axis=0)
 
subspecies = ['Setosa', 'Versicolor', 'Virginica']
 
mean_vals = pd.concat([mean_setosa, mean_versicolor, mean_virginica], axis=1)
mean_vals.columns=subspecies
 
std_vals = pd.concat([std_setosa, std_versicolor, std_virginica], axis=1)
std_vals.columns=subspecies
 
mean_vals.to_csv(r'C:\Users\hilton.netta\OneDrive - Prescient\py\gitrepo\iris_means.csv', 
                 sep=',', header=True, index=True)
std_vals.to_csv(r'C:\Users\hilton.netta\OneDrive - Prescient\py\gitrepo\iris_std.csv', 
                sep=',', header=True, index=True)


# In[ ]:




