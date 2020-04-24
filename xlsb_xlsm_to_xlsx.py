#!/usr/bin/env python
# coding: utf-8

# In[1]:


"""
Authors: Ritvik Kapila and Gauri Gupta
"""

# You will need to install python xlsb library and update pandas, if not already done, on or after 29 January, using the following command for reading xlsb files


# In[2]:


# pip install pandas --upgrade


# In[3]:


# pip install pyxlsb


# In[4]:


# Now continue with the code after updation


# In[2]:


import pandas as pd


# In[3]:


# Function for conversion of xlsb file to xlsx

def xlsb2xlsx(filename, output_filename):
    dfs = pd.read_excel(filename, engine = 'pyxlsb', sheet_name = None)
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine = 'xlsxwriter')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name = sheet_name, index = False)
    writer.save()    
    


# In[4]:


# Function for conversion of xlsm file to xlsx

def xlsm2xlsx(filename, output_filename):
    dfs = pd.read_excel(filename, sheet_name = None)
    writer = pd.ExcelWriter(output_filename + '.xlsx', engine = 'xlsxwriter')
    for sheet_name in dfs.keys():
        dfs[sheet_name].to_excel(writer, sheet_name = sheet_name, index = False)
    writer.save()    
    


# In[7]:


xlsb2xlsx('test_xlsb.xlsb', 'out_xlsb')


# In[8]:


xlsm2xlsx('test_xlsm.xlsm', 'out_xlsm')


# In[ ]:




