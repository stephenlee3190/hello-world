
# coding: utf-8

# In[1]:


import numpy as np
import xlwings as xw
import time
time.sleep(5)  
xw.sheets['Sheet1'].range('A1').value = 'Hello World'

