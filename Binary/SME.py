#!/usr/bin/env python
# coding: utf-8

# In[1]:


import tabula
import os
import logging
from logging.handlers import RotatingFileHandler


# In[2]:


directory = 'C:/Users/lenovo/Desktop/Boston consulting group/Sample DRHPs/SME/Input'


# In[3]:


for pdf in os.listdir(directory):
    logger = logging.getLogger(f'C:/Users/lenovo/Desktop/Boston consulting group/Sample DRHPs/SME/Logs/{pdf}.log')
    hdler = RotatingFileHandler(f'C:/Users/lenovo/Desktop/Boston consulting group/Sample DRHPs/SME/Logs/{pdf}.log', maxBytes=1000)
    logger.addHandler(hdler)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s  %(message)s","%Y-%m-%d %H:%M")
    hdler.setFormatter(formatter)
    logger.info(f'Opened : {pdf}')
    if 'Ushanti' in pdf:
        finance = tabula.read_pdf(directory+'/'+pdf,stream=True,pages=189,silent=True)
        management = tabula.read_pdf(directory+'/'+pdf,stream=True,pages=162,silent=True)
    
    if 'Vaxtex' in pdf:
        finance = tabula.read_pdf(directory+'/'+pdf,stream=True,pages=154,silent=True)
        management = tabula.read_pdf(directory+'/'+pdf,stream=True,pages=130,silent=True)
        
    file_name = pdf.replace('.pdf','').replace('.PDF','')
    logger.info(f'Closed : {pdf}')
    #output
    finance[0].to_excel('C:/Users/lenovo/Desktop/Boston consulting group/Sample DRHPs/SME/Output'+'/'+file_name+'_finance.xlsx',index=False)
    management[0].to_excel('C:/Users/lenovo/Desktop/Boston consulting group/Sample DRHPs/SME/Output'+'/'+file_name+'_management.xlsx',index=False)


# In[ ]:





# In[ ]:




