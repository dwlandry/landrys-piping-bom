# -*- coding: utf-8 -*-
"""
Created on Sun Aug 13 07:09:20 2017

@author: dwlan
"""
# In[1]
# Module references
import landrys_piping_bom_reader_helper as helper
import landrys_piping_bom as lpbom

# In[2]
filepath = helper.get_bom_filepath()
bom = lpbom.bom(filepath)
bom.to_excel()