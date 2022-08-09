# -*- coding: utf-8 -*-
"""
Created on Sun Aug 13 06:58:58 2017

@author: dwlan
"""
import numpy as np
import re
import tkinter as tk
from tkinter import filedialog
from fractions import Fraction

#%%
def get_bom_filepath():
    ''' () -> str
    
    This function prompts the user to select an excel file that contains the 
    piping bom.  The filepath is returned as a string object.
    '''
    root=tk.Tk()
    root.withdraw()
    xlsx_filepath = filedialog.askopenfilename(
            filetypes = [("Excel Files", "*.xls;*.xlsx")], 
            title = "SELECT PIPING BOM FILE")
    return xlsx_filepath

#%%
def convert_feet_and_inches_text_to_numerical_feet(value):
    feet = 0.0
    inches = 0.0
    try:
        if "'" in value:
            #feet = float(value[:value.find("'")])
            feet = float(sum(Fraction(s) for s in re.sub('"','',value[:value.find("'")]).split()))
        if '"' in value:
            #inches = float(value[value.find("'")+1:value.find('"')])
            inches = float(sum(Fraction(s) for s in re.sub('"','',value[value.find("'")+1:value.find('"')]).split()))
        return feet + inches / 12
    except:
        return np.NaN
#%%
def parse_multiple_sizes(size_as_str):
    ''' str -> [float]
    
    Returns sizes as a list of floats.
    '''
    #size_floats=[np.nan, np.nan, np.nan]
    results=[]
    
    if type(size_as_str) == float or type(size_as_str) == int:
        results.append(size_as_str)
        return results
    
    size_as_str = size_as_str.replace('LG','')
    size_as_str = size_as_str.replace('-', ' ')
    sizes = size_as_str.lower().split('x')
    for size in sizes:
        if "'" in size:
            size_delegate = convert_feet_and_inches_text_to_numerical_feet(size) * 12
            if size_delegate not in results:
                results.append(size_delegate) # multiply by 12 to convert back to inches
        else:
            size_delegate = float(sum(Fraction(s) for s in re.sub('"','',size).split()))
            if size_delegate not in results:
                results.append(size_delegate)
    
    return results
#%%
def apply_classification(df, search_column, search_string, category_name, sub_category_name='', only_search_unclassified_items=True):
    ''' DataFrame, str, str, str, str, bool -> DataFrame 
    
    This function searches dataframe[search_column] for the search_string and updates the category column respectively.
    A new dataframe of the matched values is returned.
    '''
    import warnings
    warnings.filterwarnings("ignore", 'This pattern has match groups')
    
    is_match = df[search_column].str.contains(search_string, case=False)
    if sub_category_name:
        matches = df[:][(is_match) & (df['sub_category'] == '')]
        matches['category'] = category_name
        matches['sub_category'] = sub_category_name
        df.loc[matches.index, 'category'] = category_name
        df.loc[matches.index, 'sub_category'] = sub_category_name
    else:
        matches = df[:][(is_match) & (df['category'] == '')] # this creates a new DataFrame
        matches['category'] = category_name
        df.loc[matches.index, 'category'] = category_name
    #else:
    #    matches = df[:][is_match]
    #    matches[category_column].apply(lambda s: s.append(category_name))
        #dataframe.loc[matches.index, category_column].apply(lambda s: s.append(category_name))
    
    return matches
#%%
def apply_end_connection_classification(df, search_column, search_string, connection_name):
    '''
    
    A dataframe item might have more than one end connection.  Store the end connections in a list.
    '''
    import warnings
    warnings.filterwarnings("ignore", 'This pattern has match groups')
    
    # since there could be more than one connection type, initiate the 
    # 'connection_type' column to empty lists, which can be appended to later. 
    if 'connection_type_list' not in df.columns.tolist():
        df['connection_type_list']=np.empty((len(df), 0)).tolist()
    
    is_match = df[search_column].str.contains(search_string, case=False)
    matches = df[:][is_match]
    #matches['connection_type'].apply(lambda s: s.append(connection_name))
    df.loc[matches.index, 'connection_type_list'].apply(lambda s: s.append(connection_name))
    
#%%    
def apply_rating_classification(df, search_column, search_string, rating_name):
    if 'rating' not in df.columns.tolist(): df['rating']=''
    
    is_match = df[search_column].str.contains(search_string, case=False)
    matches = df[:][is_match]
    #matches['connection_type'].apply(lambda s: s.append(connection_name))
    df.loc[matches.index, 'rating'] = rating_name
#%%
def apply_metallurgy_classification(df, search_column, search_string, metallurgy_general, metallurgy_specific=''):
    
    if 'metallurgy_general' not in df.columns.tolist(): 
        df['metallurgy_general']=''
        df['metallurgy_specific']=''
    
    is_match = df[search_column].str.contains(search_string, case=False)
    if metallurgy_specific:
        matches = df[:][(is_match) & (df['metallurgy_specific'] == '')]
        #matches['metallurgy_general'] = metallurgy_general
        #matches['metallurgy_specific'] = metallurgy_specific
        df.loc[matches.index, 'metallurgy_general'] = metallurgy_general
        df.loc[matches.index, 'metallurgy_specific'] = metallurgy_specific
    else:
        matches = df[:][(is_match) & (df['metallurgy_general'] == '')] # this creates a new DataFrame
        #matches['metallurgy_general'] = metallurgy_general
        df.loc[matches.index, 'metallurgy_general'] = metallurgy_general
#%%
def apply_schedule_classification(df, search_column, search_string, schedule_name):
    if 'schedule_list' not in df.columns.tolist():
        df['schedule_list']=np.empty((len(df), 0)).tolist()
    
    is_match = df[search_column].str.contains(search_string, case=False)
    matches = df[:][is_match]
    #matches['connection_type'].apply(lambda s: s.append(connection_name))
    df.loc[matches.index, 'schedule_list'].apply(lambda s: s.append(schedule_name))
#%%   
def get_first_value_from_list(list_value):
    result = ''
    if len(list_value) > 0: result = list_value[0]
    
    return result
    
    
    
    
    
    
    
    
    
    
    
    