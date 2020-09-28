#!/usr/bin/env python
# coding: utf-8

# In[341]:


input('Testing')

import pandas as pd
import difflib
import numpy as np
import datetime as dt
import os

print('Searching for your downloaded file...\n')

def preprocess(path=r'C:\Users\mcgarva1\Downloads\export.xlsx'):
    df=pd.read_excel(path)
    df=df[df['IssueYear']==2020]

    cols="""
    State	System	NOV Type	Issue Status	NOV SubType	Failure Type	Responsibility	Event Date	Awareness Date	Date Reported	Date Finalized	Description"""
    cols= cols.split('\t')
    cols[0] ='State'

    coldict={}

    for i in df.columns:
        try:
            coldict[i]=difflib.get_close_matches(i, cols)[0] 
        except IndexError:
            coldict[i]= 'No Match!'

    coldict['OperatingCenter']= 'System'
    coldict['IssueType']= 'NOV Type'
    coldict['EnforcementDate']= 'No match!'

    df=df.rename(columns=coldict)
    df=df[cols]
    df= df.sort_values(['NOV Type', 'Issue Status', 'NOV SubType', 'Responsibility']).reset_index()
    for i in df.dtypes[df.dtypes=='datetime64[ns]'].index:
        df[i]=df[i].dt.strftime('%m-%d-%y')
    return df

def format_colors(x, dw_novs= 0):
    if x['Issue Status']=='NOV Confirmed' and (x['NOV SubType']=='Health-based, not acute' or x['NOV SubType']=='Health-based, acute'): 
        return ['background-color: blue'] * len(df.columns)
    elif x['Issue Status']=='NOV Expected' or  x['Issue Status']=='NOV Pending Workgroup Review':
        return ['background-color: orange'] * len(df.columns)
    elif x['Issue Status']=='NOV Confirmed' and x['Responsibility']=='Third Party':
        return ['background-color: yellow'] * len(df.columns)
    elif x['Issue Status']=='NOV Confirmed' and x['NOV SubType'] not in ['Health-based, acute', 'Health-based, not acute', 'Monitoring', 'Reporting']:
        return ['background-color: yellow'] * len(df.columns)
    elif x['Issue Status'] in ['NOV Not Expected', 'Deemed not an NOV', 'NOV Rescinded']:
        return ['background-color: yellow'] * len(df.columns)
    else:
        return ['background-color: white'] * len(df.columns)
    
    
def sort_by_colors():
    
    df=preprocess()
    df['Sort']=1000

    for y in range(len(df)):
        if df.iloc[y]['Issue Status']=='NOV Confirmed' and (df.iloc[y]['NOV SubType']=='Health-based, not acute' or df.iloc[y]['NOV SubType']=='Health-based, acute'):
            df.iloc[y, -1]=1
        elif df.iloc[y]['Issue Status']=='NOV Expected' or  df.iloc[y]['Issue Status']=='NOV Pending Workgroup Review':
             df.iloc[y, -1]=2
        elif df.iloc[y]['Issue Status']=='NOV Confirmed' and df.iloc[y]['Responsibility']=='Third Party':
            df.iloc[y, -1]=3
        elif df.iloc[y]['Issue Status']=='NOV Confirmed' and df.iloc[y]['NOV SubType'] not in ['Health-based, acute', 'Health-based, not acute', 'Monitoring', 'Reporting']:
            df.iloc[y, -1]=4
        elif df.iloc[y]['Issue Status'] in ['NOV Not Expected', 'Deemed not an NOV', 'NOV Rescinded']:
            df.iloc[y, -1]=5
        else:
            df.iloc[y, -1]=6

    df= df.sort_values('Sort').drop(columns=['index'])
    return df.style.apply(highlight_max, axis=1)


func= sort_by_colors()
print('Located your file. Writing output...\n')
func.to_excel(os.getcwd() +'\\NOV Register {}.xlsx'.format(dt.datetime.now().strftime('%m-%d-%Y')), index=False)
print('Your file has been written to the following path', os.getcwd())



# In[302]:


sort_by_colors()


# In[337]:


cols=['Goal', 'Confirmed YTD', 'Potential', 'Expected', 'Not Expected/Not Expected To Count']
index=['Total Drinking Water NOVs','Health-Based NOVs', 'Monitoring / Reporting']

summary = pd.DataFrame(columns=cols, index=index)
summary.index.rename(dt.datetime.now().strftime('%B %d, %Y'), inplace=True)

summary.iloc[0, 0]='<=6'
summary.iloc[1, 0]='<=2'
summary.iloc[2, 0]='N/A'


# In[ ]:




