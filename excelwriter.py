#!/usr/bin/env python
# coding: utf-8

# # MapCall NOV Register Excel Writer
# ## By Aaron McGarvey
# 
# 
# ### The following script takes as input the raw downloaded file from MapCall and outputs two files: a color-formatted NOV register that can be used as a basis for reporting and a summary table. 
# 
# Instructions:
# 
# 1. Download the MapCall export as per usual procedure. Save the file in the default location -'Downloads' folder
# 2. At the top of the screen, click 'Cell'---> 'Run All'
# 3. The script will prompt you for the name of the file, i.e. export (100). Enter the name of the file and press enter. 
# 4. Once the operation is complete, look in the specified folder for your new files. 
# 
# 

# In[151]:


import pandas as pd
import difflib
import numpy as np
import datetime as dt
import os
from IPython.display import HTML


filename= input('Enter the name of your downloaded file then press enter, i.e. export (100)')
print('\nSearching for your downloaded file...\n')

def preprocess(path=os.path.expanduser("~") + '\\Downloads\\' + filename + '.xlsx'):
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
        
        
    df['Op Type']= 0

    for i in range(len(df)):
        if df.iloc[i, 2][:3].isalpha():
            df.iloc[i, -1]='MBB'
        else:
            df.iloc[i, -1]='AW'
    return df

def format_colors(x, dw_novs= 0):
    df=preprocess()
    if x['NOV Type']=='Wastewater' or x['NOV Type']=='Environmental': 
        return ['background-color: white'] * len(df.columns) 
    elif x['Issue Status']=='NOV Confirmed' and (x['NOV SubType']=='Health-based, not acute' or x['NOV SubType']=='Health-based, acute' or x['NOV SubType']=='Monitoring' or x['NOV SubType']=='Reporting'): 
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
        if df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status']=='NOV Confirmed' and (df.iloc[y]['NOV SubType']=='Health-based, not acute' or df.iloc[y]['NOV SubType']=='Health-based, acute'):
            df.iloc[y, -1]=1
        elif df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status']=='NOV Confirmed' and (df.iloc[y]['NOV SubType']=='Monitoring' or df.iloc[y]['NOV SubType']=='Reporting'):
            df.iloc[y, -1]=1
        elif df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status']=='NOV Expected' or  df.iloc[y]['Issue Status']=='NOV Pending Workgroup Review':
            df.iloc[y, -1]=2
        elif df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status']=='NOV Confirmed' and df.iloc[y]['Responsibility']=='Third Party':
            df.iloc[y, -1]=3
        elif df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status']=='NOV Confirmed' and df.iloc[y]['NOV SubType'] not in ['Health-based, acute', 'Health-based, not acute', 'Monitoring', 'Reporting']:
            df.iloc[y, -1]=4
        elif df.iloc[y]['NOV Type']=='Drinking Water' and df.iloc[y]['Issue Status'] in ['NOV Not Expected', 'Deemed not an NOV', 'NOV Rescinded']:
            df.iloc[y, -1]=5
        elif df.iloc[y]['NOV Type']=='Wastewater' or df.iloc[y]['NOV Type']=='Environmental': 
            df.iloc[y, -1]=6
        else:
            df.iloc[y, -1]=6

            
    df= df.sort_values('Sort').drop(columns=['index'])
    
    return df, df.style.apply(format_colors, axis=1)



def make_summary_table():
    cols=['Goal', 'Confirmed YTD', 'Potential/Expected']
    index=['Total Drinking Water NOVs','Health-Based NOVs', 'Monitoring / Reporting']

    summary = pd.DataFrame(columns=cols, index=index)
    summary.index.rename(dt.datetime.now().strftime('%B %d, %Y'), inplace=True)
    summary.iloc[0, 0]='<=6'
    summary.iloc[1, 0]='<=2'
    summary.iloc[2, 0]='N/A'

    dfnew= sort_by_colors()[0]

    total_dw_novs_confirmed = len(dfnew[(dfnew['NOV Type']=='Drinking Water') & (dfnew['Issue Status']=='NOV Confirmed')])
    total_dw_novs_expected = len(dfnew[(dfnew['NOV Type']=='Drinking Water') & (dfnew['Issue Status']=='NOV Expected') | (dfnew['Issue Status']=='NOV Pending Workgroup Review')])
    health_confirmed= len(dfnew[(dfnew['NOV Type']=='Drinking Water') & (dfnew['Issue Status']=='NOV Confirmed') & (dfnew['NOV SubType']=='Health-based, not acute') | (dfnew['NOV SubType']=='Health-based, acute')])
    health_expected = len(dfnew[(dfnew['NOV Type']=='Drinking Water') & ((dfnew['NOV SubType']=='Health-based, not acute') | (dfnew['NOV SubType']=='Health-based, acute')) & ((dfnew['Issue Status']=='NOV Expected') | (dfnew['Issue Status']=='NOV Pending Workgroup Review'))])
    monitoring_confirmed= len(dfnew[(dfnew['NOV Type']=='Drinking Water') & (dfnew['Issue Status']=='NOV Confirmed') & ((dfnew['NOV SubType']=='Monitoring') | (dfnew['NOV SubType']=='Reporting'))])
    monitoring_expected = len(dfnew[(dfnew['NOV Type']=='Drinking Water') & (dfnew['Issue Status']=='NOV Expected') & ((dfnew['NOV SubType']=='Monitoring') | (dfnew['NOV SubType']=='Reporting'))])

    summary.iloc[0, 1] = total_dw_novs_confirmed
    summary.iloc[0, 2]=  total_dw_novs_expected
    summary.iloc[1, 1]=  health_confirmed
    summary.iloc[1, 2]=  health_expected
    summary.iloc[2, 1]=  monitoring_confirmed
    summary.iloc[2, 2]=  monitoring_expected
    return summary



func= sort_by_colors()[1]
print('Located your file. Writing output...\n')
func.to_excel(os.getcwd() +'\\NOV Register {}.xlsx'.format(dt.datetime.now().strftime('%m-%d-%Y')), index=False)
make_summary_table().to_excel(os.getcwd() + '\\Summary Table {}.xlsx'.format(dt.datetime.now().strftime('%m-%d-%Y')))


print('Your files have been written to the following path:.......', os.getcwd(), '\n')
print('NOV Register {}.xlsx'.format(dt.datetime.now().strftime('%m-%d-%Y')))
print('Summary Table {}.xlsx'.format(dt.datetime.now().strftime('%m-%d-%Y')))



# In[ ]:




