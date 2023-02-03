# -*- coding: utf-8 -*-
"""
Created on Tue Oct  4 13:09:03 2022

@author: trpjs86
"""

import pandas as pd

# User must update file name, start date, & end date
file = pd.read_csv('CMBS Holdings & Changes 20230131 - Values.csv')
start_date = '12-14-2022'
end_date = '02-01-2023'

# Replace column headers and drop corresponding row
file.columns = file.iloc[1]
file = file.iloc[2:]

# Drop unnecessary columns
file.drop(file.columns[[1,2,3]], axis = 1, inplace = True)
file.drop(file.iloc[:,3:30], axis = 1, inplace = True)
file.drop(file.iloc[:,6:], axis = 1, inplace = True)

# Rename columns
file.columns = ['Deal', 'Updated Analyst', 'TRP Tier', 'Analyst Comments',
                'RMS Note', 'Conviction']

# Create a datetime column
file['Datetime'] = pd.to_datetime(file['RMS Note'])

# Sort by RMS notes from the last month
file = file[(file['Datetime'] >= start_date) & (file['Datetime'] <= end_date)]

# Aggregate all the rake bonds
rake_check = file[file['Analyst Comments'].str.contains('rake', case = False)]

# Drop duplicates
file = file.drop_duplicates(['Deal'])

# Create list of deals w/ recent RMS notes containing rake bonds
common = pd.merge(file, rake_check, on = ['Deal'])
common = pd.Series(common['Deal'])
common = common.drop_duplicates()

# Drop datetime
del file['Datetime']
del rake_check['Datetime']

# Output to CSV
file.to_csv('New RMS Notes Output.csv')