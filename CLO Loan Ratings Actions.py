# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 11:23:34 2023

@author: trpjs86
"""

import pandas as pd

data = pd.read_excel('Loan Ratings Actions.xlsx')

# Preserve the original file
orig = data

# Drop unnecessary columns
keep = ['issuer_name', 'seniority', 'sp', 'S&P Flag', 'Prev sp', 'moodys', 
        "Moody's Flag", 'Prev moodys', 'Fac Size']

data = data.loc[:, keep]

# Filter data by senority
data = data[data['seniority'].str.contains('2ND') == False]
del data['seniority']

# Filter data by rating change
data = data[(data['sp'] == 'B-') & (data['Prev sp'] != 'B-') | 
(data['sp'].str.contains('C') == True) & (data['Prev sp'].str.contains('C') == False) | 
(data['sp'].str.contains('D') == True) & (data['Prev sp'].str.contains('D') == False) & 
(data['Prev sp'].str.contains('C') == False) | (data['moodys'] == 'B3') & 
(data['Prev moodys'] != 'B3') | (data['moodys'].str.contains('C') == True) & 
(data['Prev moodys'].str.contains('C') == False) | (data['moodys'].str.contains('D') == True) &
(data['Prev moodys'].str.contains('D') == False) & (data['Prev moodys'].str.contains('C') == False)]

# Rename columns
data = data.rename(columns = {'sp': 'S&P Curr'})
data = data.rename(columns = {'moodys': "Moody's Curr"})
data = data.rename(columns = {'issuer_name': 'Issuer'})
data = data.rename(columns = {'Prev sp': 'Prev S&P'})
data = data.rename(columns = {'Prev moodys': "Prev Moody's"})

# Consilidate by facility size
group = data.groupby('Issuer', as_index = False)['Fac Size'].sum()
group = group.merge(data, on = 'Issuer', how = 'left')
del group['Fac Size_y']
group.insert(7, 'Fac Size_x', group.pop('Fac Size_x'))
group = group.rename(columns = {'Fac Size_x': 'Fac Size'})
group = group.drop_duplicates()
data = group

# Sort by S&P Flag
data = data.sort_values('S&P Flag', ascending = True)

# Create separate dataframes for upgrades and downgrades
downgrades = data[(data['S&P Flag'] == 'Downgrade') | (data["Moody's Flag"] 
                                                       == 'Downgrade')]

upgrades = data[(data['S&P Flag'] == 'Upgrade') | (data["Moody's Flag"] 
                                                       == 'Upgrade')]

# Isolate downgrades from S&P and Moody's
sp_down = downgrades[downgrades['S&P Flag'] == 'Downgrade']
moodys_down = downgrades[downgrades["Moody's Flag"] == 'Downgrade']

# Write to excel
with pd.ExcelWriter('Table.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    downgrades.to_excel(writer, 'Sheet1', index = False)