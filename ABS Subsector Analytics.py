# -*- coding: utf-8 -*-
"""
Created on Thu Apr 13 11:08:43 2023

@author: trpjs86
"""

import pandas as pd
import math

data = pd.read_excel('ABS Subsector Analytics.xlsx')

start_date = '01/31/2012'
end_date = '01/31/2023'

# Drop first column and create datetime column
data = data.iloc[:,1:]
data['effective_date'] = pd.to_datetime(data['effective_date'])

# Set date range
data = data[(data['effective_date'] >= start_date) & 
            (data['effective_date'] <= end_date)]

# Groupby subindex
ex_ret = data.groupby('subindex')['excess_return_1m'].mean() * 12 * 100
stdev = data.groupby('subindex')['excess_return_1m'].std() * math.sqrt(12) * 100
ir = ex_ret / stdev
skew = data.groupby('subindex')['excess_return_1m'].skew()
kurt = data.groupby('subindex')['excess_return_1m'].apply(pd.DataFrame.kurtosis)

# Aggregate all stats into one dataframe
listt = [ex_ret, stdev, ir, skew, kurt]
stats = ex_ret.to_frame()
del stats['excess_return_1m']

for i in listt:
    i = i.to_frame()
    i.reset_index(inplace = True)
    stats = pd.merge(stats, i, on = ['subindex'])
    
# Rename columns
stats.columns = ['subindex', 'excess_return', 'excess_return_vol', 
                 'excess_return_ir', 'excess_return_skew', 
                 'excess_return_kurtosis']

# Get unique subindexes
unique = data.drop_duplicates('subindex')
unique = unique.reset_index()
del unique['index']

# Get start dates for each subindex
stats = pd.merge(stats, unique[['subindex', 'effective_date']],
                 on = 'subindex', how = 'left')

# Move the start date column to index 1
stats.insert(1, 'effective_date', stats.pop('effective_date'))
stats = stats.rename(columns = {'effective_date': 'start_date'})

# Remove datetime format from start_date
stats['start_date'] = stats['start_date'].dt.strftime('%m-%d-%Y')

# Add end_date column
stats.insert(2, 'end_date', end_date, True)
stats['end_date'] = pd.to_datetime(stats['end_date'])
stats['end_date'] = stats['end_date'].dt.strftime('%m-%d-%Y')

# Drop unwanted subindexes
drop_list = ['Automobile', 'Credit', 'Home', 'Manufactured', 'Miscellaneous']
mask = stats['subindex'].str.contains('|'.join(drop_list))
stats = stats[~mask]

