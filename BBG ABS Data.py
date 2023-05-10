# -*- coding: utf-8 -*-
"""
Created on Tue May  9 16:36:11 2023

@author: trpjs86
"""

import pandas as pd
import numpy as np

data = pd.read_excel('Pivot.xlsx')

effective_date = '03/31/2023'

# Pivot the dataframe
data = data.transpose()

# Create a numeerical index and move the asset classes to the first column
data.reset_index(drop = False, inplace = True)

# Set column headers and drop first row
data.columns = data.iloc[0]
data = data.iloc[1:]

# Rename columns
data = data.rename(columns = {'Unnamed: 0': 'subindex'})
data = data.rename(columns = {'Excess Rtn': 'excess_return_1m'})

# Move the excess return column
data.insert(1, 'excess_return_1m', data.pop('excess_return_1m'))

# Add prefix to subindex column values
data['subindex'] = 'BB_' + data['subindex'].astype(str)

# Replace spaces & slashes & dashes in subindex column with underscore
data = data.replace(' ', '_', regex = True)
data['subindex'] = data['subindex'].str[:-1]
data = data.replace('/', '_', regex = True)
data = data.replace('-', '_', regex = True)

# Rename the LUABTRUU index
data['subindex'] = np.where(data['subindex'].str.contains('LUABTRUU'), 
                            'BB_US_ABS_Index', data['subindex'])

# Sort the dataframe alphabetically
data = data.sort_values('subindex')

# Add date column
data['effective_date'] = effective_date
data.insert(0, 'effective_date', data.pop('effective_date'))

# Write to excel
with pd.ExcelWriter('Pivot.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    data.to_excel(writer, 'output', index = False)