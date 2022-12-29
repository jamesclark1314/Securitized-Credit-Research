# -*- coding: utf-8 -*-
"""
Created on Tue Dec 20 15:31:34 2022

@author: trpjs86
"""

import pandas as pd

file = pd.read_table('test.txt', header = None)
file.columns = ['Column']

# Drop unnecessary rows
file = file[file['Column'].str.contains(
    'From:|To:|Subject:|Date:|Source|Alert Summary:|to view / edit alert')
    == False]
file = file[file['Column'].str.contains('that triggered|this alert|Mtge|FIFW') 
            == False]

# Remove (DSCR <= 1.1)
file['Column'] = file['Column'].replace({'(DSCR <= 1.1)':''}, regex = True)

# Isolate deal name
deals = file[file['Column'].str.contains('########')]
deals = deals['Column'].str[9:-23].to_frame()

# Isolate new update
update = file[file['Column'].str.contains('====')]
update = update['Column'].str[5:].to_frame()

# Isolate DSCR 
dscr = file[(file['Column'].str.contains('DSCR'))|(file['Column']
                                                   .str.len() <= 15) & 
            (~ file['Column'].str.contains('alert'))]
        
##################
thing = file['Column'].str.len() <= 15 & ~ file['Column'].str.contains('alert')
count = file.loc[thing]
##################

# Isolate deal %
deal_perc = file[file['Column'].str.contains('Deal%')]
deal_perc = deal_perc['Column'].str[0:-30].to_frame()

# Isolate loan name
loans = file[file['Column'].str.contains('NXTW')]
loans = loans['Column'].str[:-35].to_frame()

# Clean up cell values
mask = file['Column'].str.contains('########')
file.loc[mask, 'Column'] = deals['Column']
mask = file['Column'].str.contains('NXTW')
file.loc[mask, 'Column'] = loans['Column']
mask = file['Column'].str.contains('Deal%')
file.loc[mask, 'Column'] = deal_perc['Column']
mask = file['Column'].str.contains('====')
file.loc[mask, 'Column'] = update['Column']

# Drop any rogue alert cells
file = file[file['Column'].str.contains('alert') == False]

# Remove NXTW garbage
file['Column'] = file['Column'].replace({'NXTW LDES':''}, regex = True)

# Write to excel
with pd.ExcelWriter('BBG Alert Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    file.to_excel(writer, 'Alerts', index = False)

    