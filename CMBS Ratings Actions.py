# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 08:50:45 2023

@author: trpjs86
"""

import pandas as pd

bofa_file = pd.read_excel('Ratings Actions 3-24-23.xlsx', 
                          sheet_name = 'Week ending 3.24.2023')

holdings = pd.read_excel('BBG Alert Template.xlsx', sheet_name = 'Holdings')

# Set column headers and drop unneeded columns/rows
bofa_file.columns = bofa_file.iloc[0]
bofa_file = bofa_file.iloc[1:,0:-2]
del bofa_file['CMBX']

holdings.columns = holdings.iloc[3]
holdings = holdings.iloc[4:,1:-1]

# Drop affirmed 
bofa_file = bofa_file[bofa_file['Action'].str.contains('Affirmed') == False]

# Create a list of cusips that show up in both files
cusip_match = pd.merge(bofa_file, holdings, on = ['Cusip'])

# Drop unnecessary columns
del cusip_match['Deal_y']
del cusip_match['Class']
del cusip_match['Date']

# Create a list of deals that show up in both files
deal_match = pd.merge(bofa_file, holdings, on = ['Deal'])

# Drop unnecessary columns
del deal_match['Cusip_y']
del deal_match['Class']
del deal_match['Date']

# Only include upgrades and downgrades
deal_match = deal_match.loc[deal_match['Action'].str.contains(
    'Downgraded|Upgraded')]