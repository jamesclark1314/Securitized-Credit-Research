# -*- coding: utf-8 -*-
"""
Created on Tue Mar 21 08:50:45 2023

@author: trpjs86
"""

import pandas as pd

bofa_file = pd.read_excel('Ratings Actions 3-17-23.xlsx', 
                          sheet_name = 'Week ending 3.17.2023')

holdings = pd.read_excel('BBG Alert Template.xlsx', sheet_name = 'Holdings')

# Set column headers and drop unneeded columns/rows
bofa_file.columns = bofa_file.iloc[0]
bofa_file = bofa_file.iloc[1:,0:-2]
del bofa_file['CMBX']

holdings.columns = holdings.iloc[3]
holdings = holdings.iloc[4:,1:-1]

# Drop affirmed 
bofa_file = bofa_file[bofa_file['Action'].str.contains('Affirmed') == False]

# Create a list of deals that show up in both files
common = pd.merge(bofa_file, holdings, on = ['Cusip'])

# Drop unnecessary columns
del common['Deal_y']
del common['Class']