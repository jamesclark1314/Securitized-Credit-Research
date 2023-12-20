# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 10:54:05 2023

@author: trpjs86
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Trundl API
from analytics_platform.trundl import Trundl
trundl = Trundl()

# Bloomberg API
from xbbg import blp

# USER DEFINED VARIABLES ARE UPDATED HERE ------------------------------------

# Date range for the entire data pull
start_date = '2022-04-11'
end_date = '2023-10-01'

# Function variables
deal = 'AFHT 2019-FAIR A'
scenario = 'zeroZero'
func_start = start_date
func_end = end_date

# USER DEFINED VARIABLES END HERE --------------------------------------------

# Pull the cusip for the user-defined deal from the BBG API
cusip = blp.bdp(f'{deal} MTGE', flds = ['ID_CUSIP'])

# Pull the cusip out of the df and place in a string variable
cusip = cusip.iat[0, 0]

# DATA PULL ------------------------------------------------------------------

# Create a list containing each date between start_date and end_date
datetime_list = pd.date_range(start = start_date, end = end_date).to_list()

# Convert each item in date_list to string
date_list = []
for i in datetime_list:
    date_list.append(str(i))

# Remove the timestamps from the end of each string in date_list
test_list = []    
for i in date_list:
    test_list.append(i[:-9])
    
date_list = test_list    

# Pull all the data for given date range
all_data_list = []
for i in date_list:
    # Get data
    data = trundl.get('MMDB_CMBS_BOND_ANALYTIC', startDate = i, 
                            endDate = i)
    all_data_list.append(data)
    
# Combine all dfs from all_data into one df
data = pd.DataFrame()
for i in all_data_list:
    data = data.append(i).reset_index(drop = True)

# END DATA PULL---------------------------------------------------------------

# Set a datetime index
data['effective_date'] = pd.to_datetime(data['effective_date']).apply(
    lambda x: x.strftime('%Y-%m-%d'))

data = data.set_index(['effective_date'])

# Remove N/As and replace any instance of string 'NA' with 0
data = data.replace('NA', np.nan)
data = data.dropna()

# CHANGE DATA TYPE -----------------------------------------------------------

# Get a list of all column headers to change
num_headers = ['wal_years', 'treasury_spread', 'min_ce', 'modified_duration', 
                'swap_spread', 'yield']

# Iterate through the list and apply to necessary columns
for i in num_headers:
    data[[i]] = data[[i]].apply(pd.to_numeric)
    
# ----------------------------------------------------------------------------   

# Scenario grouping
scenarios_list = ['base', 'up', 'down', 'zeroZero', 'cpy100']
scenarios = {}

for i in scenarios_list:
    # Iterate thru list and make each list item the dict key
    scenarios[i] = data.loc[data['scenario'] == i]
    # Sort dataframes by deal first then cusip
    scenarios[i] = scenarios[i].sort_values(['bloomberg_dealname', 'cusip'], 
                                            ascending = [True, True])

# Function that creats a sliced dataframe based on user-defined inputs
def df_slicer(scenario, cusip, start, end):
    global frame
    frame = scenarios[scenario]
    frame = frame[frame['cusip'].str.contains(cusip) == True]
    frame = frame[start:end]
    return frame
df_slicer(scenario, cusip, func_start, func_end)

# Plot the results
plt.plot(frame.index, frame['treasury_spread'])
plt.ylabel('Treasury Spread')
plt.title(f'{deal} Treasury Spread')
plt.show()

