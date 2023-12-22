# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 10:54:05 2023

@author: trpjs86
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib as mpl
import seaborn as sns

# Trundl API
from analytics_platform.trundl import Trundl
trundl = Trundl()

# Bloomberg API
from xbbg import blp

# USER DEFINED VARIABLES ARE UPDATED HERE ------------------------------------

# Date range for the entire data pull
start_date = '2023-11-01'
end_date = '2023-12-21'

# Line chart variables
deal = 'CGCMT 2017-P7 AS'
scenario = 'zeroZero'
func_start = start_date
func_end = end_date

# Regression Variables
x_deal = 'CGCMT 2017-P7 AS'
y_deal = 'CGCMT 2017-P7 A4'
reg_scenario = 'zeroZero'
reg_start = start_date
reg_end = end_date

# USER DEFINED VARIABLES END HERE --------------------------------------------

# Pull the cusip for the user-defined deals from the BBG API
line_cusip = blp.bdp(f'{deal} MTGE', flds = ['ID_CUSIP'])
x_cusip = blp.bdp(f'{x_deal} MTGE', flds = ['ID_CUSIP'])
y_cusip = blp.bdp(f'{y_deal} MTGE', flds = ['ID_CUSIP'])

# Pull the cusip out of the df and place in a string variable
line_cusip = line_cusip.iat[0, 0]
x_cusip = x_cusip.iat[0, 0]
y_cusip = y_cusip.iat[0, 0]

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

# # Write to excel
# with pd.ExcelWriter('CMBS Bond Analytics 2023.xlsx', engine = 'openpyxl',
#                     mode = 'a', if_sheet_exists = 'replace') as writer:
#     data.to_excel(writer, '2023', index = False)

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

# FUNCTIONS ------------------------------------------------------------------

# Function that creats a sliced dataframe based on user-defined inputs
def line_slicer(scenario, cusip, start, end):
    global frame
    frame = scenarios[scenario]
    frame = frame[frame['cusip'].str.contains(line_cusip) == True]
    frame = frame[start:end]
    return frame
line_slicer(scenario, line_cusip, func_start, func_end)

def reg_func(scenario, x_cusip, y_cusip, start, end):
    global reg_frame
    global x_frame
    global y_frame
    # Isolate only desired scenario
    reg_frame = scenarios[scenario]
    # Separate dataframe for the x variable
    x_frame = reg_frame[reg_frame['cusip'].str.contains(x_cusip) == True]
    # Separate dataframe for the y variable
    y_frame = reg_frame[reg_frame['cusip'].str.contains(y_cusip) == True]
    # Filter based on specified date range
    x_frame = x_frame[start:end]
    y_frame = y_frame[start:end]
    # Add identifying prefix to each column name
    x_frame = x_frame.add_prefix('x_')
    y_frame = y_frame.add_prefix('y_')
    # Merge x_frame and y_frame into one df
    reg_frame = x_frame.merge(y_frame, how = 'left', 
                                  left_index = True, right_index = True)
reg_func(reg_scenario, x_cusip, y_cusip, reg_start, reg_end)

# LINE PLOT ------------------------------------------------------------------

# Define the x and y axis
x = frame.index
y = frame['treasury_spread']

# Set plot style
mpl.style.available
mpl.style.use('fivethirtyeight')
# mpl.rcParams['grid.color'] = 'black'

plt.plot(x, y)
plt.ylabel('Treasury Spread')
plt.title(f'{deal} Treasury Spread')
plt.xticks(frame.index[::5], rotation = 45)
plt.show() 

# REGRESSION PLOT ------------------------------------------------------------

# Set plot style
mpl.style.available
mpl.style.use('seaborn')

sns.regplot(x = 'x_treasury_spread', y = 'y_treasury_spread', data = reg_frame,
            line_kws = {'color': 'black'})

# Highlight the most recent data point
most_recent = reg_frame.iloc[-1:]
plt.scatter(most_recent['x_treasury_spread'], most_recent['y_treasury_spread'], 
            color = 'orange')

plt.xlabel(f'{x_deal} Treasury Spread')
plt.ylabel(f'{y_deal} Treasury Spread')
