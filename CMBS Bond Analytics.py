# -*- coding: utf-8 -*-
"""
Created on Thu Dec 14 10:54:05 2023

@author: trpjs86
"""

import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt

# API initialization
from analytics_platform.trundl import Trundl
trundl = Trundl()

# USER DEFINED VARIABLES ARE UPDATED HERE ------------------------------------

# Date range for the entire data pull
start_date = '2020-11-01'
end_date = '2020-11-30'

# Function variables
scenario = 'zeroZero'
cusip = '05223HAA0'
func_start = '2020-11-01'
func_end = '2020-11-30'

# USER DEFINED VARIABLES END HERE --------------------------------------------

# Get data
data = trundl.get('MMDB_CMBS_BOND_ANALYTIC', startDate = start_date, 
                        endDate = end_date)

# Set a datetime index
data['effective_date'] = pd.to_datetime(data['effective_date']).apply(
    lambda x: x.strftime('%Y-%m-%d'))

data = data.set_index(['effective_date'])

# Drop n/a
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

# This function specifies which scenario to use
def choose_scenario(scenario):
    global frame 
    frame = scenarios[scenario]
    return frame
choose_scenario(scenario)

# This function specifies which cusip to use
def choose_cusip(cusip):
    global frame
    frame = frame[frame['cusip'].str.contains(cusip) == True]
    return frame
choose_cusip(cusip)

# This function specifies the date range
def date_range(start, end):
    global frame
    frame = frame[start:end]
    return frame
date_range(func_start, func_end)

# Plot the results
plt.plot(frame.index, frame['treasury_spread'])
plt.show()
