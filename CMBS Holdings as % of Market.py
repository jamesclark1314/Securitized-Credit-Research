# -*- coding: utf-8 -*-
"""
Created on Tue Aug 22 16:34:02 2023

@author: trpjs86
"""
import pandas as pd
import numpy as np

# Suppress scientific notation
pd.options.display.float_format = '{:.0f}'.format

# File imports
holdings = pd.read_excel('CMBS Holdings 20230731 - Values.xlsx', 
                              sheet_name = 'Data')
sasb_market = pd.read_excel('CMBS Market Snapshot.xlsx')

# You need to run the bdp function and change tranche names in excel first - 
# Run the bdp function in current balance column, run macro, and copy/paste as
# values to a new workbook
sasb_universe = pd.read_excel('Public SASB Universe.xlsx')

# Set column headers and drop unnecessary rows/columns
holdings.columns = holdings.iloc[0]
holdings = holdings.iloc[1:,:]

sasb_market.columns = sasb_market.iloc[2]
sasb_market = sasb_market.iloc[3:,1:]

sasb_universe.columns = sasb_universe.iloc[4]
sasb_universe = sasb_universe.iloc[5:,1:]

# Drop unnecessary columns in sasb universe
keep = ['Deal Name', 'Class', 'CUSIP', 'Balance \nCurrent', 'Coupon\nType']

sasb_universe = sasb_universe.loc[:, keep]

# Rename columns
sasb_universe.columns = ['Deal', 'Class', 'CUSIP', 'Par', 'Coupon']

# Combine coupon types into just fixed and floating
sasb_universe['Coupon'] = np.where(sasb_universe['Coupon'] == 
                                    'WAC/Pass Thru', 'FIXED', sasb_universe['Coupon'])
sasb_universe['Coupon'] = np.where(sasb_universe['Coupon'].str.contains('Fixed')
                                    == True, 'FIXED', sasb_universe['Coupon'])
sasb_universe['Coupon'] = np.where(sasb_universe['Coupon'].str.contains('Variable') 
                                    == True, 'FIXED', sasb_universe['Coupon'])
sasb_universe['Coupon'] = np.where(sasb_universe['Coupon'].str.contains('Floater') 
                                    == True, 'FLOATING', sasb_universe['Coupon'])

# Replace bdp errors with 0
sasb_universe['Par'] = np.where(sasb_universe['Par'].str.contains('#') == True,
                                0, sasb_universe['Par'])

# Drop na from sasb universe and remove total row at bottom
sasb_universe = sasb_universe.dropna()
sasb_universe = sasb_universe[:-1]

# Drop unnecessary columns in market df
keep = ['Deal Name', 'Deal\nType', 'Collateral Balances\nCurrent', 'Property Type %\nMF', 
        'Property Type %\nRT', 'Property Type %\nOF', 'Property Type %\nLO',
        'Property Type %\nIN', 'Property Type %\nMH', 'Property Type %\nSS',
        'Property Type %\nMU', 'Property Type %\nOT']

sasb_market = sasb_market.loc[:, keep]

# Change column names in market df
sasb_market.columns = ['Deal', 'Partition', 'Balance', 'MF%', 'RT%', 'OF%', 'LO%', 
                    'IN%', 'MH%', 'SS%', 'MU%', 'OT%']

# Maintain full cmbs market
cmbs_market = sasb_market

# Include only conduit and sasb bonds in the cmbs market
cmbs_market = cmbs_market[cmbs_market['Partition'].str.contains(
    'SnglAsset|Conduit') == True]

# Isolate SASB bonds in sasb_market
sasb_market = sasb_market[sasb_market['Partition'].str.contains('SnglAsset') == True]

# Drop 0 balances from sasb_market
sasb_market = sasb_market[sasb_market['Balance'] != 0]

# CONVERT TO NUMERIC AND CONVERT PROP TYPE TO % FOR HOLDINGS & SASB_MARKET

# Get a list of the necessary column headers
num_headers = sasb_market.iloc[:,2:12].columns.values.tolist()

# Iterate through the list and apply to necessary columns for sasb_market
for i in num_headers:
    sasb_market[[i]] = sasb_market[[i]].apply(pd.to_numeric)
    sasb_market[[i]] = sasb_market[[i]] / 100

# I divided balance column by 100 in the loop so fixing that here
sasb_market['Balance'] = sasb_market['Balance'] * 100

# Get a list of the necessary column headers
num_headers = holdings.iloc[:,3:16].columns.values.tolist()

# Iterate through the list and apply to necessary columns for holdings
for i in num_headers:
    holdings[[i]] = holdings[[i]].apply(pd.to_numeric)
    holdings[[i]] = holdings[[i]] / 100
    
# I divided par value column by 100 in the loop so fixing that here
holdings['Par'] = holdings['Par'] * 100

# Rename balance column in sasb_market
sasb_market = sasb_market.rename(columns = {'Balance': 'Par'})

# Combine floating and variable coupon types and Dupers/LCF/A1/A2
holdings['Cpn Type'] = np.where(holdings['Cpn Type'] == 'VARIABLE', 'FIXED',
                                holdings['Cpn Type'])

holdings['Partition'] = np.where(holdings['Partition'].str.contains(
    'Duper|LCF|A1/A2') == True, 'Duper', holdings['Partition'])

# Isolate SASB vs. conduit vs. other bonds
sasb = holdings[holdings['Partition'].str.contains('SA/SB') == True]

conduit = holdings[holdings['Partition'].str.contains(
    'Triple-B|Single-A|LCF|Duper|Double-A|Junior|A1/A2|Conduit') == True]

other = holdings[holdings['Partition'].str.contains('Freddie|Interest|Net')]

# Sum par values for each bond type
sasb_par = sasb['Par'].sum()
conduit_par = conduit['Par'].sum()
other_par = other['Par'].sum()
total_par = holdings['Par'].sum()

# Bond type % weights
conduit_pct = conduit_par / total_par
sasb_pct = sasb_par / total_par
other_pct = other_par / total_par

# Add totals to new dataframe
bond_types = {'Face':[conduit_par, sasb_par, other_par, total_par],
         '% of Total': [conduit_pct, sasb_pct, other_pct, 
                        (conduit_pct + sasb_pct + other_pct)]}

bond_types = pd.DataFrame(bond_types, index = ['Conduit', 'SASB', 'Other', 'Total'])

# STRAT OUT BOND TYPES

# Create groups and sum par value for conduit
conduit_strats = conduit.groupby(['Partition']).Par.sum()

# Calculate % weight for each group
conduit_strats = pd.DataFrame(conduit_strats)
conduit_strats['% of Total'] = conduit_strats['Par'] / holdings['Par'].sum()

# Sort conduit strats in decending order
conduit_strats = conduit_strats.sort_values('% of Total', ascending = False)

# Create groups and sum par value
sasb_cpn = sasb.groupby(['Cpn Type']).Par.sum()

sasb_univ_cpn = sasb_universe.groupby(['Coupon']).Par.sum()
sasb_univ_cpn = sasb_univ_cpn.apply(pd.to_numeric)

# Calculate % weight for each group
sasb_cpn = pd.DataFrame(sasb_cpn)
sasb_cpn['% of Total'] = sasb_cpn['Par'] / holdings['Par'].sum()

sasb_univ_cpn = pd.DataFrame(sasb_univ_cpn)
sasb_univ_cpn['% of Market'] = sasb_univ_cpn['Par'] / cmbs_market['Balance'].sum()

# Add % of market column to sasb_cpn
sasb_cpn['% of Market'] = sasb_cpn['Par'] / sasb_univ_cpn['Par']

# STRAT SASB BY PROPERTY TYPE

# Create a dictionary to iterate through 
df_list = {'sasb_prop_types': sasb, 'sasb_market_prop_types': sasb_market}
result_dict = {}

for df_name, df in df_list.items():
    # Get a list of the property type column headers
    prop_headers = df.loc[:,'MF%':'OT%'].columns.values.tolist()

    # Empty list to append results to
    loop_list = []

    # Iterate through the list and calculate property type concentrations
    for i in prop_headers:
        curr_prop_type = sum(df['Par'] * df[i])
        loop_list.append(curr_prop_type)
    
    result_dict[df_name] = loop_list

# Create dataframes for property types
sasb_prop_types = {'Par': result_dict['sasb_prop_types']}
sasb_prop_types = pd.DataFrame(sasb_prop_types, index = 
                                sasb.loc[:,'MF%':'OT%'].columns.values.tolist())

sasb_market_prop_types = {'Par': result_dict['sasb_market_prop_types']}
sasb_market_prop_types = pd.DataFrame(sasb_market_prop_types, index = 
                                      sasb_market.loc[:,'MF%':'OT%'].columns.values.tolist())

# Add % of total columns
sasb_prop_types['% of Total'] = sasb_prop_types['Par'] / holdings['Par'].sum()
sasb_market_prop_types['% of Market'] = sasb_market_prop_types[
    'Par'] / cmbs_market['Balance'].sum()

# Drop unnecessary property types
sasb_prop_types = sasb_prop_types.drop('CH%')
sasb_prop_types = sasb_prop_types.drop('WH%')
sasb_prop_types = sasb_prop_types.drop('HC%')

# Sort by highest percentage
sasb_prop_types = sasb_prop_types.sort_values('% of Total', ascending = False)
sasb_market_prop_types = sasb_market_prop_types.sort_values('% of Market', ascending = False)

# Our sasb holdings as % of outstanding market
sasb_prop_types['% of Market'] = sasb_prop_types['Par'] / sasb_market_prop_types['Par']

# Add property type dfs to list to loop through
prop_type_dfs = [sasb_prop_types, sasb_market_prop_types]

for i in prop_type_dfs:
    # Rename the indexes
    i = i.rename(index = {'OF%': 'Office', 'LO%': 'Lodging', 'IN%': 'Industrial', 
                          'RT%': 'Retail', 'MF%': 'Multifamily',
                          'MH%': 'Manufactured Housing','MU%': 'Mixed Use',
                          'SS%': 'Self Storage','OT%': 'Other'}, inplace = True)
    
for i in range(len(prop_type_dfs)):  
    # Set a static index order for the property type dfs
    prop_type_dfs[i] = prop_type_dfs[i].reindex(['Office', 'Lodging', 'Industrial',
                                                 'Retail', 'Multifamily',
                                                 'Manufactured Housing', 'Mixed Use', 
                                                 'Self Storage', 'Other'],
                                                copy = False)

# Assign the values from the prop_type_dfs list to the originial variables
sasb_prop_types = prop_type_dfs[0]
sasb_market_prop_types = prop_type_dfs[1]

# Add all final dataframes to a dictionary
all_dfs = {'conduit_strats': conduit_strats, 'sasb_cpn': sasb_cpn,
           'sasb_prop_types': sasb_prop_types,
           'sasb_market_prop_types': sasb_market_prop_types,
           'sasb_univ_cpn': sasb_univ_cpn, 'bond_types': bond_types}

# Loop through the dictionary of all dataframes
for df_name, df in all_dfs.items():
    # Make the index the first column in each dataframe and move to column 0
    df['Index'] = df.index
    df.insert(0, 'Index', df.pop('Index'))
    
    # Write to excel
    with pd.ExcelWriter('CMBS Holdings as % of Outstanding.xlsx', engine = 'openpyxl',
                        mode = 'a', if_sheet_exists = 'replace') as writer:
        df.to_excel(writer, df_name, index = False)
