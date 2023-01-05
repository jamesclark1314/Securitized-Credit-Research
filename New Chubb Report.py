# -*- coding: utf-8 -*-
"""
Created on Tue Dec 13 13:14:19 2022

@author: trpjs86
"""

import pandas as pd

# Imports
universe = pd.read_excel('clo-deals-universe.xlsx')
universe = universe.sort_values(['Deal'], ascending = True)
scenarios = pd.read_excel('Scenarios.xlsx')
portfolio = pd.read_excel('Portfolio.xlsx')

# Isolate sector weights
sector = universe.loc[:,'Industry: Automotive':'Industry: Utility']

# Combine metals and mining and energy columns / broadcasting and media columns
sector['Industry: Energy/Metals & Mining'] = sector[
    'Industry: Energy'] + sector['Industry: Metals And Mining']
sector['Industry: Broadcasting & Media'] = sector[
    'Industry: Broadcasting'] + sector['Industry: Diversified Media']

# Drop columns
sector = sector.drop(columns = ['Industry: Energy',
                                'Industry: Metals And Mining',
                                'Industry: Broadcasting',
                                'Industry: Diversified Media'])

# Rename columns and add deal column
sector.columns = sector.columns.str[10:]
sector = universe['Deal'].to_frame().merge(sector, how = 'left', 
                                  left_index = True, right_index = True)

# Reorder columns in sector df
sector = sector[['Deal', 'Automotive', 'Broadcasting & Media',
                 'Cable And Satellite', 'Chemicals', 'Consumer Products',
                 'Energy/Metals & Mining', 'Financial', 'Food And Beverages',
                 'Gaming Lodging And Leisure', 'Healthcare', 'Housing',
                 'Industrials', 'Paper And Packaging', 'Retail', 'Services',
                 'Technology', 'Telecommunications', 'Transportation',
                 'Utility']]

# Calculate medians sector weights
sector_median = sector.iloc[:,1:].median()
sector_median = sector_median.to_frame()
sector_median.columns = ['Sector Median']

# Make index first column - sector_median
sector_median = sector_median.reset_index(level = 0)
sector_median.columns = ['Sector', 'Sector Median']

# Isolate indebtedness and calculate median indebtedness
indebtedness = universe.loc[:,'<250mm':'>2000mm']
indebtedness_median = pd.DataFrame(indebtedness.median(), 
                                 columns = ['Indebt Median'])

# Make index first column - indebtedness_median
indebtedness_median = indebtedness_median.reset_index(level = 0)
indebtedness_median.columns = ['Indebtedness', 'Indebt Median']

# Isolate price and calculate median price
price = universe.loc[:,'Port. Px':'<$70']
price_median = pd.DataFrame(price.median(), columns = ['Price Median'])

# Make index first column - price_median
price_median = price_median.reset_index(level = 0)
price_median.columns = ['Price', 'Price Median']

# Isolate debt structure and calculate median debt structure 
debt_structure = universe.loc[:,'LoanOnly':'1LienOnly']
debt_structure['Loan/Bond'] = 1 - (debt_structure[
    'LoanOnly'] + debt_structure['1LienOnly'])
debt_structure_median = pd.DataFrame(debt_structure.median(), 
                                    columns = ['Debt Struct Median'])

# Make index first column - debt_structure_median
debt_structure_median = debt_structure_median.reset_index(level = 0)
debt_structure_median.columns = [' Debt Structure', 'Debt Struct Median']

# Isolate ratings and calculate median ratings 
ratings = universe.loc[:,'B- Fac':'B3 CFR']
ratings_median = pd.DataFrame(ratings.median(), columns = ['Ratings Median'])

# Make index first column - ratings_median
ratings_median = ratings_median.reset_index(level = 0)
ratings_median.columns = ['Ratings', 'Ratings Median']

# Isolate collateral strats
collateral_strats = universe.loc[:,'Default%':'Def Px']

# Merge dataframes - database
data_merge = sector.merge(indebtedness, how = 'left', 
                                  left_index = True, right_index = True)
data_merge = data_merge.merge(price, how = 'left', 
                                  left_index = True, right_index = True)
data_merge = data_merge.merge(debt_structure, how = 'left', 
                                  left_index = True, right_index = True)
data_merge = data_merge.merge(ratings, how = 'left', 
                                  left_index = True, right_index = True)
data_merge = data_merge.merge(collateral_strats, how = 'left', 
                                  left_index = True, right_index = True)

# Merge dataframes - medians
all_medians = sector_median.merge(indebtedness_median, how = 'left', 
                                  left_index = True, right_index = True)
all_medians = all_medians.merge(price_median, how = 'left', 
                                  left_index = True, right_index = True)
all_medians = all_medians.merge(debt_structure_median, how = 'left', 
                                  left_index = True, right_index = True)
all_medians = all_medians.merge(ratings_median, how = 'left', 
                                  left_index = True, right_index = True)

# Drop top rows and set column headers
portfolio = portfolio.iloc[4:,:]
portfolio.columns = portfolio.iloc[0]
portfolio = portfolio.iloc[1:]

# Split scenarios and stress dfs
stress = scenarios.set_index(scenarios.columns[0])
stress = stress.loc['TRP STRESS':,:]
stress = stress.reset_index()

# Delete stress from scenarios df
scenarios = scenarios.set_index(scenarios.columns[0])
scenarios = scenarios.loc[:'MAX EXT',:]
scenarios = scenarios.reset_index()

# Isolate only necessary scenario columns
scenarios = scenarios[['Scenario Name', '#', 'Deal/Tranche ID', 
                       'Intex Deal Name', 'Bloomberg Ticker', 'Reinv End Date',
                       'Callable as of Date', 'Price', 'Disc Margin', 'WAL',
                       'Yield', 'Curr Collat Bal', 'Orig Collat Bal', 
                       'Principal Collection Account (Bal)', 'Tranche']]

# Write to excel
with pd.ExcelWriter('Chubb Report Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    data_merge.to_excel(writer, 'DB_Pull', index = False)
    
with pd.ExcelWriter('Chubb Report Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_medians.to_excel(writer, 'Medians', index = False)

with pd.ExcelWriter('Chubb Report Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    portfolio.to_excel(writer, 'Portfolio', index = False)
    
with pd.ExcelWriter('Chubb Report Template.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    scenarios.to_excel(writer, 'Scenarios', index = False)
