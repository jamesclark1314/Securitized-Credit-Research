# -*- coding: utf-8 -*-
"""
Created on Fri Jun 21 17:08:10 2024

@author: trpjs86
"""

import pandas as pd
from xbbg import blp
from datetime import datetime
import numpy as np

# User defined variables ------------------------------------------------------
tranche = 'DCOT 2019-MTC D'
discount_rate = 0.12854

# If yes, coupons are reduced by loss severity during extension period, else no
reduce_coupons = 'yes'

# Scenarios - extension in months, loss in decimal, probability in decimal
upside_ext = 0
upside_loss = 0
upside_prob = 0.05

mod_up_ext = 0
mod_up_loss = 0
mod_up_prob = 0.15

base_ext = 0
base_loss = 0
base_prob = 0.5

mod_down_ext = 0
mod_down_loss = 0
mod_down_prob = 0.2

downside_ext = 0
downside_loss = 0
downside_prob = 0.1
# ----------------------------------------------------------------------------

# Deal statistics
bond_stats = blp.bdp(f'{tranche} MTGE', flds = ['ID_CUSIP'])
bond_stats['Coupon'] = blp.bdp(f'{tranche} MTGE', flds = ['CUR_CPN'])
bond_stats['Mty_Date'] = blp.bdp(f'{tranche} MTGE', flds = ['MTG_EXP_MTY_DT'])
bond_stats['Cpn Type'] = blp.bdp(f'{tranche} MTGE', flds = ['CPN_TYP'])
bond_stats['Spread %'] = blp.bdp(f'{tranche} MTGE', flds = ['FLT_SPREAD']) / 100

# Combine coupon types into just fixed and floating
bond_stats['Cpn Type'] = np.where(bond_stats['Cpn Type'] == 'VARIABLE', 
                                  'FIXED', bond_stats['Cpn Type'])

scenarios = ['Upside', 'Mod_Upside', 'Base', 'Mod_Downside', 'Downside']
ext_scenarios = [upside_ext, mod_up_ext, base_ext, mod_down_ext, downside_ext]
loss_scenarios = [upside_loss, mod_up_loss, base_loss, mod_down_loss, downside_loss]

today = datetime.today().strftime('%Y-%m-%d')
# Function calculates the # of months between now and maturity
def time_to_maturity(today, maturity):
    d1 = datetime.strptime(today, "%Y-%m-%d")
    d2 = datetime.strptime(maturity, "%Y-%m-%d")
    global months_to_maturity
    months_to_maturity = abs((d2 - d1).days) / 30
time_to_maturity(today, bond_stats.iat[0,2].strftime('%Y-%m-%d'))

# Create a dataframe for expected cashflows across scenarios
cashflows = pd.DataFrame()
for i in range(1, round(max(upside_ext, mod_up_ext, base_ext, mod_down_ext, 
                           downside_ext) + months_to_maturity + 1)):
    cashflows = pd.concat([cashflows, pd.DataFrame({'Month':[i]})], 
                          ignore_index = True)
    
if bond_stats.iat[0,3] == 'FLOATING':
    
    # Import the 1M Term SOFR forward curve
    counter = 0
    sofr_curve = pd.DataFrame()
    for i in range(counter, 11):
        if counter == 0: # Pull in current spot rate in first iteration
            sofr_curve = pd.concat([sofr_curve, 
                                    blp.bdp('TSFR1M INDEX', 
                                            flds = ['PX_LAST']).reset_index().rename(columns = {'index': 'tenor'})])
            sofr_curve.loc[counter, 'tenor'] = 'current'
        else: # Subsequent iterations pull in the forward curve
            sofr_curve = pd.concat([sofr_curve, 
                                    blp.bdp(f'S0558P {counter}Y BLC3 CURNCY', 
                                            flds = ['PX_LAST']).reset_index().rename(columns = {'index': 'tenor'})])
            sofr_curve = sofr_curve.reset_index(drop = True)
            sofr_curve.loc[counter, 'tenor'] = f'{counter}Y'
        counter = counter + 1
    
    # Interpolate 6-month forward rates
    df = sofr_curve.copy()
    df.index = range(1, 2 * len(sofr_curve), 2) # New df with index that corresponds to numerical position of interpolated rates
    df.drop(df.tail(1).index, inplace = True)
    
    df2 = sofr_curve.copy()
    df2.index = range(0, 2 * len(sofr_curve) - 1, 2) # Original sofr_curve df with correct index positions
    
    counter = 0.5
    for index, row in df.iterrows(): # Iterate through rows of df and rename tenors and calc intermediate yields
        df.loc[index, 'tenor'] = f'{counter}Y'
        df.loc[index, 'px_last'] = (df2.loc[index - 1, 'px_last'] + df2.loc[index + 1, 'px_last']) / 2
        counter = counter + 1
        
    sofr_curve = pd.concat([df2, df]).sort_index()
        
    # Model SOFR rates by month
    sofr_curve = sofr_curve.loc[sofr_curve.index.repeat(6)].reset_index(drop = True)
    sofr_curve['month'] = sofr_curve.index + 1
    sofr_curve.insert(0, 'month', sofr_curve.pop('month'))
    
    # Add a column for implied coupon payment
    sofr_curve['pot_coup'] = sofr_curve['px_last'] + bond_stats.iat[0,4]
    
    # Model cashflows under each scenario
    for j,k,l in zip(scenarios, ext_scenarios, loss_scenarios):
        
        if reduce_coupons == 'yes':
            
            # Coupon payments through maturity
            for i in range(1, round(months_to_maturity + 1)):
                cashflows.loc[cashflows.index[i-1], j] = (sofr_curve.loc[cashflows.index[i-1], 'pot_coup'] / 12)
                
            # Coupon payments beyond maturity and through the extension period
            for i in range(round(months_to_maturity + 1), round(k + months_to_maturity)):
                cashflows.loc[cashflows.index[i-1], j] = (sofr_curve.loc[cashflows.index[i-1], 'pot_coup'] / 12) * (1 - l)
                
            # Final principal + coupon payment
            cashflows.loc[cashflows.index[round(k + months_to_maturity - 1)], 
                          j] = (100 * (1 - l)) + ((sofr_curve.loc[cashflows.index[i-1], 'pot_coup'] / 12) * (1 - l))
        else: 
            # Coupon payments through maturity + extension period
            for i in range(1, round(k + months_to_maturity)):
                cashflows.loc[cashflows.index[i-1], j] = (sofr_curve.loc[cashflows.index[i-1], 'pot_coup'] / 12)
        
            # Final principal + coupon payment
            cashflows.loc[cashflows.index[round(k + months_to_maturity - 1)], 
                          j] = (100 * (1 - l)) + (sofr_curve.loc[cashflows.index[i-1], 'pot_coup'] / 12)
            
else:
    # Model cashflows under each scenario
    for j,k,l in zip(scenarios, ext_scenarios, loss_scenarios):
        
        if reduce_coupons == 'yes':
            
            # Coupon payments through maturity
            for i in range(1, round(months_to_maturity + 1)):
                cashflows.loc[cashflows.index[i-1], j] = float(bond_stats.iat[0,1] / 12)
            
            # Coupon payments beyond maturity and through the extension period
            for i in range(round(months_to_maturity + 1), round(k + months_to_maturity)):
                cashflows.loc[cashflows.index[i-1], j] = float(bond_stats.iat[0,1] / 12) * (1 - l)
                
            # Final principal + coupon payment
            cashflows.loc[cashflows.index[round(k + months_to_maturity - 1)], 
                          j] = (100 * (1 - l)) + (float(bond_stats.iat[0,1] / 12) * (1 - l))
        else: 
            # Coupon payments through maturity + extension period
            for i in range(1, round(k + months_to_maturity)):
                cashflows.loc[cashflows.index[i-1], j] = float(bond_stats.iat[0,1] / 12)
        
            # Final principal + coupon payment
            cashflows.loc[cashflows.index[round(k + months_to_maturity - 1)], 
                          j] = (100 * (1 - l)) + float(bond_stats.iat[0,1] / 12)
        
# Calculate the PV of the cashflows across scenarios
discounted_cfs = pd.DataFrame(cashflows['Month'])
for i in scenarios:
    for j in range(0, len(cashflows[i])):
        discounted_cfs.loc[discounted_cfs.index[j], 
                           i] = cashflows.loc[cashflows.index[j],
                                              i] / ((1 + (discount_rate / 12)) 
                                                    ** cashflows.loc[cashflows.index[j], 
                                                                     'Month'])
                                                    
# Create a dataframe for the implied bond values under each scenario
bond_values = pd.DataFrame()
bond_values.index = ['Value']
for i in scenarios:
    bond_values[i] = discounted_cfs[i].sum()
    
# Calculate the probability-weighted implied price
bond_values['Prob-Weighted'] = (bond_values
                                ['Upside'] * upside_prob) + (bond_values
                                ['Mod_Upside'] * mod_up_prob) + (bond_values
                                ['Base'] * base_prob) + (bond_values
                                ['Mod_Downside'] * mod_down_prob) + (bond_values
                                ['Downside'] * downside_prob)
                                                                     
print(f"Implied Price: ${bond_values.iloc[0]['Prob-Weighted']}")