# -*- coding: utf-8 -*-
"""
Created on Thu Jun 13 14:04:32 2024

@author: trpjs86
"""

import pandas as pd
from xbbg import blp
from datetime import datetime

today = datetime.today().strftime('%Y-%m-%d')

# User defined variables ------------------------------------------------------
tranche = "DCOT 2019-MTC D"
discount_rate = 0.085

# If yes, coupons are reduced by loss severity during extension period, else no
reduce_coupons = 'yes'

# Scenarios - extension in months, loss in decimal, probability in decimal
upside_ext = 12
upside_loss = 0
upside_prob = 0.05

mod_up_ext = 24
mod_up_loss = 0
mod_up_prob = 0.15

base_ext = 36
base_loss = 0.2836
base_prob = 0.5

mod_down_ext = 48
mod_down_loss = 0.9388
mod_down_prob = 0.2

downside_ext = 1
downside_loss = 1
downside_prob = 0.1
# ----------------------------------------------------------------------------

# Deal statistics
bond_stats = blp.bdp(f'{tranche} MTGE', flds = ['ID_CUSIP'])
bond_stats['Coupon'] = blp.bdp(f'{tranche} MTGE', flds = ['MTG_FIRST_CPN'])
bond_stats['Mty_Date'] = blp.bdp(f'{tranche} MTGE', flds = ['MTG_EXP_MTY_DT'])

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
    
scenarios = ['Upside', 'Mod_Upside', 'Base', 'Mod_Downside', 'Downside']
ext_scenarios = [upside_ext, mod_up_ext, base_ext, mod_down_ext, downside_ext]
loss_scenarios = [upside_loss, mod_up_loss, base_loss, mod_down_loss, downside_loss]

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
