# -*- coding: utf-8 -*-
"""
Created on Fri Jan 27 10:04:51 2023

@author: trpjs86
"""

import pandas as pd
pd.options.mode.chained_assignment = None  
import numpy as np

# All necessary file imports

# ABS imports
abs_index = pd.read_excel('AA_ABS_Index_Contents 01312023.xlsx')
nif_abs_holdings = pd.read_excel('ABS Account Holdings 11.22.xlsx', 
                                 sheet_name = 'NIF')
snf_abs_holdings = pd.read_excel('ABS Account Holdings 11.22.xlsx', 
                                 sheet_name = 'SNF')
stb_abs_holdings = pd.read_excel('ABS Account Holdings 11.22.xlsx', 
                                 sheet_name = 'STB')
ttf_abs_holdings = pd.read_excel('ABS Account Holdings 11.22.xlsx', 
                                 sheet_name = 'TTF')

# CMBS imports
cmbs_index = pd.read_excel('BBG Port Feb 9.xlsx', sheet_name = 'CMBS')
nif_cmbs_holdings = pd.read_excel('BBG Port Feb 9.xlsx', sheet_name = 'NIF')
snf_cmbs_holdings = pd.read_excel('BBG Port Feb 9.xlsx', sheet_name = 'SNF')
stb_cmbs_holdings = pd.read_excel('BBG Port Feb 9.xlsx', sheet_name = 'STB')
ttf_cmbs_holdings = pd.read_excel('BBG Port Feb 9.xlsx', sheet_name = 'TTF')

# ABS ANALYSIS

# Create a list of all ABS dataframes
all_accounts = [abs_index, nif_abs_holdings, snf_abs_holdings,
           ttf_abs_holdings, stb_abs_holdings]

# Create empty list for finished dfs to be appended to
abs_list = []

# Loop through dfs for each account and perform the function
for i in all_accounts:

# Function performs operations on specified dataframe
    def abs_account(i):
        global df
        df = i
    
        # Set column headers and drop unnecessary rows
        df.columns = df.iloc[6]
        df = df.iloc[7:]
    
        # Calculate partition weights
        partition = df.groupby('Partition', as_index = False)['% Wgt'].sum()
        partition = partition.sort_values('% Wgt', ascending = False)
        partition = partition.reset_index(drop = True)
    
        # Calculate issuer weights
        issuers = df.groupby('ISSUER', as_index = False)['% Wgt'].sum()
        issuers = issuers.sort_values('% Wgt', ascending = False)
        issuers = issuers.reset_index(drop = True)
    
        # Replace nan S&P ratings w/ Moody's rating
        df['S&P Rating'].fillna(df["Moody's Rating"], 
                                                    inplace = True)
    
        # Replace Moody's rating in S&P column w/ equivalent S&P rating
        df['S&P Rating'] = np.where(df['S&P Rating'] == 
                                            df["Moody's Rating"], 
            np.where(df['S&P Rating'] == 'Aaa', 'AAA', 
            np.where(df['S&P Rating'] == 'Aa1', 'AA+', 
            np.where(df['S&P Rating'] == 'Aa2', 'AA',
            np.where(df['S&P Rating'] == 'A3', 'A-',
            np.where(df['S&P Rating'] == 'Aa3', 'AA-', 
            np.where(df['S&P Rating'] == 'Baa3', 'BBB-', 
            np.where(df['S&P Rating'] == 'A1', 'A+',
            np.where(df['S&P Rating'] == 'A2', 'A',
            np.where(df['S&P Rating'] == 'Baa1', 'BBB+',
            np.where(df['S&P Rating'] == 'Baa2', 'BBB',
            np.where(df['S&P Rating'] == 'Ba1', 'BBB+',
            np.where(df['S&P Rating'] == 'Ba2', 'BB',
            np.where(df['S&P Rating'] == 'Ba3', 'BBB-',
                      df['S&P Rating']))))))))))))), 
            df['S&P Rating'])
        
        # Fill rating nans with NR
        df['S&P Rating'].fillna('NR', inplace = True)
    
        # Calculate S&P rating weights
        snp_ratings = df.groupby('S&P Rating', as_index = False)['% Wgt'].sum()
        snp_ratings = snp_ratings.sort_values('% Wgt', ascending = False)
        snp_ratings = snp_ratings.reset_index(drop = True)
        
        return snp_ratings, partition, issuers
    
    return_frame = abs_account(i)
    
    abs_list.append(return_frame)
        
# Unpack the tuples in abs_list
index_ratings, index_partitions, index_issuers = abs_list[0]
index_ratings.columns = ['Index Rating', '% Wgt Index']
index_partitions.columns = ['Index Partition', '% Wgt Index']
index_issuers.columns = ['Index Issuer', '% Wgt Index']

nif_ratings, nif_partitions, nif_issuers = abs_list[1]
nif_ratings.columns = ['NIF Rating', '% Wgt NIF']
nif_partitions.columns = ['NIF Partition', '% Wgt NIF']
nif_issuers.columns = ['NIF Issuer', '% Wgt NIF']

snf_ratings, snf_partitions, snf_issuers = abs_list[2]
snf_ratings.columns = ['SNF Rating', '% Wgt SNF']
snf_partitions.columns = ['SNF Partition', '% Wgt SNF']
snf_issuers.columns = ['SNF Issuer', '% Wgt SNF']

ttf_ratings, ttf_partitions, ttf_issuers = abs_list[3]
ttf_ratings.columns = ['TTF Rating', '% Wgt TTF']
ttf_partitions.columns = ['TTF Partition', '% Wgt TTF']
ttf_issuers.columns = ['TTF Issuer', '% Wgt TTF']

stb_ratings, stb_partitions, stb_issuers = abs_list[4]
stb_ratings.columns = ['STB Rating', '% Wgt STB']
stb_partitions.columns = ['STB Partition', '% Wgt STB']
stb_issuers.columns = ['STB Issuer', '% Wgt STB']

# Merge all ratings, partitions, and issuers
merge_list = [nif_ratings, snf_ratings, ttf_ratings, stb_ratings]
all_ratings = index_ratings
for i in merge_list:
    all_ratings = all_ratings.merge(i, how = 'left', 
                                      left_index = True, right_index = True)

merge_list = [nif_partitions, snf_partitions, ttf_partitions, stb_partitions]
all_partitions = index_partitions
for i in merge_list:
    all_partitions = all_partitions.merge(i, how = 'left', 
                                      left_index = True, right_index = True)
    
merge_list = [nif_issuers, snf_issuers, ttf_issuers, stb_issuers]
all_issuers = index_issuers
for i in merge_list:
    all_issuers = all_issuers.merge(i, how = 'left', 
                                      left_index = True, right_index = True)
    
# Write to excel
with pd.ExcelWriter('ABS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_ratings.to_excel(writer, 'all_ratings', index = False)
    
with pd.ExcelWriter('ABS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_partitions.to_excel(writer, 'all_partitions', index = False)
    
with pd.ExcelWriter('ABS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_issuers.to_excel(writer, 'all_issuers', index = False)
    
# CMBS ANALYSIS

# Create a list of all CMBS dataframes
all_accounts = [cmbs_index, nif_cmbs_holdings, snf_cmbs_holdings, 
                ttf_cmbs_holdings, stb_cmbs_holdings]

# Create empty list for finished dfs to be appended to
cmbs_list = []

# Loop through dfs for each account and perform the function
for i in all_accounts:
    
    # Function performs operations on specified dataframe
    def cmbs_account(i):
        global cmbs
        cmbs = i

        # Set column headers and drop unnecessary rows
        cmbs.columns = cmbs.iloc[6]
        cmbs = cmbs.iloc[7:]
        
        # Calculate partition weights
        partition = cmbs.groupby('Partition', as_index = False)['% Wgt'].sum()
        partition = partition.sort_values('% Wgt', ascending = False)
        partition = partition.reset_index(drop = True)
    
        # Calculate issuer weights
        issuers = cmbs.groupby('ISSUER', as_index = False)['% Wgt'].sum()
        issuers = issuers.sort_values('% Wgt', ascending = False)
        issuers = issuers.reset_index(drop = True)
        
        if 'S&P Rating' in cmbs.columns:
        
            # Replace nan Moody's rating w/ S&P rating
            cmbs["Moody's Rating"].fillna(cmbs['S&P Rating'], 
                                                      inplace = True)
            
            # Replace S&P rating in Moody's column w/ equivalent Moody's rating
            cmbs["Moody's Rating"] = np.where(cmbs["Moody's Rating"] == 
                                                cmbs['S&P Rating'], 
                np.where(cmbs["Moody's Rating"] == 'AAA', 'Aaa', 
                np.where(cmbs["Moody's Rating"] == 'AA+', 'Aa1', 
                np.where(cmbs["Moody's Rating"] == 'AA', 'Aa2',
                np.where(cmbs["Moody's Rating"] == 'A-', 'A3',
                np.where(cmbs["Moody's Rating"] == 'AA-', 'Aa3', 
                np.where(cmbs["Moody's Rating"] == 'BBB-', 'Baa3', 
                np.where(cmbs["Moody's Rating"] == 'A+', 'A1',
                np.where(cmbs["Moody's Rating"] == 'A', 'A2',
                np.where(cmbs["Moody's Rating"] == 'BBB+', 'Baa1',
                np.where(cmbs["Moody's Rating"] == 'BBB', 'Baa2',
                np.where(cmbs["Moody's Rating"] == 'BBB+', 'Ba1',
                np.where(cmbs["Moody's Rating"] == 'BB', 'Ba2',
                np.where(cmbs["Moody's Rating"] == 'BBB-', 'Ba3',
                np.where(cmbs["Moody's Rating"] == 'BB-', 'Ba3',
                         cmbs["Moody's Rating"])))))))))))))), 
                cmbs["Moody's Rating"])
        
        else:
            pass
        
        if 'Fitch Rating' in cmbs.columns:
        
            # Replace nan Moody's rating w/ S&P rating
            cmbs["Moody's Rating"].fillna(cmbs['Fitch Rating'], 
                                                      inplace = True)
            
            # Replace NR Moody's rating w/ Fitch rating
            cmbs["Moody's Rating"] = np.where(cmbs["Moody's Rating"] == 'NR',
                                              cmbs['Fitch Rating'],
                                              cmbs["Moody's Rating"])
            
            # Replace Fitch rtg in Moody's column w/ equivalent Moody's rtg
            cmbs["Moody's Rating"] = np.where(cmbs["Moody's Rating"] == 
                                                cmbs['Fitch Rating'], 
                np.where(cmbs["Moody's Rating"] == 'AAA', 'Aaa', 
                np.where(cmbs["Moody's Rating"] == 'AA+', 'Aa1', 
                np.where(cmbs["Moody's Rating"] == 'AA', 'Aa2',
                np.where(cmbs["Moody's Rating"] == 'A-', 'A3',
                np.where(cmbs["Moody's Rating"] == 'AA-', 'Aa3', 
                np.where(cmbs["Moody's Rating"] == 'BBB-', 'Baa3', 
                np.where(cmbs["Moody's Rating"] == 'A+', 'A1',
                np.where(cmbs["Moody's Rating"] == 'A', 'A2',
                np.where(cmbs["Moody's Rating"] == 'BBB+', 'Baa1',
                np.where(cmbs["Moody's Rating"] == 'BBB', 'Baa2',
                np.where(cmbs["Moody's Rating"] == 'BBB+', 'Ba1',
                np.where(cmbs["Moody's Rating"] == 'BB', 'Ba2',
                np.where(cmbs["Moody's Rating"] == 'BBB-', 'Ba3',
                np.where(cmbs["Moody's Rating"] == 'BB-', 'Ba3',
                np.where(cmbs["Moody's Rating"] == 'AAA *-', 'Aaa',
                         cmbs["Moody's Rating"]))))))))))))))), 
                cmbs["Moody's Rating"])
        
        else:
            pass
        
        # Fill rating nans with NR
        cmbs["Moody's Rating"].fillna('NR', inplace = True)
    
        # Calculate S&P rating weights
        ratings = cmbs.groupby("Moody's Rating", as_index = False)['% Wgt'].sum()
        ratings = ratings.sort_values('% Wgt', ascending = False)
        ratings = ratings.reset_index(drop = True)
        
        return ratings, partition, issuers
    
    return_frame = cmbs_account(i)
    
    cmbs_list.append(return_frame)
    
# Unpack the tuples in cmbs_list
index_ratings, index_partitions, index_issuers = cmbs_list[0]
index_ratings.columns = ['Index Rating', '% Wgt Index']
index_partitions.columns = ['Index Partition', '% Wgt Index']
index_issuers.columns = ['Index Issuer', '% Wgt Index']

nif_ratings, nif_partitions, nif_issuers = cmbs_list[1]
nif_ratings.columns = ['NIF Rating', '% Wgt NIF']
nif_partitions.columns = ['NIF Partition', '% Wgt NIF']
nif_issuers.columns = ['NIF Issuer', '% Wgt NIF']

snf_ratings, snf_partitions, snf_issuers = cmbs_list[2]
snf_ratings.columns = ['SNF Rating', '% Wgt SNF']
snf_partitions.columns = ['SNF Partition', '% Wgt SNF']
snf_issuers.columns = ['SNF Issuer', '% Wgt SNF']

ttf_ratings, ttf_partitions, ttf_issuers = cmbs_list[3]
ttf_ratings.columns = ['TTF Rating', '% Wgt TTF']
ttf_partitions.columns = ['TTF Partition', '% Wgt TTF']
ttf_issuers.columns = ['TTF Issuer', '% Wgt TTF']

stb_ratings, stb_partitions, stb_issuers = cmbs_list[4]
stb_ratings.columns = ['STB Rating', '% Wgt STB']
stb_partitions.columns = ['STB Partition', '% Wgt STB']
stb_issuers.columns = ['STB Issuer', '% Wgt STB']

# Merge all ratings, partitions, and issuers
merge_list = [nif_ratings, snf_ratings, ttf_ratings, stb_ratings]
all_ratings = index_ratings
for i in merge_list:
    all_ratings = all_ratings.merge(i, how = 'left', 
                                      left_index = True, right_index = True)

merge_list = [nif_partitions, snf_partitions, ttf_partitions, stb_partitions]
all_partitions = index_partitions
for i in merge_list:
    all_partitions = all_partitions.merge(i, how = 'left', 
                                      left_index = True, right_index = True)
    
merge_list = [nif_issuers, snf_issuers, ttf_issuers, stb_issuers]
all_issuers = index_issuers
for i in merge_list:
    all_issuers = all_issuers.merge(i, how = 'left', 
                                      left_index = True, right_index = True)
    
# Lowercase the index ratings
all_ratings['Index Rating'] = np.where(all_ratings['Index Rating'] == 'AAA',
'Aaa', np.where(all_ratings['Index Rating'] == 'AA3', 'Aa3',
np.where(all_ratings['Index Rating'] == 'AA2', 'Aa2',
np.where(all_ratings['Index Rating'] == 'AA1', 'Aa1',
np.where(all_ratings['Index Rating'] == 'BAA2', 'Baa2',
np.where(all_ratings['Index Rating'] == 'BAA3', 'Baa3',
np.where(all_ratings['Index Rating'] == 'BAA1', 'Baa1', 
          all_ratings['Index Rating'])))))))

# Write to excel
with pd.ExcelWriter('CMBS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_ratings.to_excel(writer, 'all_ratings', index = False)
    
with pd.ExcelWriter('CMBS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_partitions.to_excel(writer, 'all_partitions', index = False)
    
with pd.ExcelWriter('CMBS Benchmark Analysis.xlsx', engine = 'openpyxl',
                    mode = 'a', if_sheet_exists = 'replace') as writer:
    all_issuers.to_excel(writer, 'all_issuers', index = False)
        



# # Calculate account overweight vs. benchmark 
# # stb_abs_partition = stb_abs_partition.assign(test = stb_abs_partition['STB Partition'].isin(abs_partition['ABS Index Partition'])

# # stb_abs_partition['test'] = stb_abs_partition['test'].astype(bool)        


